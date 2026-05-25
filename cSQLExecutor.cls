VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSQLExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#======================================================================================================================
'# SQL Executor - Evaluates parsed SQL statements against a cDataTable
'#======================================================================================================================

Private m_Table As cDataTable
Private m_Parser As cSQLParser
Private m_AffectedRows As Long


'#======================================================================================================================
'# Public Properties
'#======================================================================================================================
Public Property Get AffectedRows() As Long
    AffectedRows = m_AffectedRows
End Property


'#======================================================================================================================
'# Public Methods
'#======================================================================================================================
Public Function Execute(ByRef parser As cSQLParser, ByRef sourceTable As cDataTable) As Variant
    Set m_Parser = parser
    Set m_Table = sourceTable
    m_AffectedRows = 0

    Select Case parser.StatementType
        Case sqlSelect
            Dim resultTable As cDataTable
            Set resultTable = ExecuteSelect()
            Set Execute = resultTable
        Case sqlInsert
            ExecuteInsert
            Execute = m_AffectedRows
        Case sqlUpdate
            ExecuteUpdate
            Execute = m_AffectedRows
        Case sqlDelete
            ExecuteDelete
            Execute = m_AffectedRows
    End Select
End Function


'#======================================================================================================================
'# SELECT Execution
'#======================================================================================================================
Private Function ExecuteSelect() As cDataTable
    Dim hasAggregates As Boolean
    Dim hasGroupBy As Boolean
    Dim grpCols As Collection

    Set grpCols = m_Parser.GroupByColumns
    hasGroupBy = (grpCols.Count > 0)
    hasAggregates = CheckForAggregates()

    If hasGroupBy Or hasAggregates Then
        Set ExecuteSelect = ExecuteGroupedSelect()
    Else
        Set ExecuteSelect = ExecuteSimpleSelect()
    End If
End Function

Private Function ExecuteSimpleSelect() As cDataTable
    Dim tResult As cDataTable
    Dim selectCols As Collection
    Dim colHeaders As String
    Dim colNames() As String
    Dim isSelectAll As Boolean
    Dim i As Long
    Dim j As Long
    Dim rowCount As Long
    Dim expr As cSQLExprNode
    Dim rec() As Variant
    Dim numCols As Long

    Set selectCols = m_Parser.SelectColumns

    ' Determine output columns
    If selectCols.Count = 1 Then
        Dim firstCol As cSQLExprNode
        Set firstCol = selectCols.Item(1)
        If firstCol.NodeType = exStar Then
            isSelectAll = True
        End If
    End If

    If isSelectAll Then
        numCols = m_Table.ColumnCount
        ReDim colNames(1 To numCols)
        colHeaders = m_Table.HeaderList
        For i = 1 To numCols
            colNames(i) = m_Table.ItemName(i)
        Next i
    Else
        numCols = selectCols.Count
        ReDim colNames(1 To numCols)
        colHeaders = ""
        For i = 1 To numCols
            Set expr = selectCols.Item(i)
            If expr.NodeType = exColumn Then
                colNames(i) = expr.ColumnName
            Else
                colNames(i) = "Expr" & i
            End If
            If Len(expr.AliasName) > 0 Then
                colNames(i) = expr.AliasName
            End If
            If i > 1 Then colHeaders = colHeaders & ", "
            colHeaders = colHeaders & colNames(i)
        Next i
    End If

    Set tResult = New cDataTable
    tResult.DefineTable numCols, colHeaders
    tResult.Name = "SQLResult"

    ' Iterate source rows, apply WHERE filter, project columns
    rowCount = 0
    m_Table.RsMoveFirst
    Do While Not m_Table.RsEOF
        ' Check TOP limit
        If m_Parser.TopCount > -1 And rowCount >= m_Parser.TopCount Then
            Exit Do
        End If

        ' Evaluate WHERE clause
        If MatchesWhere(m_Table.RsBookmark) Then
            ReDim rec(1 To numCols)
            If isSelectAll Then
                For j = 1 To numCols
                    rec(j) = m_Table.ItemRaw(m_Table.RsBookmark, j)
                Next j
            Else
                For j = 1 To selectCols.Count
                    Set expr = selectCols.Item(j)
                    rec(j) = EvalExpr(expr, m_Table.RsBookmark)
                Next j
            End If
            tResult.RecordAdd rec
            rowCount = rowCount + 1
        End If

        m_Table.RsMoveNext
    Loop

    ' Apply DISTINCT
    If m_Parser.IsDistinct Then
        ApplyDistinct tResult, colNames
    End If

    ' Apply ORDER BY
    Dim ordCols As Collection
    Set ordCols = m_Parser.OrderByColumns
    If ordCols.Count > 0 Then
        ApplyOrderBy tResult
    End If

    m_AffectedRows = tResult.RecordCount
    Set ExecuteSimpleSelect = tResult
End Function

Private Function ExecuteGroupedSelect() As cDataTable
    Dim tResult As cDataTable
    Dim selectCols As Collection
    Dim groupCols As Collection
    Dim colHeaders As String
    Dim colNames() As String
    Dim i As Long
    Dim j As Long
    Dim expr As cSQLExprNode
    Dim numSelCols As Long

    Set selectCols = m_Parser.SelectColumns
    Set groupCols = m_Parser.GroupByColumns
    numSelCols = selectCols.Count

    ' Build output column names
    ReDim colNames(1 To numSelCols)
    colHeaders = ""
    For i = 1 To numSelCols
        Set expr = selectCols.Item(i)
        If expr.NodeType = exColumn Then
            colNames(i) = expr.ColumnName
        ElseIf expr.NodeType = exFunction Then
            colNames(i) = expr.FuncName
        Else
            colNames(i) = "Expr" & i
        End If
        If Len(expr.AliasName) > 0 Then
            colNames(i) = expr.AliasName
        End If
        If i > 1 Then colHeaders = colHeaders & ", "
        colHeaders = colHeaders & colNames(i)
    Next i

    Set tResult = New cDataTable
    tResult.DefineTable numSelCols, colHeaders
    tResult.Name = "SQLResult"

    If groupCols.Count > 0 Then
        ' GROUP BY execution
        Dim groupRows As cDataTable
        Dim groupKey As String
        Dim rowList() As Long
        Dim rowListCount As Long
        Dim rec() As Variant
        Dim keys As Collection
        Dim keyVar As Variant

        Set groupRows = New cDataTable
        groupRows.DefineTable 2, "GroupKey, RowNum"

        ' Collect rows into groups
        m_Table.RsMoveFirst
        Do While Not m_Table.RsEOF
            If MatchesWhere(m_Table.RsBookmark) Then
                groupKey = BuildGroupKey(m_Table.RsBookmark, groupCols)
                groupRows.RecordAdd groupKey, CLng(m_Table.RsBookmark)
            End If
            m_Table.RsMoveNext
        Loop

        ' Get unique keys in insertion order
        Set keys = GetUniqueKeys(groupRows)

        ' Process each group
        For Each keyVar In keys
            groupKey = CStr(keyVar)
            CollectGroupRows groupRows, groupKey, rowList, rowListCount

            ' Evaluate HAVING
            Dim havingExpr As cSQLExprNode
            Set havingExpr = m_Parser.HavingExpr
            If Not havingExpr Is Nothing Then
                If Not EvalHaving(havingExpr, rowList, rowListCount) Then
                    GoTo NextGroup
                End If
            End If

            ' Build result record
            ReDim rec(1 To numSelCols)
            For j = 1 To numSelCols
                Set expr = selectCols.Item(j)
                If expr.NodeType = exFunction Then
                    rec(j) = EvalAggregate(expr, rowList, rowListCount)
                ElseIf expr.NodeType = exColumn Then
                    rec(j) = m_Table.Item(expr.ColumnName, rowList(1))
                Else
                    rec(j) = EvalExpr(expr, rowList(1))
                End If
            Next j

            tResult.RecordAdd rec

NextGroup:
        Next keyVar
    Else
        ' Aggregates without GROUP BY - entire table is one group
        Dim allRows() As Long
        Dim allRowCount As Long
        allRowCount = 0
        ReDim allRows(1 To m_Table.RecordCount + 1)

        m_Table.RsMoveFirst
        Do While Not m_Table.RsEOF
            If MatchesWhere(m_Table.RsBookmark) Then
                allRowCount = allRowCount + 1
                allRows(allRowCount) = m_Table.RsBookmark
            End If
            m_Table.RsMoveNext
        Loop

        ReDim rec(1 To numSelCols)
        For j = 1 To numSelCols
            Set expr = selectCols.Item(j)
            If expr.NodeType = exFunction Then
                rec(j) = EvalAggregate(expr, allRows, allRowCount)
            ElseIf allRowCount > 0 And expr.NodeType = exColumn Then
                rec(j) = m_Table.Item(expr.ColumnName, allRows(1))
            Else
                rec(j) = Null
            End If
        Next j
        tResult.RecordAdd rec
    End If

    ' Apply ORDER BY
    Dim ordCols As Collection
    Set ordCols = m_Parser.OrderByColumns
    If ordCols.Count > 0 Then
        ApplyOrderBy tResult
    End If

    m_AffectedRows = tResult.RecordCount
    Set ExecuteGroupedSelect = tResult
End Function


'#======================================================================================================================
'# INSERT Execution
'#======================================================================================================================
Private Sub ExecuteInsert()
    Dim insertCols As Collection
    Dim insertVals As Collection
    Dim rowValues As Collection
    Dim rec() As Variant
    Dim i As Long
    Dim j As Long
    Dim colIdx As Long
    Dim expr As cSQLExprNode
    Dim numInsertCols As Long

    Set insertCols = m_Parser.InsertColumns
    Set insertVals = m_Parser.InsertValues
    numInsertCols = insertCols.Count

    For i = 1 To insertVals.Count
        Set rowValues = insertVals.Item(i)

        If numInsertCols > 0 Then
            ReDim rec(1 To m_Table.ColumnCount)
            For j = 1 To m_Table.ColumnCount
                rec(j) = Empty
            Next j
            For j = 1 To numInsertCols
                colIdx = m_Table.ItemIndex(CStr(insertCols.Item(j)))
                Set expr = rowValues.Item(j)
                rec(colIdx) = EvalLiteralExpr(expr)
            Next j
        Else
            ReDim rec(1 To rowValues.Count)
            For j = 1 To rowValues.Count
                Set expr = rowValues.Item(j)
                rec(j) = EvalLiteralExpr(expr)
            Next j
        End If

        m_Table.RecordAdd rec
        m_AffectedRows = m_AffectedRows + 1
    Next i
End Sub


'#======================================================================================================================
'# UPDATE Execution
'#======================================================================================================================
Private Sub ExecuteUpdate()
    Dim updateCols As Collection
    Dim updateExprs As Collection
    Dim colName As String
    Dim expr As cSQLExprNode
    Dim j As Long
    Dim numUpdCols As Long
    Dim currentRow As Long
    Dim newVal As Variant

    Set updateCols = m_Parser.UpdateColumns
    Set updateExprs = m_Parser.UpdateExprs
    numUpdCols = updateCols.Count

    m_Table.RsMoveFirst
    Do While Not m_Table.RsEOF
        currentRow = m_Table.RsBookmark
        If MatchesWhere(currentRow) Then
            For j = 1 To numUpdCols
                colName = CStr(updateCols.Item(j))
                Set expr = updateExprs.Item(j)
                newVal = EvalExpr(expr, currentRow)
                m_Table.SetItem colName, currentRow, newVal
            Next j
            m_AffectedRows = m_AffectedRows + 1
        End If
        m_Table.RsBookmark = currentRow
        m_Table.RsMoveNext
    Loop
End Sub


'#======================================================================================================================
'# DELETE Execution
'#======================================================================================================================
Private Sub ExecuteDelete()
    Dim rowsToDelete As Collection
    Dim i As Long
    Dim prevGC As Boolean

    Set rowsToDelete = New Collection

    m_Table.RsMoveFirst
    Do While Not m_Table.RsEOF
        If MatchesWhere(m_Table.RsBookmark) Then
            rowsToDelete.Add m_Table.RsBookmark
        End If
        m_Table.RsMoveNext
    Loop

    prevGC = m_Table.GarbageCollection
    m_Table.GarbageCollection = False
    For i = rowsToDelete.Count To 1 Step -1
        m_Table.RecordRemove CLng(rowsToDelete.Item(i))
        m_AffectedRows = m_AffectedRows + 1
    Next i
    m_Table.GarbageCollection = prevGC
End Sub


'#======================================================================================================================
'# WHERE Evaluation
'#======================================================================================================================
Private Function MatchesWhere(ByVal Row As Long) As Boolean
    Dim whereNode As cSQLExprNode
    Set whereNode = m_Parser.WhereExpr
    If whereNode Is Nothing Then
        MatchesWhere = True
    Else
        Dim result As Variant
        result = EvalExpr(whereNode, Row)
        If IsNull(result) Then
            MatchesWhere = False
        Else
            MatchesWhere = CBool(result)
        End If
    End If
End Function


'#======================================================================================================================
'# Expression Evaluation
'#======================================================================================================================
Private Function EvalExpr(ByRef node As cSQLExprNode, ByVal Row As Long) As Variant
    Dim leftVal As Variant
    Dim rightVal As Variant
    Dim result As Variant
    Dim i As Long

    Select Case node.NodeType
        Case exLiteral
            Select Case node.LiteralType
                Case litString
                    EvalExpr = node.StringValue
                Case litNumber
                    EvalExpr = node.NumValue
                Case litBoolean
                    EvalExpr = node.BoolValue
                Case litNull
                    EvalExpr = Null
            End Select

        Case exColumn
            EvalExpr = m_Table.Item(node.ColumnName, Row)

        Case exStar
            EvalExpr = Null

        Case exBinaryOp
            EvalExpr = EvalBinaryOp(node, Row)

        Case exUnaryOp
            leftVal = EvalExpr(node.LeftChild, Row)
            If node.Operator = "NOT" Then
                If IsNull(leftVal) Then
                    EvalExpr = Null
                Else
                    EvalExpr = Not CBool(leftVal)
                End If
            ElseIf node.Operator = "-" Then
                If IsNull(leftVal) Then
                    EvalExpr = Null
                Else
                    EvalExpr = -leftVal
                End If
            End If

        Case exBetween
            leftVal = EvalExpr(node.LeftChild, Row)
            Dim lowVal As Variant
            Dim highVal As Variant
            lowVal = EvalExpr(node.RightChild, Row)
            highVal = EvalExpr(node.ThirdChild, Row)
            If IsNull(leftVal) Or IsNull(lowVal) Or IsNull(highVal) Then
                result = Null
            Else
                result = (leftVal >= lowVal And leftVal <= highVal)
            End If
            If node.IsNegated And Not IsNull(result) Then result = Not CBool(result)
            EvalExpr = result

        Case exInList
            leftVal = EvalExpr(node.LeftChild, Row)
            If IsNull(leftVal) Then
                EvalExpr = Null
            Else
                result = False
                Dim nodeArgs As Collection
                Set nodeArgs = node.Args
                Dim argNode As cSQLExprNode
                For i = 1 To nodeArgs.Count
                    Set argNode = nodeArgs.Item(i)
                    Dim argVal As Variant
                    argVal = EvalExpr(argNode, Row)
                    If Not IsNull(argVal) Then
                        If CompareVariants(leftVal, argVal) = 0 Then
                            result = True
                            Exit For
                        End If
                    End If
                Next i
                If node.IsNegated Then result = Not CBool(result)
                EvalExpr = result
            End If

        Case exIsNull
            leftVal = EvalExpr(node.LeftChild, Row)
            result = IsNull(leftVal) Or IsEmpty(leftVal)
            If node.IsNegated Then result = Not CBool(result)
            EvalExpr = result

        Case exFunction
            EvalExpr = Null

    End Select
End Function

Private Function EvalBinaryOp(ByRef node As cSQLExprNode, ByVal Row As Long) As Variant
    Dim leftVal As Variant
    Dim rightVal As Variant
    Dim cmp As Long
    Dim result As Variant

    ' Short-circuit for AND/OR
    If node.Operator = "AND" Then
        leftVal = EvalExpr(node.LeftChild, Row)
        If Not IsNull(leftVal) Then
            If Not CBool(leftVal) Then
                EvalBinaryOp = False
                Exit Function
            End If
        End If
        rightVal = EvalExpr(node.RightChild, Row)
        If IsNull(leftVal) Or IsNull(rightVal) Then
            EvalBinaryOp = Null
        Else
            EvalBinaryOp = CBool(leftVal) And CBool(rightVal)
        End If
        Exit Function
    End If

    If node.Operator = "OR" Then
        leftVal = EvalExpr(node.LeftChild, Row)
        If Not IsNull(leftVal) Then
            If CBool(leftVal) Then
                EvalBinaryOp = True
                Exit Function
            End If
        End If
        rightVal = EvalExpr(node.RightChild, Row)
        If IsNull(leftVal) Or IsNull(rightVal) Then
            EvalBinaryOp = Null
        Else
            EvalBinaryOp = CBool(leftVal) Or CBool(rightVal)
        End If
        Exit Function
    End If

    leftVal = EvalExpr(node.LeftChild, Row)
    rightVal = EvalExpr(node.RightChild, Row)

    ' NULL propagation
    If IsNull(leftVal) Or IsNull(rightVal) Then
        EvalBinaryOp = Null
        Exit Function
    End If

    Select Case node.Operator
        Case "="
            EvalBinaryOp = (CompareVariants(leftVal, rightVal) = 0)
        Case "<>"
            EvalBinaryOp = (CompareVariants(leftVal, rightVal) <> 0)
        Case "<"
            EvalBinaryOp = (CompareVariants(leftVal, rightVal) < 0)
        Case ">"
            EvalBinaryOp = (CompareVariants(leftVal, rightVal) > 0)
        Case "<="
            EvalBinaryOp = (CompareVariants(leftVal, rightVal) <= 0)
        Case ">="
            EvalBinaryOp = (CompareVariants(leftVal, rightVal) >= 0)
        Case "LIKE"
            result = SqlLikeMatch(CStr(leftVal), CStr(rightVal))
            If node.IsNegated Then result = Not CBool(result)
            EvalBinaryOp = result
        Case "+"
            If IsNumeric(leftVal) And IsNumeric(rightVal) Then
                EvalBinaryOp = CDbl(leftVal) + CDbl(rightVal)
            Else
                EvalBinaryOp = CStr(leftVal) & CStr(rightVal)
            End If
        Case "-"
            EvalBinaryOp = CDbl(leftVal) - CDbl(rightVal)
        Case "*"
            EvalBinaryOp = CDbl(leftVal) * CDbl(rightVal)
        Case "/"
            If CDbl(rightVal) = 0 Then
                EvalBinaryOp = Null
            Else
                EvalBinaryOp = CDbl(leftVal) / CDbl(rightVal)
            End If
    End Select
End Function

Private Function EvalLiteralExpr(ByRef node As cSQLExprNode) As Variant
    Select Case node.NodeType
        Case exLiteral
            Select Case node.LiteralType
                Case litString
                    EvalLiteralExpr = node.StringValue
                Case litNumber
                    EvalLiteralExpr = node.NumValue
                Case litBoolean
                    EvalLiteralExpr = node.BoolValue
                Case litNull
                    EvalLiteralExpr = Null
            End Select
        Case exUnaryOp
            If node.Operator = "-" Then
                EvalLiteralExpr = -EvalLiteralExpr(node.LeftChild)
            End If
        Case Else
            EvalLiteralExpr = Null
    End Select
End Function


'#======================================================================================================================
'# Aggregate Evaluation
'#======================================================================================================================
Private Function EvalAggregate(ByRef node As cSQLExprNode, ByRef rows() As Long, ByVal rowCount As Long) As Variant
    Dim i As Long
    Dim val As Variant
    Dim total As Double
    Dim cnt As Long
    Dim minVal As Variant
    Dim maxVal As Variant
    Dim argExpr As cSQLExprNode
    Dim isStar As Boolean
    Dim distinctVals As Collection
    Dim nodeArgs As Collection
    Dim sVal As String

    Set nodeArgs = node.Args
    If nodeArgs.Count = 0 Then
        EvalAggregate = Null
        Exit Function
    End If

    Set argExpr = nodeArgs.Item(1)
    isStar = (argExpr.NodeType = exStar)

    If node.IsDistinctFunc Then
        Set distinctVals = New Collection
    End If

    total = 0
    cnt = 0

    Select Case node.FuncName
        Case "COUNT"
            For i = 1 To rowCount
                If isStar Then
                    If node.IsDistinctFunc Then
                        sVal = CStr(rows(i))
                        If Not CollectionContains(distinctVals, sVal) Then
                            distinctVals.Add sVal
                            cnt = cnt + 1
                        End If
                    Else
                        cnt = cnt + 1
                    End If
                Else
                    val = EvalExpr(argExpr, rows(i))
                    If Not IsNull(val) And Not IsEmpty(val) Then
                        If node.IsDistinctFunc Then
                            sVal = CStr(val)
                            If Not CollectionContains(distinctVals, sVal) Then
                                distinctVals.Add sVal
                                cnt = cnt + 1
                            End If
                        Else
                            cnt = cnt + 1
                        End If
                    End If
                End If
            Next i
            EvalAggregate = cnt

        Case "SUM"
            For i = 1 To rowCount
                val = EvalExpr(argExpr, rows(i))
                If Not IsNull(val) And IsNumeric(val) Then
                    If node.IsDistinctFunc Then
                        sVal = CStr(val)
                        If Not CollectionContains(distinctVals, sVal) Then
                            distinctVals.Add sVal
                            total = total + CDbl(val)
                            cnt = cnt + 1
                        End If
                    Else
                        total = total + CDbl(val)
                        cnt = cnt + 1
                    End If
                End If
            Next i
            If cnt = 0 Then
                EvalAggregate = Null
            Else
                EvalAggregate = total
            End If

        Case "AVG"
            For i = 1 To rowCount
                val = EvalExpr(argExpr, rows(i))
                If Not IsNull(val) And IsNumeric(val) Then
                    If node.IsDistinctFunc Then
                        sVal = CStr(val)
                        If Not CollectionContains(distinctVals, sVal) Then
                            distinctVals.Add sVal
                            total = total + CDbl(val)
                            cnt = cnt + 1
                        End If
                    Else
                        total = total + CDbl(val)
                        cnt = cnt + 1
                    End If
                End If
            Next i
            If cnt = 0 Then
                EvalAggregate = Null
            Else
                EvalAggregate = total / cnt
            End If

        Case "MIN"
            For i = 1 To rowCount
                val = EvalExpr(argExpr, rows(i))
                If Not IsNull(val) And Not IsEmpty(val) Then
                    If IsEmpty(minVal) Then
                        minVal = val
                    ElseIf CompareVariants(val, minVal) < 0 Then
                        minVal = val
                    End If
                End If
            Next i
            If IsEmpty(minVal) Then
                EvalAggregate = Null
            Else
                EvalAggregate = minVal
            End If

        Case "MAX"
            For i = 1 To rowCount
                val = EvalExpr(argExpr, rows(i))
                If Not IsNull(val) And Not IsEmpty(val) Then
                    If IsEmpty(maxVal) Then
                        maxVal = val
                    ElseIf CompareVariants(val, maxVal) > 0 Then
                        maxVal = val
                    End If
                End If
            Next i
            If IsEmpty(maxVal) Then
                EvalAggregate = Null
            Else
                EvalAggregate = maxVal
            End If

        Case Else
            EvalAggregate = Null
    End Select
End Function

Private Function EvalHaving(ByRef node As cSQLExprNode, ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim val As Variant
    val = EvalAggregateExpr(node, rows, rowCount)
    If IsNull(val) Then
        EvalHaving = False
    Else
        EvalHaving = CBool(val)
    End If
End Function

Private Function EvalAggregateExpr(ByRef node As cSQLExprNode, ByRef rows() As Long, ByVal rowCount As Long) As Variant
    Select Case node.NodeType
        Case exFunction
            EvalAggregateExpr = EvalAggregate(node, rows, rowCount)

        Case exLiteral
            Select Case node.LiteralType
                Case litString:  EvalAggregateExpr = node.StringValue
                Case litNumber:  EvalAggregateExpr = node.NumValue
                Case litBoolean: EvalAggregateExpr = node.BoolValue
                Case litNull:    EvalAggregateExpr = Null
            End Select

        Case exColumn
            If rowCount > 0 Then
                EvalAggregateExpr = m_Table.Item(node.ColumnName, rows(1))
            Else
                EvalAggregateExpr = Null
            End If

        Case exBinaryOp
            Dim leftVal As Variant
            Dim rightVal As Variant
            leftVal = EvalAggregateExpr(node.LeftChild, rows, rowCount)
            rightVal = EvalAggregateExpr(node.RightChild, rows, rowCount)

            If IsNull(leftVal) Or IsNull(rightVal) Then
                EvalAggregateExpr = Null
                Exit Function
            End If

            Select Case node.Operator
                Case "=":   EvalAggregateExpr = (CompareVariants(leftVal, rightVal) = 0)
                Case "<>":  EvalAggregateExpr = (CompareVariants(leftVal, rightVal) <> 0)
                Case "<":   EvalAggregateExpr = (CompareVariants(leftVal, rightVal) < 0)
                Case ">":   EvalAggregateExpr = (CompareVariants(leftVal, rightVal) > 0)
                Case "<=":  EvalAggregateExpr = (CompareVariants(leftVal, rightVal) <= 0)
                Case ">=":  EvalAggregateExpr = (CompareVariants(leftVal, rightVal) >= 0)
                Case "AND": EvalAggregateExpr = CBool(leftVal) And CBool(rightVal)
                Case "OR":  EvalAggregateExpr = CBool(leftVal) Or CBool(rightVal)
                Case "+":   EvalAggregateExpr = CDbl(leftVal) + CDbl(rightVal)
                Case "-":   EvalAggregateExpr = CDbl(leftVal) - CDbl(rightVal)
                Case "*":   EvalAggregateExpr = CDbl(leftVal) * CDbl(rightVal)
                Case "/"
                    If CDbl(rightVal) = 0 Then
                        EvalAggregateExpr = Null
                    Else
                        EvalAggregateExpr = CDbl(leftVal) / CDbl(rightVal)
                    End If
            End Select

        Case exUnaryOp
            Dim operandVal As Variant
            operandVal = EvalAggregateExpr(node.LeftChild, rows, rowCount)
            If node.Operator = "NOT" Then
                If IsNull(operandVal) Then
                    EvalAggregateExpr = Null
                Else
                    EvalAggregateExpr = Not CBool(operandVal)
                End If
            ElseIf node.Operator = "-" Then
                If IsNull(operandVal) Then
                    EvalAggregateExpr = Null
                Else
                    EvalAggregateExpr = -operandVal
                End If
            End If

        Case Else
            EvalAggregateExpr = Null
    End Select
End Function


'#======================================================================================================================
'# LIKE Pattern Matching (SQL % and _ wildcards)
'#======================================================================================================================
Private Function SqlLikeMatch(ByVal text As String, ByVal pattern As String) As Boolean
    Dim tLen As Long
    Dim pLen As Long
    Dim ti As Long
    Dim pi As Long
    Dim starTi As Long
    Dim starPi As Long
    Dim tc As String
    Dim pc As String

    tLen = Len(text)
    pLen = Len(pattern)
    ti = 1
    pi = 1
    starTi = 0
    starPi = 0

    Do While ti <= tLen
        If pi <= pLen Then
            pc = Mid$(pattern, pi, 1)

            If pc = "_" Then
                ti = ti + 1
                pi = pi + 1
                GoTo ContinueLikeLoop
            End If

            If pc = "%" Then
                starPi = pi
                starTi = ti
                pi = pi + 1
                GoTo ContinueLikeLoop
            End If

            tc = Mid$(text, ti, 1)
            If UCase$(tc) = UCase$(pc) Then
                ti = ti + 1
                pi = pi + 1
                GoTo ContinueLikeLoop
            End If
        End If

        If starPi > 0 Then
            pi = starPi + 1
            starTi = starTi + 1
            ti = starTi
            GoTo ContinueLikeLoop
        End If

        SqlLikeMatch = False
        Exit Function

ContinueLikeLoop:
    Loop

    Do While pi <= pLen
        If Mid$(pattern, pi, 1) <> "%" Then Exit Do
        pi = pi + 1
    Loop

    SqlLikeMatch = (pi > pLen)
End Function


'#======================================================================================================================
'# Helper Functions
'#======================================================================================================================
Private Function CompareVariants(ByVal v1 As Variant, ByVal v2 As Variant) As Long
    If IsNull(v1) And IsNull(v2) Then
        CompareVariants = 0
        Exit Function
    ElseIf IsNull(v1) Then
        CompareVariants = -1
        Exit Function
    ElseIf IsNull(v2) Then
        CompareVariants = 1
        Exit Function
    End If

    If IsNumeric(v1) And IsNumeric(v2) Then
        Dim d1 As Double
        Dim d2 As Double
        d1 = CDbl(v1)
        d2 = CDbl(v2)
        If d1 < d2 Then
            CompareVariants = -1
        ElseIf d1 > d2 Then
            CompareVariants = 1
        Else
            CompareVariants = 0
        End If
    Else
        CompareVariants = StrComp(CStr(v1), CStr(v2), vbTextCompare)
    End If
End Function

Private Function CheckForAggregates() As Boolean
    Dim selectCols As Collection
    Dim expr As cSQLExprNode
    Dim i As Long

    Set selectCols = m_Parser.SelectColumns
    For i = 1 To selectCols.Count
        Set expr = selectCols.Item(i)
        If ExprHasAggregate(expr) Then
            CheckForAggregates = True
            Exit Function
        End If
    Next i
    CheckForAggregates = False
End Function

Private Function ExprHasAggregate(ByRef node As cSQLExprNode) As Boolean
    If node Is Nothing Then
        ExprHasAggregate = False
        Exit Function
    End If

    If node.NodeType = exFunction Then
        ExprHasAggregate = True
        Exit Function
    End If

    If Not node.LeftChild Is Nothing Then
        If ExprHasAggregate(node.LeftChild) Then
            ExprHasAggregate = True
            Exit Function
        End If
    End If

    If Not node.RightChild Is Nothing Then
        If ExprHasAggregate(node.RightChild) Then
            ExprHasAggregate = True
            Exit Function
        End If
    End If

    ExprHasAggregate = False
End Function

Private Function BuildGroupKey(ByVal Row As Long, ByRef groupCols As Collection) As String
    Dim key As String
    Dim i As Long
    Dim val As Variant
    key = ""
    For i = 1 To groupCols.Count
        If i > 1 Then key = key & Chr$(0)
        val = m_Table.Item(CStr(groupCols.Item(i)), Row)
        If IsNull(val) Then
            key = key & Chr$(1)
        Else
            key = key & CStr(val)
        End If
    Next i
    BuildGroupKey = key
End Function

Private Function GetUniqueKeys(ByRef groupRows As cDataTable) As Collection
    Dim keys As Collection
    Dim gk As String
    Dim found As Boolean
    Dim existingKey As Variant

    Set keys = New Collection

    groupRows.RsMoveFirst
    Do While Not groupRows.RsEOF
        gk = CStr(groupRows.Item("GroupKey"))
        found = False
        For Each existingKey In keys
            If CStr(existingKey) = gk Then
                found = True
                Exit For
            End If
        Next existingKey
        If Not found Then
            keys.Add gk
        End If
        groupRows.RsMoveNext
    Loop

    Set GetUniqueKeys = keys
End Function

Private Sub CollectGroupRows(ByRef groupRows As cDataTable, ByVal groupKey As String, ByRef rows() As Long, ByRef rowCount As Long)
    rowCount = 0
    ReDim rows(1 To groupRows.RecordCount)

    groupRows.RsFindFirst "GroupKey", Equals, groupKey
    Do While Not groupRows.RsEOF
        rowCount = rowCount + 1
        rows(rowCount) = CLng(groupRows.Item("RowNum"))
        groupRows.RsFindNext "GroupKey", Equals, groupKey
    Loop
End Sub

Private Sub ApplyDistinct(ByRef tbl As cDataTable, ByRef colNames() As String)
    Dim seenKeys As Collection
    Dim stKey As String
    Dim j As Long
    Dim found As Boolean
    Dim existingKey As Variant

    Set seenKeys = New Collection

    tbl.GarbageCollection = False
    tbl.RsMoveFirst
    Do While Not tbl.RsEOF
        stKey = ""
        For j = 1 To UBound(colNames)
            stKey = stKey & Chr$(0) & CStr(tbl.Item(colNames(j)) & "")
        Next j
        found = False
        For Each existingKey In seenKeys
            If CStr(existingKey) = stKey Then
                found = True
                Exit For
            End If
        Next existingKey
        If found Then
            tbl.RecordRemove tbl.RsBookmark
        Else
            seenKeys.Add stKey
        End If
        tbl.RsMoveNext
    Loop
    tbl.GarbageCollection = True
End Sub

Private Sub ApplyOrderBy(ByRef tbl As cDataTable)
    Dim ordCols As Collection
    Dim ordDirs As Collection
    Dim numOrd As Long

    Set ordCols = m_Parser.OrderByColumns
    Set ordDirs = m_Parser.OrderByDirections
    numOrd = ordCols.Count

    Select Case numOrd
        Case 1
            tbl.Sort CStr(ordCols.Item(1)), LCase$(CStr(ordDirs.Item(1)))
        Case 2
            tbl.Sort CStr(ordCols.Item(1)), LCase$(CStr(ordDirs.Item(1))), _
                      CStr(ordCols.Item(2)), LCase$(CStr(ordDirs.Item(2)))
        Case 3
            tbl.Sort CStr(ordCols.Item(1)), LCase$(CStr(ordDirs.Item(1))), _
                      CStr(ordCols.Item(2)), LCase$(CStr(ordDirs.Item(2))), _
                      CStr(ordCols.Item(3)), LCase$(CStr(ordDirs.Item(3)))
        Case 4
            tbl.Sort CStr(ordCols.Item(1)), LCase$(CStr(ordDirs.Item(1))), _
                      CStr(ordCols.Item(2)), LCase$(CStr(ordDirs.Item(2))), _
                      CStr(ordCols.Item(3)), LCase$(CStr(ordDirs.Item(3))), _
                      CStr(ordCols.Item(4)), LCase$(CStr(ordDirs.Item(4)))
        Case Else
            If numOrd > 0 Then
                tbl.Sort CStr(ordCols.Item(1)), LCase$(CStr(ordDirs.Item(1)))
            End If
    End Select
End Sub

Private Function CollectionContains(ByRef col As Collection, ByVal key As String) As Boolean
    Dim item As Variant
    For Each item In col
        If CStr(item) = key Then
            CollectionContains = True
            Exit Function
        End If
    Next item
    CollectionContains = False
End Function
