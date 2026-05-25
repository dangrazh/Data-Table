VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSQLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#======================================================================================================================
'# SQL Parser - Tokenizer and Recursive Descent Parser for ANSI 92 SQL
'#======================================================================================================================

Public Enum SQLStmtType
    sqlSelect = 1
    sqlInsert = 2
    sqlUpdate = 3
    sqlDelete = 4
End Enum

Private Enum TokenType
    ttWord = 1
    ttNumber = 2
    ttString = 3
    ttOperator = 4
    ttComma = 5
    ttLParen = 6
    ttRParen = 7
    ttStar = 8
    ttDot = 9
    ttEOF = 10
End Enum


'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private m_TokTypes() As Long
Private m_TokValues() As String
Private m_TokNums() As Double
Private m_TokenCount As Long
Private m_Pos As Long
Private m_SQL As String

' Parsed statement properties
Private m_StmtType As SQLStmtType
Private m_IsDistinct As Boolean
Private m_TopCount As Long
Private m_SelectColumns As Collection
Private m_TableName As String
Private m_WhereExpr As cSQLExprNode
Private m_GroupByColumns As Collection
Private m_HavingExpr As cSQLExprNode
Private m_OrderByColumns As Collection
Private m_OrderByDirections As Collection
Private m_InsertColumns As Collection
Private m_InsertValues As Collection
Private m_UpdateColumns As Collection
Private m_UpdateExprs As Collection


'#======================================================================================================================
'# Initialization
'#======================================================================================================================
Private Sub Class_Initialize()
    m_TopCount = -1
    Set m_SelectColumns = New Collection
    Set m_GroupByColumns = New Collection
    Set m_OrderByColumns = New Collection
    Set m_OrderByDirections = New Collection
    Set m_InsertColumns = New Collection
    Set m_InsertValues = New Collection
    Set m_UpdateColumns = New Collection
    Set m_UpdateExprs = New Collection
End Sub


'#======================================================================================================================
'# Public Properties
'#======================================================================================================================
Public Property Get StatementType() As SQLStmtType
    StatementType = m_StmtType
End Property

Public Property Get IsDistinct() As Boolean
    IsDistinct = m_IsDistinct
End Property

Public Property Get TopCount() As Long
    TopCount = m_TopCount
End Property

Public Property Get SelectColumns() As Collection
    Set SelectColumns = m_SelectColumns
End Property

Public Property Get TableName() As String
    TableName = m_TableName
End Property

Public Property Get WhereExpr() As cSQLExprNode
    Set WhereExpr = m_WhereExpr
End Property

Public Property Get GroupByColumns() As Collection
    Set GroupByColumns = m_GroupByColumns
End Property

Public Property Get HavingExpr() As cSQLExprNode
    Set HavingExpr = m_HavingExpr
End Property

Public Property Get OrderByColumns() As Collection
    Set OrderByColumns = m_OrderByColumns
End Property

Public Property Get OrderByDirections() As Collection
    Set OrderByDirections = m_OrderByDirections
End Property

Public Property Get InsertColumns() As Collection
    Set InsertColumns = m_InsertColumns
End Property

Public Property Get InsertValues() As Collection
    Set InsertValues = m_InsertValues
End Property

Public Property Get UpdateColumns() As Collection
    Set UpdateColumns = m_UpdateColumns
End Property

Public Property Get UpdateExprs() As Collection
    Set UpdateExprs = m_UpdateExprs
End Property


'#======================================================================================================================
'# Public Methods
'#======================================================================================================================
Public Sub Parse(ByVal sql As String)
    m_SQL = Trim$(sql)

    ' Initialize all collections (interpreter may not run Class_Initialize)
    m_TopCount = -1
    m_IsDistinct = False
    Set m_SelectColumns = New Collection
    Set m_GroupByColumns = New Collection
    Set m_OrderByColumns = New Collection
    Set m_OrderByDirections = New Collection
    Set m_InsertColumns = New Collection
    Set m_InsertValues = New Collection
    Set m_UpdateColumns = New Collection
    Set m_UpdateExprs = New Collection
    Set m_WhereExpr = Nothing
    Set m_HavingExpr = Nothing
    m_TableName = ""

    Tokenize
    m_Pos = 1
    ParseStatement
End Sub


'#======================================================================================================================
'# Token Access Helpers (avoid UDT returns)
'#======================================================================================================================
Private Function PeekType() As Long
    PeekType = m_TokTypes(m_Pos)
End Function

Private Function PeekValue() As String
    PeekValue = m_TokValues(m_Pos)
End Function

Private Function PeekNum() As Double
    PeekNum = m_TokNums(m_Pos)
End Function

Private Sub Advance()
    If m_Pos < m_TokenCount Then
        m_Pos = m_Pos + 1
    End If
End Sub

Private Function AdvanceValue() As String
    AdvanceValue = m_TokValues(m_Pos)
    If m_Pos < m_TokenCount Then
        m_Pos = m_Pos + 1
    End If
End Function

Private Function AdvanceNum() As Double
    AdvanceNum = m_TokNums(m_Pos)
    If m_Pos < m_TokenCount Then
        m_Pos = m_Pos + 1
    End If
End Function

Private Function MatchWord(ByVal keyword As String) As Boolean
    If m_TokTypes(m_Pos) = ttWord And UCase$(m_TokValues(m_Pos)) = UCase$(keyword) Then
        MatchWord = True
        m_Pos = m_Pos + 1
    Else
        MatchWord = False
    End If
End Function

Private Sub ExpectWord(ByVal keyword As String)
    If Not MatchWord(keyword) Then
        Err.Raise vbObjectError + 602, "cSQLParser", "Expected '" & keyword & "' but found '" & m_TokValues(m_Pos) & "'"
    End If
End Sub

Private Function MatchToken(ByVal tt As Long) As Boolean
    If m_TokTypes(m_Pos) = tt Then
        MatchToken = True
        m_Pos = m_Pos + 1
    Else
        MatchToken = False
    End If
End Function

Private Sub ExpectToken(ByVal tt As Long)
    If Not MatchToken(tt) Then
        Err.Raise vbObjectError + 603, "cSQLParser", "Expected token type " & tt & " but found '" & m_TokValues(m_Pos) & "'"
    End If
End Sub

Private Function PeekWord(ByVal keyword As String) As Boolean
    PeekWord = (m_TokTypes(m_Pos) = ttWord And UCase$(m_TokValues(m_Pos)) = UCase$(keyword))
End Function

Private Function IsAggregateName(ByVal n As String) As Boolean
    Dim u As String
    u = UCase$(n)
    IsAggregateName = (u = "COUNT" Or u = "SUM" Or u = "AVG" Or u = "MIN" Or u = "MAX")
End Function


'#======================================================================================================================
'# Tokenizer
'#======================================================================================================================
Private Sub Tokenize()
    Dim i As Long
    Dim ch As String
    Dim startPos As Long
    Dim bufSize As Long
    Dim tokVal As String

    bufSize = 64
    ReDim m_TokTypes(1 To bufSize)
    ReDim m_TokValues(1 To bufSize)
    ReDim m_TokNums(1 To bufSize)
    m_TokenCount = 0

    i = 1
    Do While i <= Len(m_SQL)
        ch = Mid$(m_SQL, i, 1)

        ' Skip whitespace
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            i = i + 1

        ' Single-quoted string literal
        ElseIf ch = "'" Then
            i = i + 1
            tokVal = ""
            Do While i <= Len(m_SQL)
                ch = Mid$(m_SQL, i, 1)
                If ch = "'" Then
                    If i + 1 <= Len(m_SQL) And Mid$(m_SQL, i + 1, 1) = "'" Then
                        tokVal = tokVal & "'"
                        i = i + 2
                    Else
                        i = i + 1
                        Exit Do
                    End If
                Else
                    tokVal = tokVal & ch
                    i = i + 1
                End If
            Loop
            AddTok ttString, tokVal, 0, bufSize

        ' Number literal
        ElseIf ch >= "0" And ch <= "9" Then
            startPos = i
            Do While i <= Len(m_SQL)
                ch = Mid$(m_SQL, i, 1)
                If (ch >= "0" And ch <= "9") Or ch = "." Then
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            tokVal = Mid$(m_SQL, startPos, i - startPos)
            AddTok ttNumber, tokVal, CDbl(tokVal), bufSize

        ' Word (identifier or keyword)
        ElseIf IsAlpha(ch) Or ch = "_" Then
            startPos = i
            Do While i <= Len(m_SQL)
                ch = Mid$(m_SQL, i, 1)
                If IsAlpha(ch) Or (ch >= "0" And ch <= "9") Or ch = "_" Then
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            tokVal = Mid$(m_SQL, startPos, i - startPos)
            AddTok ttWord, tokVal, 0, bufSize

        ' Bracketed identifier [name with spaces]
        ElseIf ch = "[" Then
            i = i + 1
            startPos = i
            Do While i <= Len(m_SQL)
                ch = Mid$(m_SQL, i, 1)
                If ch = "]" Then
                    Exit Do
                End If
                i = i + 1
            Loop
            tokVal = Mid$(m_SQL, startPos, i - startPos)
            i = i + 1
            AddTok ttWord, tokVal, 0, bufSize

        ' Operators
        ElseIf ch = "<" Then
            If i + 1 <= Len(m_SQL) And Mid$(m_SQL, i + 1, 1) = ">" Then
                AddTok ttOperator, "<>", 0, bufSize
                i = i + 2
            ElseIf i + 1 <= Len(m_SQL) And Mid$(m_SQL, i + 1, 1) = "=" Then
                AddTok ttOperator, "<=", 0, bufSize
                i = i + 2
            Else
                AddTok ttOperator, "<", 0, bufSize
                i = i + 1
            End If

        ElseIf ch = ">" Then
            If i + 1 <= Len(m_SQL) And Mid$(m_SQL, i + 1, 1) = "=" Then
                AddTok ttOperator, ">=", 0, bufSize
                i = i + 2
            Else
                AddTok ttOperator, ">", 0, bufSize
                i = i + 1
            End If

        ElseIf ch = "!" Then
            If i + 1 <= Len(m_SQL) And Mid$(m_SQL, i + 1, 1) = "=" Then
                AddTok ttOperator, "<>", 0, bufSize
                i = i + 2
            Else
                Err.Raise vbObjectError + 600, "cSQLParser", "Unexpected character '!' at position " & i
            End If

        ElseIf ch = "=" Then
            AddTok ttOperator, "=", 0, bufSize
            i = i + 1

        ElseIf ch = "+" Or ch = "-" Or ch = "/" Then
            AddTok ttOperator, ch, 0, bufSize
            i = i + 1

        ElseIf ch = "*" Then
            AddTok ttStar, "*", 0, bufSize
            i = i + 1

        ElseIf ch = "," Then
            AddTok ttComma, ",", 0, bufSize
            i = i + 1

        ElseIf ch = "(" Then
            AddTok ttLParen, "(", 0, bufSize
            i = i + 1

        ElseIf ch = ")" Then
            AddTok ttRParen, ")", 0, bufSize
            i = i + 1

        ElseIf ch = "." Then
            AddTok ttDot, ".", 0, bufSize
            i = i + 1

        ElseIf ch = ";" Then
            i = i + 1

        Else
            Err.Raise vbObjectError + 601, "cSQLParser", "Unexpected character '" & ch & "' at position " & i
        End If
    Loop

    ' Add EOF token
    AddTok ttEOF, "", 0, bufSize

    ' Trim arrays to actual size
    ReDim Preserve m_TokTypes(1 To m_TokenCount)
    ReDim Preserve m_TokValues(1 To m_TokenCount)
    ReDim Preserve m_TokNums(1 To m_TokenCount)
End Sub

Private Sub AddTok(ByVal tType As Long, ByVal tValue As String, ByVal tNum As Double, ByRef bufSize As Long)
    m_TokenCount = m_TokenCount + 1
    If m_TokenCount > bufSize Then
        bufSize = bufSize * 2
        ReDim Preserve m_TokTypes(1 To bufSize)
        ReDim Preserve m_TokValues(1 To bufSize)
        ReDim Preserve m_TokNums(1 To bufSize)
    End If
    m_TokTypes(m_TokenCount) = tType
    m_TokValues(m_TokenCount) = tValue
    m_TokNums(m_TokenCount) = tNum
End Sub

Private Function IsAlpha(ByVal ch As String) As Boolean
    IsAlpha = (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")
End Function


'#======================================================================================================================
'# Statement Parsing
'#======================================================================================================================
Private Sub ParseStatement()
    If PeekType() = ttWord Then
        Select Case UCase$(PeekValue())
            Case "SELECT"
                m_StmtType = sqlSelect
                ParseSelectStmt
            Case "INSERT"
                m_StmtType = sqlInsert
                ParseInsertStmt
            Case "UPDATE"
                m_StmtType = sqlUpdate
                ParseUpdateStmt
            Case "DELETE"
                m_StmtType = sqlDelete
                ParseDeleteStmt
            Case Else
                Err.Raise vbObjectError + 604, "cSQLParser", "Expected SELECT, INSERT, UPDATE, or DELETE but found '" & PeekValue() & "'"
        End Select
    Else
        Err.Raise vbObjectError + 604, "cSQLParser", "Expected SQL statement keyword"
    End If
End Sub

Private Sub ParseSelectStmt()
    ExpectWord "SELECT"

    ' DISTINCT
    If MatchWord("DISTINCT") Then
        m_IsDistinct = True
    End If

    ' TOP n
    If MatchWord("TOP") Then
        If PeekType() = ttNumber Then
            m_TopCount = CLng(PeekNum())
            Advance
        Else
            Err.Raise vbObjectError + 605, "cSQLParser", "Expected number after TOP"
        End If
    End If

    ' Select list
    ParseSelectList

    ' FROM
    ExpectWord "FROM"
    ParseTableName

    ' WHERE
    If PeekWord("WHERE") Then
        m_Pos = m_Pos + 1
        Set m_WhereExpr = ParseExpr()
    End If

    ' GROUP BY
    If PeekWord("GROUP") Then
        m_Pos = m_Pos + 1
        ExpectWord "BY"
        ParseGroupByList
    End If

    ' HAVING
    If PeekWord("HAVING") Then
        m_Pos = m_Pos + 1
        Set m_HavingExpr = ParseExpr()
    End If

    ' ORDER BY
    If PeekWord("ORDER") Then
        m_Pos = m_Pos + 1
        ExpectWord "BY"
        ParseOrderByList
    End If
End Sub

Private Sub ParseInsertStmt()
    Dim expr As cSQLExprNode
    Dim rowValues As Collection
    Dim colName As String

    ExpectWord "INSERT"
    ExpectWord "INTO"
    ParseTableName

    ' Optional column list
    If PeekType() = ttLParen Then
        m_Pos = m_Pos + 1
        Do
            If PeekType() = ttWord Then
                m_InsertColumns.Add PeekValue()
                Advance
            Else
                Err.Raise vbObjectError + 606, "cSQLParser", "Expected column name in INSERT column list"
            End If
            If Not MatchToken(ttComma) Then Exit Do
        Loop
        ExpectToken ttRParen
    End If

    ' VALUES clause
    ExpectWord "VALUES"
    Do
        ExpectToken ttLParen
        Set rowValues = New Collection
        Do
            Set expr = ParseExpr()
            rowValues.Add expr
            If Not MatchToken(ttComma) Then Exit Do
        Loop
        ExpectToken ttRParen
        m_InsertValues.Add rowValues
        If Not MatchToken(ttComma) Then Exit Do
    Loop
End Sub

Private Sub ParseUpdateStmt()
    Dim expr As cSQLExprNode
    Dim colName As String

    ExpectWord "UPDATE"
    ParseTableName
    ExpectWord "SET"

    ' SET assignments
    Do
        If PeekType() <> ttWord Then
            Err.Raise vbObjectError + 607, "cSQLParser", "Expected column name in SET clause"
        End If
        m_UpdateColumns.Add PeekValue()
        Advance
        ExpectToken ttOperator
        Set expr = ParseExpr()
        m_UpdateExprs.Add expr
        If Not MatchToken(ttComma) Then Exit Do
    Loop

    ' WHERE
    If PeekWord("WHERE") Then
        m_Pos = m_Pos + 1
        Set m_WhereExpr = ParseExpr()
    End If
End Sub

Private Sub ParseDeleteStmt()
    ExpectWord "DELETE"
    ExpectWord "FROM"
    ParseTableName

    ' WHERE
    If PeekWord("WHERE") Then
        m_Pos = m_Pos + 1
        Set m_WhereExpr = ParseExpr()
    End If
End Sub


'#======================================================================================================================
'# List Parsing Helpers
'#======================================================================================================================
Private Sub ParseSelectList()
    Dim expr As cSQLExprNode

    ' Check for SELECT *
    If PeekType() = ttStar Then
        m_Pos = m_Pos + 1
        Set expr = New cSQLExprNode
        expr.NodeType = exStar
        m_SelectColumns.Add expr
        Exit Sub
    End If

    Do
        Set expr = ParseExpr()

        ' Check for alias
        If PeekWord("AS") Then
            m_Pos = m_Pos + 1
            expr.AliasName = PeekValue()
            Advance
        ElseIf PeekType() = ttWord And Not IsReservedWord(PeekValue()) Then
            expr.AliasName = PeekValue()
            Advance
        End If

        m_SelectColumns.Add expr
        If Not MatchToken(ttComma) Then Exit Do
    Loop
End Sub

Private Sub ParseTableName()
    If PeekType() = ttWord Then
        m_TableName = PeekValue()
        Advance
    Else
        Err.Raise vbObjectError + 608, "cSQLParser", "Expected table name"
    End If
End Sub

Private Sub ParseGroupByList()
    Do
        If PeekType() = ttWord Then
            m_GroupByColumns.Add PeekValue()
            Advance
        Else
            Err.Raise vbObjectError + 609, "cSQLParser", "Expected column name in GROUP BY"
        End If
        If Not MatchToken(ttComma) Then Exit Do
    Loop
End Sub

Private Sub ParseOrderByList()
    Dim dir As String
    Do
        If PeekType() = ttWord Then
            m_OrderByColumns.Add PeekValue()
            Advance
        Else
            Err.Raise vbObjectError + 610, "cSQLParser", "Expected column name in ORDER BY"
        End If

        dir = "ASC"
        If PeekWord("ASC") Then
            m_Pos = m_Pos + 1
            dir = "ASC"
        ElseIf PeekWord("DESC") Then
            m_Pos = m_Pos + 1
            dir = "DESC"
        End If
        m_OrderByDirections.Add dir

        If Not MatchToken(ttComma) Then Exit Do
    Loop
End Sub


'#======================================================================================================================
'# Expression Parsing (Recursive Descent)
'#======================================================================================================================
Private Function ParseExpr() As cSQLExprNode
    Set ParseExpr = ParseOrExpr()
End Function

Private Function ParseOrExpr() As cSQLExprNode
    Dim left As cSQLExprNode
    Dim node As cSQLExprNode

    Set left = ParseAndExpr()

    Do While PeekWord("OR")
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exBinaryOp
        node.Operator = "OR"
        Set node.LeftChild = left
        Set node.RightChild = ParseAndExpr()
        Set left = node
    Loop

    Set ParseOrExpr = left
End Function

Private Function ParseAndExpr() As cSQLExprNode
    Dim left As cSQLExprNode
    Dim node As cSQLExprNode

    Set left = ParseNotExpr()

    Do While PeekWord("AND")
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exBinaryOp
        node.Operator = "AND"
        Set node.LeftChild = left
        Set node.RightChild = ParseNotExpr()
        Set left = node
    Loop

    Set ParseAndExpr = left
End Function

Private Function ParseNotExpr() As cSQLExprNode
    Dim node As cSQLExprNode

    If PeekWord("NOT") Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exUnaryOp
        node.Operator = "NOT"
        Set node.LeftChild = ParseNotExpr()
        Set ParseNotExpr = node
    Else
        Set ParseNotExpr = ParseComparison()
    End If
End Function

Private Function ParseComparison() As cSQLExprNode
    Dim left As cSQLExprNode
    Dim node As cSQLExprNode
    Dim negated As Boolean
    Dim opVal As String

    Set left = ParseAddition()

    ' Check for NOT before BETWEEN, IN, LIKE
    negated = False
    If PeekWord("NOT") Then
        negated = True
        m_Pos = m_Pos + 1
    End If

    ' BETWEEN
    If PeekWord("BETWEEN") Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exBetween
        node.IsNegated = negated
        Set node.LeftChild = left
        Set node.RightChild = ParseAddition()
        ExpectWord "AND"
        Set node.ThirdChild = ParseAddition()
        Set ParseComparison = node
        Exit Function
    End If

    ' IN
    If PeekWord("IN") Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exInList
        node.IsNegated = negated
        Set node.LeftChild = left
        Set node.Args = New Collection
        ExpectToken ttLParen
        Do
            node.Args.Add ParseExpr()
            If Not MatchToken(ttComma) Then Exit Do
        Loop
        ExpectToken ttRParen
        Set ParseComparison = node
        Exit Function
    End If

    ' LIKE
    If PeekWord("LIKE") Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exBinaryOp
        node.Operator = "LIKE"
        node.IsNegated = negated
        Set node.LeftChild = left
        Set node.RightChild = ParseAddition()
        Set ParseComparison = node
        Exit Function
    End If

    ' IS [NOT] NULL
    If PeekWord("IS") Then
        m_Pos = m_Pos + 1
        Dim isNotNull As Boolean
        isNotNull = negated
        If MatchWord("NOT") Then
            isNotNull = True
        End If
        ExpectWord "NULL"
        Set node = New cSQLExprNode
        node.NodeType = exIsNull
        node.IsNegated = isNotNull
        Set node.LeftChild = left
        Set ParseComparison = node
        Exit Function
    End If

    ' If we consumed NOT but didn't find BETWEEN/IN/LIKE/IS, put it back
    If negated Then
        m_Pos = m_Pos - 1
        Set ParseComparison = left
        Exit Function
    End If

    ' Standard comparison operators
    If PeekType() = ttOperator Then
        opVal = PeekValue()
        Select Case opVal
            Case "=", "<>", "<", ">", "<=", ">="
                m_Pos = m_Pos + 1
                Set node = New cSQLExprNode
                node.NodeType = exBinaryOp
                node.Operator = opVal
                Set node.LeftChild = left
                Set node.RightChild = ParseAddition()
                Set ParseComparison = node
                Exit Function
        End Select
    End If

    Set ParseComparison = left
End Function

Private Function ParseAddition() As cSQLExprNode
    Dim left As cSQLExprNode
    Dim node As cSQLExprNode
    Dim opVal As String

    Set left = ParseMultiplication()

    Do
        If PeekType() = ttOperator Then
            opVal = PeekValue()
            If opVal = "+" Or opVal = "-" Then
                m_Pos = m_Pos + 1
                Set node = New cSQLExprNode
                node.NodeType = exBinaryOp
                node.Operator = opVal
                Set node.LeftChild = left
                Set node.RightChild = ParseMultiplication()
                Set left = node
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    Set ParseAddition = left
End Function

Private Function ParseMultiplication() As cSQLExprNode
    Dim left As cSQLExprNode
    Dim node As cSQLExprNode
    Dim curType As Long

    Set left = ParseUnary()

    Do
        curType = PeekType()
        If curType = ttStar Or (curType = ttOperator And PeekValue() = "/") Then
            Dim opStr As String
            If curType = ttStar Then
                opStr = "*"
            Else
                opStr = "/"
            End If
            m_Pos = m_Pos + 1
            Set node = New cSQLExprNode
            node.NodeType = exBinaryOp
            node.Operator = opStr
            Set node.LeftChild = left
            Set node.RightChild = ParseUnary()
            Set left = node
        Else
            Exit Do
        End If
    Loop

    Set ParseMultiplication = left
End Function

Private Function ParseUnary() As cSQLExprNode
    Dim node As cSQLExprNode

    If PeekType() = ttOperator And PeekValue() = "-" Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exUnaryOp
        node.Operator = "-"
        Set node.LeftChild = ParsePrimary()
        Set ParseUnary = node
    Else
        Set ParseUnary = ParsePrimary()
    End If
End Function

Private Function ParsePrimary() As cSQLExprNode
    Dim node As cSQLExprNode
    Dim curType As Long
    Dim curVal As String
    Dim upper As String

    curType = PeekType()
    curVal = PeekValue()

    ' Number literal
    If curType = ttNumber Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exLiteral
        node.LiteralType = litNumber
        node.NumValue = CDbl(curVal)
        node.StringValue = curVal
        Set ParsePrimary = node
        Exit Function
    End If

    ' String literal
    If curType = ttString Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exLiteral
        node.LiteralType = litString
        node.StringValue = curVal
        Set ParsePrimary = node
        Exit Function
    End If

    ' Star (for SELECT * or COUNT(*))
    If curType = ttStar Then
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exStar
        Set ParsePrimary = node
        Exit Function
    End If

    ' Parenthesized expression
    If curType = ttLParen Then
        m_Pos = m_Pos + 1
        Set node = ParseExpr()
        ExpectToken ttRParen
        Set ParsePrimary = node
        Exit Function
    End If

    ' Word: could be keyword literal, function call, or column reference
    If curType = ttWord Then
        upper = UCase$(curVal)

        ' NULL literal
        If upper = "NULL" Then
            m_Pos = m_Pos + 1
            Set node = New cSQLExprNode
            node.NodeType = exLiteral
            node.LiteralType = litNull
            Set ParsePrimary = node
            Exit Function
        End If

        ' TRUE literal
        If upper = "TRUE" Then
            m_Pos = m_Pos + 1
            Set node = New cSQLExprNode
            node.NodeType = exLiteral
            node.LiteralType = litBoolean
            node.BoolValue = True
            Set ParsePrimary = node
            Exit Function
        End If

        ' FALSE literal
        If upper = "FALSE" Then
            m_Pos = m_Pos + 1
            Set node = New cSQLExprNode
            node.NodeType = exLiteral
            node.LiteralType = litBoolean
            node.BoolValue = False
            Set ParsePrimary = node
            Exit Function
        End If

        ' Aggregate function call
        If IsAggregateName(upper) Then
            If m_Pos + 1 <= m_TokenCount Then
                If m_TokTypes(m_Pos + 1) = ttLParen Then
                    m_Pos = m_Pos + 1
                    Set node = New cSQLExprNode
                    node.NodeType = exFunction
                    node.FuncName = upper
                    Set node.Args = New Collection
                    ExpectToken ttLParen

                    ' Check for DISTINCT inside function
                    If MatchWord("DISTINCT") Then
                        node.IsDistinctFunc = True
                    End If

                    ' Check for COUNT(*)
                    If PeekType() = ttStar Then
                        m_Pos = m_Pos + 1
                        Dim starNode As cSQLExprNode
                        Set starNode = New cSQLExprNode
                        starNode.NodeType = exStar
                        node.Args.Add starNode
                    Else
                        Do
                            node.Args.Add ParseExpr()
                            If Not MatchToken(ttComma) Then Exit Do
                        Loop
                    End If

                    ExpectToken ttRParen
                    Set ParsePrimary = node
                    Exit Function
                End If
            End If
        End If

        ' Column reference (possibly table.column)
        m_Pos = m_Pos + 1
        Set node = New cSQLExprNode
        node.NodeType = exColumn
        node.ColumnName = curVal

        ' Check for table.column notation
        If PeekType() = ttDot Then
            m_Pos = m_Pos + 1
            node.TablePrefix = curVal
            If PeekType() = ttStar Then
                node.NodeType = exStar
                m_Pos = m_Pos + 1
            Else
                node.ColumnName = PeekValue()
                Advance
            End If
        End If

        Set ParsePrimary = node
        Exit Function
    End If

    Err.Raise vbObjectError + 611, "cSQLParser", "Unexpected token '" & curVal & "'"
End Function


'#======================================================================================================================
'# Reserved Word Check
'#======================================================================================================================
Private Function IsReservedWord(ByVal word As String) As Boolean
    Select Case UCase$(word)
        Case "SELECT", "FROM", "WHERE", "INSERT", "INTO", "VALUES", "UPDATE", "SET", "DELETE", _
             "ORDER", "BY", "GROUP", "HAVING", "DISTINCT", "AS", "AND", "OR", "NOT", _
             "BETWEEN", "IN", "LIKE", "IS", "NULL", "TRUE", "FALSE", "ASC", "DESC", _
             "TOP", "COUNT", "SUM", "AVG", "MIN", "MAX", "ON", "JOIN", "LEFT", "RIGHT", _
             "INNER", "OUTER", "CROSS", "UNION", "ALL", "EXISTS", "CASE", "WHEN", "THEN", _
             "ELSE", "END"
            IsReservedWord = True
        Case Else
            IsReservedWord = False
    End Select
End Function
