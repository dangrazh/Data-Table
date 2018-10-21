VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
'Option Compare Text

'# <author> Daniel Grass
'# <mail> dani.grass@bluewin.ch

'#Region
'# Public Subs, Functions and Properties
'#======================================================================================================================
'# Accessible in this class
'#======================================================================================================================
'     |> Get | --- About                :: Returns description of the class.
'     |> --------- CreateLogFile        :: Creates a logfile for output and stores it in the specified [storageDirectory].
'     |> Get | Let DirectoryPath        :: Returns the directory location of the current logfile.
'     |> Get | --- Name                 :: Returns name of the data table (Default Property).
'     |> Get | --- Version              :: Returns version string for the class [e.g. #.# (year)].


'#======================================================================================================================
'# Usage
'#======================================================================================================================
'
'
'----------------------------------------------


'#======================================================================================================================
'# References
'#======================================================================================================================
#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

#If Win64 Then
    Private Const PTR_LENGTH As Long = 8
    Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    Private Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function ArrayPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As LongPtr
    Private Declare PtrSafe Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As LongPtr, ByVal Value As Long)
    Private Declare PtrSafe Function SafeArrayRedim Lib "oleaut32" (ByVal saPtr As LongPtr, saBound As Long) As LongPtr
#Else
    Private Const PTR_LENGTH As Long = 4
    Private Declare Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    Private Declare Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As LongPtr)
    Private Declare Function ArrayPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
    Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
    Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal saPtr As Long, saBound As Long) As Long
#End If


'#======================================================================================================================
'# Dependencies to other classes
'#======================================================================================================================
' cIndex                                             :: is a single Index of a data column
' cHashTable (thourgh cIndex as well)   :: is used to access column indices & as underlying data structure
'                                                       to unique indices
' cBPlusTree (thourgh cIndex)                        :: is used as underlying data structure to non-unique indices


'#======================================================================================================================
'# Application Constants, Enumerations & Types
'#======================================================================================================================
#Const C_CON_OFFICE_APP = "XLS"   '##### CHANGE THIS TO ANYTHING ELSE THEN 'XLS' WHEN NOT USING THIS CLASS IN EXCEL #####

Const C_WANT_FREE_PERCENT = 0.1   ' 0.1 translates in 10% free space
Const C_MIN_FREE = 10             ' Min unused space when resizing
Const C_MAX_WAIT_CYCLES = 100     ' Number of iterations to check if object is locked before throwing time out error
Const C_WAIT_TIME_MS = 100        ' Number of milliseconds to wait for next check if object is locked

Private Const C_CUT_OFF As Long = 100 'move into Insertionsort if list (or sublist) has less 100 items

Private Const C_Name As String = "cDataTable"


Public Enum MatchType
    Equals = 1
    DoesNotEqual = 2
    MatchesPattern = 3
    DoesNotMatchPattern = 4
    GreaterThen = 5
    GreaterThenOrEqual = 6
    LessThen = 7
    LessThenOrEqual = 8
End Enum

Public Enum OutputType
    OverwriteIfExists = 1
    AppendIfExists = 2
    ErrorIfExists = 3
End Enum

Private Enum ResizeType
    IncludingWantFreeSpace = 1
    WithoutEmptyTrailRecords = 2
End Enum

Private Enum ColumnType
    dtNull = vbNull
    dtInteger = vbInteger
    dtLong = vbLong
    dtSingle = vbSingle
    dtDouble = vbDouble
    dtCurrency = vbCurrency
    dtDate = vbDate
    dtString = vbString
    dtObject = vbObject
    dtError = vbError
    dtBoolean = vbBoolean
    dtVariant = vbVariant
    dtDataObject = vbDataObject
    dtDecimal = vbDecimal
    dtByte = vbByte
    dtUserDefinedType = vbUserDefinedType
    dtArray = vbArray
End Enum

Private Type DataColumn
    Number As Long
    Name As String
    Type As ColumnType
    IsSorted As Boolean
    IsIndexed As Boolean
    Index As cIndex
End Type


Private Type ErrorCode
    errNumber As Long
    errDescrption As String
End Type
    
    
'#======================================================================================================================
'# API Constants, Enumerations & Types
'#======================================================================================================================

'Type Declarations needed for SafeArray hacks

'The bounds of the SafeArray
Private Type SafeArrayBound
    cElements    As Long
    lLbound      As Long
End Type

Private Type SAFEARRAY1D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As LongPtr
    Bounds(0 To 0)  As SafeArrayBound
End Type

Private Type SAFEARRAY2D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As LongPtr
    Bounds(0 To 1)  As SafeArrayBound
End Type


'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private m_Name As String                        ' The name of the data table
Private m_List() As Variant                     ' The list array.
Private m_IdxToList() As Long                   ' The index table to access the list array
Private m_ListTransposed() As Variant           ' The transposed list array
Private m_NumCols As Long                       ' Number of Columns
Private m_NumItems As Long                      ' Last index in use.
Private m_ArraySize As Long                     ' Size of the list array.
Private m_ShrinkWhen As Long                    ' Shrink if m_NumItems < this.
Private m_GarbageCount As Long                  ' The number of garbage entries.
Private m_MaxGarbage As Long                    ' Collect when m_GarbageCount > this.
Private m_FirstGarbage As Long                  ' The index of the 1st record containing garbage
Private m_LastGarbage As Long                   ' The index of the last record containing garbage
Private m_CollectGarbage As Boolean             ' Flag to determine if we should collect garbage
Private m_EnableObjects As Boolean              ' Flag to determine if we allow objects to be stored
Private m_Columns() As DataColumn               ' Array containing the data columns
Private m_SearchLastItem As Variant             ' The last item that was searched
Private m_SearchLastColumn As Long              ' Contains the column index of the column last searched
Private m_Bookmark As Long                      ' Bookmark pointing to the current record
Private m_HasHeaders As Boolean                 ' Indicates if column headers were provided
Private m_IsDirty As Boolean                    ' Indicates if the array needs cleanup and a refresh of the transposed list
Private m_IsLocked As Boolean                   ' Indicates if the class is currently being updated
Private m_BOF As Boolean                        ' Indicates that the cursor has reached the beginning of the table
Private m_EOF As Boolean                        ' Indicates that the cursor has reached the end of the table
Private m_Errors(1 To 10) As ErrorCode          ' List of all error codes

'#Region
'#======================================================================================================================
'# Class Initialization, Termination & Properties
'#======================================================================================================================


Private Sub Class_Initialize()
' ************************************************
' Class constructor.
' ************************************************

    'Debug.Print "|> Initializing:= " & Me.Name
    
    'Set list to minimal size
    ReDim m_List(1, 1)
    ReDim m_IdxToList(1 To 1)
    
    'set the default name
    m_Name = "DataTable"
    
    'initialze the boolean states
    m_HasHeaders = False
    m_IsDirty = False
    m_CollectGarbage = True
    m_EOF = False
    m_BOF = False
    m_IsLocked = False
        
    'initialize any default / start values
    m_NumCols = 1
    m_Bookmark = -1
    m_FirstGarbage = 1
    m_LastGarbage = 1
    
    'initialize error codes
    m_Errors(1).errNumber = 513:  m_Errors(1).errDescrption = "List is not empty. Range can only be added to empty list!"
    m_Errors(2).errNumber = 514:  m_Errors(2).errDescrption = "The list of Header Names does not match the number of columns you specified!"
    m_Errors(3).errNumber = 515:  m_Errors(3).errDescrption = "The no of items in the input record does not match the number of columns in the data table!"
    m_Errors(4).errNumber = 516:  m_Errors(4).errDescrption = "Assigning a value to an item out of boundaries of the data table is not possible!"
    m_Errors(5).errNumber = 517:  m_Errors(5).errDescrption = "The column name you specified could not be found in the data table!"
    m_Errors(6).errNumber = 518:  m_Errors(6).errDescrption = "Accessing an item needs either a row number or an active cursor from one of the RsXxxx methods!"
    m_Errors(7).errNumber = 519:  m_Errors(7).errDescrption = "The two defined tables do not have the same data structure. Cannot append two tables with different data strucutres!"
    m_Errors(8).errNumber = 520:  m_Errors(8).errDescrption = "No of records in Data Table exeed excel row limit! Use compression mode to output bigger data sets."
    m_Errors(9).errNumber = 901:  m_Errors(9).errDescrption = "Timeout/Deadlock - the data table is locked for update, time limit exeeded!"
    m_Errors(10).errNumber = 521:  m_Errors(10).errDescrption = "Output File already exists!"
    

End Sub


Private Sub Class_Terminate()
' ************************************************
' Class destructor.
' ************************************************
    
    Dim ptrList As Long
    Dim ptrListTransposed As Long

    'Debug.Print "|> Terminating:= " & Me.Name
    
    'empty the list
    ReDim m_List(1, 1)
    ReDim m_ListTransposed(1, 1)


End Sub

Public Property Get Version() As String
' ************************************************
' Version string of the current class.
' Contains a list of (historical) changes to the class within the comments of the procedure.
' ************************************************
   'Version = "Version 1.0.0.0 (2017)"  'Initial release
   'Version = "Version 1.0.1.0 (06/2017)"   'Including bug fixes for garbage collection and output to excel
   'Version = "Version 1.1.0.0 (06/2017)"   'Including features to make class thread save
   'Version = "Version 1.2.0.0 (12/2017)"   'Including new features on columns and 1st version of indexing
   'Version = "Version 1.2.1.0 (01/2018)"   'Including x64 compatibility
   'Version = "Version 1.2.2.1 (01/2018)"    'Fixed pattern matching and SortUnstable
   Version = "Version 1.2.2.2 (02/2018)"    'Fixed RecordAdd to handle both arrays and ParamArray
   

End Property

Public Property Get About() As String
' ***********************************************
' String that describes the current class.
' ***********************************************
    About = "Data Table Class providing advanced array based data handling. Version: " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details please contact the author."
End Property

Public Property Get ClassName() As String
' ***********************************************
' Returns the name of the class.
' ***********************************************
    
    ClassName = C_Name
    
End Property

Public Property Get Name() As String
' ***********************************************
' Returns the name of the data table.
' ***********************************************
    
    Name = m_Name
    
End Property

Public Property Let Name(Value As String)
' ***********************************************
' Sets the name of the data table.
' ***********************************************

    m_Name = Value

End Property

Public Property Get Headers() As Collection
' ***********************************************
' Returns the headers of the table as collection
' ***********************************************

    Dim cHeaders As New Collection
    Dim i As Long

    For i = LBound(m_Columns) To UBound(m_Columns)
        cHeaders.Add m_Columns(i).Name
    Next i
    
    Set Headers = cHeaders

End Property

Public Property Get HasHeaders() As Boolean
' ***********************************************
' Returns the HasHeaders state of the class.
' ***********************************************

        HasHeaders = m_HasHeaders

End Property

'Public Property Let HasHeaders(bValue As Boolean)
' ***********************************************
' Sets the HasHeaders state of the class.
' ***********************************************

'        m_HasHeaders = bValue

'End Property

Public Property Get NumItems() As Long
' ***********************************************
' Return the number of items.
' ***********************************************

    NumItems = m_NumItems

End Property

Private Property Get ArraySize() As Long
' ***********************************************
' Return the size of the array.
' ***********************************************

    ArraySize = m_ArraySize

End Property

Public Property Get NumCols() As Long
' ***********************************************
' Return the number of columns.
' ***********************************************

        NumCols = m_NumCols

End Property

Public Property Get GarbageCollection() As Boolean
' ***********************************************
' Returns the CollectGarbage state of the class.
' ***********************************************

    GarbageCollection = m_CollectGarbage

End Property

Public Property Let GarbageCollection(Value As Boolean)
' ***********************************************
' Sets the CollectGarbage state of the class.
' ***********************************************
    
    If Value = False Then
        'Set the state to false
        m_CollectGarbage = False
     Else
        'set the state to true
        m_CollectGarbage = True
        'If table is dirty, collect garbage
        If m_IsDirty = True Then
            CollectGarbage
        End If
    End If

End Property

Public Property Get ObjectStorageEnabled() As Boolean
' ***********************************************
' Returns the ObjectStorageEnabled state of the class.
' ***********************************************

    ObjectStorageEnabled = m_EnableObjects


End Property


Public Property Let ObjectStorageEnabled(Value As Boolean)
' ***********************************************
' Sets the ObjectStorageEnabled state of the class.
' ***********************************************

    m_EnableObjects = Value

End Property

Public Property Get TableSummary() As String
' ***********************************************
' This displays a summary of the table
' ## This is the default property ##
' ***********************************************

    TableSummary = "Table Name: " & m_Name & vbCrLf & _
                    "No of Cols: " & m_NumCols & vbCrLf & _
                    "No of Recs: " & m_NumItems & vbCrLf & _
                    "Obj enabled: " & m_EnableObjects

End Property


'Public Property Get SADescrPtr() As LongPtr
'' ***********************************************
'' Return the pointer to the SafeArray Descriptor
'' ***********************************************
'
'
'    SADescrPtr = getSafeArrayDescrPtr(m_List)
'
'End Property
'
'Public Property Get SAStructPtr() As LongPtr
'' ***********************************************
'' Return the pointer to the SafeArray Descriptor
'' ***********************************************
'
'
'    SAStructPtr = getSafeArrayStructPtr(m_List)
'
'End Property


Public Property Get RsEOF() As Boolean
' ***********************************************
' Return the EOF State of the recordset
' ***********************************************

    RsEOF = m_EOF

End Property

Public Property Get RsBOF() As Boolean
' ***********************************************
' Return the EOF State of the recordset
' ***********************************************

    RsBOF = m_BOF

End Property

Public Property Get RsBookmark() As Long
' ***********************************************
' Return the pointer to the current record within the recordset
' ***********************************************

    RsBookmark = m_Bookmark

End Property

Public Property Let RsBookmark(Value As Long)
' ***********************************************
' Set the pointer of the current record within the recordset
' to a specific position
' ***********************************************

    If Value > m_NumItems Then
        m_EOF = True
     Else
        m_Bookmark = Value
    End If

End Property


Public Property Get ItemName(Index As Variant) As String
' ***********************************************
' Return the name of the Item (Column Header)
' ***********************************************

    ItemName = GetColumnName(Index)

End Property


Public Property Get Item(Index As Variant, Optional Row As Long = -1) As Variant
' ***********************************************
' Return a given item at index / position
' ***********************************************


    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
        If m_EnableObjects = True Then
            If IsObject(ItemGet(Row, CStr(Index))) Then
                Set Item = ItemGet(Row, CStr(Index))
             Else
                Item = ItemGet(Row, CStr(Index))
            End If
         Else
            Item = ItemGet(Row, CStr(Index))
        End If
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        If m_EnableObjects = True Then
            If IsObject(ItemGet(m_Bookmark, CStr(Index))) Then
                Set Item = ItemGet(m_Bookmark, CStr(Index))
             Else
                Item = ItemGet(m_Bookmark, CStr(Index))
            End If
         Else
            Item = ItemGet(m_Bookmark, CStr(Index))
        End If
     Else
        MsgBox "Unhandled condition in Property Get 'Item'", vbCritical
        Exit Property
    End If


End Property


Public Property Let Item(Index As Variant, Optional Row As Long = -1, Value As Variant)
' ***********************************************
' Set the value of a given item at position / index
' ***********************************************
    
    
    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
            Call ItemSet(Row, CStr(Index), Value)
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        Call ItemSet(m_Bookmark, CStr(Index), Value)
     Else
        MsgBox "Unhandled condition in Property Let 'Item'", vbCritical
        Exit Property
    End If

End Property


Public Property Get Record(Optional Row As Long = -1) As Variant()
' ***********************************************
' Return an record from the list by position.
' ***********************************************

    Dim iCol As Long
    Dim tmpRec() As Variant
    Dim position As Long
        
        
    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
        position = Row
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        position = m_Bookmark
     Else
        MsgBox "Unhandled condition in Property Get 'Record'", vbCritical
        Exit Property
    End If
        
        
        
    If position < 1 Or position > m_NumItems Then
        ' Out of bounds. Return Null.
        Record = Null
     Else

        ReDim tmpRec(1 To m_NumCols)
        
        'For iCol = 1 To m_NumCols
        '        ' Return the record.
        '        tmpRec(iCol) = m_List(iCol, position)
        'Next iCol
        RecordRead tmpRec, position
        
        Record = tmpRec
                
    End If

End Property

Public Property Let Record(Optional Row As Long = -1, Record() As Variant)
' ***********************************************
' Update a record in the list by position.
' ***********************************************

    Dim iCol As Long
    Dim tmpRec() As Variant
    Dim position As Long
        
    
    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
        position = Row
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        position = m_Bookmark
     Else
        MsgBox "Unhandled condition in Property Let 'Record'", vbCritical
        Exit Property
    End If
        
        
    If position < 1 Or position > m_NumItems Then
        ' Out of bounds, raise error
        ' The range 513-65535 is available for user errors.
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(4).errNumber, Me.Name, m_Errors(4).errDescrption
     Else


        If UBound(Record) = UBound(m_List, 1) Then
            'handle thread safety
            CheckLockStatus
            LockForUpdate
                        
            'For iCol = 1 To m_NumCols
            '    ' Set the record values.
            '    m_List(iCol, position) = Record(iCol)
            'Next iCol
            RecordWrite position, Record, 0
        
            'handle thread safety
            ReleaseLock
            
            'mark as dirty
            m_IsDirty = True
            
        Else
            'raise error -> Not the same no of column Names as NoOfColumns specified
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + 515, Me.Name, "The no of items in the input record does not match the number of columns in the data table!"
        End If
                
    End If

    

End Property


'#Region
'#======================================================================================================================
'# Table Definition, Index Handling and Statistics
'#======================================================================================================================
Public Sub DefineTable(NoOfColumns As Long, Optional ColumnHeaders As String = "n/a", Optional NoOfRows As Long = 1)
' ***********************************************
' Define structure of array and create Header List
' ***********************************************

        Dim i As Long
        Dim aTmp As Variant

        'Size the list
        ReDim m_List(1 To NoOfColumns, 1 To NoOfRows)
        m_ArraySize = NoOfRows
        
        'fill in the Headers
        ReDim m_Columns(1 To NoOfColumns)
        
        'Name the Fields Field<n> if no field list is provided
        If ColumnHeaders = "n/a" Or InStr(ColumnHeaders, ",") = 0 Then
                For i = 1 To NoOfColumns
                        m_Columns(i).Name = "Field" & i
                        m_Columns(i).Number = i
                        m_Columns(i).IsSorted = False
                        m_Columns(i).IsIndexed = False
                        m_Columns(i).Type = dtVariant
                Next i
            
                m_NumCols = NoOfColumns
            
         ElseIf InStr(ColumnHeaders, ",") > 0 Then
                aTmp = Split(ColumnHeaders, ",")
                If UBound(aTmp) + 1 = NoOfColumns Then
                    For i = 1 To UBound(aTmp) + 1
                            m_Columns(i).Name = Trim(aTmp(i - 1))
                            m_Columns(i).Number = i
                            m_Columns(i).IsSorted = False
                            m_Columns(i).IsIndexed = False
                            m_Columns(i).Type = dtVariant
                    Next i
                    
                    m_NumCols = NoOfColumns
                    m_HasHeaders = True
                    
                Else
                    'raise error -> Not the same no of column Names as NoOfColumns specified
                    ' The range 513-65535 is available for user errors.
                    ' For class errors, you add vbObjectError to the error number.
                    Err.Raise vbObjectError + 514, Me.Name, "The list of Header Names does not match the number of columns you specified!"
                End If
        End If
        
End Sub

Public Sub AnalyzeDataTypes()
' ***********************************************
' Do random spot checks of data contents of each column
' and determine the data type of the column
' ***********************************************

    '§§§ Future feature §§§
        
    Dim iCol As Long
    Dim dblPctSampleSize As Double: dblPctSampleSize = 0.1
    Dim lSampleSize As Long
    Dim lStepSize As Long
    Dim lCount As Long: lCount = 1
    Dim tyCol As ColumnType
    
    lSampleSize = m_NumItems * dblPctSampleSize
    lStepSize = m_NumItems / lSampleSize
    
    For iCol = LBound(m_Columns) To UBound(m_Columns)
    
        Me.RsBookmark = lStepSize * lCount
        Do While Not Me.RsEOF
            'determine type & keep statistics
            '§§§ pending implementation §§§
            
         lCount = lCount + 1
         Me.RsBookmark = lStepSize * lCount
        Loop
    
        'get entry with most counts
        '§§§ pending implementation §§§
        
        'update the data type of the column
        m_Columns(iCol).Type = tyCol
        
    Next iCol
        

End Sub


Public Sub AddIndex(ColumnName As String, idxType As IndexType)
' ***********************************************
' Create an index on a given column
' ***********************************************

    Dim oIdx As New cIndex
    Dim iColIdx As Long
    
    ReDim aData(1 To m_NumItems)
    
    If m_IsDirty = True Then
        Call CleanUpTable
    End If

    iColIdx = GetColumnIndex(ColumnName)
    oIdx.IndexType = idxType
    oIdx.Create Me, ColumnName
    Set m_Columns(iColIdx).Index = oIdx
    m_Columns(iColIdx).IsIndexed = True


End Sub


'Public Sub RefreshIndices()
' ***********************************************
' Refresh all indices in the table
' ***********************************************

        '§§§ Future feature §§§

'End Sub

'Public Sub RunStats()
' ***********************************************
' Refresh all stats and indices
' ***********************************************

        '§§§ Future feature §§§

'End Sub


'#Region
'#======================================================================================================================
'# Data handling (add / remove) and list maintenance
'#======================================================================================================================
        
Public Sub TruncateTable(AskForConfirmation As Boolean)

    Dim retVal As Integer

    If AskForConfirmation = True Then
        MsgBox "This will delete all data from the table! Do you really want to continue?", vbQuestion + vbYesNoCancel, "Truncating Table"
        If retVal = vbNo Or retVal = vbCancel Then
            Exit Sub
        End If
    End If
        
    'clean the list by resizing it to a single record
    ReDim m_List(1 To m_NumCols, 1 To 1)
    'reset the number of items to 0
    m_NumItems = 0
    'reset the array size to 1
    m_ArraySize = 1
    

End Sub

Private Function ItemGet(ByVal position As Long, ByVal Index As String) As Variant
' ***********************************************
' Return an item from the list by position.
' ***********************************************

    Dim iCol As Long
    Dim vItemTmp As Variant
    'Dim bColFound As Boolean: bColFound = False

    'handle thread safety
    CheckLockStatus
    LockForUpdate


    If position < 1 Or position > m_NumItems Then
        ' Out of bounds. Return Null.
        ItemGet = Null
    Else
    
        iCol = GetColumnIndex(Index)
        
        If iCol = -1 Then
            'Column Name not found. Return Null.
            ItemGet = Null
         Else
            ' Return the item.
            'ItemGet = m_List(iCol, position)
            ItemRead vItemTmp, iCol, position
        End If
        
        If m_EnableObjects = True Then
            If IsObject(vItemTmp) Then
                Set ItemGet = vItemTmp
             Else
                ItemGet = vItemTmp
            End If
         Else
            ItemGet = vItemTmp
         End If
    End If
        
    'handle thread safety
    ReleaseLock
        
End Function

Private Sub ItemSet(ByVal position As Long, ByVal Index As String, Value As Variant)
' ***********************************************
' Set the value of an item in the list by position.
' ***********************************************

    Dim iCol As Long
    'Dim bColFound As Boolean: bColFound = False

    If position < 1 Or position > m_NumItems Then
        ' Out of bounds, raise error
        ' The range 513-65535 is available for user errors.
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(4).errNumber, Me.Name, m_Errors(4).errDescrption
    Else
    
        iCol = GetColumnIndex(Index)
        If iCol = -1 Then
            'Column Name not found, raise error
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + m_Errors(5).errNumber, Me.Name, m_Errors(5).errDescrption
         Else
            'handle thread safety
            CheckLockStatus
            LockForUpdate
         
            ' Set the value of the item
            'm_List(iCol, position) = value
            ItemWrite Value, iCol, position
            
            'handle thread safety
            ReleaseLock
            
            'mark as dirty
            m_IsDirty = True
            
        End If
                
    End If
        
End Sub
        
Public Sub RecordAddOld(ByVal aRecord As Variant)
' ***********************************************
' Add a record to the list
' ***********************************************

        Dim iCol As Long
 
        'Check if no of items in the record is in line with the number of colums in the list
        If UBound(aRecord) = UBound(m_List, 1) Then
        
            'handle thread safety
            CheckLockStatus
            LockForUpdate
    
            'increase the row index
            m_NumItems = m_NumItems + 1
            'resize the list if required
            If m_NumItems > m_ArraySize Then ResizeList (IncludingWantFreeSpace)

            'add the record to the list
            For iCol = 1 To m_NumCols
                    m_List(iCol, m_NumItems) = aRecord(iCol)
            Next iCol
            
            'set bookmark to the new record
            m_Bookmark = m_NumItems
            
            'mark data table as dirty
            m_IsDirty = True
            
            'handle thread safety
            ReleaseLock

         Else
            
            'raise error -> Not the same no of column Names as NoOfColumns specified
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + 515, Me.Name, "The no of items in the input record does not match the number of columns in the data table!"
                
        End If
        
        
End Sub

Public Sub RecordAdd(ParamArray Record() As Variant)
' ***********************************************
' Add a record to the list - based on Param Array
' ***********************************************

        Dim iCol As Long
        Dim iColOffset As Integer: iColOffset = 0
        Dim lNoOfCols As Long
        Dim aRecTmp As Variant

    
        'Check if 1st element of ParamArray is an array itself,
        'if yes make the 1st element the record
        If IsArray(Record(0)) Then
            aRecTmp = Record(0)
            Record = aRecTmp
        End If
    
    
        'Check if no of items in the record is in line with the number of colums in the list
        'Attention: the ParamArray is a 0-based array, so we need to consider this here
        lNoOfCols = (UBound(Record) - LBound(Record)) + 1
        If LBound(Record) = 0 Then
            iColOffset = -1
        End If
 
        If lNoOfCols = UBound(m_List, 1) Then
        
            'handle thread safety
            CheckLockStatus
            LockForUpdate
                
            'increase the row index
            m_NumItems = m_NumItems + 1
            'resize the list if required
            If m_NumItems > m_ArraySize Then ResizeList (IncludingWantFreeSpace)

            'add the record to the list -> can't use RecordWrite as ParamArray
            'does not work as input to other procedure
            'For iCol = 1 To m_NumCols
            '        m_List(iCol, m_NumItems) = Record(iCol + iColOffset)
            'Next iCol
            If m_EnableObjects Then
                'add the record to the list
                For iCol = 1 To m_NumCols
                        If IsObject(Record(iCol + iColOffset)) Then
                            Set m_List(iCol, m_NumItems) = Record(iCol + iColOffset)
                         Else
                            m_List(iCol, m_NumItems) = Record(iCol + iColOffset)
                        End If
                Next iCol
             Else
                'add the record to the list
                For iCol = 1 To m_NumCols
                        m_List(iCol, m_NumItems) = Record(iCol + iColOffset)
                Next iCol
            End If
                        
                        
            'mark data table as dirty
            m_IsDirty = True
            
            'handle thread safety
            ReleaseLock
                     
         Else
         
            'raise error -> Not the same no of column Names as NoOfColumns specified
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + 515, Me.Name, "The no of items in the input record does not match the number of columns in the data table!"
                
        End If
        
        
End Sub

Public Sub RecordRemove(ByVal position As Long)
' ***********************************************
' Remove an item from the list.
' ***********************************************

    Dim iCol As Long

    'handle thread safety
    CheckLockStatus
    LockForUpdate


    '*** this approach is too slow if the table has a lot of columns ****
    ' set the record to Null
    For iCol = 1 To m_NumCols
            m_List(iCol, position) = "#DELETED"
    Next iCol

'    'mark the 1st element of the record as deleted (faster approach)
'    m_List(1, position) = "#DELETED"
        
    'increase the garbage counter
    m_GarbageCount = m_GarbageCount + 1
    
    'Update the 1st and the last garbage if required
    If position < m_FirstGarbage Or m_FirstGarbage = 0 Then
        m_FirstGarbage = position
    End If
    
    If position > m_LastGarbage Or m_LastGarbage = 0 Then
        m_LastGarbage = position
    End If
    
    'Check if garbage has to be collected and do so if yes
    If m_CollectGarbage = True Then
        'collect garbage if required
        If m_GarbageCount > m_MaxGarbage Then
            CollectGarbage
        End If
    End If

    'mark data table as dirty
    m_IsDirty = True
    
    'handle thread safety
    ReleaseLock

        
End Sub


#If C_CON_OFFICE_APP = "XLS" Then

Public Sub LoadRange(InputTable As Range, TableHasHeaders As Boolean)
' ***********************************************
' Add a cell range to the empty table
' ***********************************************


        Dim iCol As Long
        Dim iColCount As Long
        
        Dim rHeader As Range
        Dim cField As Range
        Dim iRowStart As Long: iRowStart = InputTable.Row
        Dim iColStart As Long: iColStart = InputTable.Column
        Dim iCols As Long: iCols = InputTable.Columns.Count
        Dim iRows As Long: iRows = InputTable.Rows.Count
        
        Dim ListTemp() As Variant
        Dim ptrList As LongPtr
        Dim ptrListTransposed As LongPtr
        Dim ptrListTemp As LongPtr
            
        Dim want_free As Long
        
        
        Dim mTimer As New CStopWatch

        
        'handle thread safety
        CheckLockStatus
        LockForUpdate
        
        
        If m_NumItems > 0 Then
                'raise error -> only allowed with empty list
                ' The range 513-65535 is available for user errors.
                ' For class errors, you add vbObjectError to the error number.
                Err.Raise vbObjectError + 513, Me.Name, "List is not empty. Range can only be added to empty list!"
        End If
        
        'Performance Improvement: if has headers, take 1st row for header creation and then reduce range size by 1st row
        'this eliminates the need to call "CleanUpTable" (i.e. collect the garbage) at the end of the procedure
        iColCount = iCols
        ReDim m_Columns(1 To iColCount)
        If TableHasHeaders = True Then
                For iCol = 1 To iColCount
                        'm_Columns(iCol) = InputTable.Cells(1, (iColStart - 1) + iCol).value
                        m_Columns(iCol).Name = InputTable.Cells(1, iCol).Value
                        m_Columns(iCol).Number = iCol
                        m_Columns(iCol).IsSorted = False
                        m_Columns(iCol).IsIndexed = False
                        m_Columns(iCol).Type = dtVariant
                Next iCol
                Set InputTable = InputTable.Offset(1, 0).Resize(iRows - 1, iCols)
        Else
                For iCol = 1 To iColCount
                        m_Columns(iCol).Name = "Field" & iCol
                        m_Columns(iCol).Number = iCol
                        m_Columns(iCol).IsSorted = False
                        m_Columns(iCol).IsIndexed = False
                        m_Columns(iCol).Type = dtVariant
                Next iCol
        End If
        m_HasHeaders = True
        
        
        'Now load the data range without header into the array
        'mTimer.Reset
        m_List = InputTable
        'Debug.Print "Assigning the input range to the variant array took: " & mTimer.MillisecondsElapsed & " milliseconds..."
        
        m_NumItems = UBound(m_List)
        m_ArraySize = UBound(m_List)
        m_NumCols = UBound(m_List, 2)
        
        
        Call Transpose
        
        'Performance Improvement: rather than copying the transposed list into the list to
        'transpose it back further down with 'CleanUpTable', simply swap the pointers to the base array info
        'To do so, we need three pointers to the descriptor of m_List, m_ListTransposed and a temp list we need
        'as a temporary buffer to store the pointer
                    
        'Get the pointers to the SafeArray
        ptrList = getSafeArrayDescrPtr(m_List)
        ptrListTransposed = getSafeArrayDescrPtr(m_ListTransposed)
        ptrListTemp = getSafeArrayDescrPtr(ListTemp)
        
        'Swap list Transposed and List
        Call CopyMemory(ByVal ptrListTemp, ByVal ptrList, 4)
        Call CopyMemory(ByVal ptrList, ByVal ptrListTransposed, 4)
        Call CopyMemory(ByVal ptrListTransposed, ByVal ptrListTemp, 4)
        
        'clean pointer of temp list to avoid crash on extiting the programm
        Call CopyMemory(ByVal ptrListTemp, 0&, 4&)

        
        'Performance Improvement: rather than run 'CleanUpTable' to get rid of the deleted 1st row and to transpose the table,
        'set the m_IsDirty flag manually (no need to transpose as earlier on the two arrays have been swapped anyway)
        'Call CleanUpTable
        m_IsDirty = False
        
        
        '*****
        'set ShrinkWhen and MaxGarbage properties
        '*****
        want_free = C_WANT_FREE_PERCENT * m_NumItems
        If want_free < C_MIN_FREE Then
            want_free = C_MIN_FREE
        End If
                
        ' We will shrink the array if SL_NumItems < SL_ShrinkWhen.
        m_ShrinkWhen = m_NumItems - want_free
        
        ' Collect garbage when there are more than
        ' want_free garbage items in the list.
        m_MaxGarbage = want_free

        
        'handle thread safety
        ReleaseLock
        
        
End Sub

#End If

Public Sub LoadDelimTextFile(filename As String, Optional FieldDelimiter As String = ",", Optional RecordDelimiter As String = vbNewLine, Optional TableHasHeaders As Boolean = True, Optional TextQualifier As String = """")
' ***********************************************
' Add a delimiter File to the empty table
' ***********************************************

    Dim iRows As Long
    Dim iCols As Long
    Dim iCol As Long
    Dim iColCount As Long
    Dim lFn As Integer
    Dim iCellIdx As Long
    Dim iFirstCell As Long
    
    Dim stFileContent As String
    Dim stCells() As String
    
        


    If m_NumItems > 0 Then
            'raise error -> only allowed with empty list
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + 513, Me.Name, "List is not empty. Range can only be added to empty list!"
     Else
        
        'handle thread safety
        CheckLockStatus
        LockForUpdate
        
        
        lFn = VBA.FreeFile
        
        '1. Open the file and load the data into a string
        Open filename For Input As #lFn
            stFileContent = Input(LOF(lFn), #lFn)
        Close #lFn
        
        '2. parse the records
        SplitCSV stFileContent, stCells, iCols, iRows, FieldDelimiter, RecordDelimiter, TextQualifier
        
        
        '3. Define the table
        
        'Size the list and set attributes
        If TableHasHeaders = True Then
            ReDim m_List(1 To iCols, 1 To iRows - 1)
            m_ArraySize = iRows - 1
            m_NumItems = iRows - 1
         Else
            ReDim m_List(1 To iCols, 1 To iRows)
            m_ArraySize = iRows
            m_NumItems = iRows
        End If
        m_NumCols = iCols

        'set the columns array
        iColCount = iCols
        ReDim m_Columns(1 To iColCount)
            
        If TableHasHeaders = True Then
            'Debug.Print "Row " & (iCellIdx \ lngCols), "Column " & (iCellIdx Mod lngCols), "Data: " & stCells(iCellIdx)
            For iCol = 1 To iColCount
                    m_Columns(iCol).Name = stCells(iCol - 1)
                    m_Columns(iCol).Number = iCol
                    m_Columns(iCol).IsSorted = False
                    m_Columns(iCol).IsIndexed = False
                    m_Columns(iCol).Type = dtString
            Next iCol
            iFirstCell = iCol - 1
        Else
            For iCol = 1 To iColCount
                    m_Columns(iCol).Name = "Field" & iCol
                    m_Columns(iCol).Number = iCol
                    m_Columns(iCol).IsSorted = False
                    m_Columns(iCol).IsIndexed = False
                    m_Columns(iCol).Type = dtString
            Next iCol
            iFirstCell = 0
        End If
        m_HasHeaders = True
        
        
        '5. loop through all records and add them to the table
        If TableHasHeaders = True Then
            For iCellIdx = iFirstCell To UBound(stCells)
                'Debug.Print "Column " & (iCellIdx Mod lngCols), "Row " & (iCellIdx \ lngCols), "Data: " & stCells(iCellIdx)
                m_List((iCellIdx Mod iCols) + 1, (iCellIdx \ iCols)) = stCells(iCellIdx)
            Next iCellIdx
         Else
            For iCellIdx = iFirstCell To UBound(stCells)
                'Debug.Print "Column " & (iCellIdx Mod lngCols), "Row " & (iCellIdx \ lngCols), "Data: " & stCells(iCellIdx)
                m_List((iCellIdx Mod iCols) + 1, (iCellIdx \ iCols) + 1) = stCells(iCellIdx)
            Next iCellIdx
        End If
        
        'handle thread safety
        ReleaseLock
        
        'mark as dirty
        m_IsDirty = True
        
    End If
        

End Sub

Public Sub AppendToTable(ByRef tSource As cDataTable)
' ***********************************************
' Append the data of a given cDataTable object to the list
' ***********************************************

    Dim lSourceColCount As Long
    Dim lTargetColCount As Long
    Dim lSturctureCheck As Long
    Dim lHeaderCheck As Long
    Dim lNewColCount As Long
    Dim arrFieldMapping() As Variant
    Dim i As Long
    Dim k As Long
    
    lSourceColCount = tSource.NumCols
    lTargetColCount = Me.NumCols
    
    ReDim arrFieldMapping(1 To lSourceColCount, 1 To 2)
    
    '§§§ UNDER CONSTRUCTION !!! §§§
    
    
    '1. Check if structure is the same
    If lSourceColCount = lTargetColCount Then
        'the tables have the same no of columns
        'let's check if they are in the same order
        If tSource.HasHeaders = True Then
            For i = 1 To lSourceColCount
                'create the source mapping in case needed
                arrFieldMapping(i, 1) = tSource.ItemName(i)
                If tSource.ItemName(i) = Me.ItemName(i) Then
                    'Source and Target Column are the same
                    lHeaderCheck = 1
                 Else
                    lHeaderCheck = 2
                End If
            Next i
            
            For i = 1 To lSourceColCount
                For k = 1 To lSourceColCount
                    If arrFieldMapping(i, 1) = Me.ItemName(i) Then
                        arrFieldMapping(i, 2) = Me.ItemName(i)
                        Exit For
                     Else
                        arrFieldMapping(i, 2) = "No Match"
                    End If
                        
               Next k
            
                'check if we found a match
                If arrFieldMapping(i, 2) = "No Match" Then
                    'raise error -> the column names do not match
                    ' For class errors, you add vbObjectError to the error number.
                    Err.Raise vbObjectError + m_Errors(7).errNumber, Me.Name, m_Errors(7).errDescrption

                End If
            Next i
        End If
                
        lSturctureCheck = 1
     ElseIf lSourceColCount > lTargetColCount Then
        'the source has more columns than the target
        lNewColCount = lSourceColCount
        lSturctureCheck = 2
     ElseIf lSourceColCount < lTargetColCount Then
        'the source has less columns, let's append what we have
        lSturctureCheck = 3
    End If
     
                 
    '2.


    '§§§ UNDER CONSTRUCTION !!! §§§


End Sub


Private Sub ResizeList(tyResize As ResizeType)
' ***********************************************
' Resize the list to have C_WANT_FREE_PERCENT unused entries. Set m_ShrinkWhen so we know to
' resize the array when the number of used entries is too small. Set this value to resize
' the array when there is more than twice the desired amount of memory free.
' ***********************************************

    Dim want_free As Long

    want_free = C_WANT_FREE_PERCENT * m_NumItems
    If want_free < C_MIN_FREE Then
        want_free = C_MIN_FREE
    End If
        
    If tyResize = IncludingWantFreeSpace Then
            m_ArraySize = m_NumItems + want_free
     Else
            m_ArraySize = m_NumItems
    End If
    
    'prevent subscript out of range issue if m_NumItems = 0 and subsequently m_ArraySize = 0
    'set the array size to minimum value 1
    If m_ArraySize = 0 Then m_ArraySize = 1
    
    
    ReDim Preserve m_List(1 To m_NumCols, 1 To m_ArraySize)
    ReDim Preserve m_IdxToList(1 To m_ArraySize)

        
    ' We will shrink the array if SL_NumItems < SL_ShrinkWhen.
    m_ShrinkWhen = m_NumItems - want_free
    
    ' Collect garbage when there are more than
    ' want_free garbage items in the list.
    m_MaxGarbage = want_free
    
    'Debug.Print "Resizing." & vbCrLf & vbCrLf & _
        "Items:" & Str$(m_ArraySize) & vbCrLf & _
        "Unused: " & Str$(want_free) & vbCrLf & _
        "ShrinkWhen:" & Str$(m_ShrinkWhen) & vbCrLf & _
        "MaxGarbage:" & Str$(m_MaxGarbage)
                
End Sub

Private Sub CollectGarbage()
' ************************************************
' Collect the garbage.
' ************************************************

    Dim i As Long
    Dim good As Long
    Dim iCol As Long
    Dim iBookmark As Long
    
    'Debug.Print "Collecting garbage..."
            
    'check if any garbage is in the table
    If m_FirstGarbage = 1 And m_LastGarbage = 1 Then
        'no garbage in the table...
        Exit Sub
    End If
            
    'remember the position of the bookmark
    iBookmark = m_Bookmark
    
    ' The first good item goes here
    ' performance improvement -> start at the 1st record which is deleted
    '                            rather then the 1st item in the list
    good = m_FirstGarbage
    
    For i = m_FirstGarbage To m_NumItems
            
        ' If 1st field in the record is not garbage,
        ' move the record to its new location and
        If Not m_List(1, i) = "#DELETED" Then
            'Debug.Print "moving record from "; i; " to record "; good
            For iCol = 1 To m_NumCols
                m_List(iCol, good) = m_List(iCol, i)
            Next iCol
            ' if the record is the bookmark, set bookmark to new location
            If i = iBookmark Then
                'Debug.Print "moving bookmark from "; i; " to record "; good
                iBookmark = good
            End If
            good = good + 1
         Else
            ' if the record is the bookmark, set bookmark to the last good location
            ' (i.e. good - 1 due to trail increment of good when shifting items)
            If i = iBookmark Then
                'Debug.Print "moving bookmark from "; i; " to record "; good - 1
                iBookmark = good - 1
            End If
        End If
        
    Next i
    
    ' This is where the last good item is (i.e. good - 1 due to trail increment of good when shifting items)
    m_NumItems = good - 1

    ' Save the new bookmark
    m_Bookmark = iBookmark

    're-initialize the garbage pointers
    m_FirstGarbage = 1
    m_LastGarbage = 1

    ' See if we should resize the list.
    If m_NumItems < m_ShrinkWhen Then ResizeList (IncludingWantFreeSpace)
        
End Sub


Private Sub CleanUpTable()
' ************************************************
' Clean up the data table to prepare for sorting etc:
' Collect the garbage + remove free tail records
' ************************************************
        'Collect the garbage
        CollectGarbage
        
        'resize List with option to remove the free tail records
        ResizeList (WithoutEmptyTrailRecords)

        'Transpose the list
        Transpose
    
        m_IsDirty = False

End Sub


'#Region
'#======================================================================================================================
'# Thread safety
'#======================================================================================================================

Private Sub LockForUpdate()
' ************************************************
' Sets the IsLocked state to true
' ************************************************

    m_IsLocked = True

End Sub

Private Sub ReleaseLock()
' ************************************************
' Sets the IsLocked state to true
' ************************************************

    m_IsLocked = False

End Sub


Private Sub CheckLockStatus()

    Dim iWait As Integer

    If m_IsLocked = False Then
        Exit Sub
     Else
        'wait a little
        For iWait = 1 To C_MAX_WAIT_CYCLES
            Sleep (C_WAIT_TIME_MS)
            If m_IsLocked = False Then
                Exit Sub
             End If
        Next iWait
        'raise error as the table is still locked for update
        'For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(9).errNumber, Me.Name, m_Errors(9).errDescrption
    End If
        
End Sub


'#Region
'#======================================================================================================================
'# Record Navigation
'#======================================================================================================================
Public Sub RsMoveFirst()
' ***********************************************
' Moves the bookmark to the 1st valid record
' ***********************************************

    Dim lPointer As Long
    
    m_EOF = False
    m_BOF = False
    
    For lPointer = LBound(m_List, 2) To m_NumItems
        If m_List(1, lPointer) <> "#DELETED" Then
            Exit For
        End If
    Next lPointer

    m_Bookmark = lPointer


End Sub

Public Sub RsMoveLast()
' ***********************************************
' Moves the bookmark to the last valid record
' ***********************************************

    Dim lPointer As Long
    
    m_EOF = False
    m_BOF = False
    
    
    For lPointer = m_NumItems To LBound(m_List, 2) Step -1
        If m_List(1, lPointer) <> "#DELETED" Then
            Exit For
        End If
    Next lPointer

    m_Bookmark = lPointer


End Sub

Public Sub RsMoveNext()
' ***********************************************
' Moves the bookmark to the next valid record
' ***********************************************

    Dim lPointer As Long
    Dim iIdx As Long
    
    iIdx = m_Bookmark
    
    If iIdx + 1 > m_NumItems Then
        'we are passing by the last record, set EOF flag
        m_EOF = True
     Else
    
        For lPointer = iIdx + 1 To m_NumItems
            If m_List(1, lPointer) <> "#DELETED" Then
                Exit For
            End If
        Next lPointer
    
        m_Bookmark = lPointer
        
    End If

End Sub

Public Sub RsMovePrevious()
' ***********************************************
' Moves the bookmark to the next valid record
' ***********************************************

    Dim lPointer As Long
    Dim iIdx As Long
    
    iIdx = m_Bookmark
    
    If iIdx - 1 < LBound(m_List, 2) Then
        'we are passing by the first record, set BOF flag
        m_BOF = True
     Else
       
        For lPointer = iIdx - 1 To LBound(m_List, 2) Step -1
            If m_List(1, lPointer) <> "#DELETED" Then
                Exit For
            End If
        Next lPointer
    
        m_Bookmark = lPointer
    End If
    
End Sub

Public Sub RsFindFirst(Index As Variant, match As MatchType, Criteria As Variant)
' ***********************************************
' Moves the bookmark to the first valid record matching the criteria
' ***********************************************

    'Reset the bookmark to -1 to trigger search
    'from the beginning of the table
    m_Bookmark = -1
    
    'set End of File lag to false as we start search from scratch
    m_EOF = False
    
    'Find the next matching record
    RsFind Index, match, Criteria

End Sub

Public Sub RsFindNext(Index As Variant, match As MatchType, Criteria As Variant)
' ***********************************************
' Moves the bookmark to the next valid record matching the criteria
' ***********************************************

    'Find the next matching record
    RsFind Index, match, Criteria
    

End Sub


Private Sub RsFind(Index As Variant, match As MatchType, Criteria As Variant)
' ***********************************************
' Moves the bookmark to the next valid record matching the criteria
' ***********************************************

    Dim lPointer As Long
    Dim iIdx As Long
    
    'set the pointer to the beginning of the table or at the next record
    If m_Bookmark = -1 Then
        lPointer = 1
     Else
        If m_Bookmark + 1 > m_NumItems Then
            'last record reached, set EOF flag
            m_EOF = True
            Exit Sub
         Else
            RsMoveNext
            lPointer = m_Bookmark
        End If
    End If
    
    'Find the next record matching the criteria and set cursor to this record
    If Search(Criteria, match, CStr(Index), lPointer) = False Then
        'item not found
        m_EOF = True
    End If
    
    '§§§
    '§§§ CONSTRUCTION SITE!! NEED TO FIGURE OUT HOW TO RESOLVE THE CRITERIA INTO A EXECUTABLE SEARCH!! §§§
    '§§§ TEMPORARY SOLUTION IS TO USE LINEAR SEARCH ON ONE COLUMN ONLY... §§§
    '§§§
    
'    For iIdx = lPointer To m_NumItems
'    Next iIdx

    '§§§
    '§§§ END OF CONSTRUCTION SITE
    '§§§


End Sub



'#Region
'#======================================================================================================================
'# Data Selection
'#======================================================================================================================


Public Function SelectData(FieldList As String, WhereClause As String, Optional GroupBy As String = "", Optional Having As String = "", Optional SortBy As String = "") As cDataTable
' ***********************************************
' Return a new data table object containing recrods
' with the specified Field List and matching the WhereClause
' ***********************************************

    Dim dtOut As cDataTable
    
    Set dtOut = New cDataTable

    '§§§ Future feature §§§
    
    '0. Parse all parameters into a parse tree
    '1. Define Output Datetable Structure based on Field List
    '2. Evaluate WhereClause
        '2a. Check if sorted / indexed
        '2b. Find relevant records
    '5. Add each relevent record to new data table
    '6. create new data table and Apply GroupBy clause
    '7. evaluate HavingClause -> remove records for which having clause is not true
    '8. If distinct -> remove duplicates from output
    '9. Order the table by the OrderBy clause
    
    'Assign the selection result to the function output
    SelectData = dtOut

End Function


'#Region
'#======================================================================================================================
'# Data Output
'#======================================================================================================================
#If C_CON_OFFICE_APP = "XLS" Then

Public Sub DumpToRange(TargetWorksheet As Worksheet, TargetCell As Range, Optional IncludeHeader As Boolean = True, Optional CompressOnRowOverflow As Boolean = False)
' ***********************************************
' Output the array to a range on a given worksheet
' This works with MS Excel only...
' ***********************************************
On Error GoTo err_DumpToRange


    Dim iRow As Integer: iRow = TargetCell.Row
    Dim iCol As Integer: iCol = TargetCell.Column
    Dim iColIdx As Integer
    
    If m_IsDirty = True Then
        Call CleanUpTable
        'Call Transpose -> not needed as included in CleanUpTable
    End If
    
    'check if number of records exeeds excel row limit (1'048'576) we take
    'a one less, to take care of a possible header row and substract
    'the starting row, as we might start somewhere else than in row one
    If m_NumItems > (1048575 - iRow) Then
        If CompressOnRowOverflow = False Then
            'raise error as we cannot write the output to the target
            'For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + m_Errors(8).errNumber, Me.Name, m_Errors(8).errDescrption
         Else
            'write the output in a compressed manner
            WriteToCells TargetWorksheet, TargetCell, IncludeHeader, CompressOnRowOverflow
        End If
    End If
        
    
    If IncludeHeader = True Then
        For iColIdx = 1 To m_NumCols
            TargetWorksheet.Cells(iRow, iCol + (iColIdx - 1)).Value = m_Columns(iColIdx).Name
        Next iColIdx
        
        'check that table has some records to dump
        If m_NumItems = 0 Then
            'write message that table is empty in top left cell of output area
            TargetWorksheet.Cells(iRow + 1, iCol).Value = "Empty Data Table"
         Else
            TargetWorksheet.Cells(iRow + 1, iCol).Resize(UBound(m_ListTransposed), UBound(m_ListTransposed, 2)) = m_ListTransposed
        End If
     Else
        'check that table has some records to dump
        If m_NumItems = 0 Then
            'write message that table is empty in top left cell of output area
            TargetWorksheet.Cells(iRow, iCol).Value = "Empty Data Table"
         Else
            TargetWorksheet.Cells(iRow, iCol).Resize(UBound(m_ListTransposed), UBound(m_ListTransposed, 2)) = m_ListTransposed
        End If
    End If

exit_DumpToRange:
    Exit Sub
    
err_DumpToRange:
    If Err.Number = 1004 Then
        'let's try to resolve the issue by writing single values to single cells
        WriteToCells TargetWorksheet, TargetCell, IncludeHeader, CompressOnRowOverflow
        Resume exit_DumpToRange
    End If


End Sub

Private Sub WriteToCells(TargetWorksheet As Worksheet, TargetCell As Range, Optional IncludeHeader As Boolean = True, Optional CompressOnRowOverflow As Boolean = False)
' ***********************************************
' Output every single element of the array to a
' corresponding single cell in the range on a given
' worksheet. This works with MS Excel only...
' ***********************************************

    Dim iRow As Long: iRow = TargetCell.Row
    Dim iCol As Long: iCol = TargetCell.Column
    Dim iRowOut As Long: iRowOut = 0
    Dim iRowIdx As Long
    Dim iColIdx As Long
    Dim iCRType As Long
    Dim stContent As String
    Dim stContentOut As String
    Dim aContent() As String
    Dim iPos As Long
    Dim iCnt As Long
    
    Dim lCellSizeMax As Long
    Dim aCellSize() As Long
    Dim aCellContent() As String
    Dim lCellMax As Long: lCellMax = 32000
    Dim stRecDelim As String: stRecDelim = "|" & vbCr & "|"
    Dim bDumpRow As Boolean: bDumpRow = False
    
    ReDim aCellSize(1 To m_NumCols)
    ReDim aCellContent(1 To m_NumCols)
        
    'should not be the case but you never know
    If m_IsDirty = True Then
        Call CleanUpTable
    End If
    
    
    'check if number of records exeeds excel row limit (1'048'576) we take
    'a one less, to take care of a possible header row and substract
    'the starting row, as we might start somewhere else than in row one
    If m_NumItems > (1048575 - iRow) Then
        '*****
        'Version with Overflow
        '*****
            
        'compressed version or raise error
        If CompressOnRowOverflow = False Then
            'raise error as we cannot write the output to the target
            'For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + m_Errors(8).errNumber, Me.Name, m_Errors(8).errDescrption
         Else
            
            '§§§ needs major rework to include compression §§§
            
            'initialize the cell size array
            For iCnt = 1 To m_NumCols
                aCellSize(iCnt) = 0
            Next iCnt
        
            
            If IncludeHeader = True Then
                For iColIdx = 1 To m_NumCols
                    TargetWorksheet.Cells(iRow, iCol + (iColIdx - 1)).Value = m_Columns(iColIdx).Name
                Next iColIdx
                
                iRowOut = iRowOut + 1
                
            End If
                
                
            'construct the output
            For iRowIdx = 1 To m_NumItems
                
                'check if we have to dump the row or continue to append (i.e. loop through the current row and see if we
                'have enough space left to append this row or if we have to dump the row before we continue to append data
                For iColIdx = 1 To m_NumCols
                        aCellSize(iColIdx) = aCellSize(iColIdx) & Len(m_ListTransposed(iRowIdx, iColIdx) & stRecDelim)
                        If aCellSize(iColIdx) > lCellSizeMax Then
                            lCellSizeMax = aCellSize(iColIdx)
                        End If
                Next iColIdx
                
                If lCellSizeMax > lCellMax Then
                    'dumpt the record and reset the cell stats
                    For iColIdx = 1 To m_NumCols
                        TargetWorksheet.Cells(iRow + iRowOut, iCol + (iColIdx - 1)).FormulaR1C1 = "'" & aCellContent(iColIdx)
                        aCellSize(iCnt) = 0
                        aCellContent(iCnt) = ""
                    Next iColIdx
                End If
                
                'append the current row to the output (i.e. construct the output itself)
                For iColIdx = 1 To m_NumCols
                        aCellContent(iColIdx) = aCellContent(iColIdx) & m_ListTransposed(iRowIdx, iColIdx) & stRecDelim
                Next iColIdx
                
                
            Next iRowIdx
                                
            
        End If
     
     Else
     
        '*****
        'Version without Overflow
        '*****
            
        '§§§ not sure this works as version with header is the same as version without header... §§§
        
        'standard version
        If IncludeHeader = True Then
            '*****
            'Output data with header
            '*****
            
            For iColIdx = 1 To m_NumCols
                TargetWorksheet.Cells(iRow, iCol + (iColIdx - 1)).Value = m_Columns(iColIdx).Name
            Next iColIdx
            
            'dump the content cell by cell, row by row
            For iRowIdx = 1 To m_NumItems
                For iColIdx = 1 To m_NumCols
                    stContent = m_ListTransposed(iRowIdx, iColIdx)
                    If InStr(stContent, vbLf) Or InStr(stContent, vbCr) Then
                        'add a single quote at the beginning of the content to prevent err 1004 to occur
                        TargetWorksheet.Cells(iRow + (iRowIdx - 1), iCol + (iColIdx - 1)).FormulaR1C1 = "'" & m_ListTransposed(iRowIdx, iColIdx)
                     Else
                        TargetWorksheet.Cells(iRow + (iRowIdx - 1), iCol + (iColIdx - 1)).FormulaR1C1 = m_ListTransposed(iRowIdx, iColIdx)
                    End If
                Next iColIdx
            Next iRowIdx
            
         Else
            '*****
            'Output data without header
            '*****
            
            'dump the content cell by cell, row by row
            For iRowIdx = 1 To m_NumItems
                For iColIdx = 1 To m_NumCols
                    stContent = m_ListTransposed(iRowIdx, iColIdx)
                    If InStr(stContent, vbLf) Or InStr(stContent, vbCr) Then
                        'add a single quote at the beginning of the content to prevent err 1004 to occur
                        TargetWorksheet.Cells(iRow + (iRowIdx - 1), iCol + (iColIdx - 1)).FormulaR1C1 = "'" & m_ListTransposed(iRowIdx, iColIdx)
                     Else
                        TargetWorksheet.Cells(iRow + (iRowIdx - 1), iCol + (iColIdx - 1)).FormulaR1C1 = m_ListTransposed(iRowIdx, iColIdx)
                    End If
                Next iColIdx
           Next iRowIdx
                
        End If
        
        '§§§ not sure this works as version with header is the same as version without header... §§§
        
    End If
    
    

End Sub

#End If


Public Sub DumpToFile(TargetFile As String, Delimiter As String, Optional IncludeHeader As Boolean = True, Optional OutputMode As OutputType = OverwriteIfExists)
' ***********************************************
' Output the array to a text file applying any
' given delimiter
' ***********************************************
        
    Dim fn As Integer
    Dim stRecored As String
    Dim i As Long
    Dim ii As Long
    Dim stRecord As String
    Dim stHeader As Variant
    
    If m_IsDirty = True Then
        Call CleanUpTable
    End If
        
    fn = VBA.FreeFile
        
    If OutputMode = OverwriteIfExists Then
        Open TargetFile For Output As #fn
     ElseIf OutputMode = AppendIfExists Then
        Open TargetFile For Append As #fn
     ElseIf OutputMode = ErrorIfExists Then
        'check if file exists
        If Dir(TargetFile) <> "" Then
            Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
        End If
     Else
        Open TargetFile For Output As #fn
    End If
        
        'Output the header if required
        stRecord = ""
        For Each stHeader In Me.Headers
            If LenB(stRecord) = 0 Then
                    stRecord = stHeader
             Else
                    stRecord = stRecord & Delimiter & stHeader
            End If
        Next
        Print #fn, stRecord
        
        'output the records
        stRecord = ""
        For i = 1 To UBound(m_ListTransposed)
                For ii = 1 To UBound(m_ListTransposed, 2)
                        'construct the record ...
                        If LenB(stRecord) = 0 Then
                                stRecord = m_ListTransposed(i, ii)
                         Else
                                stRecord = stRecord & Delimiter & m_ListTransposed(i, ii)
                        End If
                Next ii
          Print #fn, stRecord
          stRecord = ""
        Next i

        Close #fn
        
End Sub

'#Region
'#======================================================================================================================
'# Sorting
'#======================================================================================================================

Public Sub Sort(ParamArray SortOrder() As Variant)
' ************************************************
' Sorts the list with the hybrid mergesort
' The ParamArray SortOrder is expected to be the number
' of the column to sort and a sort direction (as "asc" / "desc").
' If the sort direction is omitted, ascending sort direction
' is choosen as the default.
' ************************************************

    Dim i As Long
    Dim iColToSort As Long
    Dim iCol As Long
    Dim aScratch() As Variant
    'Dim List() As Variant
        
    
    'List = m_List
    
    'Make sure we have a clean table if it's dirty
    If m_IsDirty = True Then
        'Collect the garbage
        CollectGarbage
        'resize List with option to remove the free tail records
        ResizeList (WithoutEmptyTrailRecords)
    End If
        
    For i = LBound(SortOrder) To UBound(SortOrder)
        
        
        'Only process numeric pointers to the columns to sort
        If IsTableColumn(SortOrder(i)) Then
            iColToSort = GetColumnIndex(SortOrder(i))
            'check if we are not in the last parameter, then check the sort direction
            If i + 1 <= UBound(SortOrder) Then
                If Not IsTableColumn(SortOrder(i + 1)) Then
                    If UCase(SortOrder(i + 1)) = "ASC" Then
                        ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                        Call MergeSort(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))              '§§§ changed to 2nd dimension
                        'm_SortedColumn = iColToSort
                        SetSortedColumn iColToSort
                     ElseIf UCase(SortOrder(i + 1)) = "DESC" Then
                        ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                        Call MergeSortDesc(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))          '§§§ changed to 2nd dimension
                        'm_SortedColumn = iColToSort
                        SetSortedColumn iColToSort
                     Else
                        MsgBox "Invalid sort direction supplied for column " & GetColumnName(iColToSort) & "!", vbCritical
                        Exit Sub
                    End If
                 Else
                    ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                    Call MergeSort(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))                  '§§§ changed to 2nd dimension
                    'm_SortedColumn = iColToSort
                    SetSortedColumn iColToSort
               End If
             Else
                'we are looking at the last parameter - only process if it is a numeric pointer to the column to sort
                If IsTableColumn(SortOrder(i)) Then
                    ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                    Call MergeSort(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))                  '§§§ changed to 2nd dimension
                    'm_SortedColumn = iColToSort
                    SetSortedColumn iColToSort
                End If
            End If
        End If
        
    Next i

    'Mark table as dirty, as we have changed the order of the items
    m_IsDirty = True

End Sub

Private Sub MergeSort(List() As Variant, ColToSort As Long, Scratch() As Variant, ByVal min As Long, ByVal max As Long)
' ************************************************
' Hybrid Mergesort with:
'   - Stops recursion when there are fewer than
'       C_CUT_OFF items in the list. Finishes with
'       InsertionSort.
' ************************************************
    
    Dim middle As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long

    Dim iCols As Long
    
    'Determine the number of Columns of the list
    Dim iColLower As Long: iColLower = LBound(List)          '§§§ changed to 1st dimension
    Dim iColUpper As Long: iColUpper = UBound(List)          '§§§ changed to 1st dimension
    

    ' If the list has no more than C_CUT_OFF elements,
    ' finish it off with InsertionSort.
    If max - min < C_CUT_OFF Then
        InsertionSort List(), ColToSort, min, max
        Exit Sub
    End If
    
    ' Recursively sort the sublists.
    middle = max \ 2 + min \ 2
    MergeSort List(), ColToSort, Scratch(), min, middle
    MergeSort List(), ColToSort, Scratch(), middle + 1, max

    ' Merge the sorted lists.
    i1 = min        ' Index in list 1
    i2 = middle + 1 ' Index in list 2
    i3 = min        ' Index in merged list
    Do While i1 <= middle And i2 <= max
        If List(ColToSort, i1) <= List(ColToSort, i2) Then
            'For iCols = iColLower To iColUpper
            '    Scratch(iCols, i3) = List(iCols, i1)
            'Next iCols
            DataRowCopy List, i1, Scratch, i3
            i1 = i1 + 1
        Else
            'For iCols = iColLower To iColUpper
            '    Scratch(iCols, i3) = List(iCols, i2)
            'Next iCols
            DataRowCopy List, i2, Scratch, i3
            i2 = i2 + 1
        End If
        i3 = i3 + 1
    Loop

    ' Empty out whichever list is not already empty.
    Do While i1 <= middle
        'For iCols = iColLower To iColUpper
        '    Scratch(iCols, i3) = List(iCols, i1)
        'Next iCols
        DataRowCopy List, i1, Scratch, i3
        i1 = i1 + 1
       i3 = i3 + 1
    Loop
    Do While i2 <= max
        'For iCols = iColLower To iColUpper
        '    Scratch(iCols, i3) = List(iCols, i2)
        'Next iCols
        DataRowCopy List, i2, Scratch, i3
        i2 = i2 + 1
        i3 = i3 + 1
    Loop

    ' Move the merged list back into list.
    For i3 = min To max
        'For iCols = iColLower To iColUpper
        '    List(iCols, i3) = Scratch(iCols, i3)
        'Next iCols
        DataRowCopy Scratch, i3, List, i3
    Next i3
End Sub

Private Sub MergeSortDesc(List() As Variant, ColToSort As Long, Scratch() As Variant, ByVal min As Long, ByVal max As Long)
' ************************************************
' Mergesort with:
'   - Stops recursion when there are fewer than
'       C_CUT_OFF items in the list. Finishes with
'       InsertionSort.
' ************************************************
    
    Dim middle As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long

    Dim iCols As Long
    
    'Determine the number of Columns of the list
    Dim iColLower As Long: iColLower = LBound(List)
    Dim iColUpper As Long: iColUpper = UBound(List)
    

    ' If the list has no more than C_CUT_OFF elements,
    ' finish it off with InsertionSort.
    If max - min < C_CUT_OFF Then
        InsertionSortDesc List(), ColToSort, min, max
        Exit Sub
    End If
    
    ' Recursively sort the sublists.
    middle = max \ 2 + min \ 2
    MergeSortDesc List(), ColToSort, Scratch(), min, middle
    MergeSortDesc List(), ColToSort, Scratch(), middle + 1, max

    ' Merge the sorted lists.
    i1 = min        ' Index in list 1
    i2 = middle + 1 ' Index in list 2
    i3 = min        ' Index in merged list
    Do While i1 <= middle And i2 <= max
        If List(ColToSort, i1) >= List(ColToSort, i2) Then
            'For iCols = iColLower To iColUpper
            '    Scratch(iCols, i3) = List(iCols, i1)
            'Next iCols
            DataRowCopy List, i1, Scratch, i3
            i1 = i1 + 1
        Else
            'For iCols = iColLower To iColUpper
            '    Scratch(iCols, i3) = List(iCols, i2)
            'Next iCols
            DataRowCopy List, i2, Scratch, i3
            i2 = i2 + 1
        End If
        i3 = i3 + 1
    Loop

    ' Empty out whichever list is not already empty.
    Do While i1 <= middle
        'For iCols = iColLower To iColUpper
        '    Scratch(iCols, i3) = List(iCols, i1)
        'Next iCols
        DataRowCopy List, i1, Scratch, i3
        i1 = i1 + 1
       i3 = i3 + 1
    Loop
    Do While i2 <= max
        'For iCols = iColLower To iColUpper
        '    Scratch(iCols, i3) = List(iCols, i2)
        'Next iCols
        DataRowCopy List, i2, Scratch, i3
        i2 = i2 + 1
        i3 = i3 + 1
    Loop

    ' Move the merged list back into list.
    For i3 = min To max
        'For iCols = iColLower To iColUpper
        '    List(iCols, i3) = Scratch(iCols, i3)
        'Next iCols
        DataRowCopy Scratch, i3, List, i3
    Next i3
End Sub


Private Sub InsertionSort(List() As Variant, ColToSort As Long, min As Long, max As Long)
' ************************************************
' Standard insertion sort algorythm
' ************************************************
    
    Dim iInnerLoop As Long
    Dim iOuterLoop As Long
    Dim next_item As Variant
    Dim next_item_rec() As Variant
    
    Dim iCols As Long
    Dim iColLower As Long: iColLower = LBound(List)
    Dim iColUpper As Long: iColUpper = UBound(List)

    ReDim next_item_rec(iColLower To iColUpper, 1 To 1)
    
    For iOuterLoop = min + 1 To max
        
        'this is the next item / record
        next_item = List(ColToSort, iOuterLoop)
        'For iCols = iColLower To iColUpper
        '    next_item_rec(iCols) = List(iCols, iOuterLoop)
        'Next iCols
        DataRowCopy List, iOuterLoop, next_item_rec, 1
    
        'reorder the items
        For iInnerLoop = iOuterLoop To min + 1 Step -1
            If next_item < List(ColToSort, iInnerLoop - 1) Then
                'For iCols = iColLower To iColUpper
                '    List(iCols, iInnerLoop) = List(iCols, iInnerLoop - 1)
                'Next iCols
                DataRowMove List, iInnerLoop - 1, iInnerLoop
             Else
                Exit For
            End If
        Next iInnerLoop
        
        'insert the new item
        'For iCols = iColLower To iColUpper
        '    List(iCols, iInnerLoop) = next_item_rec(iCols)
        'Next iCols
        DataRowCopy next_item_rec, 1, List, iInnerLoop
        
    Next iOuterLoop
    
End Sub


Private Sub InsertionSortDesc(List() As Variant, ColToSort As Long, min As Long, max As Long)
' ************************************************
' Standard insertion sort algorythm
' ************************************************
    
    Dim iInnerLoop As Long
    Dim iOuterLoop As Long
    Dim next_item As Variant
    Dim next_item_rec() As Variant
    
    Dim iCols As Long
    Dim iColLower As Long: iColLower = LBound(List)
    Dim iColUpper As Long: iColUpper = UBound(List)

    ReDim next_item_rec(iColLower To iColUpper, 1 To 1)
    
    For iOuterLoop = min + 1 To max
        
        'this is the next item / record
        next_item = List(ColToSort, iOuterLoop)
        'For iCols = iColLower To iColUpper
        '    next_item_rec(iCols) = List(iCols, iOuterLoop)
        'Next iCols
        DataRowCopy List, iOuterLoop, next_item_rec, 1

        'reorder the items
        For iInnerLoop = iOuterLoop To min + 1 Step -1
            If next_item > List(ColToSort, iInnerLoop - 1) Then
                'For iCols = iColLower To iColUpper
                '    List(iCols, iInnerLoop) = List(iCols, iInnerLoop - 1)
                'Next iCols
                DataRowMove List, iInnerLoop - 1, iInnerLoop
             Else
                Exit For
            End If
        Next iInnerLoop
        
        'insert the new item
        'For iCols = iColLower To iColUpper
        '    List(iCols, iInnerLoop) = next_item_rec(iCols)
        'Next iCols
        DataRowCopy next_item_rec, 1, List, iInnerLoop
        
    Next iOuterLoop
    
End Sub

Public Sub SortUnstable(Index As Variant, Optional Direction As String = "asc")
' ***********************************************
' Pulic call of QuickSort-Algorithm
' ***********************************************

    Dim iColToSort As Long

    'Make sure we have a clean table if it's dirty
    If m_IsDirty = True Then
        'Collect the garbage
        CollectGarbage
        'resize List with option to remove the free tail records
        ResizeList (WithoutEmptyTrailRecords)
    End If
    
    If IsTableColumn(Index) Then
        iColToSort = GetColumnIndex(Index)
        If UCase(Direction) = "ASC" Then
            Call QuickSort(m_List(), iColToSort, LBound(m_List, 2), UBound(m_List, 2))
         ElseIf UCase(Direction) = "DESC" Then
            Call QuickSortDesc(m_List(), iColToSort, LBound(m_List, 2), UBound(m_List, 2))
         Else
            MsgBox "Sort direction '" & Direction & "' not known."
            Exit Sub
        End If
        'Mark table as dirty, as we have changed the order of the items
        m_IsDirty = True
     Else
        MsgBox "Sort Column '" & Index & "' does not exist in table '" & Me.Name & "'"
        Exit Sub
    End If
    

End Sub

Private Sub QuickSort(List() As Variant, ColToSort As Long, Optional ByVal min As Variant, Optional ByVal max As Variant)
' ***********************************************
' QuickSort-Algorithm
' ***********************************************

  ' Sort the entire array if the start / end boundaries are not defined
  If IsMissing(min) Then min = LBound(List, 2)
  If IsMissing(max) Then max = UBound(List, 2)

  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim H() As Variant
  Dim x As Variant
  
  Dim lCStart As Long
  Dim lCEnd As Long
  
  lCStart = LBound(List, 1)
  lCEnd = UBound(List, 1)

  i = min: j = max
  x = List(ColToSort, (min + max) / 2)

  ReDim H(1 To lCEnd, 1 To 1)

  ' Split the array
  Do

    While (List(ColToSort, i) < x): i = i + 1: Wend
    While (List(ColToSort, j) > x): j = j - 1: Wend

    If (i <= j) Then
      ' Flip value pairs
      'For k = lCStart To lCEnd
      '  H(k) = List(k, i)
      'Next k
      DataRowCopy List, i, H, 1
      'For k = lCStart To lCEnd
      '  List(k, i) = List(k, j)
      'Next k
      DataRowMove List, j, i
      'For k = lCStart To lCEnd
      '  List(k, j) = H(k)
      'Next k
      DataRowCopy H, 1, List, j
      i = i + 1: j = j - 1
    End If
  Loop Until (i > j)

  ' Recursion
  If (min < j) Then QuickSort List, ColToSort, min, j
  If (i < max) Then QuickSort List, ColToSort, i, max
  
End Sub

Private Sub QuickSortDesc(List() As Variant, ColToSort As Long, Optional ByVal min As Variant, Optional ByVal max As Variant)
' ***********************************************
' QuickSort-Algorithm
' ***********************************************

  ' Sort the entire array if the start / end boundaries are not defined
  If IsMissing(min) Then min = LBound(List, 2)
  If IsMissing(max) Then max = UBound(List, 2)

  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim H() As Variant
  Dim x As Variant
  
  Dim lCStart As Long
  Dim lCEnd As Long
  
  lCStart = LBound(List, 1)
  lCEnd = UBound(List, 1)

  i = min: j = max
  x = List(ColToSort, (min + max) / 2)

  ReDim H(1 To lCEnd, 1 To 1)

  ' Split the array
  Do

    While (List(ColToSort, i) > x): i = i + 1: Wend
    While (List(ColToSort, j) < x): j = j - 1: Wend

    If (i >= j) Then
      ' Flip value pairs
      'For k = lCStart To lCEnd
      '  H(k) = List(k, i)
      'Next k
      DataRowCopy List, i, H, 1
      'For k = lCStart To lCEnd
      '  List(k, i) = List(k, j)
      'Next k
      DataRowMove List, j, i
      'For k = lCStart To lCEnd
      '  List(k, j) = H(k)
      'Next k
      DataRowCopy H, 1, List, j
      i = i + 1: j = j - 1
    End If
  Loop Until (i < j)

  ' Recursion
  If (min > j) Then QuickSort List, ColToSort, min, j
  If (i > max) Then QuickSort List, ColToSort, i, max
  
End Sub


'#Region
'#======================================================================================================================
'# Searching
'#======================================================================================================================

Private Function Search(SearchFor As Variant, match As MatchType, Index As String, Optional Start As Long = 1) As Boolean
' ************************************************
' Function to search the list, currently supporting:
' full table search, binary search on sorted column, indexed search on uniquely indexed columns
' §§§ Future usage considering non unique indices and data types §§§
' ************************************************

    Dim lPointer As Long
    Dim iCol As Long
    Dim bFirstItem As Boolean: bFirstItem = False
    Dim lPointer2 As Long
    
    iCol = GetColumnIndex(Index)
    
    'Check if column is indexed
    If m_Columns(iCol).IsIndexed = True Then
        'Column is index, go for the index
        lPointer = IndexedSearch(SearchFor, match, Index, Start)
     Else
        'Check if column is sorted
        If m_Columns(iCol).IsSorted = True Then
            'Column is sorted, now check if we already did such a search or not
            
            'check if this item has already been searched for this column
            If (m_SearchLastColumn = iCol And m_SearchLastItem = SearchFor) Then
                'we have already searched this value in this column, just check for the next item
                Me.RsMoveNext
                If Compare(m_List(iCol, Me.RsBookmark), SearchFor, match) = True Then
                    'this is a valid match (no need to set the pointer
                    'as the bookmark is already set with with Me.RsMoveNext)
                    Search = True
                    Exit Function
                 Else
                    Search = False
                    Exit Function
                End If
             Else
                'this is the 1st search of this value in this column, use binary search to find a matching item
                lPointer = BinarySearch(SearchFor, Index)
                                
                If lPointer = 0 Then
                    'No item found, let's exit
                    Search = False
                    Exit Function
                 Else
                    'Check if this is the 1st element matching the criteria
                    lPointer2 = lPointer
                    Do While bFirstItem = False
                        lPointer2 = lPointer2 - 1
                        If Compare(m_List(iCol, lPointer2), SearchFor, match) = False Then
                            'this no longer a matching item, the last item was the 1st element in the list matching the criteria
                            lPointer = lPointer2 + 1
                            bFirstItem = True
                        End If
                    Loop
                    'store the element for later comparison
                    m_SearchLastItem = SearchFor
                    m_SearchLastColumn = iCol
                End If
            End If
         Else
            'Column is neither indexed nor sorted - use linear search
            lPointer = LinearSearch(SearchFor, match, Index, Start)
        End If
    End If

    If lPointer = 0 Then
        Search = False
     Else
        Search = True
        m_Bookmark = lPointer
    End If


End Function

Private Function IndexedSearch(SearchFor As Variant, match As MatchType, Index As String, Optional Start As Long = 1) As Long
' ************************************************
' Locate the item using the index on the column
' This function returns the pointer to the recrod, not the record itself.
' If item is not found, a 0-Pointer is returned.
' ************************************************

    Dim iCol As Long
    Dim lPointer As Long

    iCol = GetColumnIndex(Index)
    
    
    If m_Columns(iCol).Index.IndexType = UniqueIndex Then
        'we have a unique index - let's check if the item exists
        'if the item exists, return it's position in the table
        '§§§ to be implemented §§§
        lPointer = m_Columns(iCol).Index.GetNextItem(SearchFor, match, Start)
            
     ElseIf m_Columns(iCol).Index.IndexType = NonUniqueIndex Then
        'we have a non-unique index - let's traverse the index
        '§§§ to be implemented - for the time being let's use linear search §§§
        lPointer = LinearSearch(SearchFor, match, Index, Start)
        
     Else
        'Index Type is not defined properly - use linear search
        lPointer = LinearSearch(SearchFor, match, Index, Start)
        
    End If

    IndexedSearch = lPointer

End Function


Private Function LinearSearch(SearchFor As Variant, match As MatchType, Index As String, Optional Start As Long = 1) As Long
' ************************************************
' Locate the item using exhaustive search.
' This function returns the pointer to the recrod, not the record itself.
' If item is not found, a 0-Pointer is returned.
' ************************************************

    Dim Target As Variant: Target = SearchFor
    Dim i As Long
    Dim iCol As Long
    Dim NumSearches As Long
    
    iCol = GetColumnIndex(Index)
    
    For i = Start To m_NumItems
        'New code using compare function
        If Compare(m_List(iCol, i), Target, match) = True Then
            Exit For
        End If
    Next i
    
    If i > m_NumItems Then
        LinearSearch = 0    ' Not found.
    ElseIf Compare(m_List(iCol, i), Target, match) = False Then
         LinearSearch = 0    ' Not found.
    Else
        LinearSearch = i    ' Found.
    End If
    
End Function

Function BinarySearch(SearchFor As Variant, Index As String) As Long
'Optional LastEl As Variant
' ************************************************
' Locate the item using generic binary search.
' This function returns the pointer to the recrod, not the record itself.
' If item is not found, a 0-Pointer is returned.
' The list *must* be sorted on the column to search, in ascending or descending
' order (the routines finds out the sort direction).
' String are compared in case-sensitive mode.
' You can write faster procedures if you modify the procedure
' to account for a specific data type, eg.
'   Function BinarySearchL (arr() As Long, search As Long,
'  Optional lastEl As Variant) As Long
' Based on a routine published by Francesco Balena at http://www.devx.com/vb2themax/Tip/18913
' ************************************************
    
    Dim first As Long
    Dim last As Long
    Dim middle As Long
    Dim inverseOrder As Boolean
    Dim iCol As Long
    
    
    iCol = GetColumnIndex(Index)
    
    ' account for optional arguments
    'If IsMissing(lastEl) Then LastEl = m_NumItems
    
    first = LBound(m_List(2))
    'last = LastEl
    last = m_NumItems

    ' deduct direction of sorting
    inverseOrder = (m_List(iCol, first) > m_List(iCol, last))

    ' assume searches failed
    BinarySearch = 0
    
    Do
        middle = (first + last) \ 2
        If m_List(iCol, middle) = SearchFor Then
            BinarySearch = middle
            Exit Do
        ElseIf ((m_List(iCol, middle) < SearchFor) Xor inverseOrder) Then
            first = middle + 1
        Else
            last = middle - 1
        End If
    Loop Until first > last
    
End Function


'#Region
'#======================================================================================================================
'# Array handling
'#======================================================================================================================

Private Sub Transpose()
' ***********************************************
' Transpose the array using a SafeArray hack
' ***********************************************

    Dim i As Long
    Dim k As Long
    Dim m As Long
    Dim tmpArr() As Variant
    Dim acurOld() As Currency
    Dim acurNew() As Currency
    Dim lngCount As Long
    Dim lngSafeArrayList As LongPtr
    Dim lngSafeArray As LongPtr
    Dim lngData As LongPtr
    Dim lngCountX As Long
    Dim lngCountY As Long
    Dim lngLBoundX As Long
    Dim lngLBoundY As Long
    Dim SafeArray2 As SAFEARRAY2D
    
    'Dim mTimer As New cStopWatch
    
    'mTimer.Reset
    
    '§§§ Performance Improvment: This might be faster by using CopyMemory at a later stage §§§
    'copy the data from m_List to m_ListTransposed
    m_ListTransposed = m_List
    'Debug.Print "HP Copying m_List to m_ListTransposed took: " & mTimer.MillisecondsElapsed & " milliseconds..."
    
    
    'Get the pointer to the SafeArray Base of m_ListTransposed
    lngSafeArray = getSafeArrayStructPtr(m_ListTransposed)
    
    'Copy the SafeArray base in to it's structure
    CopyMemory SafeArray2, ByVal lngSafeArray, Len(SafeArray2)

    With SafeArray2
        'Get the pointer to the actual data
        lngData = .pvData
        'Get the LBound and the no of elements for each dimension
        lngCountX = .Bounds(0).cElements
        lngLBoundX = .Bounds(0).lLbound
        lngCountY = .Bounds(1).cElements
        lngLBoundY = .Bounds(1).lLbound
    End With
    
    'Stop
    
    'mTimer.Reset
    
    'Get to no of elements of the matrix
    lngCount = (lngCountX * lngCountY)

    'We need to maintain the order of the elements, so we need to create a buffer for the data
    'Each element requires 16 Byte space, so twice a Currency.
    'This does not actually hold the data itself but rather the variant info with the
    'pointers to the data.
    ReDim acurOld(1 To 2, 1 To lngCount)
    ReDim acurNew(1 To 2, 1 To lngCount)
    
    'copy the headerinfo of the data to the buffer
    CopyMemory acurOld(1, 1), ByVal lngData, lngCount * 16
    
    'swap the headerinfo  of the data and write it into the buffer acurNew
    For i = 1 To lngCountY
        For k = 1 To lngCountX
            m = m + 1
            acurNew(1, m) = acurOld(1, (k - 1) * lngCountY + i)
            acurNew(2, m) = acurOld(2, (k - 1) * lngCountY + i)
        Next
    Next
    
    'Now copy the swaped headerinfos back to their original position
    CopyMemory ByVal lngData, acurNew(1, 1), lngCount * 16

    'Now we pretend that the 1st dimension has the no of elements of the 2nd and vice versa
    CopyMemory ByVal lngSafeArray + 16, lngCountY, 4
    CopyMemory ByVal lngSafeArray + 20, lngLBoundY, 4
    CopyMemory ByVal lngSafeArray + 24, lngCountX, 4
    CopyMemory ByVal lngSafeArray + 28, lngLBoundX, 4
    
    'Debug.Print "HP Transposing the array took: " & mTimer.MillisecondsElapsed & " milliseconds..."
    
    'That's it -> the list is transposed

End Sub



Private Function DimensionCount(ByRef ArrayToCheck As Variant) As Integer
' ***********************************************
' Get the dimensions of the array
' ***********************************************
        
        'Dim lVariableType As Long
        Dim lpSADescriptor As Long
        Dim lpSafeArray As Long
        'Dim SafeArray2 As SAFEARRAY2D
        'Dim aToCheck As Variant: aToCheck = m_List
        
        If Not IsArray(ArrayToCheck) Then
                DimensionCount = -1
                Exit Function
        End If

        ' Get the tpye of the variable
        'CopyMemory lVariableType, ByVal VarPtr(ArrayToCheck), 2
    
        ' Get address of SAFEARRAY-Descriptor
        CopyMemory lpSADescriptor, ByVal VarPtr(ArrayToCheck) + 8, 4

        ' Get address of SAFEARRAY-Structure
        CopyMemory lpSafeArray, ByVal lpSADescriptor, 4
        
        If lpSafeArray <> 0 Then
            ' Get the first 2 Bytes, which contain the dimensitons of the array:
            CopyMemory DimensionCount, ByVal lpSafeArray, 2
        End If
          
End Function


'#Region
'#======================================================================================================================
'# General routines
'#======================================================================================================================

Private Sub ItemRead(ByRef vItem As Variant, ByVal iCol As Long, ByVal position As Long)
' ***********************************************
' Returns a given single value
' ***********************************************
    

    If m_EnableObjects Then
        If IsObject(m_List(iCol, position)) Then
            Set vItem = m_List(iCol, position)
         Else
            vItem = m_List(iCol, position)
        End If
     Else
        vItem = m_List(iCol, position)
    End If


End Sub


Private Sub ItemWrite(vValue As Variant, ByVal iCol As Long, ByVal position As Long)
' ***********************************************
' Writes a given single value to a given position
' ***********************************************


    If m_EnableObjects Then
        If IsObject(vValue) Then
            Set m_List(iCol, position) = vValue
         Else
            m_List(iCol, position) = vValue
        End If
     Else
        m_List(iCol, position) = vValue
    End If


End Sub

Private Sub RecordRead(ByRef returnRec() As Variant, ByVal position As Long)
' ***********************************************
' Returns a given single row of data
' ***********************************************

    Dim iCol As Long
    
    If m_EnableObjects Then
        For iCol = 1 To m_NumCols
            ' Return the record.
            If IsObject(m_List(iCol, position)) Then
                Set returnRec(iCol) = m_List(iCol, position)
             Else
                returnRec(iCol) = m_List(iCol, position)
            End If
        Next iCol
     Else
        For iCol = 1 To m_NumCols
            ' Return the record.
            returnRec(iCol) = m_List(iCol, position)
        Next iCol
    End If

End Sub


Private Sub RecordWrite(position As Long, Record() As Variant, iColOffset As Long)
' ***********************************************
' Writes a given single row of data to a given position
' ***********************************************


    Dim iCol As Long

    If m_EnableObjects Then
        'add the record to the list
        For iCol = 1 To m_NumCols
                If IsObject(Record(iCol + iColOffset)) Then
                    Set m_List(iCol, position) = Record(iCol + iColOffset)
                 Else
                    m_List(iCol, position) = Record(iCol + iColOffset)
                End If
        Next iCol
     Else
        'add the record to the list
        For iCol = 1 To m_NumCols
                m_List(iCol, position) = Record(iCol + iColOffset)
        Next iCol
    End If


End Sub

Private Sub DataRowCopy(ByRef Source() As Variant, idxFrom As Long, ByRef Target() As Variant, idxTo As Long)
' ***********************************************
' Copies a given single row of data from a given position
' in a source array into a given target array
' ***********************************************

    Dim iCol As Long

    If m_EnableObjects Then
        'add the record to the list
        For iCol = 1 To m_NumCols
                If IsObject(Source(iCol, idxFrom)) Then
                    Set Target(iCol, idxTo) = Source(iCol, idxFrom)
                 Else
                    Target(iCol, idxTo) = Source(iCol, idxFrom)
                End If
        Next iCol
     Else
        'add the record to the list
        For iCol = 1 To m_NumCols
                Target(iCol, idxTo) = Source(iCol, idxFrom)
        Next iCol
    End If


End Sub

Private Sub DataRowMove(ByRef List() As Variant, idxFrom As Long, idxTo As Long)
' ***********************************************
' Moves a given single row of data from a given source
' position to a given target position in an array
' ***********************************************


    Dim iCol As Long

    If m_EnableObjects Then
        'add the record to the list
        For iCol = 1 To m_NumCols
                If IsObject(List(iCol, idxFrom)) Then
                    Set List(iCol, idxTo) = List(iCol, idxFrom)
                 Else
                    List(iCol, idxTo) = List(iCol, idxFrom)
                End If
        Next iCol
     Else
        'add the record to the list
        For iCol = 1 To m_NumCols
                List(iCol, idxTo) = List(iCol, idxFrom)
        Next iCol
    End If



End Sub



Private Function GetColumnIndex(ByVal Index As String) As Long
' ***********************************************
' Get the column number of a given index
' Returns either the number provided (as Long rather than String)
' or the index based on the name of the column provided
' ***********************************************

    Dim iCol As Long: iCol = -1

    'initialize with Default not found = -1
    GetColumnIndex = iCol
    
    If IsNumeric(Index) Then
        'Use the Column index provided
        iCol = CLng(Index)
        GetColumnIndex = iCol
     Else
        'Find the column index based on the field name
        For iCol = 1 To m_NumCols
            If m_Columns(iCol).Name = Index Then
                    GetColumnIndex = iCol
                    Exit For
            End If
        Next iCol
    End If
    
End Function

Private Function GetColumnName(ByVal Index As Long) As String
' ***********************************************
' Get the column name of a given index
' ***********************************************

    GetColumnName = m_Columns(Index).Name

End Function

Private Function IsTableColumn(ByVal Index As String) As Boolean
' ***********************************************
' Checks if the column Name
' ***********************************************
    
    If GetColumnIndex(Index) = -1 Then
        IsTableColumn = False
     Else
        IsTableColumn = True
    End If


End Function

Private Sub SetSortedColumn(Index As Long)
' ***********************************************
' Sets the "IsSorted" flag for a given column
' ***********************************************


    Dim i As Long

    For i = LBound(m_Columns) To UBound(m_Columns)
        If i = Index Then
            m_Columns(i).IsSorted = True
         Else
            m_Columns(i).IsSorted = False
        End If
    Next i

End Sub


'#Region
'#======================================================================================================================
'# Service Routines
'#======================================================================================================================

Private Function Compare(Value As Variant, ByVal CompareWith As Variant, Comparison As MatchType) As Boolean
' ***********************************************
' Compares Value1 (value) with Value2 (CompareWith) and
' returns TRUE if specific comparison is positive
' ***********************************************
        
    Dim stValue As String
    Dim stCompareWith As String
    Dim lLenCW As Long
    Dim aMatchWords() As String

    Compare = False

    Select Case Comparison
    
        Case Equals
            If Value = CompareWith Then Compare = True
        Case DoesNotEqual
            If Value <> CompareWith Then Compare = True
        Case GreaterThen
            If Value > CompareWith Then Compare = True
        Case GreaterThenOrEqual
            If Value >= CompareWith Then Compare = True
        Case LessThen
            If Value < CompareWith Then Compare = True
        Case LessThenOrEqual
            If Value <= CompareWith Then Compare = True
        Case MatchesPattern, DoesNotMatchPattern
            stValue = Value
            stCompareWith = CompareWith
            Compare = PatternComparison(Value, CompareWith, Comparison)
        Case Else
            'nothing to do
    End Select

End Function

Private Function PatternComparison(ByVal Value As String, ByVal CompareWith As String, Comparison As MatchType) As Boolean
' ***********************************************
' Compares Value1 (value) with Value2 (CompareWith) and
' returns TRUE if specific comparison is positive
' ***********************************************

    Dim aMatchWords() As String
    Dim i As Long
    Dim iMatch As Long: iMatch = 0
    Dim iPrevMatch As Long: iPrevMatch = 1
    Dim matchWord As Boolean: matchWord = False
    Dim beginsWith As Boolean: beginsWith = False
    Dim endsWith As Boolean: endsWith = False
    
    PatternComparison = False
    
    'check if begins with condition
    If Left$(CompareWith, 1) <> "*" Then
        beginsWith = True
     Else
        CompareWith = Mid$(CompareWith, 2)
    End If
    
    'check if ends with condition
    If Right$(CompareWith, 1) <> "*" Then
        endsWith = True
     Else
        CompareWith = Left$(CompareWith, Len(CompareWith) - 1)
    End If
    
    'get array of single patterns
    aMatchWords = Split(CompareWith, "*")
    
    'outer check - compare begins with
    If beginsWith = True Then
       If Left$(Value, Len(aMatchWords(LBound(aMatchWords)))) = aMatchWords(LBound(aMatchWords)) Then
            If Comparison = MatchesPattern Then
                'Continue...
             Else
                'matchStartEnd = True
                matchWord = True
                'exit here
                GoTo Final_sequence
             End If
        Else
            If Comparison = MatchesPattern Then
                'exit here
                GoTo Final_sequence
             Else
                'exit here
                GoTo Final_sequence
            End If
       End If
    End If
        
    'outer check - compare ends with
    If endsWith = True Then
       If Right$(Value, Len(aMatchWords(UBound(aMatchWords)))) = aMatchWords(UBound(aMatchWords)) Then
            If Comparison = MatchesPattern Then
                'Continue...
             Else
                'matchStartEnd = True
                matchWord = True
                'exit here
                GoTo Final_sequence
             End If
        Else
            If Comparison = MatchesPattern Then
                'exit here
                GoTo Final_sequence
             Else
                'exit here
                GoTo Final_sequence
            End If
       End If
    End If

    
    'main loop of single patterns
    For i = LBound(aMatchWords) To UBound(aMatchWords)
        iMatch = InStr(iPrevMatch, Value, aMatchWords(i))
        If iMatch > 0 Then
            matchWord = True
            iPrevMatch = iMatch + 1
         Else
            matchWord = False
            Exit For
        End If
    Next i


Final_sequence:
    'final check
    If matchWord = True Then
        If Comparison = MatchesPattern Then
            PatternComparison = True
         Else
            PatternComparison = False
        End If
     Else
        If Comparison = MatchesPattern Then
            PatternComparison = False
         Else
            PatternComparison = True
        End If
    End If


End Function



Private Function getSafeArrayDescrPtr(ByRef varInput As Variant) As LongPtr

    Dim lngSafeArrayDescr As Long
    
    '1st pointer only - get the pointer to the Safearray descriptor of input array
    CopyMemory lngSafeArrayDescr, ByVal (VarPtr(varInput) + 8), 4
    
    getSafeArrayDescrPtr = lngSafeArrayDescr

End Function

Private Function getSafeArrayStructPtr(ByRef varInput As Variant) As LongPtr

    Dim lngSafeArrayDescr As Long
    Dim lngSafeArrayStruct As Long
    
    '1st pointer - get the pointer to the Safearray desciptor of input array
    CopyMemory lngSafeArrayDescr, ByVal (VarPtr(varInput) + 8), 4
    
    '2nd pointer - get the pointer to the Safearray desciptor of input array
    Call CopyMemory(lngSafeArrayStruct, ByVal lngSafeArrayDescr, 4&)
    
    getSafeArrayStructPtr = lngSafeArrayStruct

End Function

Private Function GetVarType(myVar) As String

' ---------------------------------------------------------------
' Written By Shanmuga Sundara Raman for http://vbadud.blogspot.com
' ---------------------------------------------------------------

'§§§ kept here for the sake of having all data types in one place, maybe used at a later stage §§§

    If VarType(myVar) = vbNull Then
        GetVarType = "Null (no valid data) "
    ElseIf VarType(myVar) = vbInteger Then
        GetVarType = "Integer "
    ElseIf VarType(myVar) = vbLong Then
        GetVarType = "Long integer "
    ElseIf VarType(myVar) = vbSingle Then
        GetVarType = "Single-precision floating-point number "
    ElseIf VarType(myVar) = vbDouble Then
        GetVarType = "Double-precision floating-point number "
    ElseIf VarType(myVar) = vbCurrency Then
        GetVarType = "Currency value "
    ElseIf VarType(myVar) = vbDate Then
        GetVarType = "Date value "
    ElseIf VarType(myVar) = vbString Then
        GetVarType = "String "
    ElseIf VarType(myVar) = vbObject Then
        GetVarType = "Object "
    ElseIf VarType(myVar) = vbError Then
        GetVarType = "Error value "
    ElseIf VarType(myVar) = vbBoolean Then
        GetVarType = "Boolean value "
    ElseIf VarType(myVar) = vbVariant Then
        GetVarType = "Variant (used only with arrays of variants) "
    ElseIf VarType(myVar) = vbDataObject Then
         GetVarType = "A data access object "
    ElseIf VarType(myVar) = vbDecimal Then
        GetVarType = "Decimal value "
    ElseIf VarType(myVar) = vbByte Then
        GetVarType = "Byte value "
    ElseIf VarType(myVar) = vbUserDefinedType Then
        GetVarType = "Variants that contain user-defined types "
    ElseIf VarType(myVar) = vbArray Then
        GetVarType = "Array "
    Else
        GetVarType = VarType(myVar)
    End If

End Function


Private Sub SplitCSV(Expression As String, ResultSplit() As String, ResultCols As Long, ResultRows As Long, Optional ColumnDelimiter As String = ",", Optional RowDelimiter As String = vbNewLine, Optional Quote As String = """")
'*************************************************************
' returns one dimensional zero based string array in ResultSplit containing parsed CSV cells
' - ResultCols (in/out) number of columns; if positive on input the CSV data is fixed to given number of columns
' - ResultRows (out) number of rows
' - (C): VBForums / User: Merri - http://www.vbforums.com/showthread.php?604031-Fast-amp-correctly-reading-CSV-parser-(Comma-separated-values)
'*************************************************************

    Dim CSV() As Integer, HeaderCSV(0 To 5) As LongPtr, lngCSV As LongPtr
    ' general variables that we need
    Dim intColumn As Integer, intQuote As Integer, lngRow As Long, strRow As String
    Dim lngExpLen As Long, lngRowLen As Long
    Dim blnQuote As Boolean, lngA As Long, lngB As Long, lngC As Long, lngCount As Long, lngResults() As Long
    ' some dummy variables that we happen to need
    Dim Compare As VbCompareMethod, SafeArrayBound(1) As Long
    ' length information
    lngExpLen = LenB(Expression)
    lngRowLen = LenB(RowDelimiter)
    ' validate lengths
    If lngExpLen > 0 And lngRowLen > 0 Then
        ' column delimiter
        If LenB(ColumnDelimiter) Then intColumn = AscW(ColumnDelimiter): ColumnDelimiter = Left$(ColumnDelimiter, 1) Else intColumn = 44: ColumnDelimiter = ","
        ' quote character
        If LenB(Quote) Then intQuote = AscW(Quote): Quote = Left$(Quote, 1) Else intQuote = 34: Quote = """"
        ' maximum number of results
        ReDim lngResults(0 To (lngExpLen \ lngRowLen))
        ' prepare CSV array
        HeaderCSV(0) = 1
        HeaderCSV(1) = 2
        HeaderCSV(3) = StrPtr(Expression)
        HeaderCSV(4) = Len(Expression)
        ' assign Expression data to the Integer array
        lngCSV = ArrayPtr(CSV)
        PutMem4 lngCSV, CLng(VarPtr(HeaderCSV(0)))
        ' find first row delimiter, see if within quote or not
        lngA = InStrB(1, Expression, RowDelimiter, Compare)
        Do Until (lngA And 1) Or (lngA = 0)
            lngA = InStrB(lngA + 1, Expression, RowDelimiter, Compare)
        Loop
        lngB = InStrB(1, Expression, Quote, Compare)
        Do Until (lngB And 1) Or (lngB = 0)
            lngB = InStrB(lngB + 1, Expression, Quote, Compare)
        Loop
        Do While lngA > 0
            If lngA + lngRowLen <= lngB Or lngB = 0 Then
                lngResults(lngCount) = lngA
                lngA = InStrB(lngA + lngRowLen, Expression, RowDelimiter, Compare)
                Do Until (lngA And 1) Or (lngA = 0)
                    lngA = InStrB(lngA + 1, Expression, RowDelimiter, Compare)
                Loop
                If lngCount Then
                    lngCount = lngCount + 1
                Else
                    ' calculate number of resulting columns if invalid number of columns
                    If ResultCols < 1 Then
                        ResultCols = 1
                        intColumn = AscW(ColumnDelimiter)
                        For lngC = 0 To (lngResults(0) - 1) \ 2
                            If blnQuote Then
                                If CSV(lngC) <> intQuote Then Else blnQuote = False
                            Else
                                Select Case CSV(lngC)
                                    Case intQuote
                                        blnQuote = True
                                    Case intColumn
                                        ResultCols = ResultCols + 1
                                End Select
                            End If
                        Next lngC
                    End If
                    lngCount = 1
                End If
            Else
                lngB = InStrB(lngB + 2, Expression, Quote, Compare)
                Do Until (lngB And 1) Or (lngB = 0)
                    lngB = InStrB(lngB + 1, Expression, Quote, Compare)
                Loop
                If lngB Then
                    lngA = InStrB(lngB + 2, Expression, RowDelimiter, Compare)
                    Do Until (lngA And 1) Or (lngA = 0)
                        lngA = InStrB(lngA + 1, Expression, RowDelimiter, Compare)
                    Loop
                    If lngA Then
                        lngB = InStrB(lngB + 2, Expression, Quote, Compare)
                        Do Until (lngB And 1) Or (lngB = 0)
                            lngB = InStrB(lngB + 1, Expression, Quote, Compare)
                        Loop
                    End If
                End If
            End If
        Loop
        lngResults(lngCount) = lngExpLen + 1
        ' number of rows
        ResultRows = lngCount + 1
        ' string array items to return
        ReDim Preserve ResultSplit(0 To ResultRows * ResultCols - 1)
        ' first row
        lngCount = 0
        strRow = LeftB$(Expression, lngResults(0) - 1)
        HeaderCSV(3) = StrPtr(strRow)
        lngC = 0
        blnQuote = False
        For lngB = 0 To (lngResults(0) - 1) \ 2
            If blnQuote Then
                Select Case CSV(lngB)
                    Case intQuote
                        If CSV(lngB + 1) = intQuote Then
                            ' skip next char (quote)
                            lngB = lngB + 1
                            ' add quote char
                            CSV(lngC) = intQuote
                            lngC = lngC + 1
                        Else
                            blnQuote = False
                        End If
                    Case Else
                        ' add this char
                        If lngB > lngC Then CSV(lngC) = CSV(lngB)
                        lngC = lngC + 1
                End Select
            Else
                Select Case CSV(lngB)
                    Case intQuote
                        blnQuote = True
                    Case intColumn
                        ' add this column
                        ResultSplit(lngCount) = Left$(strRow, lngC)
                        ' max column reached?
                        lngCount = lngCount + 1
                        If lngCount >= ResultCols Then Exit For
                        ' start filling column string buffer from start (strRow)
                        lngC = 0
                    Case Else
                        ' add this char
                        If lngB > lngC Then CSV(lngC) = CSV(lngB)
                        lngC = lngC + 1
                End Select
            End If
        Next lngB
        ' add last column item?
        If lngCount < ResultCols Then ResultSplit(lngCount) = Left$(strRow, lngC - 1)
        ' rows after first
        For lngA = 1 To ResultRows - 1
            ' start index for columns
            lngRow = lngA * ResultCols
            lngCount = 0
            strRow = MidB$(Expression, lngResults(lngA - 1) + lngRowLen, lngResults(lngA) - lngResults(lngA - 1) - lngRowLen)
            HeaderCSV(3) = StrPtr(strRow)
            lngC = 0
            blnQuote = False
            For lngB = 0 To (lngResults(lngA) - lngResults(lngA - 1) - lngRowLen) \ 2
                If blnQuote Then
                    Select Case CSV(lngB)
                        Case intQuote
                            If CSV(lngB + 1) = intQuote Then
                                ' skip next char (quote)
                                lngB = lngB + 1
                                ' add quote char
                                CSV(lngC) = intQuote
                                lngC = lngC + 1
                            Else
                                blnQuote = False
                            End If
                        Case Else
                            ' add this char
                            CSV(lngC) = CSV(lngB)
                            lngC = lngC + 1
                    End Select
                Else
                    Select Case CSV(lngB)
                        Case intQuote
                            blnQuote = True
                        Case intColumn
                            ' add this column
                            ResultSplit(lngRow + lngCount) = Left$(strRow, lngC)
                            ' max column reached?
                            lngCount = lngCount + 1
                            If lngCount >= ResultCols Then Exit For
                            ' start filling column string buffer from start (strRow)
                            lngC = 0
                        Case Else
                            ' add this char
                            If lngB > lngC Then CSV(lngC) = CSV(lngB)
                            lngC = lngC + 1
                    End Select
                End If
            Next lngB
            ' add last column item?
            If lngCount < ResultCols Then ResultSplit(lngRow + lngCount) = Left$(strRow, lngC - 1)
        Next lngA
        ' clean up CSV array
        PutMem4 lngCSV, 0
    Else
        ResultCols = 0
        ResultRows = 0
        ' clean any possible data that exists in the passed string array (like if it is multidimensional)
        If Not Not ResultSplit Then Erase ResultSplit
        ' mysterious IDE error fix
        'Debug.Assert App.hInstance
        ' reset to one element, one dimension
        ReDim ResultSplit(0 To 0)
        ' custom redimension: remove the items (this duplicates the VB6 Split behavior)
        SafeArrayRedim Not Not ResultSplit, SafeArrayBound(0)
    End If
End Sub
