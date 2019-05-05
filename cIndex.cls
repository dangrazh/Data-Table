VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'# <author> Daniel Grass
'# <mail> dani.grass@bluewin.ch

'#Region
'# Public Subs, Functions and Properties
'#======================================================================================================================
'# About this class
'#======================================================================================================================
' Index class module
'
' This class provides a database index like encapsulation of either an idex with unique keys (based on Hash Table class)
' or an idex with non unique keys (based on B+ Tree class).
'
'     |> Get | --- About                ::
'     |> --------- CreateLogFile        ::
'     |> Get | Let DirectoryPath        ::
'     |> Get | --- Name                 ::
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
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)


'#======================================================================================================================
'# Dependencies to other classes
'#======================================================================================================================
' cDataTable    :: is a single Index of a data column
' cHashTable    :: is used as underlying data structure to unique indices
' cBPlusTree    :: is used as underlying data structure to non-unique indices


'#======================================================================================================================
'# Application Constants, Enumerations & Types
'#======================================================================================================================

'the class name
Private Const C_Name As String = "cIndex"

Public Enum IndexType
    UniqueIndex = 1
    NonUniqueIndex = 2
    NotDefined = 3
End Enum

Private Enum MatchType
    Equals = 1
    DoesNotEqual = 2
    MatchesPattern = 3
    DoesNotMatchPattern = 4
    GreaterThen = 5
    GreaterThenOrEqual = 6
    LessThen = 7
    LessThenOrEqual = 8
End Enum



'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private m_IndexUnique As cHashTable             'The actual index itself - if it's a unique index
'Private m_IndexNonUniuqe As cBPlusTree          'The actual index itself - if it's a non unique index
'Private m_Index As Object                       'The generic representation of the index itself
Private m_IndexType As IndexType                'The type of the index
Private m_Name As String



'#Region
'#======================================================================================================================
'# Class Initialization, Termination & Properties
'#======================================================================================================================

Private Sub Class_Initialize()
' ************************************************
' Class constructor.
' ************************************************

    'Debug.Print "|> Initializing:= " & Me.Name
    'Initialize Defaults
    m_IndexType = NotDefined
    
    
End Sub

Private Sub Class_Terminate()
' ************************************************
' Class destructor.
' ************************************************

    'Debug.Print "|> Terminating:= " & Me.Name
    Set m_IndexUnique = Nothing
    'Set m_IndexNonUniuqe = Nothing
    'Set m_Index = Nothing
    

End Sub


Public Property Get Version() As String
' ************************************************
' Version string of the current class.
' Contains a list of (historical) changes to the class within the comments of the procedure.
' ************************************************
   Version = "Version 1.0 (03/2017)"

End Property

Public Property Get About() As String
' ***********************************************
' String that describes the current class.
' ***********************************************
    About = "This class provides a database index like encapsulation of either an idex with unique keys (based on Hash Table class) or  an idex with non unique keys (based on B+ Tree class). Version: " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details please contact the author."
End Property

Public Property Get ClassName() As String
' ***********************************************
' Returns the name of the class.
' ***********************************************
    
    Name = C_Name
    
End Property

Public Property Let Name(Value As String)
' ***********************************************
' Sets the name of the Index.
' ***********************************************

    m_Name = Value
    
End Property


Public Property Get Name() As String
' ***********************************************
' Returns the name of the index.
' ***********************************************

    Name = m_Name
    
End Property


Public Property Get IndexType() As IndexType
' ***********************************************
' Returns the type of the index.
' ***********************************************

    IndexType = m_IndexType
    
End Property

Public Property Let IndexType(Value As IndexType)
' ***********************************************
' Returns the type of the index.
' ***********************************************

    m_IndexType = Value
    
End Property


Public Sub Create(ByRef DataTable As cDataTable, Column As String)
' ***********************************************
' Create the actual index.
' ***********************************************

        Dim i As Long
        Dim stValue As String
        
        If m_IndexType = UniqueIndex Then
            Set m_IndexUnique = New cHashTable
            m_IndexUnique.SetSize DataTable.NumItems * 1.5
            DataTable.RsMoveFirst
            Do While Not DataTable.RsEOF
                stValue = DataTable.Item(Column)
                If m_IndexUnique.Exists(stValue) Then
                    Err.Raise vbObjectError + 514, Me.Name, "Value '" & DataTable.Item(Column) & "' is not unique! Failed to create unique index."
                 Else
                    m_IndexUnique.Add stValue, DataTable.RsBookmark
                End If
             DataTable.RsMoveNext
            Loop
         ElseIf m_IndexType = NonUniqueIndex Then
            '§§§ to be implemented §§§
         Else
            Err.Raise vbObjectError + 513, Me.Name, "Can not create index without a defined index type!"
        End If

End Sub

Public Sub ItemAdd(Value As String, Position As Long)
' ***********************************************
' Add a single item in the index.
' ***********************************************

    If m_IndexType = UniqueIndex Then
        If m_IndexUnique.Exists(Value) Then
            Err.Raise vbObjectError + 514, Me.Name, "Value '" & Value & "' is not unique! Failed to update unique index."
         Else
            m_IndexUnique.Add Value, Position
        End If
     ElseIf m_IndexType = NonUniqueIndex Then
        '§§§ to be implemented §§§
     Else
        Err.Raise vbObjectError + 513, Me.Name, "Can not create index without a defined index type!"
    End If



End Sub

Public Sub ItemRemove(Value As String, Position As Long)
' ***********************************************
' Add a single item in the index.
' ***********************************************

    If m_IndexType = UniqueIndex Then
        If m_IndexUnique.Exists(Value) Then
            m_IndexUnique.Remove Value
         Else
            Err.Raise vbObjectError + 514, Me.Name, "Value '" & Value & "' not found in index! Failed to delete from index."
        End If
     ElseIf m_IndexType = NonUniqueIndex Then
        '§§§ to be implemented §§§
     Else
        Err.Raise vbObjectError + 513, Me.Name, "Can not create index without a defined index type!"
    End If



End Sub

Public Sub ItemUpdate(ValueOld As String, ValueNew As String, Position As Long)
' ***********************************************
' Update a single item in the index.
' ***********************************************

    If m_IndexType = UniqueIndex Then
        If m_IndexUnique.Exists(ValueNew) Then
            Err.Raise vbObjectError + 514, Me.Name, "Value '" & ValueNew & "' is not unique! Failed to update unique index."
         Else
            m_IndexUnique.Remove ValueOld
            m_IndexUnique.Add ValueNew, Position
        End If
     ElseIf m_IndexType = NonUniqueIndex Then
        '§§§ to be implemented §§§
     Else
        Err.Raise vbObjectError + 513, Me.Name, "Can not create index without a defined index type!"
    End If


End Sub


Public Function GetNextItem(SearchFor As Variant, match As Long, Optional Start As Long = 1) As Long
' ***********************************************
' Get the pointer to the next item in the index
' ***********************************************
    
    Dim stSearchFor As String
    
    stSearchFor = CStr(SearchFor)

    If m_IndexType = UniqueIndex Then
        If m_IndexUnique.Exists(stSearchFor) Then
            GetNextItem = m_IndexUnique.Item(stSearchFor)
         Else
            GetNextItem = 0
        End If
     ElseIf m_IndexType = NonUniqueIndex Then
        '§§§ to be implemented §§§
    End If


End Function
