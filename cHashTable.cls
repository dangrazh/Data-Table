VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'# <author> Francesco Balena posted on http://www.devx.com/vb2themax/Tip/19307
'           (with modifications and bug fixes by Daniel Grass)

'#Region
'# Public Subs, Functions and Properties
'#======================================================================================================================
'# About this class
'#======================================================================================================================
' HASHTABLE class module
'
' This class implements a hashtable, a structure that offers many of the features of a collectior or dictionary, and is often
' even faster than the built-in collection.
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
'   Dim ht As New cHashTable
'   ht.SetSize 10000           ' initial number of slots (the higher, the better)
'
'   ' enforce case-insensitive key search
'   ht.IgnoreCase = True
'
'   ' add values
'   ht.Add "key", value        ' add a value associated to a key
'
'   ' count how many values are in the table
'   Print ht.Count
'
'   ' read/write a value
'   Print ht("key")
'   ht("key") = newValue
'
'   ' remove a value
'   ht.Remove "key"
'   ' remove all values
'   ht.RemoveAll
'   ' check whether a value exists
'   If ht.Exists("key") Then ...
'
'   ' get the array of keys and values
'   Dim keys() As Variant, values() As Variant
'   keys() = ht.Keys
'   values() = ht.Values
'
'----------------------------------------------


'#======================================================================================================================
'# References
'#======================================================================================================================
#If Win64 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
#End If


'#======================================================================================================================
'# Application Constants, Enumerations & Types
'#======================================================================================================================

'String encoding constants
Const STRING_BASE = 37
Const ASC_A = 65        ' ASCII code for "A"


' default values
Const DEFAULT_HASHSIZE = 1024
Const DEFAULT_LISTSIZE = 2048
Const DEFAULT_CHUNKSIZE = 1024

Private Const C_Name As String = "cHashTable.cls"

Private Type SlotType
    Key As String
    Value As Variant
    nextItem As Long      ' 0 if last item
End Type


'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Dim hashTbl() As Long                   ' for each hash code this array holds the first element in slotTable()
                                        ' with the corresponding hash code
Dim slotTable() As SlotType             ' the array that holds the data
Dim FreeNdx As Long                     ' pointer to first free slot
Dim m_HashSize As Long                  ' size of hash table
Dim m_ListSize As Long                  ' size of slot table
Dim m_ChunkSize As Long                 ' chunk size
Dim m_Count As Long                     ' items in the slot table
Private m_IgnoreCase As Boolean         ' member variable for IgnoreCase property



'#Region
'#======================================================================================================================
'# Class Initialization, Termination & Properties
'#======================================================================================================================

Private Sub Class_Initialize()
' ************************************************
' Class constructor.
' ************************************************

    'Debug.Print "|> Initializing:= " & Me.Name

    ' initialize the tables at default size
    SetSize DEFAULT_HASHSIZE, DEFAULT_LISTSIZE, DEFAULT_CHUNKSIZE
    
End Sub

Private Sub Class_Terminate()
' ************************************************
' Class destructor.
' ************************************************

    'Debug.Print "|> Terminating:= " & Me.Name
    
    'empty the list
    Me.RemoveAll

End Sub


Public Property Get Version() As String
' ************************************************
' Version string of the current class.
' Contains a list of (historical) changes to the class within the comments of the procedure.
' ************************************************
   'Version = "Version 1.0 (07/2001)"  'Posted on http://www.devx.com/vb2themax/Tip/19307
   Version = "Version 1.1 (05/2018)"  'Added KeysAndValues procedure, fixed bug in Values procedure to output objects to array
                                      'fixed a bug in the Hash function which uccurs under certain circumstances when german umlauts (הצü)
                                      'and other special characters happen to be on the 1st place of a quad of characters

End Property

Public Property Get About() As String
' ***********************************************
' String that describes the current class.
' ***********************************************
    About = "Hash Table Class that offers many of the features of a collection or dictionary, and is often even faster than the built-in collection. Version: " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details please contact the author."
End Property

Public Property Get Name() As String
' ***********************************************
' Returns the name of the class.
' ***********************************************
    Name = C_Name
End Property




Property Get IgnoreCase() As Boolean
' ***********************************************
' Returns the IgnoreCase Flag (True if keys are
' searched in case-unsensitive mode)
' ***********************************************
    
    IgnoreCase = m_IgnoreCase
    
End Property

Property Let IgnoreCase(ByVal newValue As Boolean)
' ***********************************************
' True if keys are searched in case-unsensitive mode
' this can be assigned to only when the hash table is empty
' ***********************************************
    
    
    If m_Count Then
        Err.Raise 1001, , "The Hash Table isn't empty"
    End If
    m_IgnoreCase = newValue
End Property



Sub SetSize(ByVal HashSize As Long, Optional ByVal ListSize As Long, Optional ByVal ChunkSize As Long)
' ***********************************************
' initialize the hash table
' ***********************************************


    ' provide defaults
    If ListSize <= 0 Then ListSize = m_ListSize
    If ChunkSize <= 0 Then ChunkSize = m_ChunkSize
    ' save size values
    m_HashSize = HashSize
    m_ListSize = ListSize
    m_ChunkSize = ChunkSize
    m_Count = 0
    ' rebuild tables
    FreeNdx = 0
    ReDim hashTbl(0 To HashSize - 1) As Long
    ReDim slotTable(0) As SlotType
    ExpandSlotTable m_ListSize
End Sub

' check whether an item is in the hash table

Function Exists(Key As String) As Boolean
    Exists = GetSlotIndex(Key) <> 0
End Function

' add a new element to the hash table

Sub Add(Key As String, Value As Variant)
    Dim ndx As Long, Create As Boolean
    
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    Create = True
    ndx = GetSlotIndex(Key, Create)
    
    If Create Then
        ' the item was actually added
        If IsObject(Value) Then
            Set slotTable(ndx).Value = Value
        Else
            slotTable(ndx).Value = Value
        End If
    Else
        ' raise error "This key is already associated with an item of this
        ' collection"
        Err.Raise 457
    End If
End Sub

' the value associated to a key
' (empty if not found)

Property Get Item(Key As String) As Variant
    Dim ndx As Long
    ' get the index to the slot where the value is
    ndx = GetSlotIndex(Key)
    If ndx = 0 Then
        ' return Empty if not found
    ElseIf IsObject(slotTable(ndx).Value) Then
        Set Item = slotTable(ndx).Value
    Else
        Item = slotTable(ndx).Value
    End If
End Property

Property Let Item(Key As String, Value As Variant)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(Key, True)
    ' store the value
    slotTable(ndx).Value = Value
End Property

Property Set Item(Key As String, Value As Object)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(Key, True)
    ' store the value
    Set slotTable(ndx).Value = Value
End Property

' remove an item from the hash table

Sub Remove(Key As String)
    Dim ndx As Long, HCode As Long, LastNdx As Long
    ndx = GetSlotIndex(Key, False, HCode, LastNdx)
    ' raise error if no such element
    If ndx = 0 Then Err.Raise 5
    
    If LastNdx Then
        ' this isn't the first item in the slotTable() array
        slotTable(LastNdx).nextItem = slotTable(ndx).nextItem
    ElseIf slotTable(ndx).nextItem Then
        ' this is the first item in the slotTable() array
        ' and is followed by one or more items
        hashTbl(HCode) = slotTable(ndx).nextItem
    Else
        ' this is the only item in the slotTable() array
        ' for this hash code
        hashTbl(HCode) = 0
    End If
    
    ' put the element back in the free list
    slotTable(ndx).nextItem = FreeNdx
    FreeNdx = ndx
    ' we have deleted an item
    m_Count = m_Count - 1
    
End Sub

' remove all items from the hash table

Sub RemoveAll()
    SetSize m_HashSize, m_ListSize, m_ChunkSize
End Sub

' the number of items in the hash table

Property Get Count() As Long
    Count = m_Count
End Property

' the array of all keys

Property Get Keys() As Variant()
    Dim i As Long, ndx As Long
    Dim n As Long
    ReDim res(0 To m_Count - 1) As Variant
    
    For i = 0 To m_HashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            res(n) = slotTable(ndx).Key
            n = n + 1
            ndx = slotTable(ndx).nextItem
        Loop
    Next
        
    ' assign to the result
    Keys = res()
End Property

' the array of all values

Property Get Values() As Variant()
    Dim i As Long, ndx As Long
    Dim n As Long
    ReDim res(0 To m_Count - 1) As Variant
    
    For i = 0 To m_HashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            If IsObject(slotTable(ndx).Value) Then
                Set res(n) = slotTable(ndx).Value
             Else
                res(n) = slotTable(ndx).Value
            End If
            n = n + 1
            ndx = slotTable(ndx).nextItem
        Loop
    Next
        
    ' assign to the result
    Values = res()
End Property

Property Get KeysAndValues() As Variant()

    Dim i As Long, ndx As Long
    Dim n As Long
    ReDim res(0 To m_Count - 1, 0 To 1) As Variant
    
    For i = 0 To m_HashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            res(n, 0) = slotTable(ndx).Key
            If IsObject(slotTable(ndx).Value) Then
                Set res(n, 1) = slotTable(ndx).Value
             Else
                res(n, 1) = slotTable(ndx).Value
            End If
            n = n + 1
            ndx = slotTable(ndx).nextItem
        Loop
    Next
        
    ' assign to the result
    KeysAndValues = res()

End Property

' expand the slotTable() array

Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
    
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As SlotType
    ' create the linked list of free items
    For i = newFreeNdx To UBound(slotTable)
        slotTable(i).nextItem = i + 1
    Next
    ' overwrite the last (wrong) value
    slotTable(UBound(slotTable)).nextItem = FreeNdx
    ' we now know where to pick the first free item
    FreeNdx = newFreeNdx
End Sub

' return the hash code of a string

Private Function HashCode(Key As String) As Long
    Dim lastEl As Long, i As Long
    Dim stMid As String
    Dim iStart As Long
    
    ' copy ansi codes into an array of long
    lastEl = (Len(Key) - 1) \ 4
    ReDim codes(lastEl) As Long
    ' this also converts from Unicode to ANSI
    CopyMemory codes(0), ByVal Key, Len(Key)
    
    ' XOR the ANSI codes of all characters
    For i = 0 To lastEl
        If codes(i) < 0 Then
            'this is the fix for the bug when special characters on the 1st position
            'of the quad of characters lead to negative interger representations
            iStart = i * 4 + 1
            If iStart + 4 > Len(Key) Then
                stMid = Mid$(Key, iStart)
             Else
                stMid = Mid$(Key, iStart, 4)
            End If
            HashCode = HashCode Xor StringToLng(stMid, Len(stMid) + 1)
         Else
            HashCode = HashCode Xor codes(i)
        End If
    Next
    
End Function

' get the index where an item is stored or 0 if not found
' if Create = True the item is created
'
' on exit Create=True only if a slot has been actually created

Private Function GetSlotIndex(ByVal Key As String, Optional Create As Boolean, _
    Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long
    
    ' raise error if invalid key
    If Len(Key) = 0 Then Err.Raise 1001, , "Invalid key"
    
    ' keep case-unsensitiveness into account
    If m_IgnoreCase Then Key = UCase$(Key)
    ' get the index in the hashTbl() array
    HCode = HashCode(Key) Mod m_HashSize
    ' get the pointer to the slotTable() array
    ndx = hashTbl(HCode)
    
    ' exit if there is no item with that hash code
    Do While ndx
        ' compare key with actual value
        If slotTable(ndx).Key = Key Then Exit Do
        ' remember last pointer
        LastNdx = ndx
        ' check the next item
        ndx = slotTable(ndx).nextItem
    Loop
    
    ' create a new item if not there
    If ndx = 0 And Create Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, Key, HCode, LastNdx
    Else
        ' signal that no item has been created
        Create = False
    End If
    ' this is the return value
    GetSlotIndex = ndx

End Function

' return the first free slot

Private Function GetFreeSlot() As Long
    ' allocate new memory if necessary
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize
    ' use the first slot
    GetFreeSlot = FreeNdx
    ' update the pointer to the first slot
    FreeNdx = slotTable(GetFreeSlot).nextItem
    ' signal this as the end of the linked list
    slotTable(GetFreeSlot).nextItem = 0
    ' we have one more item
    m_Count = m_Count + 1
End Function

' assign a key and value to a given slot

Private Sub PrepareSlot(ByVal Index As Long, ByVal Key As String, _
    ByVal HCode As Long, ByVal LastNdx As Long)
    ' assign the key
    ' keep case-sensitiveness into account
    If m_IgnoreCase Then Key = UCase$(Key)
    slotTable(Index).Key = Key
    
    If LastNdx Then
        ' this is the successor of another slot
        slotTable(LastNdx).nextItem = Index
    Else
        ' this is the first slot for a given hash code
        hashTbl(HCode) = Index
    End If
End Sub


Function StringToLng(txt As String, full_len As Integer) As Long
' ************************************************
' Convert a string into a double encoding.
' full_len gives the full length of the string.
' For example, "AX" as a three character string
' would have full_len = 3.
' ************************************************

Dim strlen As Integer
Dim i As Integer
Dim Value As Double
Dim ch As String * 1
    
    strlen = Len(txt)
    If strlen > full_len Then strlen = full_len
    
    Value = 0#
    For i = 1 To strlen
        ch = Mid$(txt, i, 1)
        Value = Value * STRING_BASE + Asc(ch) - ASC_A + 1
    Next i

    For i = strlen + 1 To full_len
        Value = Value * STRING_BASE
    Next i

    StringToLng = Value
End Function

