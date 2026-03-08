VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cARTNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'# <author> Daniel Grass (generated with Claude Code)
'# Internal node class for the Adaptive Radix Tree (cAdaptiveRadixTree).
'# Reference: "The Adaptive Radix Tree: ARTful Indexing for Main-Memory Databases"
'#            by Viktor Leis, Alfons Kemper, Thomas Neumann (ICDE 2013)
'#
'# Four inner node types (Node4 / Node16 / Node48 / Node256) differ only in how
'# they map a byte value to a child pointer. Leaf nodes store the full key string
'# (for collision detection) and a sorted array of row positions.
'#
'# Path compression: shared prefixes among all children are stored in Partial().
'# PartialLen bytes are stored inline; the full prefix is always kept (no cap).

'#======================================================================================================================
'# Node-type constants  (public so cAdaptiveRadixTree can reference them through an instance)
'#======================================================================================================================
Const ART_LEAF    As Integer = 0
Const ART_NODE4   As Integer = 4
Const ART_NODE16  As Integer = 16
Const ART_NODE48  As Integer = 48
Const ART_NODE256 As Integer = 256

'#======================================================================================================================
'# Public fields
'#======================================================================================================================

' Which kind of node this is
Public NodeType As Integer

' Path compression  ----------------------------------------------------------
' PartialLen bytes shared by all keys in this subtree (stored at this node
' so traversal can skip past them without descending).
Public PartialLen As Integer
Private Partial()  As Byte     ' dynamically sized; valid indices 0 .. PartialLen-1

' Number of children (internal nodes only)
Public NumChildren As Integer

' Node4 / Node16: parallel arrays  ------------------------------------------
' Keys(i) holds the byte value that routes to Children(i).
Private Keys()     As Integer  ' byte values 0-255, dimensioned 0..3 or 0..15
Private Children() As Object   ' cARTNode references, same size as Keys

' Node48: 256-slot byte?slot map  --------------------------------------------
' KeyIndex(b) = 0          ? no child for byte b
' KeyIndex(b) = s+1 (>0)  ? child is Children(s)  (s is 0-based slot)
Private KeyIndex(0 To 255) As Integer

' Node48 / Node256 use Children() directly (48 or 256 slots).
' Node256 does not use Keys() or KeyIndex(); Children(b) is the child for byte b.

' Leaf fields  ----------------------------------------------------------------
Public LeafKey       As String  ' full key string (for exact-match confirmation)
Public PositionCount As Long    ' number of valid positions stored
Private Positions()   As Long    ' sorted ascending, 1-based; valid 1 .. PositionCount


'#======================================================================================================================
'# Class lifecycle
'#======================================================================================================================

Private Sub Class_Initialize()
    NodeType = ART_LEAF
    PartialLen = 0
    NumChildren = 0
    PositionCount = 0
End Sub

Private Sub Class_Terminate()
    ' Let VBA reference-counting handle child nodes
End Sub


'#======================================================================================================================
'# Node initializers  (call exactly one after New cARTNode)
'#======================================================================================================================

Public Sub InitAsLeaf(Key As String)
    NodeType = ART_LEAF
    LeafKey = Key
    PositionCount = 0
End Sub

Public Sub InitAsNode4()
    NodeType = ART_NODE4
    NumChildren = 0
    ReDim Keys(0 To 3)
    ReDim Children(0 To 3)
End Sub

Public Sub InitAsNode16()
    NodeType = ART_NODE16
    NumChildren = 0
    ReDim Keys(0 To 15)
    ReDim Children(0 To 15)
End Sub

Public Sub InitAsNode48()
    Dim i As Integer
    NodeType = ART_NODE48
    NumChildren = 0
    For i = 0 To 255
        KeyIndex(i) = 0
    Next i
    ReDim Children(0 To 47)
End Sub

Public Sub InitAsNode256()
    NodeType = ART_NODE256
    NumChildren = 0
    ReDim Children(0 To 255)
End Sub


'#======================================================================================================================
'# Properties
'#======================================================================================================================

Public Property Get IsLeaf() As Boolean
    IsLeaf = (NodeType = ART_LEAF)
End Property

Public Property Get IsFull() As Boolean
    Select Case NodeType
        Case ART_NODE4:   IsFull = (NumChildren >= 4)
        Case ART_NODE16:  IsFull = (NumChildren >= 16)
        Case ART_NODE48:  IsFull = (NumChildren >= 48)
        Case Else:        IsFull = False   ' Node256 never full; Leaf irrelevant
    End Select
End Property


'#======================================================================================================================
'# Child-pointer operations
'#======================================================================================================================

Public Function FindChild(byteVal As Integer) As Object
' Returns the child node for ByteVal, or Nothing if absent.
    Dim i As Integer
    Select Case NodeType
        Case ART_NODE4, ART_NODE16
            For i = 0 To NumChildren - 1
                If Keys(i) = byteVal Then
                    Set FindChild = Children(i)
                    Exit Function
                End If
            Next i
            Set FindChild = Nothing

        Case ART_NODE48
            If KeyIndex(byteVal) > 0 Then
                Set FindChild = Children(KeyIndex(byteVal) - 1)
            Else
                Set FindChild = Nothing
            End If

        Case ART_NODE256
            Set FindChild = Children(byteVal)   ' may be Nothing

        Case Else
            Set FindChild = Nothing
    End Select
End Function


Public Sub AddChild(byteVal As Integer, ChildNode As Object)
' Adds a child for ByteVal. Caller must ensure the node is not full.
    Dim i As Integer
    Select Case NodeType
        Case ART_NODE4, ART_NODE16
            Keys(NumChildren) = byteVal
            Set Children(NumChildren) = ChildNode
            NumChildren = NumChildren + 1

        Case ART_NODE48
            ' Find first free slot (Children(i) Is Nothing)
            For i = 0 To 47
                If Children(i) Is Nothing Then
                    KeyIndex(byteVal) = i + 1      ' 1-based
                    Set Children(i) = ChildNode
                    NumChildren = NumChildren + 1
                    Exit Sub
                End If
            Next i

        Case ART_NODE256
            Set Children(byteVal) = ChildNode
            NumChildren = NumChildren + 1
    End Select
End Sub


Public Sub RemoveChild(byteVal As Integer)
' Removes the child for ByteVal (if present).
    Dim i As Integer
    Dim j As Integer
    Select Case NodeType
        Case ART_NODE4, ART_NODE16
            For i = 0 To NumChildren - 1
                If Keys(i) = byteVal Then
                    ' Shift remaining entries left
                    For j = i To NumChildren - 2
                        Keys(j) = Keys(j + 1)
                        Set Children(j) = Children(j + 1)
                    Next j
                    Set Children(NumChildren - 1) = Nothing
                    NumChildren = NumChildren - 1
                    Exit Sub
                End If
            Next i

        Case ART_NODE48
            If KeyIndex(byteVal) > 0 Then
                Set Children(KeyIndex(byteVal) - 1) = Nothing
                KeyIndex(byteVal) = 0
                NumChildren = NumChildren - 1
            End If

        Case ART_NODE256
            If Not Children(byteVal) Is Nothing Then
                Set Children(byteVal) = Nothing
                NumChildren = NumChildren - 1
            End If
    End Select
End Sub


'#======================================================================================================================
'# Node growth / shrink
'#======================================================================================================================

Public Function GrowNode() As Object
' Creates the next-larger node type and copies all children into it.
' Returns the new node (caller replaces the old reference with it).
    Dim newNode As cARTNode
    Dim i As Integer
    Select Case NodeType
        Case ART_NODE4
            Set newNode = New cARTNode
            newNode.InitAsNode16
            CopyPartialTo newNode
            For i = 0 To NumChildren - 1
                newNode.AddChild Keys(i), Children(i)
            Next i
            Set GrowNode = newNode

        Case ART_NODE16
            Set newNode = New cARTNode
            newNode.InitAsNode48
            CopyPartialTo newNode
            For i = 0 To NumChildren - 1
                newNode.AddChild Keys(i), Children(i)
            Next i
            Set GrowNode = newNode

        Case ART_NODE48
            Set newNode = New cARTNode
            newNode.InitAsNode256
            CopyPartialTo newNode
            For i = 0 To 255
                If KeyIndex(i) > 0 Then
                    newNode.AddChild i, Children(KeyIndex(i) - 1)
                End If
            Next i
            Set GrowNode = newNode

        Case Else
            Set GrowNode = Nothing    ' Node256 cannot grow further
    End Select
End Function


Public Function ShrinkNode() As Object
' Creates the next-smaller node type and copies all children into it.
    Dim newNode As cARTNode
    Dim i As Integer
    Select Case NodeType
        Case ART_NODE16
            Set newNode = New cARTNode
            newNode.InitAsNode4
            CopyPartialTo newNode
            For i = 0 To NumChildren - 1
                newNode.AddChild Keys(i), Children(i)
            Next i
            Set ShrinkNode = newNode

        Case ART_NODE48
            Set newNode = New cARTNode
            newNode.InitAsNode16
            CopyPartialTo newNode
            For i = 0 To 255
                If KeyIndex(i) > 0 Then
                    newNode.AddChild i, Children(KeyIndex(i) - 1)
                End If
            Next i
            Set ShrinkNode = newNode

        Case ART_NODE256
            Set newNode = New cARTNode
            newNode.InitAsNode48
            CopyPartialTo newNode
            For i = 0 To 255
                If Not Children(i) Is Nothing Then
                    newNode.AddChild i, Children(i)
                End If
            Next i
            Set ShrinkNode = newNode

        Case Else
            Set ShrinkNode = Nothing
    End Select
End Function


'#======================================================================================================================
'# Path-compression helpers
'#======================================================================================================================

Public Sub SetPartialFromBytes(ByRef Bytes() As Byte, StartIdx As Integer, Length As Integer)
' Stores Length bytes from Bytes(StartIdx..) as this node's compressed prefix.
    Dim i As Integer
    PartialLen = Length
    If Length > 0 Then
        ReDim Partial(0 To Length - 1)
        For i = 0 To Length - 1
            Partial(i) = Bytes(StartIdx + i)
        Next i
    End If
End Sub


Public Sub CopyPartialTo(ByRef Target As cARTNode)
' Copies this node's path-compression data to Target.
    If PartialLen > 0 Then
        Target.SetPartialFromBytes Partial, 0, PartialLen
    Else
        Target.PartialLen = 0
    End If
End Sub


Public Function CheckPrefix(ByRef keyBytes() As Byte, Depth As Integer) As Integer
' Returns the number of bytes in this node's Partial() that match KeyBytes
' starting at position Depth.  Returns PartialLen if all match.
    Dim i As Integer
    For i = 0 To PartialLen - 1
        If Depth + i > UBound(keyBytes) Then
            CheckPrefix = i     ' Key ran out before prefix ended
            Exit Function
        End If
        If Partial(i) <> keyBytes(Depth + i) Then
            CheckPrefix = i     ' Mismatch at position i
            Exit Function
        End If
    Next i
    CheckPrefix = PartialLen    ' All prefix bytes matched
End Function


Public Function GetDivergingByte(matchLen As Integer) As Integer
' Returns the byte at position matchLen in the Partial() array.
' Used during path-compression splits before the partial is trimmed.
    GetDivergingByte = CInt(Partial(matchLen))
End Function


Public Sub TrimPartial(TrimCount As Integer)
' Removes the first TrimCount bytes from Partial() (shift-left in place).
    Dim newLen As Integer
    Dim i As Integer
    newLen = PartialLen - TrimCount
    If newLen < 0 Then newLen = 0
    PartialLen = newLen
    If newLen > 0 Then
        For i = 0 To newLen - 1
            Partial(i) = Partial(i + TrimCount)
        Next i
        ReDim Preserve Partial(0 To newLen - 1)
    End If
End Sub


'#======================================================================================================================
'# Leaf position-list operations  (only valid when NodeType = ART_LEAF)
'#======================================================================================================================

Public Sub InsertPosition(Position As Long)
' Inserts Position into the sorted Positions() array (duplicates are ignored).
    Dim lo As Long
    Dim hi As Long
    Dim mid As Long
    Dim i As Long

    If PositionCount = 0 Then
        PositionCount = 1
        ReDim Positions(1 To 1)
        Positions(1) = Position
        Exit Sub
    End If

    ' Binary search for the insertion point
    lo = 1
    hi = PositionCount
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If Positions(mid) = Position Then
            Exit Sub            ' Already present, nothing to do
        ElseIf Positions(mid) < Position Then
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop

    ' lo is the insertion point; shift elements right and insert
    PositionCount = PositionCount + 1
    ReDim Preserve Positions(1 To PositionCount)
    For i = PositionCount To lo + 1 Step -1
        Positions(i) = Positions(i - 1)
    Next i
    Positions(lo) = Position
End Sub


Public Function RemovePosition(Position As Long) As Boolean
' Removes Position from Positions(). Returns True if it was found and removed.
    Dim lo As Long
    Dim hi As Long
    Dim mid As Long
    Dim i As Long

    If PositionCount = 0 Then
        RemovePosition = False
        Exit Function
    End If

    lo = 1
    hi = PositionCount
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If Positions(mid) = Position Then
            ' Shift remaining elements left
            For i = mid To PositionCount - 1
                Positions(i) = Positions(i + 1)
            Next i
            PositionCount = PositionCount - 1
            If PositionCount > 0 Then
                ReDim Preserve Positions(1 To PositionCount)
            End If
            RemovePosition = True
            Exit Function
        ElseIf Positions(mid) < Position Then
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop

    RemovePosition = False
End Function


Public Function UpdatePosition(OldPos As Long, NewPos As Long) As Boolean
' Replaces OldPos with NewPos in the sorted list. Returns True if OldPos was found.
    If RemovePosition(OldPos) Then
        InsertPosition NewPos
        UpdatePosition = True
    Else
        UpdatePosition = False
    End If
End Function


Public Function GetFirstPosition(StartAt As Long) As Long
' Returns the smallest position >= StartAt, or 0 if none exists.
    Dim lo As Long
    Dim hi As Long
    Dim mid As Long
    Dim result As Long

    If PositionCount = 0 Then
        GetFirstPosition = 0
        Exit Function
    End If

    lo = 1
    hi = PositionCount
    result = 0
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If Positions(mid) >= StartAt Then
            result = Positions(mid)
            hi = mid - 1     ' try to find a smaller one
        Else
            lo = mid + 1
        End If
    Loop

    GetFirstPosition = result
End Function

