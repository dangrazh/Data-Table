VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdaptiveRadixTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'# <author> Daniel Grass (generated with Claude Code)
'# Adaptive Radix Tree: maps string keys to sorted lists of Long row positions.
'# Designed as the non-unique index backing store for cIndex / cDataTable.
'#
'# Each key can have multiple associated positions (non-unique).
'# Positions within a key are kept sorted ascending for efficient range queries.
'#
'# Key encoding: strings are converted to ANSI byte sequences via
'# StrConv(key, vbFromUnicode).  Characters outside the ANSI range (Unicode > 255)
'# are mapped to '?' by StrConv – this is acceptable for typical Excel data.
'#
'# Reference: "The Adaptive Radix Tree: ARTful Indexing for Main-Memory Databases"
'#            Viktor Leis, Alfons Kemper, Thomas Neumann  (ICDE 2013)
'#
'# Node-type thresholds (same as the paper):
'#   Node4   -> Node16  when 5th child added   (IsFull at 4)
'#   Node16  -> Node48  when 17th child added   (IsFull at 16)
'#   Node48  -> Node256 when 49th child added   (IsFull at 48)
'#   Node16  <- Node4   when <=3 children remain after delete
'#   Node48  <- Node16  when <=12 children remain after delete
'#   Node256 <- Node48  when <=37 children remain after delete


'#======================================================================================================================
'# Private variables
'#======================================================================================================================
Private m_Root       As Object    ' cARTNode (may be Nothing when empty)
Private m_ItemCount  As Long      ' total (key, position) pairs stored
Private m_LeafCount  As Long      ' number of distinct keys


'#======================================================================================================================
'# Class lifecycle
'#======================================================================================================================

Private Sub Class_Initialize()
    Set m_Root = Nothing
    m_ItemCount = 0
    m_LeafCount = 0
End Sub

Private Sub Class_Terminate()
    Set m_Root = Nothing
End Sub


'#======================================================================================================================
'# Public properties
'#======================================================================================================================

Public Property Get ItemCount() As Long
' Total number of (key, position) pairs stored.
    ItemCount = m_ItemCount
End Property

Public Property Get LeafCount() As Long
' Number of distinct keys stored.
    LeafCount = m_LeafCount
End Property


'#======================================================================================================================
'# Public interface
'#======================================================================================================================

Public Sub Add(Key As String, Position As Long)
' Inserts the (Key, Position) pair. If Key already exists, Position is added to
' that key's sorted position list (duplicate positions are silently ignored).
    Dim keyBytes() As Byte
    keyBytes = KeyToBytes(Key)

    If m_Root Is Nothing Then
        ' First entry: create root leaf directly
        Dim firstLeaf As New cARTNode
        firstLeaf.InitAsLeaf Key
        firstLeaf.InsertPosition Position
        Set m_Root = firstLeaf
        m_ItemCount = 1
        m_LeafCount = 1
        Exit Sub
    End If

    Dim newRoot As Object
    Set newRoot = InsertRec(m_Root, keyBytes, 0, Key, Position)
    If Not newRoot Is Nothing Then
        Set m_Root = newRoot
    End If
End Sub


Public Sub Remove(Key As String, Position As Long)
' Removes the (Key, Position) pair.  If the key has no positions left, the key
' leaf is deleted from the tree.  No error if the pair does not exist.
    If m_Root Is Nothing Then Exit Sub

    Dim keyBytes() As Byte
    keyBytes = KeyToBytes(Key)

    ' Special-case: root is a leaf
    If m_Root.IsLeaf() Then
        If m_Root.LeafKey = Key Then
            If m_Root.RemovePosition(Position) Then
                m_ItemCount = m_ItemCount - 1
                If m_Root.PositionCount = 0 Then
                    m_LeafCount = m_LeafCount - 1
                    Set m_Root = Nothing
                End If
            End If
        End If
        Exit Sub
    End If

    Dim removed     As Boolean
    Dim leafDeleted As Boolean
    Dim newRoot     As Object

    Set newRoot = RemoveRec(m_Root, keyBytes, 0, Key, Position, removed, leafDeleted)

    If removed Then
        m_ItemCount = m_ItemCount - 1
        If leafDeleted Then m_LeafCount = m_LeafCount - 1
    End If

    If Not newRoot Is Nothing Then Set m_Root = newRoot

    ' If all items were removed, release root (handles rare empty-internal-node case)
    If m_LeafCount = 0 Then Set m_Root = Nothing
End Sub


Public Function GetFirstPosition(Key As String, StartAt As Long) As Long
' Returns the smallest position stored for Key that is >= StartAt.
' Returns 0 if Key is not found or no position satisfies the constraint.
    Dim leaf As Object
    Set leaf = FindLeaf(Key)
    If leaf Is Nothing Then
        GetFirstPosition = 0
    Else
        GetFirstPosition = leaf.GetFirstPosition(StartAt)
    End If
End Function


Public Sub UpdatePosition(Key As String, OldPos As Long, NewPos As Long)
' Replaces OldPos with NewPos in Key's position list.
' Used by cDataTable during garbage collection (record compaction).
    Dim leaf As Object
    Set leaf = FindLeaf(Key)
    If Not leaf Is Nothing Then
        leaf.UpdatePosition OldPos, NewPos
    End If
End Sub


Public Sub UpdateKey(OldKey As String, NewKey As String, Position As Long)
' Moves Position from OldKey to NewKey.
' Used by cDataTable when a cell value changes in an indexed column.
    Remove OldKey, Position
    Add NewKey, Position
End Sub


Public Sub RemoveAll()
' Clears the entire tree.
    Set m_Root = Nothing
    m_ItemCount = 0
    m_LeafCount = 0
End Sub


'#======================================================================================================================
'# Private: key encoding
'#======================================================================================================================

Private Function KeyToBytes(Key As String) As Byte()
' Converts Key to an ANSI byte array for byte-by-byte trie traversal.
' An empty key is represented as a single null byte so the tree can still store it.
    Dim b() As Byte
    If Len(Key) = 0 Then
        ReDim b(0 To 0)
        b(0) = 0
    Else
        b = StrConv(Key, vbFromUnicode)
    End If
    KeyToBytes = b
End Function


'#======================================================================================================================
'# Private: leaf lookup
'#======================================================================================================================

Private Function FindLeaf(Key As String) As Object
' Traverses the tree and returns the leaf node for Key, or Nothing.
    If m_Root Is Nothing Then
        Set FindLeaf = Nothing
        Exit Function
    End If

    Dim keyBytes() As Byte
    keyBytes = KeyToBytes(Key)

    Dim node    As Object
    Dim Depth   As Integer
    Dim byteVal As Integer

    Set node = m_Root
    Depth = 0

    Do While Not node Is Nothing
        If node.IsLeaf() Then
            If node.LeafKey = Key Then
                Set FindLeaf = node
            Else
                Set FindLeaf = Nothing
            End If
            Exit Function
        End If

        ' Verify path-compression prefix
        If node.PartialLen > 0 Then
            If node.CheckPrefix(keyBytes, Depth) < node.PartialLen Then
                Set FindLeaf = Nothing
                Exit Function
            End If
            Depth = Depth + node.PartialLen
        End If

        ' Descend via next byte
        If Depth > UBound(keyBytes) Then
            byteVal = 0
        Else
            byteVal = CInt(keyBytes(Depth))
        End If

        Set node = node.FindChild(byteVal)
        Depth = Depth + 1
    Loop

    Set FindLeaf = Nothing
End Function


'#======================================================================================================================
'# Private: recursive insert
'#======================================================================================================================

Private Function InsertRec(ByVal node   As Object, _
                            ByRef keyBytes() As Byte, _
                            ByVal Depth    As Integer, _
                            ByVal FullKey  As String, _
                            ByVal Position As Long) As Object
' Recursively inserts (FullKey, Position) into the subtree rooted at node.
' Returns a replacement node when the current node must be replaced (e.g. after
' a leaf split or a node-type upgrade); returns Nothing for in-place changes.

    Dim i As Integer

    ' ── Case 1: current node is a leaf ───────────────────────────────────────
    If node.IsLeaf() Then

        If node.LeafKey = FullKey Then
            ' Same key: just add the position to this leaf
            node.InsertPosition Position
            m_ItemCount = m_ItemCount + 1
            Set InsertRec = Nothing
            Exit Function
        End If

        ' Different key: split this leaf by creating a Node4 above both leaves.
        Dim existBytes() As Byte
        existBytes = KeyToBytes(node.LeafKey)

        ' Compute the common prefix length at the current depth
        Dim commonLen As Integer
        commonLen = 0
        Do
            If Depth + commonLen > UBound(keyBytes) Then Exit Do
            If Depth + commonLen > UBound(existBytes) Then Exit Do
            If keyBytes(Depth + commonLen) <> existBytes(Depth + commonLen) Then Exit Do
            commonLen = commonLen + 1
        Loop

        ' Build the split node with the common prefix as path compression
        Dim splitNode As New cARTNode
        splitNode.InitAsNode4
        If commonLen > 0 Then
            splitNode.SetPartialFromBytes keyBytes, Depth, commonLen
        End If

        ' Diverging bytes for old and new leaves
        Dim existDepth As Integer: existDepth = Depth + commonLen
        Dim newDepth   As Integer: newDepth   = Depth + commonLen

        Dim existByte As Integer
        Dim newByte   As Integer

        If existDepth > UBound(existBytes) Then
            existByte = 0       ' existing key is a prefix → use null terminator
        Else
            existByte = CInt(existBytes(existDepth))
        End If

        If newDepth > UBound(keyBytes) Then
            newByte = 0         ' new key is a prefix → use null terminator
        Else
            newByte = CInt(keyBytes(newDepth))
        End If

        ' Attach old leaf and new leaf as children of split node
        splitNode.AddChild existByte, node

        Dim newLeaf As New cARTNode
        newLeaf.InitAsLeaf FullKey
        newLeaf.InsertPosition Position
        splitNode.AddChild newByte, newLeaf

        m_ItemCount = m_ItemCount + 1
        m_LeafCount = m_LeafCount + 1
        Set InsertRec = splitNode
        Exit Function
    End If

    ' ── Case 2: current node is an internal node ──────────────────────────────

    ' Check whether the key matches the compressed path prefix
    If node.PartialLen > 0 Then
        Dim matchLen As Integer
        matchLen = node.CheckPrefix(keyBytes, Depth)

        If matchLen < node.PartialLen Then
            ' Mismatch in compressed path → split the internal node

            Dim splitNode2 As New cARTNode
            splitNode2.InitAsNode4

            ' New node's prefix = the matching part
            If matchLen > 0 Then
                splitNode2.SetPartialFromBytes keyBytes, Depth, matchLen
            End If

            ' Byte that diverges for the existing subtree
            Dim oldByte As Integer
            oldByte = node.GetDivergingByte(matchLen)

            ' Trim the existing node's prefix: remove first (matchLen+1) bytes
            node.TrimPartial matchLen + 1

            ' Attach existing subtree under its diverging byte
            splitNode2.AddChild oldByte, node

            ' Byte that diverges for the new key
            Dim newByte2 As Integer
            If Depth + matchLen > UBound(keyBytes) Then
                newByte2 = 0
            Else
                newByte2 = CInt(keyBytes(Depth + matchLen))
            End If

            ' Create and attach new leaf
            Dim newLeaf2 As New cARTNode
            newLeaf2.InitAsLeaf FullKey
            newLeaf2.InsertPosition Position
            splitNode2.AddChild newByte2, newLeaf2

            m_ItemCount = m_ItemCount + 1
            m_LeafCount = m_LeafCount + 1
            Set InsertRec = splitNode2
            Exit Function
        End If

        ' All prefix bytes matched: advance depth
        Depth = Depth + node.PartialLen
    End If

    ' Determine the routing byte for the next level
    Dim byteVal As Integer
    If Depth > UBound(keyBytes) Then
        byteVal = 0
    Else
        byteVal = CInt(keyBytes(Depth))
    End If

    Dim child As Object
    Set child = node.FindChild(byteVal)

    If child Is Nothing Then
        ' No child for this byte → create a new leaf
        Dim newLeaf3 As New cARTNode
        newLeaf3.InitAsLeaf FullKey
        newLeaf3.InsertPosition Position

        If node.IsFull() Then
            ' Upgrade node type first, then add
            Dim grownNode As Object
            Set grownNode = node.GrowNode()
            grownNode.AddChild byteVal, newLeaf3
            m_ItemCount = m_ItemCount + 1
            m_LeafCount = m_LeafCount + 1
            Set InsertRec = grownNode
        Else
            node.AddChild byteVal, newLeaf3
            m_ItemCount = m_ItemCount + 1
            m_LeafCount = m_LeafCount + 1
            Set InsertRec = Nothing
        End If

    Else
        ' Recurse into the existing child
        Dim newChild As Object
        Set newChild = InsertRec(child, keyBytes, Depth + 1, FullKey, Position)

        If Not newChild Is Nothing Then
            ' Child was replaced (split or grown): update parent's reference.
            ' This is always a replace (not net-add), so no capacity check needed.
            node.RemoveChild byteVal
            node.AddChild byteVal, newChild
        End If

        Set InsertRec = Nothing
    End If
End Function


'#======================================================================================================================
'# Private: recursive remove
'#======================================================================================================================

Private Function RemoveRec(ByVal node   As Object, _
                            ByRef keyBytes() As Byte, _
                            ByVal Depth    As Integer, _
                            ByVal FullKey  As String, _
                            ByVal Position As Long, _
                            ByRef Removed     As Boolean, _
                            ByRef LeafDeleted As Boolean) As Object
' Recursively removes (FullKey, Position) from the subtree rooted at node.
' Removed     is set True if the position was found and removed.
' LeafDeleted is set True if the leaf that held the position is now empty.
' Returns a replacement node when the current node must be replaced (shrink),
' otherwise returns Nothing.

    Removed     = False
    LeafDeleted = False
    Set RemoveRec = Nothing

    ' ── Leaf ─────────────────────────────────────────────────────────────────
    If node.IsLeaf() Then
        If node.LeafKey <> FullKey Then Exit Function
        Removed = node.RemovePosition(Position)
        If Removed And node.PositionCount = 0 Then
            LeafDeleted = True
        End If
        Exit Function
    End If

    ' ── Internal node ─────────────────────────────────────────────────────────

    ' Verify path-compression prefix
    If node.PartialLen > 0 Then
        If node.CheckPrefix(keyBytes, Depth) < node.PartialLen Then
            Exit Function
        End If
        Depth = Depth + node.PartialLen
    End If

    ' Find routing byte
    Dim byteVal As Integer
    If Depth > UBound(keyBytes) Then
        byteVal = 0
    Else
        byteVal = CInt(keyBytes(Depth))
    End If

    Dim child As Object
    Set child = node.FindChild(byteVal)
    If child Is Nothing Then Exit Function

    ' Recurse
    Dim childRemoved     As Boolean
    Dim childLeafDeleted As Boolean
    Dim newChild         As Object

    Set newChild = RemoveRec(child, keyBytes, Depth + 1, FullKey, Position, _
                             childRemoved, childLeafDeleted)

    Removed     = childRemoved
    LeafDeleted = childLeafDeleted

    If childLeafDeleted Then
        ' Remove the now-empty leaf from this node
        node.RemoveChild byteVal

        ' Check whether this node should shrink
        If ShouldShrink(node) Then
            Set RemoveRec = node.ShrinkNode()
        End If
    ElseIf Not newChild Is Nothing Then
        ' Child was replaced by a shrunk node
        node.RemoveChild byteVal
        node.AddChild byteVal, newChild
    End If
End Function


'#======================================================================================================================
'# Private: node-shrink threshold
'#======================================================================================================================

Private Function ShouldShrink(node As Object) As Boolean
' Returns True when the node has fallen below the threshold for its type.
    Select Case node.NodeType
        Case 16:   ShouldShrink = (node.NumChildren <= 3)
        Case 48:   ShouldShrink = (node.NumChildren <= 12)
        Case 256:  ShouldShrink = (node.NumChildren <= 37)
        Case Else: ShouldShrink = False
    End Select
End Function
