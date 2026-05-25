VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSQLExprNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SQLExprType
    exLiteral = 1
    exColumn = 2
    exBinaryOp = 3
    exUnaryOp = 4
    exFunction = 5
    exBetween = 6
    exInList = 7
    exIsNull = 8
    exStar = 9
End Enum

Public Enum SQLLiteralType
    litString = 1
    litNumber = 2
    litBoolean = 3
    litNull = 4
End Enum

Public NodeType As SQLExprType
Public LiteralType As SQLLiteralType
Public StringValue As String
Public NumValue As Double
Public BoolValue As Boolean
Public ColumnName As String
Public TablePrefix As String
Public Operator As String
Public FuncName As String
Public LeftChild As cSQLExprNode
Public RightChild As cSQLExprNode
Public ThirdChild As cSQLExprNode
Public Args As Collection
Public IsNegated As Boolean
Public AliasName As String
Public IsDistinctFunc As Boolean

Private Sub Class_Initialize()
    Set Args = New Collection
End Sub
