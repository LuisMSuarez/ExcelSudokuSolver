VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'List of properties
Private Head As node
Private Tail As node
Public Length As Integer
Public Sub Add(node As node)
  If Tail Is Nothing Then
    Set Tail = node
    Set Head = node
  Else
    Set Tail.NextNode = node
    Set node.PrevNode = Tail
    Set Tail = node
  End If
  Length = Length + 1
End Sub
Public Function PopNode() As node
  If Length < 1 Then
    Err.Raise vbObjectError + 513, "List", "Cannot pop a node on an empty list!"
  End If
  
  Set PopNode = Tail
  Set Tail = PopNode.PrevNode
  Length = Length - 1
End Function


