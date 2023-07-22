Attribute VB_Name = "View"
Option Explicit
Public Function Load() As node
  Set Load = New node
  Dim row As Integer
  Dim col As Integer
  
  For row = 1 To 9
    For col = 1 To 9
      Call Load.SetValue(row, col, Cells(row, col))
    Next col
  Next row
End Function
Public Sub Render(puzzle As node, nodeList As List, totalNodes As Integer, bestSolution As Integer)
  Dim row As Integer
  Dim col As Integer
  Dim value As Integer
    
  For row = 1 To 9
    For col = 1 To 9
      value = puzzle.GetValue(row, col)
      If value = 0 Then
        Cells(row, col + 11) = ""
      Else
        Cells(row, col + 11) = value
      End If
    Next col
  Next row
  
  'Print stats
  Cells(11, 1) = "Total number of nodes: " & totalNodes
  Cells(12, 1) = "Current number of nodes: " & nodeList.Length
  Cells(13, 1) = "% Complete: " & 100 * bestSolution \ (9 * 9) & "%"
  DoEvents
End Sub


