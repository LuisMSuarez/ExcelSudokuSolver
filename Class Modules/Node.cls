VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'List of properties
Private Values(1 To 9, 1 To 9) As Integer
Public NextNode As node
Public PrevNode As node
Private filledValues As Integer
Private Sub Class_Initialize()
  filledValues = 0
End Sub
Public Function GetValue(row As Integer, col As Integer) As Integer
  GetValue = Values(row, col)
End Function
Public Sub SetValue(row As Integer, col As Integer, value As Integer)
  Values(row, col) = value
  If value <> 0 Then
    filledValues = filledValues + 1
  End If
End Sub
Public Function Clone() As node
  Set Clone = New node
  Dim row As Integer
  Dim col As Integer
  
  For row = 1 To 9
    For col = 1 To 9
      Call Clone.SetValue(row, col, Values(row, col))
    Next col
  Next row
End Function
Public Function IsValidSudoku(pos As position) As Boolean
  Dim value As Integer
  Dim row As Integer
  Dim col As Integer
  Dim Map(1 To 9) As Boolean
  
  'scan row where new value was inserted
  Call ResetMap(Map)
  For col = 1 To 9
    value = Values(pos.row, col)
    If value <> 0 Then
      If Map(value) Then
        'Duplicate number found
        IsValidSudoku = False
        Exit Function
      Else
        Map(value) = True
      End If
    End If
  Next col
  
  'scan col where new value was inserted
  Call ResetMap(Map)
  For row = 1 To 9
    value = Values(row, pos.col)
    If value <> 0 Then
      If Map(value) Then
        'Duplicate number found
        IsValidSudoku = False
        Exit Function
      Else
        Map(value) = True
      End If
    End If
  Next row
  
  'scan square where new value was inserted
  Dim squareCornerRow As Integer
  Dim squareCornerCol As Integer
  squareCornerRow = pos.row - (pos.row - 1) Mod 3
  squareCornerCol = pos.col - (pos.col - 1) Mod 3
  Call ResetMap(Map)
    
  For row = 0 To 2
    For col = 0 To 2
      value = Values(row + squareCornerRow, col + squareCornerCol)
      If value <> 0 Then
        If Map(value) Then
          IsValidSudoku = False
          Exit Function
        Else
          Map(value) = True
        End If
      End If
    Next col
  Next row
  
  'no violations found, so it's a valid Sudoku
  IsValidSudoku = True
End Function
Public Function IsSolvedSudoku() As Boolean
  IsSolvedSudoku = filledValues = 9 * 9
End Function
Public Function FilledValueCount() As Integer
  FilledValueCount = filledValues
End Function
Public Function NextEmptyPosition() As position
  Set NextEmptyPosition = New position
  Dim row As Integer
  Dim col As Integer

  For row = 1 To 9
    For col = 1 To 9
      If Values(row, col) = 0 Then
        NextEmptyPosition.row = row
        NextEmptyPosition.col = col
        Exit Function
      End If
    Next col
  Next row
End Function
Private Sub ResetMap(Map() As Boolean)
  Dim i As Integer
  For i = 1 To 9
    Map(i) = False
  Next i
End Sub






