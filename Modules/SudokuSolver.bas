Attribute VB_Name = "SudokuSolver"
Option Explicit
Public Sub Sudoku()
    Dim initialPuzzle As node
    Set initialPuzzle = View.Load()
     
    Dim nodeList As New List
    Call nodeList.Add(initialPuzzle)
    Dim totalNodes As Integer
    totalNodes = 1
    
    Dim currentPuzzle As node
    Dim solved As Boolean
    solved = False
        
    Dim bestSolution As Integer
    bestSolution = initialPuzzle.FilledValueCount
    Call View.Render(initialPuzzle, nodeList, totalNodes, bestSolution)
    
    Do While nodeList.Length > 0 And Not solved
        Set currentPuzzle = nodeList.PopNode
        Dim position As position
        Set position = currentPuzzle.NextEmptyPosition()
        
        Dim value As Integer
        For value = 1 To 9
          Dim newPuzzle As node
          Set newPuzzle = currentPuzzle.Clone
          Call newPuzzle.SetValue(position.row, position.col, value)
          If newPuzzle.IsValidSudoku(position) Then
            If newPuzzle.FilledValueCount > bestSolution Then
              Call View.Render(newPuzzle, nodeList, totalNodes, bestSolution)
              bestSolution = newPuzzle.FilledValueCount
            End If
            If newPuzzle.IsSolvedSudoku Then
              Call View.Render(newPuzzle, nodeList, totalNodes, bestSolution)
              solved = True
            End If
            Call nodeList.Add(newPuzzle)
            totalNodes = totalNodes + 1
          End If
        Next value
    Loop
End Sub
