Attribute VB_Name = "basScramble"
Option Explicit

Public Sub ScramblePuzzle()
      Dim X As Integer, I As Integer
      Dim col%, row%, rndNum%
      'Used to place the digit blocks in a random order.
      Dim colPuzzle As New Collection
      
      'Add digits to collection.
      For X = 1 To PuzzleSize ^ 2
            colPuzzle.Add X, Str(X)
      Next
      
      'Re-dim the Board array.
      ReDim Board(1 To PuzzleSize, 1 To PuzzleSize) As Integer
         
      For col = 1 To PuzzleSize
            For row = 1 To PuzzleSize
                  'Generate randomizer.
                  Randomize Timer
                  'Get a random number.
                  rndNum = colPuzzle.Item(Int(Rnd * (colPuzzle.Count - 1) + 1))
                  'Remove it from the collection.
                  colPuzzle.Remove Str(rndNum)
                  'Set the random number to Board array.
                  Board(row, col) = rndNum
            Next row
      Next col
      Set colPuzzle = Nothing
End Sub
