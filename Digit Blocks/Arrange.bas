Attribute VB_Name = "basArrange"
Option Explicit

Sub ArrangeDigits(PuzzleSize As Integer, MaxNum As Byte)
      Dim X%, Y%, Index%
      
      'Set random numbers to block type
      For X = 1 To PuzzleSize ^ 2
            Randomize Timer
            With Block(X)
                  .LeftDigit = Int(Rnd * MaxNum)
                  .RightDigit = Int(Rnd * MaxNum)
                  .TopDigit = Int(Rnd * MaxNum)
                  .BottomDigit = Int(Rnd * MaxNum)
            End With
      Next X
      
      'Set left-right digits to be the same.
      For X = 1 To ((PuzzleSize ^ 2) - 1)
            Block(X + 1).LeftDigit = Block(X).RightDigit
      Next X
                        
      'Set top-bottom digits to be the same.
      For X = 1 To PuzzleSize
            For Y = 1 To (PuzzleSize - 1)
                  'Pattern
                  '---------------------
                  '1,2,3...
                  '4,5,6...
                  '---------------------
                  'Index = 1,4,2,5,3,6...
                  Index = (Y * PuzzleSize) + X - PuzzleSize
                  Block(Index).BottomDigit = Block(Index + PuzzleSize).TopDigit
            Next Y
      Next X
      
      'Scramble top digits of the first row
      'You can scramble to your heart's content.
      For X = 1 To PuzzleSize
            Randomize Timer
            Block(X).TopDigit = Int(Rnd * MaxNum)
      Next X
      
      'Scramble bottom digits of the last row
      'Just scramble it, dude!!
      For X = (PuzzleSize * (PuzzleSize - 1)) + 1 To (PuzzleSize ^ 2)
            Randomize Timer
            Block(X).BottomDigit = Int(Rnd * MaxNum)
      Next X
      
      'Scramble left digits of the first column
      'Yeah, I am doing a little scrambling job as a professional !!
      For X = 1 To 1
            For Y = 1 To PuzzleSize
                  'The first column digits.
                  Index = (Y * PuzzleSize) + X - PuzzleSize
                  Randomize Timer
                  Block(Index).LeftDigit = Int(Rnd * MaxNum)
            Next Y
      Next X
      
      'Scramble right digits of the last column
      'Here we go again !!
      For X = PuzzleSize To PuzzleSize
            For Y = 1 To PuzzleSize
                  'The last column digits
                  Index = (Y * PuzzleSize) + X - PuzzleSize
                  Randomize Timer
                  Block(Index).RightDigit = Int(Rnd * MaxNum)
            Next Y
      Next X
      
End Sub
