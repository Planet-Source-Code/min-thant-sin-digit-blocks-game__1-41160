Attribute VB_Name = "basInitPuzzle"
Option Explicit

Public Sub InitPuzzle(PuzzleSize As Integer)
      Dim X%, Y%, Index%
      
      bSolving = False
      bGameStart = True
      frmMain.picWinBox.Visible = False
      
      ReDim Block(1 To PuzzleSize ^ 2) As DIGITBLOCK
      ReDim Square(1 To PuzzleSize ^ 2) As SQUAREBOARD
                  
      With frmMain
            For X = 1 To 36
                  .picSquare(X).Visible = False
                  .picBlock(X).Visible = False
                  .picBlock(X).ZOrder 0
            Next X
            .picBackground.Picture = .Picture
            .Picture = LoadPicture("")
            .Picture = .picBackground.Picture
      End With
            
      Call ScramblePuzzle
      
      Dim Delay As Double
      
      For Y = 1 To PuzzleSize
            Delay = Timer
            For X = 1 To PuzzleSize
                  Index = ((Y * PuzzleSize) + X) - PuzzleSize
                  
                  With Square(Index)
                        .row = Y
                        .col = X
                        .Index = frmMain.picSquare(Index).Index
                        .HasBlock = False
                  End With
                                                      
                  Dim nWidth%, nHeight%
                  
                  'Position picture box controls
                  With frmMain
                        .picSquare(Index).Move ((X - 1) * BlockWidth) + (BlockWidth / 10), (Y * BlockHeight) - (BlockWidth / 4)
                        .picBlock(Board(X, Y)).Left = .picSquare(1).Left + (BlockWidth * PuzzleSize) + (X * BlockWidth) - (BlockWidth - 5)
                        .picBlock(Board(X, Y)).Top = .picSquare(1).Top + (Y - 1) * BlockHeight
                        Block(Index).Index = .picBlock(Index).Index
                        Block(Index).HomeX = .picBlock(Index).Left
                        Block(Index).HomeY = .picBlock(Index).Top
                        
                        nWidth = PuzzleSize * BlockWidth
                        nHeight = PuzzleSize * BlockHeight
                        .picWinBox.Move .picSquare(1).Left, .picSquare(1).Top, nWidth, nHeight
                        .imgSmile.Move (.picWinBox.Width - .imgSmile.Width) / 2, (.picWinBox.Height - .imgSmile.Height) / 2
                  End With
            Next X
      Next Y
      
      With frmMain.picSquare(1)
            NewLineX = .Left - 1
            NewLineY = .Top - 1
            NewLineWidth = .Left + BlockWidth * PuzzleSize
            NewLineHeight = .Top + BlockHeight * PuzzleSize
      End With
      frmMain.Line (NewLineX, NewLineY)-(NewLineWidth, NewLineHeight), vbBlack, B
      
      For X = 1 To PuzzleSize ^ 2
            frmMain.picBlock(X).Visible = True
            frmMain.picSquare(X).Visible = True
      Next X
End Sub
