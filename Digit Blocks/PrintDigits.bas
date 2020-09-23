Attribute VB_Name = "basPrintDigits"
Option Explicit

Public Sub PrintDigits(PuzzleSize As Integer)
      Dim X As Integer
      Const OffsetX = 1
      Const OffsetY = 1
      
      'Now we get down to business...
      For X = 1 To PuzzleSize ^ 2
            With frmMain.picBlock(X)
                  .Picture = LoadPicture("")
                  If bSolving Then
                        .Picture = frmMain.picFixedBlock.Picture
                  Else
                        .Picture = frmMain.pctBlock.Picture
                  End If
                  
                  '""""""""""""""""""""""
                  'White color digits
                  '""""""""""""""""""""""
                  .ForeColor = vbWhite
                  
                  'Start printing...
                  'Left digit
                  .CurrentX = LeftDigitX - OffsetX
                  .CurrentY = LeftDigitY - OffsetY
                  frmMain.picBlock(X).Print Block(X).LeftDigit
                  
                  'Right digit
                  .CurrentX = RightDigitX - OffsetX
                  .CurrentY = RightDigitY - OffsetY
                  frmMain.picBlock(X).Print Block(X).RightDigit
                  
                  'Top digit
                  .CurrentX = TopDigitX - OffsetX
                  .CurrentY = TopDigitY - OffsetY
                  frmMain.picBlock(X).Print Block(X).TopDigit
                  
                  'Bottom digit
                  .CurrentX = BottomDigitX - OffsetX
                  .CurrentY = BottomDigitY - OffsetY
                  frmMain.picBlock(X).Print Block(X).BottomDigit
                  
                  '=================
                  'Black color digits
                  '=================
                  .ForeColor = vbBlack
                  
                  'Start printing...
                  'Left digit
                  .CurrentX = LeftDigitX
                  .CurrentY = LeftDigitY
                  frmMain.picBlock(X).Print Block(X).LeftDigit
                  
                  'Right digit
                  .CurrentX = RightDigitX
                  .CurrentY = RightDigitY
                  frmMain.picBlock(X).Print Block(X).RightDigit
                  
                  'Top digit
                  .CurrentX = TopDigitX
                  .CurrentY = TopDigitY
                  frmMain.picBlock(X).Print Block(X).TopDigit
                  
                  'Bottom digit
                  .CurrentX = BottomDigitX
                  .CurrentY = BottomDigitY
                  frmMain.picBlock(X).Print Block(X).BottomDigit
            End With
      Next X
End Sub
