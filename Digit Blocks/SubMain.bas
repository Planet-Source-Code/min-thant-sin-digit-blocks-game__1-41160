Attribute VB_Name = "basMain"
Option Explicit

Sub Main()
      Dim X As Integer
      
      frmMain.Visible = False
      'Load picture box controls
      For X = 2 To 36
            'Load digits block picture boxes.
            Load frmMain.picBlock(X)
            'Load square picture boxes.
            Load frmMain.picSquare(X)
      Next X
      
      With frmMain
            .picShadow.Visible = False
            .picWinBox.Visible = False
            
            BlockWidth = .picBlock(1).ScaleWidth
            BlockHeight = .picBlock(1).ScaleHeight
      End With
      
      'Start with puzzle size  3
      PuzzleSize = 3
      MaxDigit = 9
      
      Call InitTextMeasures
      Call InitPuzzle(PuzzleSize)
      Call ArrangeDigits(PuzzleSize, MaxDigit)
      Call PrintDigits(PuzzleSize)
      frmMain.Show
      
End Sub

Sub InitTextMeasures()
      With frmMain.picBlock(1)
            'I set the digit width and height to be
            'the width and height of "0".
            DigitWidth = frmMain.picBlock(1).TextWidth("0")
            DigitHeight = frmMain.picBlock(1).TextHeight("0")
            
            'The following lengthy statements are used to
            'print the digits on the picture box control (picBlock).
            LeftDigitX = DigitWidth / 5
            LeftDigitY = (.Height - DigitHeight) / 2
                  
            RightDigitX = .Width - (DigitWidth * 2.25)
            RightDigitY = (.Height - DigitHeight) / 2
                  
            TopDigitX = (.Width - DigitWidth) / 2 - (DigitWidth / 3.5)
            TopDigitY = DigitHeight / 10
                  
            BottomDigitX = (.Width - DigitWidth) / 2 - (DigitWidth / 3.5)
            BottomDigitY = .Height - DigitHeight
      End With
End Sub
