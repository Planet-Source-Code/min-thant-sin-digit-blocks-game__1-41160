Attribute VB_Name = "basVariable"
Option Explicit

Type RECT
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
End Type

Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
'Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Type SQUAREBOARD
      Index As Integer
      HasBlock As Boolean
      CurBlock As Integer
      row As Integer
      col As Integer
      HomeX As Integer
      HomeY As Integer
End Type

Public Type DIGITBLOCK
      Index As Integer
      CurSquare As Integer
      HomeX As Single
      HomeY As Single
      
      LeftDigit As Integer
      RightDigit As Integer
      TopDigit As Integer
      BottomDigit As Integer
End Type

Public Square() As SQUAREBOARD
Public Block() As DIGITBLOCK
Public Board() As Integer

Public PuzzleSize As Integer

Public BlockWidth As Single
Public BlockHeight As Single

'These are for printing digits.
Public DigitWidth%, DigitHeight%
Public LeftDigitX!, LeftDigitY!
Public RightDigitX!, RightDigitY!
Public TopDigitX!, TopDigitY!
Public BottomDigitX!, BottomDigitY!

'These New_ variables are used to
'draw the box around the squares board.
Public NewLineX As Single, NewLineY As Single
Public NewLineWidth As Single, NewLineHeight As Single

'Maximum digit is 9.
'Minimum, of course, is 1.
Public MaxDigit As Byte

Public bGameStart As Boolean
Public bSolving As Boolean
Public SoundPath As String
