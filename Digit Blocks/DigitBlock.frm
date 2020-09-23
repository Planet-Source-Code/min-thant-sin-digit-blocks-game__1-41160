VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tetravex  (minsin999@hotmail.com)"
   ClientHeight    =   5700
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8745
   ForeColor       =   &H00000000&
   Icon            =   "DigitBlock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picWinBox 
      Height          =   1515
      Left            =   75
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   6
      Top             =   1125
      Visible         =   0   'False
      Width           =   1515
      Begin VB.Image imgSmile 
         Height          =   480
         Left            =   450
         Picture         =   "DigitBlock.frx":08CA
         Top             =   525
         Width           =   480
      End
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   3600
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   5
      Top             =   150
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picFixedBlock 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   1725
      Picture         =   "DigitBlock.frx":0D0C
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox pctBlock 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   900
      Picture         =   "DigitBlock.frx":1BAA
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   3
      Top             =   150
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picSquare 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   1
      Left            =   2550
      Picture         =   "DigitBlock.frx":2A48
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   2
      Top             =   150
      Width           =   765
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   1
      Left            =   75
      Picture         =   "DigitBlock.frx":38E6
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   0
      Top             =   150
      Width           =   765
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   4425
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "Sol&ve"
         Shortcut        =   ^S
      End
      Begin VB.Menu sepArrange 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuChooseSize 
         Caption         =   "&Size        "
         Begin VB.Menu mnuSize 
            Caption         =   "&2x2"
            Index           =   2
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&3x3"
            Index           =   3
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&4x4"
            Index           =   4
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&5x5"
            Index           =   5
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&6x6"
            Index           =   6
         End
      End
      Begin VB.Menu sepOptions 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChooseDigits 
         Caption         =   "&Digits"
         Begin VB.Menu mnuDigit 
            Caption         =   "&6"
            Index           =   6
         End
         Begin VB.Menu mnuDigit 
            Caption         =   "&7"
            Index           =   7
         End
         Begin VB.Menu mnuDigit 
            Caption         =   "&8"
            Index           =   8
         End
         Begin VB.Menu mnuDigit 
            Caption         =   "&9"
            Index           =   9
         End
         Begin VB.Menu mnuDigit 
            Caption         =   "&10"
            Index           =   10
         End
      End
   End
   Begin VB.Menu mnuHelpRelated 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Tetravex..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************
'Thanks to those people for their useful source codes that I
'use in some parts of this program.
'************************************************************

Dim tmpLeft As Single, tmpTop As Single
Dim OldX As Single, OldY As Single
Dim OldSquare As Integer

Dim OldDigitIndex As Byte
Dim OldSizeIndex As Byte

Dim bCanMove As Boolean
Dim bFromSquare As Boolean

Private Sub Form_Load()
      'Range from 1 to 9 inclusive.
      mnuDigit(9).Checked = True
      OldDigitIndex = mnuDigit(9).Index
      
      'Size is 3x3.
      mnuSize(3).Checked = True
      OldSizeIndex = mnuSize(3).Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
      Dim X As Integer
      'Clean up your lousy mess.
      For X = 2 To PuzzleSize ^ 2
            Unload picBlock(X)
            Unload picSquare(X)
      Next X
End Sub

Private Sub imgSmile_Click()
      Call StartNewGame
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Written with Visual Basic 6.0", vbInformation
End Sub

Private Sub mnuExit_Click()
      Unload Me
End Sub

Private Sub mnuDigit_Click(Index As Integer)
      mnuDigit(OldDigitIndex).Checked = False
      mnuDigit(Index).Checked = True
      OldDigitIndex = Index
      
      MaxDigit = Index
      Call StartNewGame
End Sub

Private Sub mnuNewGame_Click()
      Call StartNewGame
End Sub

Sub ShowAllBlocks(bShow As Boolean)
      Dim X As Integer
      For X = 1 To PuzzleSize ^ 2
            picBlock(X).Visible = bShow
      Next X
End Sub

Private Sub mnuSize_Click(Index As Integer)
      mnuSize(OldSizeIndex).Checked = False
      mnuSize(Index).Checked = True
      OldSizeIndex = Index
      
      PuzzleSize = Index
      Call StartNewGame
End Sub

Private Sub mnuSolve_Click()
      Dim X%, Y%
      Dim Delay#
      
      If Not bGameStart Then Exit Sub
      
      bSolving = True
      bGameStart = False
      Call PrintDigits(PuzzleSize)
      
      For X = 1 To PuzzleSize ^ 2
            Delay = Timer
            picBlock(X).ZOrder 0
            picBlock(X).Move picSquare(X).Left, picSquare(X).Top
            Square(X).HasBlock = True
            Square(X).CurBlock = Block(X).Index
            Do Until Timer > Delay + 0.01
                  DoEvents
            Loop
            DoEvents
      Next X
      
      mnuSolve.Enabled = False
End Sub

Sub StartNewGame()
      bSolving = False
      bGameStart = False
      mnuSolve.Enabled = True
      
      Call InitPuzzle(PuzzleSize)
      Call ArrangeDigits(PuzzleSize, MaxDigit)
      Call PrintDigits(PuzzleSize)
End Sub

Private Sub picBlock_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      If (Button <> vbLeftButton And Button <> vbRightButton) Then Exit Sub
      
      If bGameStart = False Then Exit Sub
      
      Dim I As Integer
                  
      bCanMove = False
      'Call ShowCursor(False)
      
      bFromSquare = False
      For I = 1 To PuzzleSize ^ 2
            If Overlap(picSquare(I), picBlock(Index)) Then
                  OldSquare = picSquare(I).Index
                  Square(OldSquare).CurBlock = Block(Index).Index
                  Square(OldSquare).HasBlock = False
                  bFromSquare = True
                  Exit For
            End If
      Next I
      
      OldX = X + 2
      OldY = Y + 2
            
      picShadow.Visible = True
      picShadow.ZOrder 0
      picBlock(Index).ZOrder 0
      
      With picBlock(Index)
            Block(Index).HomeX = .Left
            Block(Index).HomeY = .Top
      End With
      
      'ShowCursor False
      bCanMove = True
End Sub

Private Sub picBlock_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      'If Button <> 1 And Button <> 2 Then Exit Sub
      If (Button <> vbLeftButton And Button <> vbRightButton) Then Exit Sub
      If bGameStart = False Then Exit Sub
      
      If bCanMove Then
            With picBlock(Index)
                  tmpLeft = (X + picBlock(Index).Left) - OldX
                  tmpTop = (Y + picBlock(Index).Top) - OldY
                  If tmpLeft < 0 Then tmpLeft = 0
                  If tmpLeft > (Me.ScaleWidth - picBlock(Index).ScaleWidth) Then tmpLeft = (Me.ScaleWidth - picBlock(Index).ScaleWidth)
                  If tmpTop < 0 Then tmpTop = 0
                  If tmpTop > (Me.ScaleHeight - picBlock(Index).ScaleHeight) Then tmpTop = (Me.ScaleHeight - picBlock(Index).ScaleHeight)
                  
                  .Left = tmpLeft
                  .Top = tmpTop
                  picShadow.Move .Left + 2, .Top + 2
            End With
      End If
End Sub

Private Sub picBlock_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      If (Button <> vbLeftButton And Button <> vbRightButton) Then Exit Sub
      
      If bGameStart = False Then Exit Sub
      
      Dim bInSquare As Boolean
      Dim I As Integer, j As Integer
      Static DestSquare, CurrentBlock As Integer
      Dim bMoveItNow As Boolean
      Dim bTouch As Boolean
      
      'ShowCursor True
      
      bInSquare = False
      For I = 1 To PuzzleSize ^ 2
            'If the block is over one of the squares...
            If Overlap(picSquare(I), picBlock(Index)) Then
                  'Check this square has already a block or not...
                  'If this square has a block,
                  'we can't move the block to that square.
                  If Square(I).HasBlock Then
                        'Check if the block has been moved from
                        'one of the squares...
                        'If so...
                        If bFromSquare Then
                              'Move it back to its original square
                              picBlock(Index).Move picSquare(OldSquare).Left, picSquare(OldSquare).Top
                              'picBlock(Index).Move block(index).HomeX,block(index).HomeY
                              
                              'Set the square HasBlock property to True.
                              Square(OldSquare).HasBlock = True
                              Square(OldSquare).CurBlock = Block(Index).Index
                        'Otherwise..
                        Else
                              'Just move it back to its original place
                              picBlock(Index).Move Block(Index).HomeX, Block(Index).HomeY
                        End If
                              bCanMove = False
                              picShadow.Visible = False
                              Exit Sub
                              
                  'This square has no block.
                  Else
                        'This block is over a blank square.
                        'Set the indicator to true.
                        bInSquare = True
                        'And mark the destination square.
                        DestSquare = picSquare(I).Index
                        'No need to check the other squares.
                        Exit For
                  End If
                  
            End If
      Next I
      
      bTouch = False
      For I = 1 To PuzzleSize ^ 2
            If Collided(picSquare(I), picBlock(Index)) Then
                  bTouch = True
                  Exit For
            End If
      Next I
      
      If bInSquare Then
            bMoveItNow = True
            For I = 1 To PuzzleSize ^ 2
                  If Square(I).HasBlock Then
                        CurrentBlock = Square(I).CurBlock
                                                
                        'Determine if they have the same row or
                        'column with the destination square.
                        Select Case _
                        Abs(Square(DestSquare).Index - Square(I).Index)
                        Case 1
                              If Square(DestSquare).row = Square(I).row Then
                                    If Square(DestSquare).Index > Square(I).Index Then
                                          If Block(Index).LeftDigit <> Block(CurrentBlock).RightDigit Then
                                                bMoveItNow = False
                                          End If
                                    Else
                                          If Block(Index).RightDigit <> Block(CurrentBlock).LeftDigit Then
                                                bMoveItNow = False
                                          End If
                                    End If
                              End If
                        Case PuzzleSize
                              If Square(DestSquare).col = Square(I).col Then
                                    If Square(DestSquare).Index > Square(I).Index Then
                                          If Block(Index).TopDigit <> Block(CurrentBlock).BottomDigit Then
                                                bMoveItNow = False
                                          End If
                                    Else
                                          If Block(Index).BottomDigit <> Block(CurrentBlock).TopDigit Then
                                                bMoveItNow = False
                                          End If
                                    End If
                              End If
                        End Select
                  End If
            Next I
            
            If bMoveItNow Then
                  picBlock(Index).Move picSquare(DestSquare).Left, picSquare(DestSquare).Top
                  Block(Index).CurSquare = Square(DestSquare).Index
                  Square(DestSquare).CurBlock = Block(Index).Index
                  Square(DestSquare).HasBlock = True
            Else
                  If bFromSquare Then
                        Square(OldSquare).HasBlock = True
                        Square(OldSquare).CurBlock = Block(Index).Index
                        picBlock(Index).Move Block(Index).HomeX, Block(Index).HomeY
                  Else
                        picBlock(Index).Move Block(Index).HomeX, Block(Index).HomeY
                  End If
                  
            End If
                  
      Else
            If bTouch Then
                  If bFromSquare Then
                        Square(OldSquare).HasBlock = True
                        Square(OldSquare).CurBlock = Block(Index).Index
                        picBlock(Index).Move Block(Index).HomeX, Block(Index).HomeY
                  Else
                        picBlock(Index).Move Block(Index).HomeX, Block(Index).HomeY
                  End If
            Else
                  picBlock(Index).Move picShadow.Left, picShadow.Top
            End If
            
      End If
      
      bCanMove = False
      picShadow.Visible = False
      
      If (WinGame And bGameStart) Then
            picWinBox.Visible = True
            picWinBox.ZOrder 0
            bGameStart = False
            mnuSolve.Enabled = False
      End If
End Sub

Function WinGame() As Boolean
      Dim X%, Y%
      
      WinGame = True
      For X = 1 To PuzzleSize ^ 2
            If Square(X).HasBlock = False Then
                  WinGame = False
                  Exit For
            End If
      Next X
End Function

Private Sub picWinBox_Click()
      Call StartNewGame
End Sub
