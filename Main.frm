VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Pong Ouin"
   ClientHeight    =   4968
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   6612
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBall 
      Interval        =   50
      Left            =   3720
      Top             =   480
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Ball 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3120
      Top             =   2400
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Line NetLine 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   216
      X2              =   216
      Y1              =   0
      Y2              =   328
   End
   Begin VB.Shape Player2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   6240
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Player1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New Game"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuGamePause 
         Caption         =   "&Pause game"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameAbort 
         Caption         =   "&Abort game"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuGameOptions 
         Caption         =   "&Game options"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGraphicOptions 
         Caption         =   "G&raphic options"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim PlayGame, BallOut, Direction As Boolean
Dim Yincrement As Integer
Dim Score(2) As Byte

'Keep track of keys pressed
Dim iKeys As Integer

Private Sub Form_Load()
   'Position computer's racket on X axis
   Player2.Left = Me.ScaleWidth - 2 * Player2.Width
   'Position computer's score
   lblScore(1).Left = Me.ScaleWidth - 8 * Player2.Width
   'Set default game options
   WinPoints = 15
   Speed1 = 10
   Speed2 = 10
   StartDir = True
   NewBall = False
   'Set default graphic options
   Racket1Size = 6
   Racket2Size = 6
   BallSize = 3
   BallShape = True
   HumanPlayer1 = True
   HumanPlayer2 = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   Set frmMain = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If PlayGame Then
        Select Case KeyCode
            Case 81: iKeys = (iKeys Or 1)          'Q
            Case 65: iKeys = (iKeys Or 2)          'A
            Case 113: iKeys = (iKeys Or 1)         'q
            Case 97: iKeys = (iKeys Or 2)          'a
            Case vbKeyUp: iKeys = (iKeys Or 4)     'UP
            Case vbKeyDown: iKeys = (iKeys Or 8)   'DOWN
        End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If PlayGame Then
        Select Case KeyCode
            Case 81: iKeys = (iKeys And Not 1)
            Case 65: iKeys = (iKeys And Not 2)
            Case 113: iKeys = (iKeys And Not 1)
            Case 97: iKeys = (iKeys And Not 2)
            Case vbKeyUp: iKeys = (iKeys And Not 4)
            Case vbKeyDown: iKeys = (iKeys And Not 8)
        End Select
    End If
End Sub

Private Sub Form_Resize()
   'Resize net line
   NetLine.X1 = Me.ScaleWidth / 2
   NetLine.X2 = NetLine.X1
   NetLine.Y1 = 0
   NetLine.Y2 = Me.ScaleHeight
   'Position computer's racket on X axis
   Player2.Left = Me.ScaleWidth - 2 * Player2.Width
   'Position computer's score
   lblScore(1).Left = Me.ScaleWidth - 8 * Player2.Width
   'Reposition rackets and ball if out of screen
   If (Player1.Top > Me.ScaleHeight - Player1.Height) Then Player1.Top = Me.ScaleHeight - Player1.Height
   If (Player2.Top > Me.ScaleHeight - Player2.Height) Then Player2.Top = Me.ScaleHeight - Player2.Height
   If (Ball.Top > Me.ScaleHeight - Ball.Height) Then Ball.Top = Me.ScaleHeight - Ball.Height
End Sub

Private Sub mnuGameNew_Click()
   'Set rackets and ball size
   Player1.Height = 10 * Racket1Size
   Player2.Height = 10 * Racket2Size
   Ball.Height = 5 * BallSize
   Ball.Width = 5 * BallSize
   If BallShape = True Then Ball.Shape = vbShapeSquare Else Ball.Shape = vbShapeCircle
   'Make  rackets and ball visible
   Player1.Visible = True
   Player2.Visible = True
   Ball.Visible = True
   'Initialise rackets and ball position
   Ball.Left = (Me.ScaleWidth - Ball.Width) / 2
   Ball.Top = (Me.ScaleHeight - Ball.Height) / 2
   Player1.Top = Me.ScaleHeight / 2 - Player1.Height / 2
   Player2.Top = Me.ScaleHeight / 2 - Player2.Height / 2
   PlayGame = True
   mnuGamePause.Enabled = True
   mnuGameAbort.Enabled = True
   Score(1) = 0
   Score(2) = 0
   lblScore(0).Caption = 0
   lblScore(1).Caption = 0
   Direction = StartDir
   BallOut = False
   Yincrement = Int(10 * Rnd) + 1
End Sub

Private Sub mnuGamePause_Click()
   Call PauseGame
End Sub

Private Sub mnuGameAbort_Click()
   Call GameOver
End Sub

Private Sub mnuGameExit_Click()
   Unload Me
End Sub

Private Sub mnuGameOptions_Click()
   If PlayGame Then
      Call PauseGame
      frmGameOptions.Show vbModal
      Call PauseGame
   Else
      frmGameOptions.Show vbModal
   End If
End Sub

Private Sub mnuGraphicOptions_Click()
   If PlayGame Then
      Call PauseGame
      frmGraphicOptions.Show vbModal
      Call PauseGame
   Else
      frmGraphicOptions.Show vbModal
   End If
   'Set rackets and ball size
   Player1.Height = 10 * Racket1Size
   Player2.Height = 10 * Racket2Size
   Ball.Height = 5 * BallSize
   Ball.Width = 5 * BallSize
   If BallShape = True Then Ball.Shape = vbShapeSquare Else Ball.Shape = vbShapeCircle
End Sub

Private Sub tmrBall_Timer()
   'Game paused? Exit
   If Not PlayGame Then Exit Sub
   'Check player 1 keys
   If HumanPlayer1 Then
      'Player 1 up key = Q
      If (iKeys And 1) = 1 Then
         If Player1.Top >= 10 Then Player1.Top = Player1.Top - 10
      End If
      'Player 1 down key = A
      If (iKeys And 2) = 2 Then
         If Player1.Top <= (Me.ScaleHeight - Player1.Height - 10) Then Player1.Top = Player1.Top + 10
      End If
   End If
   'Check player 2 keys
   If HumanPlayer2 Then
      'Player 2 up key = UP
      If (iKeys And 4) = 4 Then
         If Player2.Top >= 10 Then Player2.Top = Player2.Top - 10
      End If
      'Player 2 down key = DOWN
      If (iKeys And 8) = 8 Then
         If Player2.Top <= (Me.ScaleHeight - Player2.Height - 10) Then Player2.Top = Player2.Top + 10
      End If
   End If
   'Ball passing computer's racket?
   If (Ball.Left + Ball.Width) >= Player2.Left Then
      'Ball touches racket?
      If Ball.Top >= (Player2.Top - Ball.Height / 2) And Ball.Top <= (Player2.Top + Player2.Height + Ball.Height / 2) Then
         'Yes: invert ball direction
         Direction = True
         'Set ball deviation function of position of ball on racket
         Yincrement = ((Ball.Top + Ball.Height / 2) - (Player2.Top + Player2.Height / 2)) / 5
      Else
         'No: flag ball out
         BallOut = True
      End If
   End If
   'Ball passing player's racket?
   If Ball.Left < Player1.Left + Player1.Width Then
      'Ball touches racket?
      If Ball.Top > (Player1.Top - Ball.Height / 2) And Ball.Top <= (Player1.Top + Player1.Height + Ball.Height / 2) Then
         'Yes: invert ball direction
         Direction = False
         'Set ball deviation function of position of ball on racket
         Yincrement = ((Ball.Top + Ball.Height / 2) - (Player1.Top + Player1.Height / 2)) / 5
      Else
         'No: flag ball out
         BallOut = True
      End If
   End If
   'Computer moves its racket accordingly to ball position
   If Not HumanPlayer1 Then
      If Ball.Top < Player1.Top Then Player1.Top = Player1.Top - Speed1
      If Ball.Top > Player1.Top + Player1.Height Then Player1.Top = Player1.Top + Speed1
   End If
   If Not HumanPlayer2 Then
      If Ball.Top < Player2.Top Then Player2.Top = Player2.Top - Speed2
      If Ball.Top > Player2.Top + Player2.Height Then Player2.Top = Player2.Top + Speed2
   End If
   'Ball out?
   If BallOut Then
      If Ball.Left <= Me.ScaleHeight / 2 Then
         'Computer scores
         Score(2) = Score(2) + 1
         lblScore(1) = Score(2)
         Ball.Left = ScaleWidth / 2
         If NewBall Then
            Direction = False
         Else
            Direction = True
         End If
      Else
         'Player scores
         Score(1) = Score(1) + 1
         lblScore(0) = Score(1)
         Ball.Left = ScaleWidth / 2
         If NewBall Then
            Direction = True
         Else
            Direction = False
         End If
      End If
      BallOut = False
      If Score(1) >= WinPoints Or Score(2) >= WinPoints Then GameOver
   End If
   'Ball touches upper or lower wall: invert bounce
   If Ball.Top <= 0 Or Ball.Top >= Me.ScaleHeight - Ball.Height Then Yincrement = -Yincrement
   'Update ball Y position
   Ball.Top = Ball.Top + Yincrement
   'Update ball X position
   If Direction Then
      Ball.Left = Ball.Left - 10
   Else
      Ball.Left = Ball.Left + 10
   End If
End Sub

Private Sub GameOver()
   PlayGame = False
   mnuGamePause.Enabled = False
   mnuGameAbort.Enabled = False
   If Score(1) = WinPoints Then
      MsgBox "Player 1 wins!", vbExclamation, "Game over"
   Else
      MsgBox "Player 2 wins!", vbExclamation, "Game over"
   End If
   Player1.Visible = False
   Player2.Visible = False
   Ball.Visible = False
End Sub

Private Sub PauseGame()
   PlayGame = Not PlayGame
   If PlayGame Then
      mnuGamePause.Caption = "&Pause game"
   Else
      mnuGamePause.Caption = "&Resume game"
   End If
End Sub
