VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGameOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game options"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   2040
      Picture         =   "frmOptions.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit without validating changes"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   855
      Left            =   120
      Picture         =   "frmOptions.frx":114C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Validate changes and exit"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Game options"
      Height          =   4095
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Choose who receives the ball when a point is scored"
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox chkPlayer2 
         Caption         =   "Player 2 is computer"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Choose whether player 2 is human or the computer"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkPlayer1 
         Caption         =   "Player 1 is human"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Choose whether player 1 is human or the computer"
         Top             =   240
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkBall 
         Caption         =   "Winner receives the ball"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Choose who receives the ball when a point is scored"
         Top             =   3720
         Width           =   3255
      End
      Begin MSComctlLib.Slider sldPoints 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Points needed to win the game"
         Top             =   2880
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   30
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CheckBox chkStart 
         Caption         =   "Player 1 receives the ball first"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Choose who receives the ball first when the game begins"
         Top             =   3360
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin MSComctlLib.Slider sldSpeed1 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "1=Dummy ... 15=Expert"
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   15
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider sldSpeed2 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "1=Dummy ... 15=Expert"
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   15
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblSpeed2 
         Caption         =   "Player 2 speed:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblPoints 
         Caption         =   "Points needed to win the game:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label lblSpeed1 
         Caption         =   "Player 1 speed:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmGameOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   If HumanPlayer1 = True Then
      chkPlayer1.Value = 1
      chkPlayer1.Caption = "Player 1 is human"
   Else
      chkPlayer1.Value = 0
      chkPlayer1.Caption = "Player 1 is computer"
   End If
   If HumanPlayer2 = True Then
      chkPlayer2.Value = 1
      chkPlayer2.Caption = "Player 2 is human"
   Else
      chkPlayer2.Value = 0
      chkPlayer2.Caption = "Player 2 is computer"
   End If
   sldSpeed1.Value = Speed1
   sldSpeed2.Value = Speed2
   sldPoints.Value = WinPoints
   If NewBall = True Then
      chkBall.Value = 4
      chkBall.Caption = "Winner receives the ball"
   Else
      chkBall.Value = 0
      chkBall.Caption = "Loser receives the ball"
   End If
   If StartDir = True Then
      chkStart.Value = 1
      chkStart.Caption = "Player 1 receives the ball first"
   Else
      chkStart.Value = 0
      chkStart.Caption = "Player 1 receives the ball first"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   Set frmGameOptions = Nothing
End Sub

Private Sub chkPlayer1_Click()
   If chkPlayer1.Value = 1 Then
      chkPlayer1.Caption = "Player 1 is human"
   Else
      chkPlayer1.Caption = "Player 1 is computer"
   End If
End Sub

Private Sub chkPlayer2_Click()
   If chkPlayer2.Value = 1 Then
      chkPlayer2.Caption = "Player 2 is human"
   Else
      chkPlayer2.Caption = "Player 2 is computer"
   End If
End Sub

Private Sub chkBall_Click()
   If chkBall.Value = 1 Then
      chkBall.Caption = "Winner receives the ball"
   Else
      chkBall.Caption = "Loser receives the ball"
   End If
End Sub

Private Sub chkStart_Click()
   If chkStart.Value = 1 Then
      chkStart.Caption = "Player 1 receives the ball first"
   Else
      chkStart.Caption = "Player 2 receives the ball first"
   End If
End Sub

Private Sub cmdOK_Click()
   If chkPlayer1.Value = 1 Then
      HumanPlayer1 = True
   Else
      HumanPlayer1 = False
   End If
   If chkPlayer2.Value = 1 Then
      HumanPlayer2 = True
   Else
      HumanPlayer2 = False
   End If
   Speed1 = sldSpeed1.Value
   Speed2 = sldSpeed2.Value
   WinPoints = sldPoints.Value
   If chkBall.Value = 1 Then
      NewBall = True
   Else
      NewBall = False
   End If
   If chkStart.Value = 1 Then
      StartDir = True
   Else
      StartDir = False
   End If
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub
