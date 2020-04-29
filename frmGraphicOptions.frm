VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGraphicOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graphic options"
   ClientHeight    =   4575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3990
   Icon            =   "frmGraphicOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   2040
      Picture         =   "frmGraphicOptions.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit without validating changes"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   855
      Left            =   120
      Picture         =   "frmGraphicOptions.frx":114C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Validate changes and exit"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Game options"
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox chkBallShape 
         Caption         =   "Square ball"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Choose whther the ball will be rounded or squared"
         Top             =   2640
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin MSComctlLib.Slider sldRacket1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "1=Tiny ... 15=Extra large"
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   15
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider sldRacket2 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "1=Tiny ... 15=Extra large"
         Top             =   1320
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   15
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider sldBallSize 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "1=Tiny ... 10=Extra large"
         Top             =   2040
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblBallSize 
         Caption         =   "Ball size:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblRacket2 
         Caption         =   "Player 2 racket size:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblRacket1 
         Caption         =   "Player 1 racket size:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmGraphicOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   sldRacket1.Value = Racket1Size
   sldRacket2.Value = Racket2Size
   sldBallSize.Value = BallSize
   If BallShape = True Then
      chkBallShape.Value = 1
      chkBallShape.Caption = "Square ball"
   Else
      chkBallShape.Value = 0
      chkBallShape.Caption = "Round ball"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   Set frmGraphicOptions = Nothing
End Sub

Private Sub chkBallShape_Click()
   If chkBallShape.Value = 1 Then
      chkBallShape.Caption = "Square ball"
   Else
      chkBallShape.Caption = "Round ball"
   End If
End Sub

Private Sub cmdOK_Click()
   Racket1Size = sldRacket1.Value
   Racket2Size = sldRacket2.Value
   BallSize = sldBallSize.Value
   If chkBallShape.Value = 1 Then
      BallShape = True
   Else
      BallShape = False
   End If
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub
