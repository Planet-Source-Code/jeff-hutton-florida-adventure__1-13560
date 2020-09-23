VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   4755
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7935
   ClipControls    =   0   'False
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3281.986
   ScaleMode       =   0  'User
   ScaleWidth      =   7451.375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmHelp.frx":08CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3360
      TabIndex        =   0
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Label lblStarting 
      Caption         =   "To start a game, click ""Game"", then ""New Game"".  Good luck!"
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   3480
      Width           =   6495
   End
   Begin VB.Label lblStartHeader 
      Caption         =   "Starting a Game:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   8
      Top             =   3240
      Width           =   6525
   End
   Begin VB.Label lblMovement 
      Caption         =   "Use the arrow keys to move around the game board.  Use the mouse to select game options and menu options."
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   2520
      Width           =   6495
   End
   Begin VB.Label lblMovementHeader 
      Caption         =   "Movement:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   6525
   End
   Begin VB.Label lblScoring 
      Caption         =   $"frmHelp.frx":1194
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   7324.603
      Y1              =   2898.915
      Y2              =   2898.915
   End
   Begin VB.Label lblScoringHeader 
      Caption         =   "Scoring:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   6525
   End
   Begin VB.Label lblObjectHead 
      Caption         =   "Object:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   4965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   7338.689
      Y1              =   2898.915
      Y2              =   2898.915
   End
   Begin VB.Label lblObject 
      Caption         =   $"frmHelp.frx":1225
      Height          =   705
      Left            =   1050
      TabIndex        =   4
      Top             =   480
      Width           =   6405
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub
