VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Select"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   2760
      Picture         =   "frmSelect.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   240
      Picture         =   "frmSelect.frx":6BBE
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.OptionButton optBush 
      Caption         =   "George W. Bush"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.OptionButton optGore 
      Caption         =   "Al Gore"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      Caption         =   "Choose your Player, then click ""Start"""
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub cmdStart_Click()
  'Looks at status of option buttons
  If optGore = True Then
    intPlayerChoice = 0
    Unload Me
    FrmMain.Initialize 'starts game
   ElseIf optBush = True Then
    intPlayerChoice = 1
    Unload Me
    FrmMain.Initialize 'starts game
   Else: MsgBox "You must choose a player.  Try Again.", , "Player Selection"
  End If
  
End Sub

Private Sub Form_Load()
  optGore = False
  optBush = False
End Sub
