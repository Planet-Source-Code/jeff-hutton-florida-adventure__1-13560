VERSION 5.00
Begin VB.Form frmHighScore 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Score"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmHighScore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtInitials 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblHighScore 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Congratulations!!   You got a High Score!!  Enter your initials!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkay_Click()
  Dim blnHalt As Boolean
  If txtInitials.Text <> "" Then 'Text must not be blank
    blnHalt = False
    HiScore(5).strInitials = txtInitials.Text
    'Little routine to sort high scores
    For X = 1 To 4
        For Y = (X + 1) To 5
            If HiScore(X).intScore < HiScore(Y).intScore Then
                Dim PlaceHolder As HighScore
                PlaceHolder = HiScore(X)
                HiScore(X) = HiScore(Y)
                HiScore(Y) = PlaceHolder
            End If
        Next Y
    Next X
    Else
      'Whoops - gotta enter something
      MsgBox "You must enter your initials.", , "Enter Initials"
      blnHalt = True
      txtInitials.SetFocus
  End If
  If blnHalt = False Then
   FrmMain.UpdateHighScores 'calls sub to update high scores on main form
   Unload Me
  End If
End Sub
