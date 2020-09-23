VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "George and Al's Florida Election Adventure!"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10980
   Icon            =   "FRMMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FRMMAIN.frx":08CA
   ScaleHeight     =   7875
   ScaleWidth      =   10980
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   7200
   End
   Begin VB.Shape shpRectangle 
      BackColor       =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblScore3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   17
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblScore5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblScore4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lblInitials3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label lblInitials4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lblInitials5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblScore2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lblInitials2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label lblScore1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblInitials1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblInitials 
      Alignment       =   2  'Center
      Caption         =   "Intitials"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblHighScores 
      Alignment       =   2  'Center
      Caption         =   "High Scores:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      Caption         =   "Time Remaining:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblCountdown 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   9
      Left            =   1680
      Picture         =   "FRMMAIN.frx":499E4
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   8
      Left            =   1200
      Picture         =   "FRMMAIN.frx":4A1A6
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   7
      Left            =   840
      Picture         =   "FRMMAIN.frx":4A968
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   6
      Left            =   2160
      Picture         =   "FRMMAIN.frx":4B12A
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   5
      Left            =   2640
      Picture         =   "FRMMAIN.frx":4B8EC
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   4
      Left            =   3120
      Picture         =   "FRMMAIN.frx":4C0AE
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   3
      Left            =   3600
      Picture         =   "FRMMAIN.frx":4C870
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   2
      Left            =   4080
      Picture         =   "FRMMAIN.frx":4D032
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   1
      Left            =   4560
      Picture         =   "FRMMAIN.frx":4D7F4
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgRBallot 
      Height          =   480
      Index           =   0
      Left            =   5160
      Picture         =   "FRMMAIN.frx":4DFB6
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   10
      Left            =   840
      Picture         =   "FRMMAIN.frx":4E778
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   9
      Left            =   1320
      Picture         =   "FRMMAIN.frx":4EF3A
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   8
      Left            =   1800
      Picture         =   "FRMMAIN.frx":4F6FC
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   7
      Left            =   2280
      Picture         =   "FRMMAIN.frx":4FEBE
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   6
      Left            =   2760
      Picture         =   "FRMMAIN.frx":50680
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   5
      Left            =   3240
      Picture         =   "FRMMAIN.frx":50E42
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   4
      Left            =   3720
      Picture         =   "FRMMAIN.frx":51604
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   3
      Left            =   4200
      Picture         =   "FRMMAIN.frx":51DC6
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   2
      Left            =   4680
      Picture         =   "FRMMAIN.frx":52588
      Top             =   1320
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   1
      Left            =   5040
      Picture         =   "FRMMAIN.frx":52D4A
      Top             =   1320
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDBallot 
      Height          =   480
      Index           =   0
      Left            =   960
      Picture         =   "FRMMAIN.frx":5350C
      Top             =   1920
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image PlayerImage 
      Height          =   855
      Left            =   7200
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgGoreLeft 
      Height          =   990
      Left            =   3360
      Picture         =   "FRMMAIN.frx":53CCE
      Top             =   3240
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imgBushLeft 
      Height          =   960
      Left            =   2760
      Picture         =   "FRMMAIN.frx":54A58
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGoreRight 
      Height          =   990
      Left            =   7560
      Picture         =   "FRMMAIN.frx":5569A
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imgBushRight 
      Height          =   960
      Left            =   7320
      Picture         =   "FRMMAIN.frx":56424
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblUScore 
      Caption         =   "Votes:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "George and Al's Florida Election Adventure!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuGameHigh 
         Caption         =   "High &Scores"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intPlayerFace As Integer 'direction of image
Private intP As Integer
Private intHiScore As Integer 'current high score
Private intN As Integer 'stores values of multiple images - ballots
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private strPlaySound As String
'variable and function above used to play wav sound files using windows mulitmedia libraries

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'this is the primary sub of the whole app
'it moves the players image across the screen and detects collisions between objects
If blnGameOn = False Then Exit Sub
Select Case KeyCode
    Case vbKeyUp 'did the user press the up arrow?
      PlayerImage.Top = PlayerImage.Top - 100 'moves image
        If PlayerImage.Top <= 0 Then 'detects boundaries
        PlayerImage.Top = PlayerImage.Top + 100
        End If
        intN = 0
        For intN = 0 To 9 'collision detection with ballot images
        If PlayerImage.Left + PlayerImage.Width > imgDBallot(intN).Left Then
          If PlayerImage.Left < imgDBallot(intN).Left + imgDBallot(intN).Width Then
            If PlayerImage.Top < imgDBallot(intN).Top + imgDBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgDBallot(intN).Top Then
              Scoring 'goes to scoring subroutine
              imgDBallot(intN).Visible = False
              imgDBallot(intN).Left = Int(Rnd * 7000) 'randomly relocates ballot
              imgDBallot(intN).Top = Int(Rnd * 7000)
              imgDBallot(intN).Visible = True
             End If
            End If
          End If
         End If
        Next intN
        intN = 0
        For intN = 0 To 9
        If PlayerImage.Left + PlayerImage.Width > imgRBallot(intN).Left Then
          If PlayerImage.Left < imgRBallot(intN).Left + imgRBallot(intN).Width Then
            If PlayerImage.Top < imgRBallot(intN).Top + imgRBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgRBallot(intN).Top Then
              Scoring_Other
              imgRBallot(intN).Visible = False
              imgRBallot(intN).Left = Int(Rnd * 7000)
              imgRBallot(intN).Top = Int(Rnd * 7000)
              imgRBallot(intN).Visible = True
             End If
            End If
          End If
         End If
        Next intN
    Case vbKeyDown 'did the user press the down arrow?
      PlayerImage.Top = PlayerImage.Top + 100
        If PlayerImage.Top >= 6950 Then
        PlayerImage.Top = PlayerImage.Top - 100
        End If
      intN = 0
      For intN = 0 To 9
      If PlayerImage.Left + PlayerImage.Width > imgDBallot(intN).Left Then
          If PlayerImage.Left < imgDBallot(intN).Left + imgDBallot(intN).Width Then
            If PlayerImage.Top < imgDBallot(intN).Top + imgDBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgDBallot(intN).Top Then
              Scoring
              imgDBallot(intN).Visible = False
              imgDBallot(intN).Left = Int(Rnd * 7000)
              imgDBallot(intN).Top = Int(Rnd * 7000)
              imgDBallot(intN).Visible = True
             End If
            End If
          End If
         End If
       Next intN
       intN = 0
        For intN = 0 To 9
        If PlayerImage.Left + PlayerImage.Width > imgRBallot(intN).Left Then
          If PlayerImage.Left < imgRBallot(intN).Left + imgRBallot(intN).Width Then
            If PlayerImage.Top < imgRBallot(intN).Top + imgRBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgRBallot(intN).Top Then
              Scoring_Other
              imgRBallot(intN).Visible = False
              imgRBallot(intN).Left = Int(Rnd * 7000)
              imgRBallot(intN).Top = Int(Rnd * 7000)
              imgRBallot(intN).Visible = True
             End If
            End If
          End If
         End If
        Next intN
    Case vbKeyLeft 'same for left arrow
      intPlayerFace = 2 'switches images to point left
      PlayerPicture 'goes to sub that displays image
      PlayerImage.Left = PlayerImage.Left - 100
        If PlayerImage.Left <= 0 Then
        PlayerImage.Left = PlayerImage.Left + 100
        End If
     intN = 0
     For intN = 0 To 9
     If PlayerImage.Left + PlayerImage.Width > imgDBallot(intN).Left Then
          If PlayerImage.Left < imgDBallot(intN).Left + imgDBallot(intN).Width Then
            If PlayerImage.Top < imgDBallot(intN).Top + imgDBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgDBallot(intN).Top Then
              Scoring
              imgDBallot(intN).Visible = False
              imgDBallot(intN).Left = Int(Rnd * 7000)
              imgDBallot(intN).Top = Int(Rnd * 7000)
              imgDBallot(intN).Visible = True
             End If
            End If
          End If
         End If
       Next intN
       intN = 0
        For intN = 0 To 9
        If PlayerImage.Left + PlayerImage.Width > imgRBallot(intN).Left Then
          If PlayerImage.Left < imgRBallot(intN).Left + imgRBallot(intN).Width Then
            If PlayerImage.Top < imgRBallot(intN).Top + imgRBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgRBallot(intN).Top Then
              Scoring_Other
              imgRBallot(intN).Visible = False
              imgRBallot(intN).Left = Int(Rnd * 7000)
              imgRBallot(intN).Top = Int(Rnd * 7000)
              imgRBallot(intN).Visible = True
             End If
            End If
          End If
         End If
        Next intN
    Case vbKeyRight 'and right arrow
      intPlayerFace = 1 'switches images to point right
      PlayerPicture 'goes to sub that displays image
      PlayerImage.Left = PlayerImage.Left + 100
        If PlayerImage.Left >= 8025 Then
        PlayerImage.Left = PlayerImage.Left - 100
        End If
      intN = 0
      For intN = 0 To 9
      If PlayerImage.Left + PlayerImage.Width > imgDBallot(intN).Left Then
          If PlayerImage.Left < imgDBallot(intN).Left + imgDBallot(intN).Width Then
            If PlayerImage.Top < imgDBallot(intN).Top + imgDBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgDBallot(intN).Top Then
              Scoring
              imgDBallot(intN).Visible = False
              imgDBallot(intN).Left = Int(Rnd * 7000)
              imgDBallot(intN).Top = Int(Rnd * 7000)
              imgDBallot(intN).Visible = True
             End If
            End If
          End If
         End If
       Next intN
       intN = 0
        For intN = 0 To 9
        If PlayerImage.Left + PlayerImage.Width > imgRBallot(intN).Left Then
          If PlayerImage.Left < imgRBallot(intN).Left + imgRBallot(intN).Width Then
            If PlayerImage.Top < imgRBallot(intN).Top + imgRBallot(intN).Height Then
             If PlayerImage.Top + PlayerImage.Height > imgRBallot(intN).Top Then
              Scoring_Other
              imgRBallot(intN).Visible = False
              imgRBallot(intN).Left = Int(Rnd * 7000)
              imgRBallot(intN).Top = Int(Rnd * 7000)
              imgRBallot(intN).Visible = True
             End If
            End If
          End If
         End If
        Next intN
    End Select
End Sub

Private Sub Form_Load()
  'on loading of main form, plays hail to the chief
  strMidi = "2hail.mid"
  strSongCmd = "play " + App.Path + "\" + strMidi
  strPlaySong = mciSendString(strSongCmd, 0&, 0, 0&)
  'Sets up initial high scores and displays them
  HiScore(1).strInitials = "AAA"
  lblInitials1.Caption = HiScore(1).strInitials
  HiScore(1).intScore = 500
  lblScore1.Caption = HiScore(1).intScore
  HiScore(2).strInitials = "BBB"
  lblInitials2.Caption = HiScore(2).strInitials
  HiScore(2).intScore = 400
  lblScore2.Caption = HiScore(2).intScore
  HiScore(3).strInitials = "CCC"
  lblInitials3.Caption = HiScore(3).strInitials
  HiScore(3).intScore = 300
  lblScore3.Caption = HiScore(3).intScore
  HiScore(4).strInitials = "DDD"
  lblInitials4.Caption = HiScore(4).strInitials
  HiScore(4).intScore = 200
  lblScore4.Caption = HiScore(4).intScore
  HiScore(5).strInitials = "EEE"
  lblInitials5.Caption = HiScore(5).strInitials
  HiScore(5).intScore = 100
  lblScore5.Caption = HiScore(5).intScore
End Sub

Private Sub Form_Terminate()
  'makes sure we stop playing a song when we quit
  strSongCmd = "stop " + App.Path + "\" + strMidi
  strPlaySong = mciSendString(strSongCmd, 0&, 0, 0&)
End Sub


Public Sub PlayerPicture()
  'switches between right and left facing images depending
  'on whether or not user selected Gore or Bush
  If intPlayerFace = 1 And intPlayerChoice = 0 Then
  PlayerImage = ImgGoreRight.Picture
    ElseIf intPlayerFace = 2 And intPlayerChoice = 0 Then
    PlayerImage = imgGoreLeft.Picture
    ElseIf intPlayerFace = 1 And intPlayerChoice = 1 Then
    PlayerImage = imgBushRight.Picture
    ElseIf intPlayerFace = 2 And intPlayerChoice = 1 Then
    PlayerImage = imgBushLeft.Picture
  End If
End Sub

Public Sub Initialize()
  'stops hail to the chief
  strSongCmd = "stop " + App.Path + "\" + strMidi
  strPlaySong = mciSendString(strSongCmd, 0&, 0, 0&)
  intPlayerFace = 1
  lblCountdown.Caption = 45
  lblScore.Caption = 0
  intN = 0
  blnGameOn = True
  strMidi = "lowrider.mid" 'new in-game song
  strSongCmd = "play " + App.Path + "\" + strMidi
  strPlaySong = mciSendString(strSongCmd, 0&, 0, 0&)
  'next if then routine picks image depending on player selected
  'on select form
  If intPlayerChoice = 0 Then
    PlayerImage.Picture = ImgGoreRight.Picture
    ElseIf intPlayerChoice = 1 Then
    PlayerImage.Picture = imgBushRight.Picture
    End If
  PlayerImage.Visible = True
  'randomizes 10 democrat ballots and 10 republican ballots
  For intN = 0 To 9
    imgDBallot(intN).Visible = True
    imgDBallot(intN).Left = Int(Rnd * 7000)
    imgDBallot(intN).Top = Int(Rnd * 7000)
  Next intN
  For intN = 0 To 9
    imgRBallot(intN).Visible = True
    imgRBallot(intN).Left = Int(Rnd * 7000)
    imgRBallot(intN).Top = Int(Rnd * 7000)
  Next intN
  tmrCountdown.Enabled = True
End Sub

Public Sub Scoring()
 'scoring sub - plays sounds and scores according to
 'Gore player selection
 If intPlayerChoice = 0 Then
   strPlaySound = sndPlaySound("drip.wav", 1)
   lblScore.Caption = lblScore.Caption + 10
 Else
   strPlaySound = sndPlaySound("DOH!.wav", 1)
   lblScore.Caption = lblScore.Caption - 10
 End If
End Sub

Public Sub Scoring_Other()
 'scoring sub - plays sounds and scores according to
 'Bush player selection
If intPlayerChoice = 1 Then
   strPlaySound = sndPlaySound("drip.wav", 1)
   lblScore.Caption = lblScore.Caption + 10
 Else
   strPlaySound = sndPlaySound("DOH!.wav", 1)
   lblScore.Caption = lblScore.Caption - 10
 End If
End Sub

Private Sub mnuAbout_Click()
  Load frmAbout
  frmAbout.Show
End Sub

Private Sub mnuGameExit_Click()
  'Exits game and form
  strSongCmd = "stop " + App.Path + "\" + strMidi
  strPlaySong = mciSendString(strSongCmd, 0&, 0, 0&)
  Unload Me
End Sub

Private Sub mnuGameHigh_Click()
  For intP = 1 To 5
    Print HiScore(intP).strInitials, HiScore(intP).intScore
  Next intP
End Sub

Private Sub mnuGameNew_Click()
  Load frmSelect
  frmSelect.Show
End Sub

Private Sub mnuHelp_Click()
  Load frmHelp
  frmHelp.Show
End Sub

Private Sub tmrCountdown_Timer()
  'Timer sub for game countdown
  lblCountdown.Caption = Val(lblCountdown) - 1
  If Val(lblCountdown) <= 0 Then
    EndGame
  End If
End Sub

Public Sub EndGame()
tmrCountdown.Enabled = False
blnGameOn = False
strSongCmd = "stop " + App.Path + "\" + strMidi
  strPlaySong = mciSendString(strSongCmd, 0&, 0, 0&)
'Did you get a high score? If so - capture and display the High Score form
If lblScore.Caption >= HiScore(5).intScore Then
  HiScore(5).intScore = lblScore.Caption
  Load frmHighScore
  frmHighScore.Show
  Else: MsgBox "Game Over.  Try Again.", , "Time's Up!"
End If
End Sub

Public Sub UpdateHighScores()
  'Just updates all the labels with the new high scores
  lblInitials1.Caption = HiScore(1).strInitials
  lblScore1.Caption = HiScore(1).intScore
  lblInitials2.Caption = HiScore(2).strInitials
  lblScore2.Caption = HiScore(2).intScore
  lblInitials3.Caption = HiScore(3).strInitials
  lblScore3.Caption = HiScore(3).intScore
  lblInitials4.Caption = HiScore(4).strInitials
  lblScore4.Caption = HiScore(4).intScore
  lblInitials5.Caption = HiScore(5).strInitials
  lblScore5.Caption = HiScore(5).intScore
End Sub
