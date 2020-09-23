Attribute VB_Name = "Module1"
'These variables are used between forms so it must be
'in a module
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Global strMidi As String
Global strSongCmd As String
Global strPlaySong As String
'three variables and function above used to play midi files using windows dll
Global intPlayerChoice As Integer
Type HighScore
  strInitials As String
  intScore As Integer
 End Type
Global HiScore(1 To 5) As HighScore
Global blnGameOn As Boolean

