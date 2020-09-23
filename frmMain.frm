VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Nuke - The Best MP3 Player Of Them All"
   ClientHeight    =   4350
   ClientLeft      =   1365
   ClientTop       =   1275
   ClientWidth     =   6555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6555
   Begin VB.PictureBox picSkin 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   0
      Width           =   6615
      Begin VB.Timer tmrScroll 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6120
         Top             =   1080
      End
      Begin VB.CommandButton cmdFileOpt 
         Caption         =   "File Options"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdListOpt 
         Caption         =   "List Options"
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ListBox playList 
         Height          =   1035
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   29
         Top             =   3120
         Width           =   6135
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<|"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "|>"
         Height          =   375
         Left            =   4560
         TabIndex        =   27
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   375
         Left            =   2160
         TabIndex        =   26
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   3360
         TabIndex        =   25
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Left            =   960
         TabIndex        =   24
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   5400
         TabIndex        =   23
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdID3 
         Caption         =   "ID3 Tag"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdOpt 
         Caption         =   "Options"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.HScrollBar scrollBalance 
         Height          =   255
         Left            =   5040
         Max             =   2
         TabIndex        =   6
         Top             =   1200
         Value           =   1
         Width           =   1095
      End
      Begin VB.HScrollBar scrollVolume 
         Height          =   255
         Left            =   1800
         Max             =   0
         Min             =   -4000
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.HScrollBar scrollPos 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   4
         Top             =   1680
         Width           =   6135
      End
      Begin VB.Label lblPlayList 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Play List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblSongInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label txtMode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblMode 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.Label txtFrequency 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblFrequency 
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblBiteRate 
         BackStyle       =   0  'Transparent
         Caption         =   "Biterate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label txtBiteRate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblBalanceLeft 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   15
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblBalanceRight 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblBalanceCenter 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   13
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblMaxVolume 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblMinVolume 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblDurationTime 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblCurrentTime 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   6840
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   6720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer scrollSetter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   720
   End
   Begin VB.Label lblArtist 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblSongName 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPlayback 
         Caption         =   "Playback"
         Begin VB.Menu mnuPlay 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuStop 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuNext 
            Caption         =   "Next"
         End
         Begin VB.Menu mnuPrevious 
            Caption         =   "Previous"
         End
      End
      Begin VB.Menu Break 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "Credits"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMisc 
      Caption         =   ""
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Play List"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Play List"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Files To The Playlist"
      End
      Begin VB.Menu mnuAddDir 
         Caption         =   "Add A Dir To The Playlist"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Files From The Playlist"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear The Playlist"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Str As String
Option Explicit
Private m_blnClose As Boolean
Private m_blnMin As Boolean
Private m_blnMax As Boolean
Dim ThisFile As String
Dim s As Integer
Dim dta As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub SavePlaylist(sFilename As String)
    Dim iFilenum As Integer
    Dim iCnt As Integer
    
    iFilenum = FreeFile
    
    Open sFilename For Output As #iFilenum
    
    For iCnt = 1 To playList.ListCount
        Print #iFilenum, playList.List(iCnt - 1)
    Next iCnt
    
    Close #iFilenum
End Sub


Private Sub MP3Info()
ThisFile = MediaPlayer1.FileName
 ReadMP3Header (ThisFile)
    With MP3HeaderInfo
    Dim Result
    txtBiterate.Caption = .Bitrate
    txtFrequency.Caption = .Frequency
    txtMode.Caption = .Mode
    End With
End Sub
Private Sub CurrentTime()
Dim sec As Integer
Dim Secs As Integer
Dim min As Single
Dim Ecsit As Boolean
Secs = MediaPlayer1.CurrentPosition
min = Secs / 60
min = Int(min)
sec = Secs - (min * 60)
If min > 9 Then
Ecsit = True
End If
If sec < 10 Then
If Ecsit = False Then
lblCurrentTime.Caption = "0" & min & ":0" & sec
Else
lblCurrentTime.Caption = min & ":0" & sec
End If
Else
If Ecsit = False Then
lblCurrentTime.Caption = "0" & min & ":" & sec
Else
lblCurrentTime.Caption = min & ":" & sec
End If
End If
End Sub
Private Sub DurationTime()
Dim sec As Integer
Dim Secs As Integer
Dim min As Single
Dim Ecsit As Boolean
Secs = MediaPlayer1.Duration
min = Secs / 60
min = Int(min)
sec = Secs - (min * 60)
If min > 9 Then
Ecsit = True
End If
If sec < 10 Then
If Ecsit = False Then
lblDurationTime.Caption = "0" & min & ":0" & sec
Else
lblDurationTime.Caption = min & ":0" & sec
End If
Else
If Ecsit = False Then
lblDurationTime.Caption = "0" & min & ":" & sec
Else
lblDurationTime.Caption = min & ":" & sec
End If
End If
End Sub
Private Sub cmdAdd_Click()
End Sub

Private Sub cmdClear_Click()
End Sub

Private Sub cmdFile_Click()

End Sub

Private Sub cmdFileOpt_Click()
Me.PopupMenu Me.mnuFile
End Sub

Private Sub cmdID3_Click()
frmID3Tag.Show
End Sub

Private Sub cmdMisc_Click()
End Sub

Private Sub cmdListOpt_Click()
Me.PopupMenu Me.mnuMisc
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err:
If playList.ListIndex >= (playList.ListCount - 1) Then
    playList.ListIndex = 0
Else
    playList.ListIndex = playList.ListIndex + 1
End If
Call cmdPlay_Click
Err:
Exit Sub
End Sub

Private Sub cmdOpen_Click()
Call cmdClear_Click
On Error GoTo fileOpenErr
cdOpen.CancelError = True
cdOpen.Flags = &H4& Or &H100& Or cdlOFNPathMustExist Or cdlOFNFileMustExist
cdOpen.DialogTitle = "Select File To Open"
cdOpen.Filter = "MP3 Files (*.mp3)|*.mp3"
cdOpen.ShowOpen
playList.Clear
playList.AddItem cdOpen.FileName
Dim Index As Integer
On Error Resume Next
Index = playList.Text
playList.ListIndex = Index
If Index >= 0 Then playList.Selected(Index) = True
Call cmdPlay_Click
fileOpenErr:
Exit Sub
End Sub

Private Sub cmdOpt_Click()
frmOptions.Show
End Sub

Private Sub cmdPause_Click()
If MediaPlayer1.PlayState = 1 Then
MediaPlayer1.Play
ElseIf MediaPlayer1.PlayState = 2 Then
MediaPlayer1.Pause
Else
End If
End Sub

Private Sub cmdPlay_Click()
frmID3Tag.txtArtist.Text = ""
frmID3Tag.txtTitle.Text = ""
frmID3Tag.txtAlbum.Text = ""
frmID3Tag.txtYear.Text = ""
frmID3Tag.txtComment.Text = ""
If playList.ListIndex = -1 Then
MsgBox "Select A File From The Playlist To Play", vbExclamation, "Error"
MediaPlayer1.Stop
scrollSetter.Enabled = False
scrollPos.Enabled = False
lblCurrentTime.Caption = ""
lblDurationTime.Caption = ""
lblSongName.Caption = ""
lblArtist.Caption = ""
Else
On Local Error Resume Next
Err.Clear
If MediaPlayer1.PlayState = 1 Then
MediaPlayer1.Play
Else
MediaPlayer1.AutoStart = False
MediaPlayer1.FileName = playList.Text
Call MP3Info
Dim FName As String
    FName = MediaPlayer1.FileName


    If FName = "" Then
       frmID3Tag.lblStatus.Caption = "No filename given In command line!"
        Exit Sub
    End If


    If Dir(FName) = "" Then
        frmID3Tag.lblStatus.Caption = "File given In command line was Not found!"
        Exit Sub
    End If
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Dim strInput As String
    Open FName For Binary As FileNum
    


    If LOF(FileNum) < 128 Then
        frmID3Tag.lblStatus.Caption = "File To short For ID3-Tag!"
        Exit Sub
    End If
    
    Seek FileNum, LOF(FileNum) - 127
    strInput = Space(3)
    Get FileNum, , strInput


    If strInput <> "TAG" Then
        frmID3Tag.lblStatus.Caption = "No ID3-Tag found!"
        Exit Sub
    End If
    
    strInput = Space(30)
    Get FileNum, , strInput
    frmID3Tag.txtTitle.Text = strInput
    lblSongName.Caption = strInput
    
    strInput = Space(30)
    Get FileNum, , strInput
    frmID3Tag.txtArtist.Text = strInput
    lblArtist.Caption = strInput
    
    strInput = Space(30)
    Get FileNum, , strInput
    frmID3Tag.txtAlbum.Text = strInput
    
    strInput = Space(4)
    Get FileNum, , strInput
    frmID3Tag.txtYear.Text = strInput
    
    strInput = Space(30)
    Get FileNum, , strInput
    frmID3Tag.txtComment.Text = strInput
    Close FileNum
lblSongInfo = lblArtist.Caption & " - " & lblSongName.Caption
MediaPlayer1.Play
End If
Call DurationTime
scrollSetter.Enabled = True
scrollPos.Enabled = True
End If
frmPopUp.lblSong.Caption = lblSongName.Caption
frmPopUp.lblArtist.Caption = lblArtist.Caption
If frmOptions.chkPopUp.Value = 1 Then
frmPopUp.Show
Else
End If
tmrScroll.Enabled = True
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err:
If playList.ListIndex <= 0 Then
    playList.ListIndex = (playList.ListCount - 1)
Else
    playList.ListIndex = playList.ListIndex - 1
End If
Call cmdPlay_Click
Err:
Exit Sub
End Sub
Private Sub cmdStop_Click()
lblSongInfo.Caption = ""
MediaPlayer1.Stop
scrollSetter.Enabled = False
scrollPos.Value = 0
scrollPos.Enabled = False
lblDurationTime.Caption = ""
lblCurrentTime.Caption = ""
txtBiterate.Caption = ""
txtFrequency.Caption = ""
txtMode.Caption = ""
lblSongName.Caption = ""
lblArtist.Caption = ""
tmrScroll.Enabled = False
End Sub
Function AddToINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long

    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    AddToINI = (lRet)
End Function
Function GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
    Dim sBuffer As String, lRet As Long
    sBuffer = String$(255, 0)
    lRet = GetPrivateProfileString(sSection, sKey, "", sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
       
        If sDefault <> "" Then AddToINI sSection, sKey, sDefault, sIniFile
        GetFromINI = sDefault
    Else
       
        GetFromINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
End Function


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
frmOptions.Show
frmOptions.Hide
   'Skin Loader
   On Error Resume Next
picSkin.Picture = LoadPicture(frmOptions.txtSkinDir)
lblBiteRate.ForeColor = frmOptions.txtTextColour.Text
lblFrequency.ForeColor = frmOptions.txtTextColour.Text
lblMode.ForeColor = frmOptions.txtTextColour.Text
lblMinVolume.ForeColor = frmOptions.txtTextColour.Text
lblMaxVolume.ForeColor = frmOptions.txtTextColour.Text
lblBalanceLeft.ForeColor = frmOptions.txtTextColour.Text
lblBalanceCenter.ForeColor = frmOptions.txtTextColour.Text
lblBalanceRight.ForeColor = frmOptions.txtTextColour.Text
lblPlayList.ForeColor = frmOptions.txtTextColour.Text

   m_blnClose = True
    m_blnMin = True
    m_blnMax = True
       m_blnMax = Not m_blnMax
 EnableMaxButton frmMain.hWnd, m_blnMax
 Dim i As Long
    i = waveOutGetNumDevs()
    If i > 0 Then
        
    Else
        MsgBox "Your system can not play sound Files. No sound card was detected", vbCritical, "Sound Card Test"
    End If
scrollVolume.Value = 0
scrollBalance.Value = 1
Me.Show
Me.Refresh
With nid
.cbSize = Len(nid)
.hWnd = Me.hWnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = frmMain.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 5

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu Me.mnuMain
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = x
Else
msg = x / Screen.TwipsPerPixelX
End If
Select Case msg
Case WM_LBUTTONUP
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hWnd)
Me.Show
Case WM_LBUTTONDBLCLK
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hWnd)
Me.Show
Case WM_RBUTTONUP
Result = SetForegroundWindow(Me.hWnd)
Me.PopupMenu Me.mnuMain
End Select
End Sub

Private Sub Form_Resize()
Dim Result As Long
If Me.WindowState = vbMinimized Then Me.Hide
If Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
If Me.WindowState = vbNormal Then Me.Height = "4750"
If Me.WindowState = vbNormal Then Me.Width = "6645"
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub lblBalanceCenter_Click()
scrollBalance.Value = 1
End Sub
Private Sub lblBalanceLeft_Click()
scrollBalance.Value = 0
End Sub
Private Sub lblBalanceRight_Click()
scrollBalance.Value = 2
End Sub

Private Sub lblDurationTime_Change()
frmID3Tag.txtLength.Text = lblDurationTime.Caption
End Sub

Private Sub lblMaxVolume_Click()
scrollVolume.Value = 0
End Sub

Private Sub lblMinVolume_Click()
scrollVolume.Value = -4000
End Sub

Private Sub lblSongInfo_DblClick()
frmID3Tag.Show
End Sub

Private Sub mnuAdd_Click()
frmAddFile.Show
End Sub

Private Sub mnuAddDir_Click()
frmAddDir.Show
End Sub

Private Sub mnuClear_Click()
playList.Clear
End Sub

Private Sub mnuCredits_Click()
frmCredits.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNext_Click()
Call cmdNext_Click
End Sub

Private Sub mnuOpen_Click()
Dim File As String
   cdOpen.FileName = ""
   cdOpen.Filter = "Nuke Playlist's (*.npl)|*.npl"
   cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
File = cdOpen.FileName
Dim A As String
Dim x As String
On Error GoTo Error
Open File For Input As #1
Do Until EOF(1)
Input #1, A$
playList.AddItem A$
Loop
Close 1
Dim Index As Integer
On Error Resume Next
Index = playList.Text
playList.ListIndex = Index
If Index >= 0 Then playList.Selected(Index) = True
Call cmdPlay_Click
Exit Sub
Error:
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show
End Sub

Private Sub mnuPause_Click()
Call cmdPause_Click
End Sub

Private Sub mnuPlay_Click()
Call cmdPlay_Click
End Sub

Private Sub mnuPrevious_Click()
Call cmdPrevious_Click
End Sub

Private Sub mnuRemove_Click()
If playList.ListIndex = -1 Then
MsgBox "No File Selected", vbExclamation, "Error"
Else
playList.RemoveItem playList.ListIndex
End If
End Sub

Private Sub mnuRestore_Click()
Dim Result As Long
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hWnd)
Me.Show
End Sub

Private Sub mnuSave_Click()
cdSave.FileName = ""
    cdSave.Filter = "Nuke Playlist's (*.npl)|*.npl"
    cdSave.ShowSave
    If cdSave.FileName <> "" Then
        SavePlaylist cdSave.FileName
    End If
End Sub

Private Sub mnuStop_Click()
Call cmdStop_Click
End Sub
Private Sub picSkin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu Me.mnuMain
End Sub

Private Sub playList_DblClick()
Call cmdPlay_Click
End Sub

Private Sub playList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iIndex As Integer
    Dim iCnt As Integer

    If KeyCode = vbKeyDelete Then
     playList.RemoveItem playList.ListIndex
    End If
End Sub

Private Sub playList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCnt As Integer
    For iCnt = 1 To Data.Files.Count
        If InStr(1, Data.Files(iCnt), ".") <> 0 Then
                playList.AddItem Data.Files(iCnt)
            End If
    Next iCnt
End Sub

Private Sub scrollBalance_Change()
If scrollBalance.Value = 0 Then
MediaPlayer1.Balance = -4000
End If
If scrollBalance.Value = 1 Then
MediaPlayer1.Balance = 0
End If
If scrollBalance.Value = 2 Then
MediaPlayer1.Balance = 4000
End If
End Sub
Private Sub scrollPos_Change()
Dim MyValue As String
If scrollPos.Value = scrollPos.Max Then
scrollPos.Value = scrollPos.min
lblCurrentTime.Caption = "00:00"
lblDurationTime.Caption = "00:00"

If frmOptions.optRandom.Value = True Then
Randomize Timer
 MyValue = Int((playList.ListCount * Rnd))
    playList.ListIndex = MyValue
    Call cmdPlay_Click
Else
If frmOptions.optRepeatOne.Value = True Then
Call cmdPlay_Click
Else
If playList.ListIndex = playList.ListCount - 1 Then
Call cmdStop_Click
Else
If frmOptions.optRepeatAll.Value = True Then
'function not added yet
MsgBox "Repeat All"
Else
If frmOptions.optStandard.Value = True Then
On Error Resume Next
Call cmdNext_Click

End If
End If
End If
End If
End If
End If
End Sub

Private Sub scrollPos_Scroll()
MediaPlayer1.CurrentPosition = scrollPos.Value
End Sub

Private Sub scrollSetter_Timer()
If MediaPlayer1.PlayState = mpClosed Or mpStopped Then
scrollSetter.Enabled = False
Else
Call CurrentTime
scrollPos.Max = MediaPlayer1.Duration
scrollPos.Value = MediaPlayer1.CurrentPosition
End If
End Sub

Private Sub scrollVolume_Change()
If scrollVolume.Value = -4000 Then
MediaPlayer1.Volume = scrollVolume.Value
End If
If scrollVolume.Value = 0 Then
MediaPlayer1.Volume = scrollVolume.Value
End If
End Sub

Private Sub scrollVolume_Scroll()
MediaPlayer1.Volume = scrollVolume.Value
End Sub


Private Sub tmrPlay_Timer()
On Error Resume Next
If MediaPlayer1.Duration - MediaPlayer1.CurrentPosition < 2 Then
MediaPlayer1.Stop
Else
If playList.ListIndex + 1 < playList.ListCount Then

MediaPlayer1.FileName = playList.Text
MediaPlayer1.Play
Else
If MediaPlayer1.Duration - MediaPlayer1.CurrentPosition > 2 Then
Exit Sub
End If
End If
End If
End Sub

Private Sub tmrScroll_Timer()
    dta = frmID3Tag.txtArtist.Text & " - " & frmID3Tag.txtTitle.Text & " * " & frmID3Tag.txtLength & " * " & Space$(55)
    s = s + 1
    lblSongInfo.Caption = Mid(dta, 1, s)
    If Len(lblSongInfo.Caption) >= 56 Then lblSongInfo.Caption = Right(lblSongInfo.Caption, 300)


    If s = Len(dta) Then
        lblSongInfo.Caption = ""
        s = 0
    End If
End Sub

Private Sub tmrWritin_Timer()

End Sub


Private Sub txtBiteRate_Change()
frmID3Tag.txtBiterate.Text = txtBiterate.Caption
End Sub

Private Sub txtFrequency_Change()
frmID3Tag.txtFreq.Text = txtFrequency.Caption
End Sub

Private Sub txtMode_Change()
frmID3Tag.txtMode.Text = txtMode.Caption
End Sub

