VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optRandom 
      BackColor       =   &H00808080&
      Caption         =   "Random"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   4440
      Width           =   1815
   End
   Begin VB.OptionButton optRepeatAll 
      BackColor       =   &H00808080&
      Caption         =   "Repeat All"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   4080
      Width           =   1815
   End
   Begin VB.OptionButton optRepeatOne 
      BackColor       =   &H00808080&
      Caption         =   "Repeat One"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3720
      Width           =   1815
   End
   Begin VB.OptionButton optStandard 
      BackColor       =   &H00808080&
      Caption         =   "Standard Play"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3360
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtTextColour 
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdSkinBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtSkinDir 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   4800
      Width           =   735
   End
   Begin VB.CheckBox chkOnTop 
      BackColor       =   &H00808080&
      Caption         =   "Always On Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkPopUp 
      BackColor       =   &H00808080&
      Caption         =   "Song Popups"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Label lblPlayOptions 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Play Options"
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
      Left            =   480
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblAppearance 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Appearance"
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
      Left            =   480
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblOptions 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Main Option"
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
      Left            =   480
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Text Colour"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Path"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   5175
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = hWnd_NOTOPMOST
    End If
    SetWindowPos myfrm.hWnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


Private Sub chkOnTop_Click()
If chkOnTop.Value = 1 Then AlwaysOnTop frmMain, True
If chkOnTop.Value = 0 Then AlwaysOnTop frmMain, False
If chkOnTop.Value = 1 Then AlwaysOnTop frmOptions, True
If chkOnTop.Value = 0 Then AlwaysOnTop frmOptions, False
If chkOnTop.Value = 1 Then AlwaysOnTop frmAddFile, True
If chkOnTop.Value = 0 Then AlwaysOnTop frmAddFile, False
If chkOnTop.Value = 1 Then AlwaysOnTop frmCredits, True
If chkOnTop.Value = 0 Then AlwaysOnTop frmCredits, False
If chkOnTop.Value = 1 Then AlwaysOnTop frmID3Tag, True
If chkOnTop.Value = 0 Then AlwaysOnTop frmID3Tag, False
If chkOnTop.Value = 1 Then AlwaysOnTop frmAddDir, True
If chkOnTop.Value = 0 Then AlwaysOnTop frmAddDir, False
If chkOnTop.Value = 1 Then AlwaysOnTop frmSkin, True
If chkOnTop.Value = 0 Then AlwaysOnTop frmSkin, False
End Sub
Private Sub cmdHide_Click()
Call cmdRefresh_Click
Me.Hide
End Sub
Private Sub cmdRefresh_Click()
 'Skin Loader
   On Error Resume Next
frmMain.picSkin.Picture = LoadPicture(frmOptions.txtSkinDir)
frmMain.lblBiteRate.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblFrequency.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblMode.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblMinVolume.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblMaxVolume.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblBalanceLeft.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblBalanceCenter.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblBalanceRight.ForeColor = frmOptions.txtTextColour.Text
frmMain.lblPlayList.ForeColor = frmOptions.txtTextColour.Text
frmMain.Refresh
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
AddToINI "Options", "Popups", chkPopUp.Value, App.Path & "\settings.ini"
AddToINI "Options", "AlwaysOnTop", chkOnTop.Value, App.Path & "\settings.ini"
AddToINI "Options", "SkinDir", txtSkinDir.Text, App.Path & "\settings.ini"
AddToINI "Options", "TxtColour", txtTextColour.Text, App.Path & "\settings.ini"
AddToINI "Play Options", "Random", optRandom.Value, App.Path & "\settings.ini"
AddToINI "Play Options", "Repeat One", optRepeatOne.Value, App.Path & "\settings.ini"
AddToINI "Play Options", "Repeat All", optRepeatAll.Value, App.Path & "\settings.ini"
AddToINI "Play Options", "Standard", optStandard.Value, App.Path & "\settings.ini"
Call cmdRefresh_Click
Call cmdHide_Click
End Sub

Private Sub cmdSkinBrowse_Click()
frmSkin.Show
End Sub


Private Sub Form_Load()
On Error Resume Next
RemoveMenus
sValue = GetFromINI("Options", "Popups", "", App.Path & "\settings.ini")
chkPopUp.Value = sValue
sValue = GetFromINI("Options", "AlwaysOnTop", "", App.Path & "\settings.ini")
chkOnTop.Value = sValue
sValue = GetFromINI("Options", "SkinDir", "", App.Path & "\settings.ini")
txtSkinDir.Text = sValue
sValue = GetFromINI("Options", "TxtColour", "", App.Path & "\settings.ini")
txtTextColour.Text = sValue
sValue = GetFromINI("Play Options", "Random", "", App.Path & "\settings.ini")
optRandom.Value = sValue
sValue = GetFromINI("Play Options", "Repeat One", "", App.Path & "\settings.ini")
optRepeatOne.Value = sValue
sValue = GetFromINI("Play Options", "Repeat All", "", App.Path & "\settings.ini")
optRepeatAll.Value = sValue
sValue = GetFromINI("Play Options", "Standard", "", App.Path & "\settings.ini")
optStandard.Value = sValue
Call cmdRefresh_Click
End Sub

Private Sub RemoveMenus()
Dim hMenu As Long
hMenu = GetSystemMenu(hWnd, False)
DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub
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
Function AddToINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    AddToINI = (lRet)
End Function

