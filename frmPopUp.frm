VERSION 5.00
Begin VB.Form frmPopUp 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Nuke - Current Song"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPopUp 
      Interval        =   500
      Left            =   2760
      Top             =   120
   End
   Begin VB.Label lblArtist 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblSong 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DirectionIsUp As Boolean

Private Sub Form_Click()
frmMain.Show
frmMain.WindowState = vbNormal
End Sub

Private Sub Form_Load()
Me.Top = Screen.Height + 10
Me.Left = Screen.Width - (Me.Width + 100)
DirectionIsUp = True
End Sub

Private Sub tmrPopUp_Timer()
 tmrPopUp.Interval = 10
  If DirectionIsUp Then
    Me.Top = Me.Top - 50
    If (Me.Top <= Screen.Height - (Me.Height - 10)) Then
      tmrPopUp.Interval = 3000
      DirectionIsUp = False
    End If
  Else
    Me.Top = Me.Top + 50
    If Me.Top >= Screen.Height + 10 Then
      tmrPopUp.Enabled = False
      Unload Me
    End If
  End If
End Sub

