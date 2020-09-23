VERSION 5.00
Begin VB.Form frmID3Tag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ID3 Tag"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmID3Tag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMode 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   20
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtFreq 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtBiterate 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtLength 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtComment 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtAlbum 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtArtist 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblMode 
      BackColor       =   &H00808080&
      Caption         =   "Mode"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblFreq 
      BackColor       =   &H00808080&
      Caption         =   "Frequency"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblRate 
      BackColor       =   &H00808080&
      Caption         =   "Biterate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblLength 
      BackColor       =   &H00808080&
      Caption         =   "Length"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblComment 
      BackColor       =   &H00808080&
      Caption         =   "Comment"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00808080&
      Caption         =   "Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblAlbum 
      BackColor       =   &H00808080&
      Caption         =   "Album"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00808080&
      Caption         =   "Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblArtist 
      BackColor       =   &H00808080&
      Caption         =   "Artist"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblId3 
      BackColor       =   &H00808080&
      Caption         =   "ID3 Tag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmID3Tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Dim ThisFile As String
Dim FileType As String
Private Sub cmdHide_Click()
Me.Hide
End Sub

Private Sub Form_Load()
RemoveMenus
End Sub

Private Sub RemoveMenus()
Dim hMenu As Long
hMenu = GetSystemMenu(hWnd, False)
DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

