VERSION 5.00
Begin VB.Form frmAddFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Selected Mp3's To Playlist"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAddFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
frmMain.playList.AddItem Dir1.Path & "\" & File1.FileName
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_DblClick()
Call cmdAdd_Click
End Sub

