VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Your Skin Directory And File"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
frmOptions.txtSkinDir = Dir1.Path & "\" & File1.FileName
frmSkin.Hide
End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub


