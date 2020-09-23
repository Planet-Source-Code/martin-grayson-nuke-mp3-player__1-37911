VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Credits"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      Picture         =   "frmCredits.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblPsc 
      Caption         =   "Please visit  http://www.planet-source-code.com  "
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "E-Mail: martin_grayson@hotmail.com"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblName 
      Caption         =   "Nuke - The Best MP3 Player Of Them All v1.4  Designed and programmed by Martin Grayson  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblStuff 
      Caption         =   $"frmCredits.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
