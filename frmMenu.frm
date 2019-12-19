VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   4080
      Picture         =   "frmMenu.frx":2C974
      ScaleHeight     =   1875
      ScaleWidth      =   20190
      TabIndex        =   5
      Top             =   -120
      Width           =   20250
   End
   Begin VB.PictureBox Picture3 
      Height          =   10000
      Left            =   0
      Picture         =   "frmMenu.frx":3EF2B
      ScaleHeight     =   9945
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   1800
      Width           =   1700
      Begin VB.Image imgAdd 
         Height          =   720
         Left            =   360
         MouseIcon       =   "frmMenu.frx":6B89F
         MousePointer    =   99  'Custom
         Picture         =   "frmMenu.frx":6C169
         Top             =   360
         Width           =   720
      End
      Begin VB.Image imgUpdate 
         Height          =   720
         Left            =   480
         MouseIcon       =   "frmMenu.frx":6D033
         MousePointer    =   99  'Custom
         Picture         =   "frmMenu.frx":6D8FD
         Top             =   3840
         Width           =   720
      End
      Begin VB.Image imgDelete 
         Height          =   600
         Left            =   480
         MouseIcon       =   "frmMenu.frx":6E7C7
         MousePointer    =   99  'Custom
         Picture         =   "frmMenu.frx":6F091
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Information"
         BeginProperty Font 
            Name            =   "Eras Light ITC"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblDelete 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Information"
         BeginProperty Font 
            Name            =   "Eras Light ITC"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblUpdate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Update Information"
         BeginProperty Font 
            Name            =   "Eras Light ITC"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Image imgClose 
         Height          =   600
         Left            =   480
         MouseIcon       =   "frmMenu.frx":6F4D3
         MousePointer    =   99  'Custom
         Picture         =   "frmMenu.frx":6FD9D
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   600
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close Program"
         BeginProperty Font 
            Name            =   "Eras Light ITC"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   6480
         Width           =   900
      End
   End
   Begin VB.Image Image6 
      Height          =   1815
      Left            =   0
      Picture         =   "frmMenu.frx":70A67
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4080
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub imgAdd_Click()
    frmAdd.Show vbModal
End Sub

Private Sub imgDelete_Click()
    frmDelete.Show vbModal
End Sub

Private Sub imgClose_Click()
    End
End Sub

Private Sub imgUpdate_Click()
    frmUpdate.Show vbModal
End Sub

Private Sub lblAdd_Click()
    frmAdd.Show vbModal
End Sub

Private Sub lblClose_Click()
    End
End Sub

Private Sub lblDelete_Click()
    frmDelete.Show vbModal
End Sub

Private Sub lblUpdate_Click()
    frmSearchUpdate.Show vbModal
End Sub

