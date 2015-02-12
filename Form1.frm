VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Layout Test - Twip"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   2115
      TabIndex        =   9
      Tag             =   "HV"
      Top             =   2040
      Width           =   2175
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Tag             =   "HV"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Tag             =   "WX"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Tag             =   "W"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Tag             =   "WX"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Tag             =   "W"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Tag             =   "XW"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Tag             =   "W"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   2400
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   8
      Tag             =   "VR"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Tag             =   "H"
      Top             =   1080
      Width           =   2175
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1875
         TabIndex        =   7
         Tag             =   "H"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Tag             =   "BR"
      Top             =   3720
      Width           =   1350
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Ok"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Tag             =   "BR"
      Top             =   3720
      Width           =   1395
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   2400
      TabIndex        =   1
      Tag             =   "R"
      Top             =   225
      Width           =   2760
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Tag             =   "H"
      Top             =   225
      Width           =   2265
   End
   Begin VB.Label Label3 
      Caption         =   "List 2"
      Height          =   240
      Left            =   2415
      TabIndex        =   5
      Tag             =   "X"
      Top             =   30
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "List 1"
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myLayout As New clsLayout

Private Sub cmdexit_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
   myLayout.SetLayout Me
   Form2.Show
End Sub

Private Sub Form_Resize()
   Call myLayout.RedrawLayout(True)
End Sub
