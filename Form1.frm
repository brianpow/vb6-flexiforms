VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Layout test"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Tag             =   "Bottom Right"
      Top             =   1170
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1890
      TabIndex        =   2
      Tag             =   "fixBottom fixRight"
      Top             =   1170
      Width           =   1395
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   2400
      TabIndex        =   1
      Tag             =   " varWidth varVertical fixRight"
      Top             =   225
      Width           =   2280
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Tag             =   "varWidth varVertical"
      Top             =   225
      Width           =   2265
   End
   Begin VB.Label Label3 
      Caption         =   "List 2"
      Height          =   240
      Left            =   2415
      TabIndex        =   5
      Tag             =   "varWidth fixRight"
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

Private Sub Form_Load()
   myLayout.SetLayout Me
End Sub

Private Sub Form_Resize()
   myLayout.RedrawLayout
End Sub
