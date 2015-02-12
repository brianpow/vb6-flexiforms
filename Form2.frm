VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Layout Test - Mixed Scale Mode"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   StartUpPosition =   3  'Windows Default
   Tag             =   "varFull"
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2280
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   2
      Tag             =   "H"
      Top             =   120
      Width           =   2415
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   2115
         TabIndex        =   3
         Tag             =   "H"
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Tag             =   "H"
            Top             =   120
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Tag             =   "F"
      Top             =   960
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private myLayout As New clsLayout

Private Sub cmdexit_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
   myLayout.SetLayout Me
   Debug.Print "Form_Load(): " & PropertyExist(Picture1, "Container", ReadableProperties)
End Sub

Private Sub Form_Resize()
   myLayout.RedrawLayout
End Sub

