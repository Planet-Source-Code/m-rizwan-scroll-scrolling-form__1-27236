VERSION 5.00
Object = "*\ARiz_Scroll.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Riz_Scroll.RizScroll RizScroll1 
      Height          =   720
      Left            =   3720
      Top             =   2520
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   -360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    RizScroll1.Scrolling_Horizontal = True
    RizScroll1.Scrolling_Vertical = True
    RizScroll1.StartScroll Me
End Sub
