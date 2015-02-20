VERSION 5.00
Begin VB.Form frmSlpashI 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   2670
   ClientTop       =   0
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   Picture         =   "frmSlpashI.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrCargando 
      Interval        =   1
      Left            =   9960
      Top             =   8160
   End
   Begin VB.PictureBox pctCargando 
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   1000
      Height          =   255
      Left            =   720
      ScaleHeight     =   1000
      ScaleMode       =   0  'User
      ScaleWidth      =   3500
      TabIndex        =   0
      Top             =   8280
      Width           =   9135
   End
   Begin VB.Timer Splash 
      Interval        =   3500
      Left            =   9840
      Top             =   120
   End
End
Attribute VB_Name = "frmSlpashI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public X As Integer

Private Sub Form_Load()
X = 0
End Sub

Private Sub Splash_Timer()
MainMenu.Show
Unload Me
End Sub

Private Sub tmrCargando_Timer()
X = X + 1

pctCargando.PSet (X, 50), RGB(0, 255, 0)

End Sub
