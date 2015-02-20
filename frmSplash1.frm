VERSION 5.00
Begin VB.Form frmSplash1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   3500
      Left            =   9720
      Top             =   120
   End
   Begin VB.Timer tmrBarCarg 
      Interval        =   15
      Left            =   9840
      Top             =   7560
   End
   Begin VB.PictureBox pctBarCarg 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   150
      Height          =   255
      Left            =   240
      ScaleHeight     =   1000
      ScaleMode       =   0  'User
      ScaleWidth      =   3633.545
      TabIndex        =   0
      Top             =   8040
      Width           =   9855
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\archived\images\MenuCMM.jpg")
End Sub

Private Sub tmrBarCarg_Timer()
Static nbar As Integer

nbar = nbar + 20

pctBarCarg.PSet (nbar, 500), RGB(0, 255, 0)

End Sub

Private Sub tmrSplash_Timer()
Fondo.Show
Fondo.Enabled = False
MainMenu.Show
Unload Me
End Sub
