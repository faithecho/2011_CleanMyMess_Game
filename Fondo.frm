VERSION 5.00
Begin VB.Form Fondo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   750
   ClientTop       =   0
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Salir 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   480
   End
End
Attribute VB_Name = "Fondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Salir_Timer()
Static cont As Integer

cont = cont + Salir.Interval

Select Case cont
    Case Is = 1000
        MenuS.Hide
    Case Is = 1500
        MainMenu.Hide
    Case Is = 2500
        End
End Select

End Sub
