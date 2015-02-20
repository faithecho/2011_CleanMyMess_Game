VERSION 5.00
Begin VB.Form MenuGame 
   Appearance      =   0  'Flat
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   -285
   ClientTop       =   870
   ClientWidth     =   4440
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label lblImage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Texto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Load 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblReturn 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Height          =   6855
      Left            =   -1080
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "MenuGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lblImage_Click()

With Game
    .charChat.Visible = 0
    .charName.Visible = 0
    .charNameFondo.Visible = 0
    .Skip.Visible = 0
    .Cortina.Enabled = 1
    .QuitMenu.Visible = 0
    .OptionsGame.Visible = 0
    .SaveGame.Visible = 0
    .TiempoJugado.Visible = 0
    .Enabled = 1
End With

Unload Me

End Sub

Private Sub lblMainMenu_Click()

MenuEspecial 7

End Sub

'Regresar del menu raro q hice
Private Sub lblReturn_Click()

'Form de juego habilitado
Game.Enabled = True

'Tiempo de juego renaudado
Game.TimeJuego.Enabled = True

'Descarga del menu
Unload Me

End Sub

Private Sub Load_Click()
'Esconde el Form
Fondo.Show

'Carga el menu de Load
LoadGameMenu.Show
LoadGameMenu.FondoEliminar.Visible = 0
LoadGameMenu.Eliminar.Visible = False

'Descarga el menu
Unload Me

End Sub


