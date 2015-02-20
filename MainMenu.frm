VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form MainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WMPLibCtl.WindowsMediaPlayer MusicFondo 
      Height          =   420
      Left            =   14760
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   420
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   741
      _cy             =   741
   End
   Begin VB.Label Extras 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Extras"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   10
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   13200
      TabIndex        =   9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Creditos 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   8
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label StartGame 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   7
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   12600
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Salir 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   5
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   13200
      TabIndex        =   4
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label LoadGame 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   13200
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Options 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   1
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   13200
      TabIndex        =   0
      Top             =   6960
      Width           =   1815
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'################################################################
'################################################################
'################################################################
'####            Clean My Mess v 0.1 (c)2012                 ####
'####            ===========================                 ####
'####                                                        ####
'####     Codigo libre, cualquier modificacion, validar.     ####
'####                                                        ####
'################################################################
'################################################################
'####   Equipo:                                              ####
'####   =======                                              ####
'####                                                        ####
'####       Programacion: Kevin N. Encinas V.                ####
'####           Lenguaje: BASIC                              ####
'####           Entorno: Visual Basic 6                      ####
'####                                                        ####
'################################################################
'################################################################
'################################################################


Private Sub Creditos_Click()
CreditosMenu.Show
Me.Hide
End Sub

Private Sub Extras_Click()
'
'Menu de Extras
'
ExtrasMenu.Show
Me.Hide
End Sub


Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\archived\images\MenuCMM.jpg")

OptionsMenu.Show
Game.Show
Game.Hide
OptionsMenu.Hide

End Sub


Private Sub LoadGame_Click()
'Menu de Carga de Partida

LoadGameMenu.Show
Me.Hide
End Sub

Private Sub Options_Click()
'
'Opciones
'
OptionsMenu.Show
LoadSave_Opciones "L"
Me.Enabled = False
Me.Hide

End Sub

Private Sub Salir_Click()
'
'Salir
'

Me.Enabled = False
MenuEspecial (1)
End Sub

Private Sub StartGame_Click()
'
'StartNewGame
'
Game.Show
Load OptionsMenu
Game.TimeJuego.Enabled = True
LoadSave_Opciones "L"
Unload OptionsMenu

Game.seg = 0
Game.min = 0
Game.hora = 0

Game.pMusica.Enabled = True
Game.pMusica.URL = App.Path & "\archived\sounds\music\music1.mp3"

Unload Me

End Sub
