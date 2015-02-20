VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form Game 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   -135
   ClientTop       =   -135
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer contrlMusica 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12840
      Top             =   6480
   End
   Begin VB.TextBox charChat 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   7200
      Width           =   14175
   End
   Begin VB.Timer TimeTexto 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   13800
      Top             =   6480
   End
   Begin VB.Timer TimeJuego 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13320
      Top             =   6480
   End
   Begin VB.Label charName 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label charNameFondo 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   6360
      Width           =   2535
   End
   Begin WMPLibCtl.WindowsMediaPlayer pFX 
      Height          =   375
      Left            =   14760
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
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
      _cx             =   661
      _cy             =   661
   End
   Begin VB.Label Opicion 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5040
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Opicion 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5040
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label OptFondo 
      BackColor       =   &H000040C0&
      Height          =   855
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label OptFondo 
      BackColor       =   &H000040C0&
      Height          =   855
      Index           =   0
      Left            =   4920
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Skip 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   14520
      TabIndex        =   7
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label TiempoJugado 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10920
      TabIndex        =   5
      Top             =   8640
      Width           =   2055
   End
   Begin VB.Label OptionsGame 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Options"
      Height          =   255
      Left            =   13200
      TabIndex        =   2
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label SaveGame 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Save"
      Height          =   255
      Left            =   13920
      TabIndex        =   1
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label QuitMenu 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Menu"
      Height          =   255
      Left            =   14640
      TabIndex        =   0
      Top             =   8640
      Width           =   615
   End
   Begin AlphaImageControl.aicAlphaImage Cortina 
      Height          =   8625
      Left            =   -240
      TabIndex        =   16
      Top             =   -2760
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   15214
      Image           =   "Game.frx":0000
      Scaler          =   1
      Enabled         =   0   'False
      Props           =   0
   End
   Begin AlphaImageControl.aicAlphaImage imgActor2 
      Height          =   8985
      Left            =   7920
      TabIndex        =   9
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   15849
      Image           =   "Game.frx":0018
      Scaler          =   1
   End
   Begin AlphaImageControl.aicAlphaImage imgActor1 
      Height          =   9360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   16510
      Image           =   "Game.frx":0030
      Scaler          =   3
   End
   Begin AlphaImageControl.aicAlphaImage imFondo 
      Height          =   14160
      Left            =   -1200
      TabIndex        =   15
      Top             =   0
      Width           =   22680
      _ExtentX        =   40005
      _ExtentY        =   24977
      Image           =   "Game.frx":0048
      Scaler          =   1
      Props           =   0
   End
   Begin WMPLibCtl.WindowsMediaPlayer pMusica 
      Height          =   375
      Left            =   14280
      TabIndex        =   17
      Top             =   6480
      Width           =   375
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
      _cx             =   661
      _cy             =   661
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public seg As Byte, min As Byte, hora As Byte

Private Sub contrlMusica_Timer()

If (pMusica.playState = wmppsStopped) Then
    pMusica.Controls.play
End If

End Sub

Private Sub Cortina_Click(ByVal Button As Integer)

If (Button = vbKeyRButton) Then
    Cortina.Enabled = False
    charChat.Visible = 1
    charName.Visible = 1
    charNameFondo.Visible = 1
    QuitMenu.Visible = 1
    TiempoJugado.Visible = 1
    OptionsGame.Visible = 1
    Skip.Visible = 1
    SaveGame.Visible = 1
    MenuGame.Show
    Game.Enabled = 0
End If

End Sub

Private Sub Form_Load()

Me.imFondo.LoadImage_FromFile App.Path & "\archived\images\bg\cap1\bg_Cortinas.jpg"

'Velocidad del texto incial
TimeTexto.Interval = OptionsMenu.VelTexto.Value

End Sub

Private Sub OptionsGame_Click()
'
'Abrir Menu opciones
'

OptionsMenu.Show
LoadSave_Opciones "L"
Me.TimeJuego.Enabled = 0

End Sub

Private Sub QuitMenu_Click()
'Menu especial de control
Me.Enabled = False

MenuGame.Show
Me.TimeJuego.Enabled = False

End Sub

Private Sub SaveGame_Click()

'
'cambia para formar el Guardar
'
With LoadGameMenu
    .Show
    .Enabled = True
    .Titulo.Caption = "Guardar"
    .Eliminar.Visible = False
    .FondoEliminar.Visible = False
End With

Game.Hide

TimeJuego.Enabled = False

End Sub

'
Private Sub TimeJuego_Timer()
'###############################################

seg = seg + 1

If (seg = 60) Then
    min = min + 1
    seg = 0
    If min = 60 Then
        hora = hora + 1
        min = 0
    End If
End If

'#################################################

'Tiempo de juego
TiempoJugado.Caption = Str(hora) & " : " & Str(min) & " : " & Str(seg)
    
End Sub

