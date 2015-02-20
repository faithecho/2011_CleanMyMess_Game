VERSION 5.00
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form OptionsMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox FullScreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   9360
      TabIndex        =   26
      Top             =   5760
      Width           =   255
   End
   Begin VB.HScrollBar musicaVol 
      Height          =   615
      Left            =   3480
      Max             =   100
      TabIndex        =   19
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CheckBox ContAdult 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   15
      Top             =   4680
      Width           =   255
   End
   Begin VB.HScrollBar VelTexto 
      Height          =   615
      Left            =   3120
      Max             =   1000
      Min             =   100
      TabIndex        =   8
      Top             =   7680
      Value           =   300
      Width           =   3855
   End
   Begin VB.HScrollBar FXVol 
      Height          =   615
      Left            =   3480
      Max             =   100
      TabIndex        =   3
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Pantalla Completa"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   25
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label Label18 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   9120
      TabIndex        =   24
      Top             =   5520
      Width           =   3975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Volumen Musica"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   21
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   20
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Back 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "<< Volver"
      Height          =   255
      Left            =   14520
      TabIndex        =   0
      Top             =   8640
      Width           =   735
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   4215
      Left            =   12240
      TabIndex        =   18
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7435
      Image           =   "OptionsMenu.frx":0000
      Scaler          =   1
   End
   Begin VB.Label ResetAll 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Resetear juego entero"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   16
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Deshabilitar contenido adulto"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   13
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Label Label12 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   8280
      TabIndex        =   14
      Top             =   4440
      Width           =   5535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Velocidad de Texto"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   7080
      Width           =   4335
   End
   Begin VB.Label Label10 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   2640
      TabIndex        =   12
      Top             =   6960
      Width           =   4815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   10
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   9
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   7
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Volumen  FX"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   4920
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label Label14 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   9120
      TabIndex        =   17
      Top             =   6600
      Width           =   3975
   End
   Begin VB.Label Label17 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   2640
      TabIndex        =   23
      Top             =   2880
      Width           =   4815
   End
End
Attribute VB_Name = "OptionsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()

If Not (MainMenu.Enabled) Then
    MainMenu.Show
    MainMenu.Enabled = True
    Unload Me
Else
    Game.Show
    Game.TimeJuego.Enabled = 1
    Unload Me
End If

End Sub

Private Sub ContAdult_Click()

'Cambio de opciones
LoadSave_Opciones "S"

End Sub

Private Sub Form_Load()
'Carga de datos
LoadSave_Opciones "L"
End Sub

Private Sub FXVol_Change()

'Cambio de opciones
LoadSave_Opciones "S"

Game.pFX.settings.volume = FXVol.Value

End Sub

Private Sub musicaVol_Change()

'Cambio de opciones
LoadSave_Opciones "S"

Game.pMusica.settings.volume = musicaVol.Value

End Sub

Private Sub VelTexto_Change()

'Cambio de opciones
LoadSave_Opciones "S"

End Sub

