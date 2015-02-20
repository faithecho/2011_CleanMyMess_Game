VERSION 5.00
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form ExtrasMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin AlphaImageControl.aicAlphaImage imgStory 
      Height          =   1335
      Left            =   3480
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      Image           =   "ExtrasMenu.frx":0000
      Scaler          =   1
   End
   Begin AlphaImageControl.aicAlphaImage imgAnimacion 
      Height          =   2535
      Left            =   10320
      TabIndex        =   13
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4471
      Image           =   "ExtrasMenu.frx":0018
      Scaler          =   1
   End
   Begin AlphaImageControl.aicAlphaImage imgImagen 
      Height          =   2535
      Left            =   6120
      TabIndex        =   12
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4471
      Image           =   "ExtrasMenu.frx":0030
      Scaler          =   1
   End
   Begin AlphaImageControl.aicAlphaImage imgMusica 
      Height          =   2535
      Left            =   1920
      TabIndex        =   11
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4471
      Image           =   "ExtrasMenu.frx":0048
      Scaler          =   1
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Animaciones"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   10080
      TabIndex        =   10
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Pics"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   5880
      TabIndex        =   9
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5160
      TabIndex        =   8
      Top             =   6960
      Width           =   6855
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      Height          =   1815
      Left            =   3240
      TabIndex        =   7
      Top             =   6720
      Width           =   9015
   End
   Begin VB.Label Label7 
      BackColor       =   &H000040C0&
      Height          =   3855
      Left            =   9840
      TabIndex        =   6
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   3855
      Left            =   5640
      TabIndex        =   5
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Music"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Height          =   3855
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Back 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "<< Volver"
      Height          =   255
      Left            =   14520
      TabIndex        =   2
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "EXTRAS"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "ExtrasMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()
MainMenu.Show
Me.Hide
End Sub

Private Sub Form_Load()
Cargar_Extras
End Sub

