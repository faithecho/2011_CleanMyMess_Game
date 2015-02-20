VERSION 5.00
Begin VB.Form CreditosMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDITS"
   ClientHeight    =   3585
   ClientLeft      =   3270
   ClientTop       =   2640
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8925
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "dibujos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "historia,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "diseño de personajes,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "(C) 2012"
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Kevin Nahuel Encinas Vargas"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Programación (BASIC),"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "CreditosMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Fondo.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Fondo.Enabled = True
MainMenu.Show
End Sub
