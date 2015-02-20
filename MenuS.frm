VERSION 5.00
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form MenuS 
   Appearance      =   0  'Flat
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   3675
   ClientTop       =   2310
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AlphaImageControl.aicAlphaImage imgMenu 
      Height          =   1335
      Left            =   6840
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Image           =   "MenuS.frx":0000
      Scaler          =   1
   End
   Begin VB.Label No 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4800
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Si 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Si"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Titulo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Height          =   1575
      Left            =   6720
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   1815
      Left            =   6600
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "MenuS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub No_Click()

Select Case (Titulo.Caption)
    
    'Termina el programa
    Case Is = "¿Desea salir de Clean My Mess?"
        MainMenu.Enabled = True
        
    'Sobreescritura del archivo
    Case Is = "¿Desea sobreescribir la partida?"
        LoadGameMenu.Enabled = True
        
    'Eliminar partida
    Case Is = "¿Desea eliminar su partida?"
        LoadGameMenu.Enabled = True
        
    'Reseteo del juego
    Case Is = "Se volveran los valores a 0. ¿Seguir?"
        OptionsMenu.Enabled = True
        
    Case Is = "¿Empezar otro juego?"
        MainMenu.Enabled = True
        
    Case Is = "¿Desea ir al Main Menu?"
        MenuGame.Show
        Unload Me
        
End Select

Unload Me

End Sub

Private Sub Si_Click()

Select Case (Titulo.Caption)
    
    'Termina el programa
    Case Is = "¿Desea salir de Clean My Mess?"
        Fondo.Salir.Enabled = True
    
    'Sobreescritura del archivo
    Case Is = "¿Desea sobreescribir la partida?"
        Yes_No 2
    
    'Eliminar partida
    Case Is = "¿Desea eliminar su partida?"
        Yes_No 3
    
    'Reseteo del juego
    Case Is = "Se volveran los valores a 0. ¿Seguir?"
        Yes_No 4
    
    'Guardar
    Case Is = "Partida guardada con exito!!"
        Yes_No 5
    
    'Aparencia del menu
    Case Is = "No hay datos guardados"
        LoadGameMenu.Enabled = True
        MenuS.Si.Left = 2040
        MenuS.Si.Width = 975
        MenuS.Si.Caption = "Si"
        MenuS.No.Visible = True
        Unload Me
        
    'Nuevo juego con uno cargado
    Case Is = "¿Empezar otro juego?"
        Yes_No 6
    
    Case Is = "¿Desea ir al Main Menu?"
        MainMenu.Show
        Unload Game
        Unload Me
        
End Select

    
End Sub


