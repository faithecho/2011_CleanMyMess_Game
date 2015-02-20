VERSION 5.00
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form LoadGameMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Label SaLo 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   7560
      TabIndex        =   16
      Top             =   8040
      Width           =   4455
   End
   Begin VB.Label SaLo 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3240
      TabIndex        =   15
      Top             =   8040
      Width           =   4095
   End
   Begin VB.Label SaLo 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7560
      TabIndex        =   14
      Top             =   4680
      Width           =   4335
   End
   Begin AlphaImageControl.aicAlphaImage imgPartida 
      Height          =   1815
      Index           =   3
      Left            =   8760
      TabIndex        =   13
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3201
      Image           =   "LoadGameMenu.frx":0000
      Scaler          =   1
   End
   Begin AlphaImageControl.aicAlphaImage imgPartida 
      Height          =   1815
      Index           =   2
      Left            =   4320
      TabIndex        =   12
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3201
      Image           =   "LoadGameMenu.frx":0018
      Scaler          =   1
   End
   Begin AlphaImageControl.aicAlphaImage imgPartida 
      Height          =   1815
      Index           =   1
      Left            =   8760
      TabIndex        =   11
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3201
      Image           =   "LoadGameMenu.frx":0030
      Scaler          =   1
   End
   Begin AlphaImageControl.aicAlphaImage imgPartida 
      Height          =   1815
      Index           =   0
      Left            =   4200
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3201
      Image           =   "LoadGameMenu.frx":0048
      Scaler          =   1
   End
   Begin VB.Label Label12 
      BackColor       =   &H000080FF&
      Height          =   2295
      Left            =   8520
      TabIndex        =   9
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Height          =   2295
      Left            =   4080
      TabIndex        =   8
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Height          =   2295
      Left            =   3960
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label SaLo 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Top             =   4680
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Height          =   2295
      Left            =   8520
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Back 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "<< Volver"
      Height          =   255
      Left            =   14400
      TabIndex        =   4
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label Eliminar 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label FondoEliminar 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   13320
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Titulo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Partidas"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "LoadGameMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Back_Click()
Dim i As Byte

For i = 0 To 3
    SaLo(i).BackColor = &H40C0&
Next i

'Va para atras
If (Titulo.Caption = "Guardar") Then
    Titulo.Caption = "Partidas"
    Eliminar.Visible = True
    FondoEliminar.Visible = True
    Unload LoadGameMenu
    Game.Show
    MenuGame.Show
Else
    If (Not (Game.Enabled) Or Not (LoadGameMenu.Eliminar.Visible)) Then
        Game.Show
        MenuGame.Show
    Else
        MainMenu.Show
    End If
End If

Unload Me

End Sub

Private Sub Eliminar_Click()
Dim i As Byte, carg As Boolean

Me.Enabled = False

i = 1

Do: DoEvents
    
    If (SaLo(i - 1).BackColor = RGB(0, 255, 0)) Then
        carg = HayDatos(i)
        If (carg) Then
            i = 5
           'Elimina el slot que quiere
            MenuS.Show
            MenuS.Titulo.Caption = "¿Desea eliminar su partida?"
        Else
            MenuEspecial 8
            i = 5
        End If
    Else
        i = i + 1
        
        If (i = 5) Then
            Me.Enabled = True
        End If
    
    End If
    
Loop While Not (i = 5)

End Sub

Private Sub Form_Load()
conteoSaves
End Sub

Private Sub SaLo_Click(Index As Integer)
Dim i As Byte

SaLo(Index).BackColor = RGB(0, 255, 0)

For i = 0 To 3
    If Not (i = Index) Then
        SaLo(i).BackColor = &H40C0&
    End If
Next i

End Sub

Private Sub SaLo_DblClick(Index As Integer)

'Guarda o carga
If (Titulo.Caption = "Guardar") Then
    'Guarda
    Salvar Index
Else
    'Carga
    Cargar Index

End If

End Sub
