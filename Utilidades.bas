Attribute VB_Name = "Utilidades"
Option Explicit

'Estructura de datos para Cargar y Salvar
Public Type DatosSave
    hrsJugado As Byte
    minJugado As Byte
    segJugado As Byte
    nroLineaTexto As Integer
    numCap As Byte
    subCap As String * 20
    imFondo As String * 15
    imActor1 As String * 15
    imActor2 As String * 15
    lnMusica As Byte
    lnSonido As Byte
End Type

'Estructura de datos para el control de Imagenes
Public Type DatosImagenes
    nroCap As Byte
    ID_Image As String * 2
    urlImage As String * 15
End Type

'Estructura de datos para el control de Sonidos
Public Type DatosSonidos
    nroCap As Byte
    ID_Sound As String * 2
    urlSound As String * 15
End Type

'Estructura de datos para los Extras
Public Type DatosExtras
    nroCap As Byte
    tipoArch As String * 1
    urlArchivo As String * 15
    subCap As String * 20
End Type


'########################################
'Estructura de datos para las opciones ##
Public Type DatosOpciones             '##
    contAdulto As Byte                '##
    fScreen As Byte                   '##
    volFondo As Byte                  '##
    volFX As Byte                     '##
    VelTexto As Integer               '##
End Type                              '##
'########################################

'Estructura de datos para el Texto
Public Type DataTexto
    charPersonaje As String * 10
    nroReg As Integer
    txtTexto As String * 64
End Type


'###############################################################
'### Variable publica necesaria para el control de los slots ###
'### y otra para el juego cargado.                           ###
'###############################################################

Public bCargado As Boolean
Public txtPerChat As String
'###############################################################
'###############################################################


Public Function HayDatos(ByVal x As Byte) As Boolean
Dim dato As DatosSave

Open App.Path & "\saves\Slot" & Trim(Str(x)) & ".sav" For Random As #3 Len = Len(dato)
    dato.nroLineaTexto = LOF(3) / Len(dato)
    
    If (dato.nroLineaTexto = 0) Then
        Close #3
        Kill App.Path & "\saves\Slot" & Trim(Str(x)) & ".sav"
        HayDatos = False
    Else
        Close #3
        HayDatos = True
    End If
End Function

Public Sub Juego()

End Sub

Public Sub Load_Animation()

End Sub

'
'
'
Public Sub Load_FX(ByRef Cap As Byte, Optional ByRef lnFX As Byte)
Static regFX As Byte

If (lnFX <> 0) Then

End If

End Sub

'
'Carga de Imagenes
'
Public Sub Load_Imagenes(ByRef Cap As Byte, ByRef tImagen As String)


End Sub


'
'Tipo de Sonido
'
Public Sub Load_Sonido(ByRef nCap As Byte, ByRef tSonido As String)

'Segun el Sonido
If (tSonido = "fx") Then
    Load_FX nCap
Else
    
End If

End Sub

Public Sub CambioPantalla(n As Byte)
If (n = 0) Then
    
Else

End If

End Sub

'Segun el paramtro carga o salva las opciones
Public Sub LoadSave_Opciones(ByVal L_S As String)
Dim optMias As DatosOpciones

If (L_S = "L") Then
    Open App.Path & "\controls\Opciones.cnf" For Random As #65 Len = Len(optMias)
        Get #65, 1, optMias
        
        'Añado las opciones a mis nuevos
        OptionsMenu.ContAdult.Value = optMias.contAdulto
        OptionsMenu.musicaVol.Value = optMias.volFondo
        OptionsMenu.FXVol.Value = optMias.volFX
        OptionsMenu.VelTexto.Value = optMias.VelTexto
        OptionsMenu.FullScreen.Value = optMias.fScreen
        
        CambioPantalla (optMias.fScreen)
        
    Close #65
Else
    Open App.Path & "\controls\Opciones.cnf" For Random As #66 Len = Len(optMias)
        'Levanto los cambios
        optMias.contAdulto = OptionsMenu.ContAdult.Value
        optMias.volFondo = OptionsMenu.musicaVol.Value
        optMias.volFX = OptionsMenu.FXVol.Value
        optMias.VelTexto = OptionsMenu.VelTexto.Value
        optMias.fScreen = OptionsMenu.FullScreen.Value
        
        'Y cambio en mi programa
        Game.TimeTexto.Interval = optMias.VelTexto
        Game.pMusica.settings.volume = optMias.volFondo
        Game.pFX.settings.volume = optMias.volFX
        
        Put #66, 1, optMias
        
        CambioPantalla (optMias.fScreen)
        
    Close #66
End If

End Sub

'
'Carga los extras
'
'##################################################################
'### Funcionamiento: a medida que el jugador vaya desbloqueando ###
'### nuevos extras (imagenes, sonidos, animaciones, historia),  ###
'### estos se iran guardando en el archivo Mis_Extras.xtr       ###
'### o se cargaran. Esto de decide por el parametro opcional de ###
'### bNuevo, q tomando los valores de "True" guardara un nuevo  ###
'### extra desbloqueado.                                        ###
'##################################################################
'
Public Sub Cargar_Extras(Optional ByVal bNuevo As Boolean = False)
Dim misExtras As DatosExtras, i As Integer

If Not (bNuevo) Then
    
    i = 1
    
    'Carga los extras
    Open App.Path & "\controls\Mis_Extras.xtr" For Random As #25 Len = Len(misExtras)
        Do: DoEvents
            Get #25, i, misExtras
            
            If Not (EOF(25)) Then
                Select Case Trim(misExtras.tipoArch)
                    Case Is = "#"
                    Case Is = "$"
                    Case Is = "!"
                    Case Is = "%"
                End Select
                
                i = i + 1
                
            End If
            
        Loop While Not (EOF(25))
    Close #25

Else

    'Guarda un nuevo extra
    Open App.Path & "\controls\Mis_Extras.xtr" For Random As #25 Len = Len(misExtras)
        i = LOF(25) / Len(misExtras) + 1
    Close #25

End If

End Sub


'Carga los archivos llenos y los no llenos
'
'#############################################################
'### Funcionamiento: Determina los archivos de partida que ###
'### tienen algo en su interior y luego muestra la         ###
'### informacion importante de esa partida en los labels   ###
'###                   de los slots.                       ###
'#############################################################
'
Public Sub conteoSaves()

Dim i As Byte, Datos As DatosSave

For i = 1 To 4
    
    Open App.Path & "\saves\Slot" & Trim(Str(i)) & ".sav" For Random As #30 Len = Len(Datos)
        If (LOF(30) / Len(Datos) <> 0) Then
            Get #30, 1, Datos
            LoadGameMenu.SaLo(i - 1).FontSize = 11
            LoadGameMenu.SaLo(i - 1).Caption = Str(Datos.hrsJugado) & " : " & Str(Datos.minJugado) & " : " & Str(Datos.segJugado) & "  Capitulo: " & Str(Datos.numCap) & " / " & Trim(Datos.subCap)
        End If
    Close #30
    
    If LoadGameMenu.SaLo(i - 1).Caption = "Empty" Then
        LoadGameMenu.SaLo(i - 1).FontSize = 14
        Kill (App.Path & "\saves\Slot" & Trim(Str(i)) & ".sav")
    End If
    
Next i

End Sub

Public Sub MenuEspecial(ByRef x As Byte)

MenuS.Show

Select Case (x)
    'Salir
    Case Is = 1
        MenuS.Titulo.Caption = "¿Desea salir de Clean My Mess?"
    
    'Overwrite de partida guardada
    Case Is = 2
        LoadGameMenu.Enabled = False
        MenuS.Titulo.Caption = "¿Desea sobreescribir la partida?"
        
    'Eliminacion de partida
    Case Is = 3
        LoadGameMenu.Enabled = False
        MenuS.Titulo.Caption = "¿Desea eliminar su partida?"
    
    'Reseteo del juego
    Case Is = 4
        OptionsMenu.Enabled = False
        MenuS.Titulo.Caption = "Se volveran los valores a 0. ¿Seguir?"
    
    'Finalizar proceso
    Case Is = 5
        '#################### Aspecto de Menu ######################
        LoadGameMenu.Enabled = False                            '###
        MenuS.Titulo.Caption = "Partida guardada con exito!!"   '###
        MenuS.No.Visible = False                                '###
        MenuS.Si.Caption = "Continuar"                          '###
        MenuS.Si.Width = 2415                                   '###
        MenuS.Si.Left = 2760                                    '###
        '###########################################################
        
    Case Is = 6
        'Nuevo juego dentro de uno cargado
        MainMenu.Enabled = False
        MenuS.Show
        MenuS.Titulo.Caption = "¿Empezar otro juego?"
    
    Case Is = 7
        'Volver al MainMenu
        MenuS.Show
        Unload MenuGame
        MenuS.Titulo.Caption = "¿Desea ir al Main Menu?"
    
    Case Is = 8
        MenuS.Titulo.Caption = "No hay datos guardados"   '###
        MenuS.No.Visible = False                                '###
        MenuS.Si.Caption = "Continuar"                          '###
        MenuS.Si.Width = 2415                                   '###
        MenuS.Si.Left = 2760
    
        
End Select

End Sub


'
'Sobreescribir
'
'########################################################################
'### Funcionamiento: Sobreescribe un archivo de partida (Slot1.sav,   ###
'### etc) Siempre y cuando exista el archivo y contega algo.          ###
'########################################################################
'
Public Sub Reescribir(ByVal x As Single)
Dim misDatos As DatosSave, misDatosSal As DatosSave

Open App.Path & "\saves\Slot" & Trim(Str(x)) & ".sav" For Random As #13 Len = Len(misDatos)
    Open App.Path & "\saves\Slot" & Trim(Str(x)) & ".tmp" For Random As #14 Len = Len(misDatos)
        
        Get #13, 1, misDatos

      '###################### Guardar Tiempo ###################
        misDatosSal.hrsJugado = misDatos.hrsJugado + Game.hora
        misDatosSal.minJugado = misDatos.minJugado + Game.min
        misDatosSal.segJugado = misDatos.segJugado + Game.seg
        If (misDatosSal.segJugado >= 60) Then
            misDatosSal.segJugado = misDatosSal.segJugado - 60
            misDatosSal.minJugado = misDatosSal.minJugado + 1
        End If
        If (misDatosSal.minJugado >= 60) Then
            misDatosSal.minJugado = misDatosSal.minJugado - 60
            misDatosSal.hrsJugado = misDatosSal.hrsJugado + 1
        End If
      '########################################################
         
        misDatosSal.numCap = x
        Put #14, 1, misDatosSal
        
    Close #13
   
   '################## Renombre y eliminacion del archivo ####################################################
    Kill App.Path & "\saves\Slot" & Trim(Str(x)) & ".sav"
   '##########################################################################################################

Close #14

Name App.Path & "\saves\Slot" & Trim(Str(x)) & ".tmp" As App.Path & "\saves\Slot" & Trim(Str(x)) & ".sav"

End Sub
'
'Guardar partida
'
'##########################################################
'### Funcionamiento: Si no existe un archivo de partida ###
'### (Slot1.sav, Slot2.sav, etc), este lo crea y guarda ###
'###                  ¡NO SOBREESCRIBE!                 ###
'##########################################################
'
Public Sub Salvar(ByVal x As Byte)
Dim Datos As DatosSave

Open App.Path & "\saves\Slot" & Trim(Str(x + 1)) & ".sav" For Random As #48 Len = Len(Datos)
    Datos.nroLineaTexto = LOF(48) / Len(Datos)
    If (Datos.nroLineaTexto = 0) Then
        'Nueva partida
        Put #48, 1, Datos
        
        Close #48
        conteoSaves
        
        Unload LoadGameMenu
        Game.Show
        MenuGame.Show
        
    Else
        
        'Sobre-escribir
        MenuEspecial 2
        Close #48
    
    End If
    
End Sub
'
'Cargar las partidas y empezar a jugar
'
'
'#####################################################################
'### Funcionamiento: abre los archivos de partida (Slot1.sav, etc) ###
'###     carga los datos en la partida, luego inicia el juego.     ###
'#####################################################################
'
Public Sub Cargar(ByVal x As Byte)
Dim parDatos As DatosSave

If Not (LoadGameMenu.SaLo(x).Caption = "Empty") Then
    Open App.Path & "\saves\Slot" & Trim(Str(x + 1)) & ".sav" For Random As #34
        
        Get #34, 1, parDatos
        
        'Cargo mis datos a mi Form del Juego
        Game.TiempoJugado.Caption = Str(parDatos.hrsJugado) & " : " & Str(parDatos.minJugado) & " : " & Str(parDatos.segJugado)
        Game.imgActor1.LoadImage_FromFile App.Path & "\archived\images\cap" & Trim(Str(parDatos.numCap)) & "\" & Trim(parDatos.imFondo)
        Game.imgActor1.LoadImage_FromFile App.Path & "\archived\images\cap" & Trim(Str(parDatos.numCap)) & "\" & Trim(parDatos.imActor1)
        Game.imgActor1.LoadImage_FromFile App.Path & "\archived\images\cap" & Trim(Str(parDatos.numCap)) & "\" & Trim(parDatos.imActor2)
    Close #34
    
    Game.Show
    Unload MenuGame
    Unload MenuS
    Unload LoadGameMenu
    Game.Enabled = True
    Game.TimeJuego.Enabled = True
    
Else
    '######## #########
    MenuS.Show
    MenuS.Titulo.Caption = "No hay datos guardados"
    MenuS.Si.Caption = "Continuar"
    MenuS.Si.Width = 2415
    MenuS.Si.Left = 2760
    MenuS.No.Visible = False
    LoadGameMenu.Enabled = False
    '######## #########
End If

End Sub



Public Sub Yes_No(ByVal nroMenu As Byte)
Dim i As Byte

Select Case (nroMenu)
    'Sobreescribir el juego
    Case Is = 2
        i = 0
        Do: DoEvents
            If (LoadGameMenu.SaLo(i).BackColor = RGB(0, 255, 0)) Then
                Reescribir (i + 1)
                i = 4
            Else
                i = i + 1
            End If
        Loop While Not (i = 4)
        Unload MenuS
        Unload LoadGameMenu
        Game.Show
        MenuGame.Show
    'Eliminar Partida
    Case Is = 3
        i = 0
        Do: DoEvents
            If (LoadGameMenu.SaLo(i).BackColor = RGB(0, 255, 0)) Then
                Kill App.Path & "\saves\Slot" & Trim(Str(i + 1)) & ".sav"
                LoadGameMenu.SaLo(i).Caption = "Empty"
                LoadGameMenu.SaLo(i).FontSize = 14
                i = 4
            Else
                i = i + 1
            End If
        Loop While Not (i = 4)
        Unload MenuS
        LoadGameMenu.Enabled = True
        LoadGameMenu.Show
    'Reseteo del juego
    Case Is = 4
    
    
    'Partida sobreescrita
    Case Is = 5
        '###### Aspecto de menu #######
        MenuS.Hide                  '##
        LoadGameMenu.Enabled = True '##
        Game.Show                   '##
        MenuS.Si.Left = 2040        '##
        MenuS.Si.Width = 975        '##
        MenuS.Si.Caption = "Si"     '##
        MenuS.No.Visible = True     '##
        LoadGameMenu.Hide           '##
        '##############################
    
    'Nuevo juego
    Case Is = 6
        MainMenu.Enabled = True
        

End Select

End Sub


