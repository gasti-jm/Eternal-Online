Attribute VB_Name = "modGeneral"
Option Explicit
'Lectura y escritura para los bloc de notas
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public RUTA_INIT As String '// Ruta especifica.
Public PROCESS_SELECTED As Byte '// Que elegimos?

' Cabecera - Argentum Online ==========================================
Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type
'======================================================================

'// Graficos.ind
Public Type GrhData
    sx As Integer
    sy As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Long
    speed As Single
    active As Boolean
End Type
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    'HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Public Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Long
End Type

'Armas.ind
Public Type tIndiceArmas
    Dir(1 To 4) As Integer
End Type
'Escudos.ind
Public Type tIndiceEscudos
    Dir(1 To 4) As Integer
End Type

'Lista de las animaciones de los escudos
Public Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Particles.ind
Public Type ParticlesStream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortintR As Byte
    colortintG As Byte
    colortintB As Byte
   
    speed As Single
    life_counter As Long
End Type

'Config.ind
Public Type tGameIni
    ResolutionX As Long
    ResolutionY As Long
    AccountName As String '// por seguridad guardamos unicamente el nombre.
    FullScreen As Boolean '¿Estoy en pantalla completa?
    Sounds As Boolean
    Music As Boolean
    SoundVolume As Byte
    MusicVolume As Byte
    CursorGraphic As Boolean
    VSYNC As Byte
End Type

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub
Public Sub Comprimir(Process As Byte)
    Select Case Process
        Case 1
            Call cGRHS
        Case 2
            Call cBodys
        Case 3
            Call cWeapons
        Case 4
            Call cHelmets
        Case 5
            Call cHeads
        Case 6
            Call cParticles
        Case 7
            Call cFXs
        Case 8
            Call cShields
        Case 9
            Call cCFG
        Case Else
            MsgBox "ERROR: NO SE HA SELECIONADO EL TIPO DE COMPRESION", vbCritical, "ERROR"
            Exit Sub
    End Select
End Sub
Public Sub Descomprimir(Process As Byte)
        Select Case Process
        Case 1
            Call dGRHS
        Case 2
            Call dBodys
        Case 3
            Call dWeapons
        Case 4
            Call dHelmets
        Case 5
            Call dHeads
        Case 6
            Call dParticles
        Case 7
            Call dFXs
        Case 8
            Call dShields
        Case 9
            Call dCFG
        Case Else
            MsgBox "ERROR: NO SE HA SELECIONADO EL TIPO DE COMPRESION", vbCritical, "ERROR"
            Exit Sub
    End Select
End Sub
Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
    writeprivateprofilestring Main, var, value, file
End Sub
Function GetVar(ByVal file As String, ByVal Main As String, ByVal var As String) As String
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function
Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
   
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(Mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = Mid$(Text, LastPos + 1)
    End If
End Function
