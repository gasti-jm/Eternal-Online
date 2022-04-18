Attribute VB_Name = "modComprimir"
Option Explicit
Public Function cGRHS()

End Function
Public Function cBodys()

End Function
Public Function cWeapons()
On Local Error Resume Next

    Dim loopC As Long ' // para el for
    Dim MiCabecera As tCabecera
    Dim N As Integer
    Dim MisArmas() As tIndiceArmas
    Dim NumWeaponAnims As Integer
    
    If LenB(Dir(RUTA_INIT & "\Armas.dat", vbArchive)) = 0 Then
        MsgBox "Se requiere Armas.dat en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    NumWeaponAnims = Val(GetVar(RUTA_INIT & "\Armas.dat", "INIT", "NumArmas"))

    ReDim MisArmas(1 To NumWeaponAnims) As tIndiceArmas
    For loopC = 1 To NumWeaponAnims
        MisArmas(loopC).Dir(1) = Val(GetVar(RUTA_INIT & "\Armas.dat", "ARMA" & loopC, "Dir1"))
        MisArmas(loopC).Dir(2) = Val(GetVar(RUTA_INIT & "\Armas.dat", "ARMA" & loopC, "Dir2"))
        MisArmas(loopC).Dir(3) = Val(GetVar(RUTA_INIT & "\Armas.dat", "ARMA" & loopC, "Dir3"))
        MisArmas(loopC).Dir(4) = Val(GetVar(RUTA_INIT & "\Armas.dat", "ARMA" & loopC, "Dir4"))
    Next loopC
    
    If LenB(Dir(RUTA_INIT & "\Armas.dat", vbArchive)) <> 0 Then
        Kill RUTA_INIT & "\Armas.dat"
    End If
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open RUTA_INIT & "\Armas.ind" For Binary As #N
        Put #N, , MiCabecera
        Put #N, , NumWeaponAnims
        
        For loopC = 1 To NumWeaponAnims
            Put #N, , MisArmas(loopC).Dir(1)
            Put #N, , MisArmas(loopC).Dir(2)
            Put #N, , MisArmas(loopC).Dir(3)
            Put #N, , MisArmas(loopC).Dir(4)
        Next loopC
    Close #N
    
    MsgBox "Indexación completada!", vbOKOnly
End Function
Public Function cHelmets()

End Function
Public Function cHeads()

End Function
Public Function cParticles()
'================================================
' ESTO SE VA A HACER LARGO LPM!
' Autor: ZenitraM
' Last Modification: 11/07/2020
'================================================
Dim StreamFile As String
Dim loopC As Long '// para el for
Dim i As Long '// otro for gd.
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
Dim TotalStreams As Integer
Dim MiCabecera As tCabecera '// pablito querido S2
Dim StreamData() As ParticlesStream
Dim N As Integer

    StreamFile = RUTA_INIT & "\Particles.ini"

    '¿Hay archivo?
    If LenB(Dir(StreamFile, vbArchive)) = 0 Then
        MsgBox "Se requiere Particles.ini en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    '// Total num of particles.
    TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
    ReDim StreamData(1 To TotalStreams) As ParticlesStream
    
    For loopC = 1 To TotalStreams
        StreamData(loopC).name = GetVar(StreamFile, Val(loopC), "Name")
        StreamData(loopC).NumOfParticles = GetVar(StreamFile, Val(loopC), "NumOfParticles")
        StreamData(loopC).x1 = GetVar(StreamFile, Val(loopC), "X1")
        StreamData(loopC).y1 = GetVar(StreamFile, Val(loopC), "Y1")
        StreamData(loopC).x2 = GetVar(StreamFile, Val(loopC), "X2")
        StreamData(loopC).y2 = GetVar(StreamFile, Val(loopC), "Y2")
        StreamData(loopC).angle = GetVar(StreamFile, Val(loopC), "Angle")
        StreamData(loopC).vecx1 = GetVar(StreamFile, Val(loopC), "VecX1")
        StreamData(loopC).vecx2 = GetVar(StreamFile, Val(loopC), "VecX2")
        StreamData(loopC).vecy1 = GetVar(StreamFile, Val(loopC), "VecY1")
        StreamData(loopC).vecy2 = GetVar(StreamFile, Val(loopC), "VecY2")
        StreamData(loopC).life1 = GetVar(StreamFile, Val(loopC), "Life1")
        StreamData(loopC).life2 = GetVar(StreamFile, Val(loopC), "Life2")
        StreamData(loopC).friction = GetVar(StreamFile, Val(loopC), "Friction")
        StreamData(loopC).spin = GetVar(StreamFile, Val(loopC), "Spin")
        StreamData(loopC).spin_speedL = GetVar(StreamFile, Val(loopC), "Spin_SpeedL")
        StreamData(loopC).spin_speedH = GetVar(StreamFile, Val(loopC), "Spin_SpeedH")
        StreamData(loopC).AlphaBlend = GetVar(StreamFile, Val(loopC), "AlphaBlend")
        StreamData(loopC).gravity = GetVar(StreamFile, Val(loopC), "Gravity")
        StreamData(loopC).grav_strength = GetVar(StreamFile, Val(loopC), "Grav_Strength")
        StreamData(loopC).bounce_strength = GetVar(StreamFile, Val(loopC), "Bounce_Strength")
        StreamData(loopC).XMove = GetVar(StreamFile, Val(loopC), "XMove")
        StreamData(loopC).YMove = GetVar(StreamFile, Val(loopC), "YMove")
        StreamData(loopC).move_x1 = GetVar(StreamFile, Val(loopC), "move_x1")
        StreamData(loopC).move_x2 = GetVar(StreamFile, Val(loopC), "move_x2")
        StreamData(loopC).move_y1 = GetVar(StreamFile, Val(loopC), "move_y1")
        StreamData(loopC).move_y2 = GetVar(StreamFile, Val(loopC), "move_y2")
        StreamData(loopC).life_counter = GetVar(StreamFile, Val(loopC), "life_counter")
        StreamData(loopC).speed = Val(GetVar(StreamFile, Val(loopC), "Speed"))
        'StreamData(loopc).grh_resize = Val(GetVar(StreamFile, Val(loopc), "resize"))
        'StreamData(loopc).grh_resizex = Val(GetVar(StreamFile, Val(loopc), "rx"))
        'StreamData(loopc).grh_resizey = Val(GetVar(StreamFile, Val(loopc), "ry"))
        StreamData(loopC).NumGrhs = GetVar(StreamFile, Val(loopC), "NumGrhs")
       
        ReDim StreamData(loopC).grh_list(1 To StreamData(loopC).NumGrhs)
        GrhListing = GetVar(StreamFile, Val(loopC), "Grh_List")
       
        For i = 1 To StreamData(loopC).NumGrhs
            StreamData(loopC).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
            StreamData(loopC).grh_list(i - 1) = StreamData(loopC).grh_list(i - 1)
            
        For ColorSet = 1 To 4
            TempSet = GetVar(StreamFile, Val(loopC), "ColorSet" & ColorSet)
            StreamData(loopC).colortintR = General_Field_Read(1, TempSet, 44)
            StreamData(loopC).colortintG = General_Field_Read(2, TempSet, 44)
            StreamData(loopC).colortintB = General_Field_Read(3, TempSet, 44)
        Next ColorSet
    Next loopC
    
    If LenB(Dir(StreamFile, vbArchive)) <> 0 Then
        Kill StreamFile
    End If
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open RUTA_INIT & "\Particles.ind" For Binary As #N
        Put #N, , MiCabecera
        Put #N, , TotalStreams

        For loopC = 1 To TotalStreams
            Put #N, , StreamData(loopC).name
            Put #N, , StreamData(loopC).NumOfParticles
            Put #N, , StreamData(loopC).x1
            Put #N, , StreamData(loopC).y1
            Put #N, , StreamData(loopC).x2
            Put #N, , StreamData(loopC).y2
            Put #N, , StreamData(loopC).angle
            Put #N, , StreamData(loopC).vecx1
            Put #N, , StreamData(loopC).vecx2
            Put #N, , StreamData(loopC).vecy1
            Put #N, , StreamData(loopC).vecy2
            Put #N, , StreamData(loopC).life1
            Put #N, , StreamData(loopC).life2
            Put #N, , StreamData(loopC).friction
            Put #N, , StreamData(loopC).spin
            Put #N, , StreamData(loopC).spin_speedL
            Put #N, , StreamData(loopC).spin_speedH
            Put #N, , StreamData(loopC).AlphaBlend
            Put #N, , StreamData(loopC).gravity
            Put #N, , StreamData(loopC).grav_strength
            Put #N, , StreamData(loopC).bounce_strength
            Put #N, , StreamData(loopC).XMove
            Put #N, , StreamData(loopC).YMove
            Put #N, , StreamData(loopC).move_x1
            Put #N, , StreamData(loopC).move_x2
            Put #N, , StreamData(loopC).move_y1
            Put #N, , StreamData(loopC).move_y2
            Put #N, , StreamData(loopC).life_counter
            Put #N, , StreamData(loopC).speed
            'Put #N, , StreamData(loopc).grh_resize
            'Put #N, , StreamData(loopc).grh_resizex
            'Put #N, , StreamData(loopc).grh_resizey
            Put #N, , StreamData(loopC).NumGrhs
       
            'ReDim StreamData(loopC).grh_list(1 To StreamData(loopC).NumGrhs)
            Put #N, , GrhListing
       
            For i = 1 To StreamData(loopC).NumGrhs
                Put #N, , StreamData(loopC).grh_list(i)
            Next i
                'StreamData(loopC).grh_list(i - 1) = StreamData(loopC).grh_list(i - 1)
                
            For ColorSet = 1 To 4
                Put #N, , TempSet
                Put #N, , StreamData(loopC).colortintR
                Put #N, , StreamData(loopC).colortintG
                Put #N, , StreamData(loopC).colortintB
            Next ColorSet
        Next loopC
    Close #N
    
    MsgBox "Indexación completada!", vbOKOnly

End Function
Public Function cFXs()

End Function
Public Function cShields()
On Local Error Resume Next

    Dim loopC As Long ' // para el for
    Dim MiCabecera As tCabecera
    Dim N As Integer
    Dim MisEscudos() As tIndiceEscudos
    Dim NumShieldAnims As Integer
    
    If LenB(Dir(RUTA_INIT & "\Escudos.dat", vbArchive)) = 0 Then
        MsgBox "Se requiere Escudos.dat en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    NumShieldAnims = Val(GetVar(RUTA_INIT & "\Escudos.dat", "INIT", "NumEscudos"))

    ReDim MisEscudos(1 To NumShieldAnims) As tIndiceEscudos
    For loopC = 1 To NumShieldAnims
        MisEscudos(loopC).Dir(1) = Val(GetVar(RUTA_INIT & "\Escudos.dat", "ESC" & loopC, "Dir1"))
        MisEscudos(loopC).Dir(2) = Val(GetVar(RUTA_INIT & "\Escudos.dat", "ESC" & loopC, "Dir2"))
        MisEscudos(loopC).Dir(3) = Val(GetVar(RUTA_INIT & "\Escudos.dat", "ESC" & loopC, "Dir3"))
        MisEscudos(loopC).Dir(4) = Val(GetVar(RUTA_INIT & "\Escudos.dat", "ESC" & loopC, "Dir4"))
    Next loopC
    
    If LenB(Dir(RUTA_INIT & "\Escudos.dat", vbArchive)) <> 0 Then
        Kill RUTA_INIT & "\Escudos.dat"
    End If
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open RUTA_INIT & "\Escudos.ind" For Binary As #N
        Put #N, , MiCabecera
        Put #N, , NumShieldAnims
        
        For loopC = 1 To NumShieldAnims
            Put #N, , MisEscudos(loopC).Dir(1)
            Put #N, , MisEscudos(loopC).Dir(2)
            Put #N, , MisEscudos(loopC).Dir(3)
            Put #N, , MisEscudos(loopC).Dir(4)
        Next loopC
    Close #N
    
    MsgBox "Indexación completada!", vbOKOnly
End Function
Public Function cCFG()
On Local Error Resume Next
    If LenB(Dir(RUTA_INIT & "\Config.ini", vbArchive)) = 0 Then
        MsgBox "Se requiere Config.ini en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim MiCabecera As tCabecera
    Dim Config_Inicio As tGameIni
    Dim N As Integer
    
    Config_Inicio.CursorGraphic = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "CursorGraphic")
    Config_Inicio.ResolutionX = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "ResolutionX")
    Config_Inicio.ResolutionY = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "ResolutionY")
    Config_Inicio.FullScreen = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "FullScreen")
    Config_Inicio.Sounds = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "Sounds")
    Config_Inicio.Music = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "Music")
    Config_Inicio.SoundVolume = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "SoundVolume")
    Config_Inicio.MusicVolume = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "MusicVolume")
    Config_Inicio.VSYNC = GetVar(RUTA_INIT & "\Config.ini", "GameCFG", "VSYNC")
    
    If LenB(Dir(RUTA_INIT & "\Config.ind", vbArchive)) <> 0 Then
        Kill RUTA_INIT & "\Config.ind"
    End If
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open RUTA_INIT & "\Config.ind" For Binary As #N
    Put #N, , MiCabecera
    Put #N, , Config_Inicio
    Close #N
    
    MsgBox "Indexación completada!", vbOKOnly
End Function
