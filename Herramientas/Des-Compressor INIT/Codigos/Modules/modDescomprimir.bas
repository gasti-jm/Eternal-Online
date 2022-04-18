Attribute VB_Name = "modDescomprimir"
Option Explicit
Public Function dGRHS()

End Function
Public Function dBodys()

End Function
Public Function dWeapons()
'==================================
' Descomprimimos las armas
'==================================
On Local Error Resume Next
    If LenB(Dir(RUTA_INIT & "\Armas.ind", vbArchive)) = 0 Then
        MsgBox "Se requiere Armas.ind en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim LoopC As Long
    Dim MiCabecera As tCabecera
    Dim MisArmas() As tIndiceArmas
    Dim NumWeaponAnims As Integer
    Dim N As Integer
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open RUTA_INIT & "\Armas.ind" For Binary As #N
        Get #N, , MiCabecera
        Get #N, , NumWeaponAnims
        ReDim MisArmas(1 To NumWeaponAnims) As tIndiceArmas '// NAZI
        For LoopC = 1 To NumWeaponAnims
            Get #N, , MisArmas(LoopC).Dir(1)
            Get #N, , MisArmas(LoopC).Dir(2)
            Get #N, , MisArmas(LoopC).Dir(3)
            Get #N, , MisArmas(LoopC).Dir(4)
        Next LoopC
    Close #N
    
    If LenB(Dir(RUTA_INIT & "\Armas.dat", vbArchive)) <> 0 Then
        Kill RUTA_INIT & "\Armas.dat"
    End If
    
    Call WriteVar(RUTA_INIT & "\Armas.dat", "INIT", "NumArmas", NumWeaponAnims)
    For LoopC = 1 To NumWeaponAnims
        Call WriteVar(RUTA_INIT & "\Armas.dat", "ARMA" & LoopC, "Dir1", MisArmas(LoopC).Dir(1))
        Call WriteVar(RUTA_INIT & "\Armas.dat", "ARMA" & LoopC, "Dir2", MisArmas(LoopC).Dir(2))
        Call WriteVar(RUTA_INIT & "\Armas.dat", "ARMA" & LoopC, "Dir3", MisArmas(LoopC).Dir(3))
        Call WriteVar(RUTA_INIT & "\Armas.dat", "ARMA" & LoopC, "Dir4", MisArmas(LoopC).Dir(4))
    Next LoopC
    
    MsgBox "Extracción completada!", vbOKOnly
End Function
Public Function dHelmets()

End Function
Public Function dHeads()

End Function
Public Function dParticles()

End Function
Public Function dFXs()

End Function
Public Function dShields()
'==================================
' Descomprimimos los escudos
'==================================
On Local Error Resume Next

    If LenB(Dir(RUTA_INIT & "\Escudos.ind", vbArchive)) = 0 Then
        MsgBox "Se requiere Escudos.ind en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim LoopC As Long
    Dim MiCabecera As tCabecera
    Dim MisEscudos() As tIndiceEscudos
    Dim NumShieldAnims As Integer
    Dim N As Integer
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open RUTA_INIT & "\Escudos.ind" For Binary As #N
        Get #N, , MiCabecera
        Get #N, , NumShieldAnims
        ReDim MisEscudos(1 To NumShieldAnims) As tIndiceEscudos '// NAZI
        For LoopC = 1 To NumShieldAnims
            Get #N, , MisEscudos(LoopC).Dir(1)
            Get #N, , MisEscudos(LoopC).Dir(2)
            Get #N, , MisEscudos(LoopC).Dir(3)
            Get #N, , MisEscudos(LoopC).Dir(4)
        Next LoopC
    Close #N
    
    If LenB(Dir(RUTA_INIT & "\Escudos.dat", vbArchive)) <> 0 Then
        Kill RUTA_INIT & "\Escudos.dat"
    End If
    
    Call WriteVar(RUTA_INIT & "\Escudos.dat", "INIT", "NumEscudos", NumShieldAnims)
    For LoopC = 1 To NumShieldAnims
        Call WriteVar(RUTA_INIT & "\Escudos.dat", "ESC" & LoopC, "Dir1", MisEscudos(LoopC).Dir(1))
        Call WriteVar(RUTA_INIT & "\Escudos.dat", "ESC" & LoopC, "Dir2", MisEscudos(LoopC).Dir(2))
        Call WriteVar(RUTA_INIT & "\Escudos.dat", "ESC" & LoopC, "Dir3", MisEscudos(LoopC).Dir(3))
        Call WriteVar(RUTA_INIT & "\Escudos.dat", "ESC" & LoopC, "Dir4", MisEscudos(LoopC).Dir(4))
    Next LoopC
    
    MsgBox "Extracción completada!", vbOKOnly
End Function
Public Function dCFG()
On Local Error Resume Next
    If LenB(Dir(RUTA_INIT & "\Config.ind", vbArchive)) = 0 Then
        MsgBox "Se requiere Config.ind en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim MiCabecera As tCabecera
    Dim Config_Inicio As tGameIni
    Dim N As Integer
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open RUTA_INIT & "\Config.ind" For Binary As #N
        Get #N, , MiCabecera
        Get #N, , Config_Inicio
    Close #N
    
    If LenB(Dir(RUTA_INIT & "\Config.ini", vbArchive)) <> 0 Then
        Kill RUTA_INIT & "\Config.ini"
    End If
    
    ' CONFIG
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "CursorGraphic", Config_Inicio.CursorGraphic)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "ResolutionX", Config_Inicio.ResolutionX)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "ResolutionY", Config_Inicio.ResolutionY)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "FullScreen", Config_Inicio.FullScreen)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "Sounds", Config_Inicio.Sounds)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "Music", Config_Inicio.Music)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "SoundVolume", Config_Inicio.SoundVolume)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "MusicVolume", Config_Inicio.MusicVolume)
    Call WriteVar(RUTA_INIT & "\Config.ini", "GameCFG", "VSYNC", Config_Inicio.VSYNC)
    
    MsgBox "Extracción completada!", vbOKOnly
End Function
