Attribute VB_Name = "Modo_0_13_0"
Option Explicit
Private Type tCabecera0130 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type
Private Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Long
    Speed As Single
End Type
Private Type tIndiceCabeza
    Head(1 To 4) As Long
End Type
Private Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type
Private Type tIndiceFx
    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer
End Type
Private GrhData As GrhData
Private MiCabecera0130 As tCabecera0130

Private Sub IniciarCabecera(ByRef Cabecera As tCabecera0130)
    Cabecera.desc = "Argentum Online by Noland Studios. Des-Indexador Universal (c) GS-Zone 2021 - http://www.gs-zone.org"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function Indexar0130() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Handle As Integer
    Dim grhCount As Long
    Dim fileVersion As Long
    Dim Leer As New clsIniReader
    Dim Datos As String
    Dim DatoR() As String
    Dim tF As Integer
    
    Indexar0130 = False
    Handle = FreeFile()

    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    grhCount = Val(Leer.GetValue("INIT", "NumGrh"))
    fileVersion = Val(Leer.GetValue("INIT", "Version"))
    If (fileVersion = 0) Then
        'MsgBox "El valor de 'Version' es invalido!", vbCritical
        'Exit Function
        fileVersion = 1
    ElseIf (grhCount > 200000 Or grhCount <= 0) Then
        MsgBox "La valor de 'NumGrh' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents
    
    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    
    Seek Handle, 1
    
    Put Handle, , fileVersion
    Put Handle, , grhCount
    
    For Grh = 1 To grhCount
        GrhData.sX = 0
        GrhData.sY = 0
        GrhData.pixelWidth = 0
        GrhData.pixelHeight = 0
        GrhData.FileNum = 0
        GrhData.NumFrames = 0
        GrhData.Speed = 0
        'Erase GrhData.Frames()
        
        Datos = Leer.GetValue("Graphics", "Grh" & Grh)
        If LenB(Datos) <> 0 Then
            DatoR() = Split(Datos, "-")
            If DatoR(0) > 1 Then
                Put Handle, , Grh
                GrhData.NumFrames = Val(DatoR(0))
                Put Handle, , GrhData.NumFrames
                ReDim GrhData.Frames(1 To GrhData.NumFrames)
                tF = 1
                While Not GrhData.NumFrames < tF
                    GrhData.Frames(tF) = Val(DatoR(tF))
                    Put Handle, , GrhData.Frames(tF)
                    tF = tF + 1
                Wend
                GrhData.Speed = Val(DatoR(tF))
                Put Handle, , GrhData.Speed
            ElseIf DatoR(0) = 1 Then
                Put Handle, , Grh
                GrhData.NumFrames = Val(DatoR(0))
                Put Handle, , GrhData.NumFrames
                GrhData.FileNum = Val(DatoR(1))
                Put Handle, , GrhData.FileNum
                GrhData.sX = Val(DatoR(2))
                Put Handle, , GrhData.sX
                GrhData.sY = Val(DatoR(3))
                Put Handle, , GrhData.sY
                GrhData.pixelWidth = Val(DatoR(4))
                Put Handle, , GrhData.pixelWidth
                GrhData.pixelHeight = Val(DatoR(5))
                Put Handle, , GrhData.pixelHeight
            End If
        End If
    Next
    
    Close Handle
    
    Indexar0130 = True
Exit Function

ErrorHandler:
    
End Function

Public Function Desindexar0130() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim Handle As Integer
    Dim handleW As Integer
    Dim fileVersion As Long
    Dim Datos As String
    
    Desindexar0130 = False
    'Open files
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    Open IniPathD & GraphicsFileD For Binary Access Write As handleW
    
    Seek Handle, 1
    
    'Get file version
    Get Handle, , fileVersion
    
    'Get number of grhs
    Get Handle, , grhCount
    If (grhCount > 200000 Or grhCount <= 0 Or fileVersion < 0) Then
        MsgBox "Indice incompatible!", vbCritical
        Desindexar0130 = False
        Close handleW
        Close Handle
        Exit Function
    End If
    'Resize arrays
    'ReDim GrhData(1 To grhCount) As GrhData
    
    Put handleW, , "[INIT]" & vbCrLf & "NumGrh=" & grhCount & vbCrLf & "Version=" & fileVersion & vbCrLf & vbCrLf
    Put handleW, , "[Graphics]" & vbCrLf

    Get Handle, , Grh
    While Not EOF(Handle) And (Grh <> 0 And Grh <= grhCount)
        With GrhData
            'Get number of frames
            Get Handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            Datos$ = ""
            ReDim .Frames(1 To GrhData.NumFrames)
            
            If .NumFrames > 1 Then
                Datos$ = CStr(.NumFrames)
            
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get Handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                    Datos$ = Datos$ & "-" & CStr(.Frames(Frame))
                Next Frame
                
                Get Handle, , .Speed
                If .Speed <= 0 Then GoTo ErrorHandler
                
                Datos$ = Datos$ & "-" & CStr(.Speed)
            Else
                'Read in normal GRH data
                Get Handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get Handle, , .sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get Handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get Handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get Handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
    
                .Frames(1) = Grh
                Datos$ = "1-" & CStr(.FileNum) & "-" & CStr(.sX) & "-" & CStr(.sY) & "-" & CStr(.pixelWidth) & "-" & CStr(.pixelHeight)
            End If
        End With
        If LenB(Datos$) <> 0 Then
            Put handleW, , "Grh" & CStr(Grh) & "=" & Datos$ & vbCrLf
        End If
        GrhData.FileNum = 0
        GrhData.NumFrames = 0
        Get Handle, , Grh
    Wend
    
    Close Handle
    Close handleW
    
    Desindexar0130 = True
Exit Function

ErrorHandler:
    
End Function


Public Function DesindexarCabezas0130() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Long
    Dim Numheads As Integer
    Dim MisCabezas As tIndiceCabeza
    
    DesindexarCabezas0130 = False
    Call IniciarCabecera(MiCabecera0130)
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    If (NoUsarCabecera = False) Then Get Handle, , MiCabecera0130
    Get Handle, , Numheads
    
    If Numheads <= 0 Then
        MsgBox "Indice incompatible!", vbCritical
        Close Handle
        Exit Function
    End If
    
    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents
    Open IniPathD & GraphicsFileD For Binary Access Write As handleW
    Put handleW, , "[INIT]" & vbCrLf & "NumHeads=" & Numheads & vbCrLf & vbCrLf

    For I = 1 To Numheads
        Get Handle, , MisCabezas
        Put handleW, , "[HEAD" & I & "]" & vbCrLf
        Put handleW, , "Head1=" & MisCabezas.Head(1) & vbTab & " ' arriba" & vbCrLf
        Put handleW, , "Head2=" & MisCabezas.Head(2) & vbTab & " ' derecha" & vbCrLf
        Put handleW, , "Head3=" & MisCabezas.Head(3) & vbTab & " ' abajo" & vbCrLf
        Put handleW, , "Head4=" & MisCabezas.Head(4) & vbTab & " ' izq" & vbCrLf & vbCrLf
    Next I
    Close Handle
    Close handleW
    
    DesindexarCabezas0130 = True
Exit Function

ErrorHandler:
End Function

Public Function IndexarCabezas0130() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Long
    Dim Numheads As Integer
    Dim MisCabezas As tIndiceCabeza
    Dim Leer As New clsIniReader

    
    IndexarCabezas0130 = False
    Call IniciarCabecera(MiCabecera0130)
    Handle = FreeFile()
    
    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    Numheads = Val(Leer.GetValue("INIT", "NumHeads"))
    If (Numheads > 200000 Or Numheads <= 0) Then
        MsgBox "La valor de 'NumHeads' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents

    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    If (NoUsarCabecera = False) Then Put Handle, , MiCabecera0130
    Put Handle, , Numheads
    
    For I = 1 To Numheads
        MisCabezas.Head(1) = Val(Leer.GetValue("HEAD" & I, "Head1"))
        MisCabezas.Head(2) = Val(Leer.GetValue("HEAD" & I, "Head2"))
        MisCabezas.Head(3) = Val(Leer.GetValue("HEAD" & I, "Head3"))
        MisCabezas.Head(4) = Val(Leer.GetValue("HEAD" & I, "Head4"))
        Put Handle, , MisCabezas
    Next I
    Close Handle
    
    IndexarCabezas0130 = True
Exit Function

ErrorHandler:
End Function


Public Function DesindexarCuerpos0130() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos As tIndiceCuerpo
    
    DesindexarCuerpos0130 = False
    Call IniciarCabecera(MiCabecera0130)
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    If (NoUsarCabecera = False) Then Get Handle, , MiCabecera0130
    Get Handle, , NumCuerpos
    
    If NumCuerpos <= 0 Then
        MsgBox "Indice incompatible!", vbCritical
        Close Handle
        Exit Function
    End If
    
    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents
    Open IniPathD & GraphicsFileD For Binary Access Write As handleW
    Put handleW, , "[INIT]" & vbCrLf & "NumBodies=" & NumCuerpos & vbCrLf & vbCrLf

    For I = 1 To NumCuerpos
        Get Handle, , MisCuerpos
        Put handleW, , "[BODY" & I & "]" & vbCrLf
        If MisCuerpos.Body(1) > 300000000 Then
            MsgBox "Los cuerpos no se encuentran en formato v0.12.1/0.13.x.", vbCritical + vbOKOnly
            Close Handle
            Exit Function
        End If
        Put handleW, , "Walk1=" & MisCuerpos.Body(1) & vbTab & " ' arriba" & vbCrLf
        Put handleW, , "Walk2=" & MisCuerpos.Body(2) & vbTab & " ' derecha" & vbCrLf
        Put handleW, , "Walk3=" & MisCuerpos.Body(3) & vbTab & " ' abajo" & vbCrLf
        Put handleW, , "Walk4=" & MisCuerpos.Body(4) & vbTab & " ' izq" & vbCrLf
        Put handleW, , "HeadOffsetX=" & MisCuerpos.HeadOffsetX & vbCrLf
        Put handleW, , "HeadOffsetY=" & MisCuerpos.HeadOffsetY & vbCrLf & vbCrLf
    Next I
    Close Handle
    Close handleW
    
    DesindexarCuerpos0130 = True
Exit Function

ErrorHandler:
End Function

Public Function IndexarCuerpos0130() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos As tIndiceCuerpo
    Dim Leer As New clsIniReader

    IndexarCuerpos0130 = False
    Call IniciarCabecera(MiCabecera0130)
    Handle = FreeFile()
    
    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    NumCuerpos = Val(Leer.GetValue("INIT", "NumBodies"))
    If (NumCuerpos > 200000 Or NumCuerpos <= 0) Then
        MsgBox "La valor de 'NumBodies' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents

    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    If (NoUsarCabecera = False) Then Put Handle, , MiCabecera0130
    Put Handle, , NumCuerpos
    
    For I = 1 To NumCuerpos
        MisCuerpos.Body(1) = Val(Leer.GetValue("BODY" & I, "Walk1"))
        MisCuerpos.Body(2) = Val(Leer.GetValue("BODY" & I, "Walk2"))
        MisCuerpos.Body(3) = Val(Leer.GetValue("BODY" & I, "Walk3"))
        MisCuerpos.Body(4) = Val(Leer.GetValue("BODY" & I, "Walk4"))
        MisCuerpos.HeadOffsetX = Val(Leer.GetValue("BODY" & I, "HeadOffsetX"))
        MisCuerpos.HeadOffsetY = Val(Leer.GetValue("BODY" & I, "HeadOffsetY"))
        Put Handle, , MisCuerpos
    Next I
    Close Handle
    
    IndexarCuerpos0130 = True
Exit Function

ErrorHandler:
End Function



Public Function DesindexarFX0130() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Long
    Dim NumFX As Integer
    Dim MisFXs As tIndiceFx
    
    DesindexarFX0130 = False
    Call IniciarCabecera(MiCabecera0130)
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    If (NoUsarCabecera = False) Then Get Handle, , MiCabecera0130
    Get Handle, , NumFX
    
    If NumFX <= 0 Then
        MsgBox "Indice incompatible!", vbCritical
        Close Handle
        Exit Function
    End If
    
    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents
    Open IniPathD & GraphicsFileD For Binary Access Write As handleW
    Put handleW, , "[INIT]" & vbCrLf & "NumFxs=" & NumFX & vbCrLf & vbCrLf

    For I = 1 To NumFX
        Get Handle, , MisFXs
        Put handleW, , "[FX" & I & "]" & vbCrLf
        Put handleW, , "Animacion=" & MisFXs.Animacion & vbCrLf
        Put handleW, , "OffsetX=" & MisFXs.OffsetX & vbCrLf
        Put handleW, , "OffsetY=" & MisFXs.OffsetY & vbCrLf & vbCrLf
    Next I
    Close Handle
    Close handleW
    
    DesindexarFX0130 = True
Exit Function

ErrorHandler:
End Function

Public Function IndexarFX0130() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Long
    Dim NumFX As Integer
    Dim MisFXs As tIndiceFx
    Dim Leer As New clsIniReader

    IndexarFX0130 = False
    Call IniciarCabecera(MiCabecera0130)
    Handle = FreeFile()
    
    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    NumFX = Val(Leer.GetValue("INIT", "NumFxs"))
    If (NumFX > 200000 Or NumFX <= 0) Then
        MsgBox "La valor de 'NumFxs' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents

    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    If (NoUsarCabecera = False) Then Put Handle, , MiCabecera0130
    Put Handle, , NumFX
    
    For I = 1 To NumFX
        MisFXs.Animacion = Val(Leer.GetValue("FX" & I, "Animacion"))
        MisFXs.OffsetX = Val(Leer.GetValue("FX" & I, "OffsetX"))
        MisFXs.OffsetY = Val(Leer.GetValue("FX" & I, "OffsetY"))
        Put Handle, , MisFXs
    Next I
    Close Handle
    
    IndexarFX0130 = True
Exit Function

ErrorHandler:
End Function

