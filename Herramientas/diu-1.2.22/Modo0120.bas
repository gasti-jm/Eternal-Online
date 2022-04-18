Attribute VB_Name = "Modo_0_12_0"
Option Explicit
Private Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type
Private Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type
Private Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type
Private Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type
Private Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type
Private GrhData As GrhData
Private MiCabecera As tCabecera

Private Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.desc = "Argentum Online by Noland Studios. Des-Indexador Universal (c) GS-Zone 2021 - http://www.gs-zone.org"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function Indexar0120() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Integer
    Dim grhCount As Integer
    Dim tempint As Integer
    Dim Handle As Integer
    Dim Leer As New clsIniReader
    Dim Datos As String
    Dim DatoR() As String
    Dim tF As Integer
    
    Indexar0120 = False
    Handle = FreeFile()
    Call IniciarCabecera(MiCabecera)

    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    grhCount = Val(Leer.GetValue("INIT", "NumGrh"))
    If (grhCount > 32767 Or grhCount <= 0) Then
        MsgBox "La valor de 'NumGrh' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents
    
    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    
    Seek Handle, 1
    
    If (NoUsarCabecera = False) Then Put Handle, , MiCabecera
    Put Handle, , tempint
    Put Handle, , tempint
    Put Handle, , tempint
    Put Handle, , tempint
    Put Handle, , tempint
    
    For Grh = 1 To grhCount
        GrhData.sX = 0
        GrhData.sY = 0
        GrhData.pixelWidth = 0
        GrhData.pixelHeight = 0
        GrhData.FileNum = 0
        GrhData.NumFrames = 0
        GrhData.Speed = 0
        
        Datos = Leer.GetValue("Graphics", "Grh" & Grh)
        If LenB(Datos) <> 0 Then
            DatoR() = Split(Datos, "-")
            If DatoR(0) > 1 Then
                Put Handle, , Grh
                GrhData.NumFrames = Val(DatoR(0))
                Put Handle, , GrhData.NumFrames
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
        If Grh = 32767 Then Exit For
    Next
    
    Close Handle
    
    Indexar0120 = True
Exit Function

ErrorHandler:
    
End Function



Public Function Desindexar0120() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Integer
    Dim Frame As Integer
    Dim grhCount As Integer
    Dim Handle As Integer
    Dim handleW As Integer
    Dim Datos As String
    Dim tempint As Integer
    
    Desindexar0120 = False
    'Open files
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    Call IniciarCabecera(MiCabecera)
    grhCount = 32767
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    
    Seek Handle, 1
    
    If (NoUsarCabecera = False) Then Get Handle, , MiCabecera
    Get Handle, , tempint
    Get Handle, , tempint
    Get Handle, , tempint
    Get Handle, , tempint
    Get Handle, , tempint

    If (MiCabecera.CRC > 100 And MiCabecera.MagicWord > 10) Then
        MsgBox "Indice incompatible!", vbCritical
        Desindexar0120 = False
        Close Handle
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents

    Open IniPathD & GraphicsFileD For Binary Access Write As handleW

    Put handleW, , "[INIT]" & vbCrLf & "NumGrh=" & grhCount & vbCrLf & vbCrLf
    Put handleW, , "[Graphics]" & vbCrLf

    Get Handle, , Grh
    If Grh < 0 Or Grh = 0 Then
        MsgBox "Indice incompatible!", vbCritical
        Close handleW
        Close Handle
        Desindexar0120 = False
        Exit Function
    End If
    While Not EOF(Handle) And (Grh <> 0 And Grh <= grhCount)
        With GrhData
            'Get number of frames
            Get Handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            Datos$ = ""
            'ReDim .Frames(1 To GrhData.NumFrames)
            If (.NumFrames > 25) Then
                MsgBox "Indice incompatible!", vbCritical
                Desindexar0120 = False
                Close handleW
                Close Handle
                Exit Function
            End If
            
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
    
    Desindexar0120 = True
Exit Function

ErrorHandler:
    
End Function

Public Function DesindexarCabezas0120() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Integer
    Dim Numheads As Integer
    Dim MisCabezas As tIndiceCabeza
    
    DesindexarCabezas0120 = False
    Call IniciarCabecera(MiCabecera)
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    If (NoUsarCabecera = False) Then Get Handle, , MiCabecera
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
    
    DesindexarCabezas0120 = True
Exit Function

ErrorHandler:
End Function

Public Function IndexarCabezas0120() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Integer
    Dim Numheads As Integer
    Dim MisCabezas As tIndiceCabeza
    Dim Leer As New clsIniReader

    
    IndexarCabezas0120 = False
    Call IniciarCabecera(MiCabecera)
    Handle = FreeFile()
    
    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    Numheads = Val(Leer.GetValue("INIT", "NumHeads"))
    If (Numheads > 32767 Or Numheads <= 0) Then
        MsgBox "La valor de 'NumHeads' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents

    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    If (NoUsarCabecera = False) Then Put Handle, , MiCabecera
    Put Handle, , Numheads
    
    For I = 1 To Numheads
        MisCabezas.Head(1) = Val(Leer.GetValue("HEAD" & I, "Head1"))
        MisCabezas.Head(2) = Val(Leer.GetValue("HEAD" & I, "Head2"))
        MisCabezas.Head(3) = Val(Leer.GetValue("HEAD" & I, "Head3"))
        MisCabezas.Head(4) = Val(Leer.GetValue("HEAD" & I, "Head4"))
        Put Handle, , MisCabezas
    Next I
    Close Handle
    
    IndexarCabezas0120 = True
Exit Function

ErrorHandler:
End Function


Public Function DesindexarCuerpos0120() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Integer
    Dim NumCuerpos As Integer
    Dim MisCuerpos As tIndiceCuerpo
    
    DesindexarCuerpos0120 = False
    Call IniciarCabecera(MiCabecera)
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    If (NoUsarCabecera = False) Then Get Handle, , MiCabecera
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
        Put handleW, , "Walk1=" & MisCuerpos.Body(1) & vbTab & " ' arriba" & vbCrLf
        Put handleW, , "Walk2=" & MisCuerpos.Body(2) & vbTab & " ' derecha" & vbCrLf
        Put handleW, , "Walk3=" & MisCuerpos.Body(3) & vbTab & " ' abajo" & vbCrLf
        Put handleW, , "Walk4=" & MisCuerpos.Body(4) & vbTab & " ' izq" & vbCrLf
        Put handleW, , "HeadOffsetX=" & MisCuerpos.HeadOffsetX & vbCrLf
        Put handleW, , "HeadOffsetY=" & MisCuerpos.HeadOffsetY & vbCrLf & vbCrLf
    Next I
    Close Handle
    Close handleW
    
    DesindexarCuerpos0120 = True
Exit Function

ErrorHandler:
End Function

Public Function IndexarCuerpos0120() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Integer
    Dim NumCuerpos As Integer
    Dim MisCuerpos As tIndiceCuerpo
    Dim Leer As New clsIniReader

    IndexarCuerpos0120 = False
    Call IniciarCabecera(MiCabecera)
    Handle = FreeFile()
    
    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    NumCuerpos = Val(Leer.GetValue("INIT", "NumBodies"))
    If (NumCuerpos > 32767 Or NumCuerpos <= 0) Then
        MsgBox "La valor de 'NumBodies' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents

    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    If (NoUsarCabecera = False) Then Put Handle, , MiCabecera
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
    
    IndexarCuerpos0120 = True
Exit Function

ErrorHandler:
End Function



Public Function DesindexarFX0120() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Integer
    Dim NumFX As Integer
    Dim MisFXs As tIndiceFx
    
    DesindexarFX0120 = False
    Call IniciarCabecera(MiCabecera)
    Handle = FreeFile()
    handleW = FreeFile() + 1
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    If (NoUsarCabecera = False) Then Get Handle, , MiCabecera
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
    
    DesindexarFX0120 = True
Exit Function

ErrorHandler:
End Function

Public Function IndexarFX0120() As Boolean
On Error GoTo ErrorHandler:
    Dim Handle As Integer
    Dim handleW As Integer
    Dim I As Integer
    Dim NumFX As Integer
    Dim MisFXs As tIndiceFx
    Dim Leer As New clsIniReader

    IndexarFX0120 = False
    Call IniciarCabecera(MiCabecera)
    Handle = FreeFile()
    
    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    NumFX = Val(Leer.GetValue("INIT", "NumFxs"))
    If (NumFX > 32767 Or NumFX <= 0) Then
        MsgBox "La valor de 'NumFxs' es invalido!", vbCritical
        Exit Function
    End If

    If LenB(Dir(IniPathD & GraphicsFileD)) <> 0 Then Call Kill(IniPathD & GraphicsFileD)
    DoEvents

    Open IniPathD & GraphicsFileD For Binary Access Write As Handle
    If (NoUsarCabecera = False) Then Put Handle, , MiCabecera
    Put Handle, , NumFX
    
    For I = 1 To NumFX
        MisFXs.Animacion = Val(Leer.GetValue("FX" & I, "Animacion"))
        MisFXs.OffsetX = Val(Leer.GetValue("FX" & I, "OffsetX"))
        MisFXs.OffsetY = Val(Leer.GetValue("FX" & I, "OffsetY"))
        Put Handle, , MisFXs
    Next I
    Close Handle
    
    IndexarFX0120 = True
Exit Function

ErrorHandler:
End Function
