Attribute VB_Name = "modDX8FIFO"
Option Explicit



Sub CargarCabezas()
    Dim n As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    Open DirIndex & "Heads.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
End Sub

Sub CargarCascos()
    Dim n As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    Open DirIndex & "Helmets.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
End Sub

Sub CargarCuerpos()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If Not FileExist(DirIndex & "Bodys.ind", vbArchive) Then
        MsgBox "Falta el archivo 'Bodys.ind' en " & DirIndex, vbCritical
        End
    End If
    
    Dim n As Integer
    Dim i As Integer
    
    n = FreeFile
    Open DirIndex & "Bodys.ind" For Binary Access Read As #n
        'cabecera
        Get #n, , MiCabecera
        'num de cabezas
        Get #n, , NumBodies
        
        'Resize array
        ReDim BodyData(1 To NumBodies) As BodyData
        ReDim MisCuerpos(1 To NumBodies) As tIndiceCuerpo
        
        For i = 1 To NumBodies
            Get #n, , MisCuerpos(i)
            
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0

            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        Next i
    Close #n
Exit Sub
Fallo:
    'MsgBox "Error al intentar cargar el Cuerpo " & i & " de Bodys.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Sub CargarFxs()
    Dim n As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    n = FreeFile()
    Open DirIndex & "Fxs.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #n, , FxData(i)
    Next i
    
    Close #n
End Sub

Sub CargarTips()
    Dim n As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    n = FreeFile
    Open DirIndex & "Tips.ayu" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #n, , Tips(i)
    Next i
    
    Close #n
End Sub

Sub CargarArrayLluvia()
    Dim n As Integer
    Dim i As Long
    Dim Nu As Integer
    
    n = FreeFile()
    Open DirIndex & "fk.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #n, , bLluvia(i)
    Next i
    
    Close #n
End Sub

Public Function LoadGrhData() As Boolean
On Local Error Resume Next
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim Handle As Integer
    Dim fileVersion As Long
    'Open files
    Handle = FreeFile()
    Open DirIndex & "Graphics.ind" For Binary Access Read As Handle
    Get Handle, , fileVersion
    
    Get Handle, , grhCount
    
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(Handle)
        Get Handle, , Grh
        
        With GrhData(Grh)
           ' GrhData(Grh).active = True
            Get Handle, , .NumFrames
            If .NumFrames <= 0 Then Resume Next
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                For Frame = 1 To .NumFrames
                    Get Handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        Resume Next
                    End If
                Next Frame
                
                Get Handle, , .Speed
                
                'GrhData(Grh).speed = GrhData(Grh).speed * 2
                
                If .Speed <= 0 Then Resume Next
                
                .PixelHeight = GrhData(.Frames(1)).PixelHeight
                If .PixelHeight <= 0 Then Resume Next
                
                .PixelWidth = GrhData(.Frames(1)).PixelWidth
                If .PixelWidth <= 0 Then Resume Next
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then Resume Next
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then Resume Next
            Else
                Get Handle, , .FileNum
                If .FileNum <= 0 Then Resume Next
                
                Get Handle, , GrhData(Grh).sX
                If .sX < 0 Then Resume Next
                
                Get Handle, , .sY
                If .sY < 0 Then Resume Next
                
                Get Handle, , .PixelWidth
                If .PixelWidth <= 0 Then Resume Next
                
                Get Handle, , .PixelHeight
                If .PixelHeight <= 0 Then Resume Next
                
                .TileWidth = .PixelWidth / 32
                .TileHeight = .PixelHeight / 32
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close Handle

    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function
