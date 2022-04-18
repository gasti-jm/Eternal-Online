Attribute VB_Name = "Grh"
Option Explicit

Private Const GRH_DAT_FILE As String = "Graphics.ind"
Private Const OLD_FORMAT_HEADER As String = "Argentum Online by Noland-Studios."
Private Const OLD_FORMAT_INIT_FILE As String = "Inicio.con"

Public Type GrhData
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

Private Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    Fx As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public GrhData() As GrhData

Public AllGrhData As Long

Public FileVersion As Long


Public Function LoadGrhData(ByVal path As String) As Boolean
On Error GoTo ErrHandler
    Dim handle As Integer
    Dim MiCabecera As tCabecera
    
    'Set initial size
    ReDim GrhData(0) As GrhData
    
    handle = FreeFile()
    
    If path = vbNullString Then Exit Function
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    If Not FileExists(path & GRH_DAT_FILE) Then
        MsgBox "El archivo " & path & GRH_DAT_FILE & " no existe, asegurese de seleccionar uno."
        Exit Function
    End If
    
    Open path & GRH_DAT_FILE For Binary Access Read Lock Write As handle
    
    'Check file format! (The crappy header had to have some use after all!)
    Get handle, , MiCabecera
    
    If Left$(MiCabecera.Desc, Len(OLD_FORMAT_HEADER)) = OLD_FORMAT_HEADER Then
        LoadGrhData = LoadGrhDataOld(handle, NumberOfGrhs(path))
        
        'No version available in old file format
        FileVersion = -1
    Else
        'We dont' have header, move back to the beginning
        Seek handle, 1
        
        LoadGrhData = LoadGrhDataNew(handle)
    End If
    
    Close handle
Exit Function

ErrHandler:
    Close handle
    
    MsgBox "Se ha encontrado un error al cargar Graphics.ind." & vbCrLf _
        & "Asegurate de que el archivo Graphics.ind se encuentre en la carpeta " _
        & Config.InitPath
End Function

Private Function LoadGrhDataOld(ByVal handle As Integer, ByVal totalGrhs As Long) As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Integer
    Dim Frame As Long
    Dim TempInt As Integer
    Dim max As Integer
    
    max = -1
    
    'Resize array
    ReDim GrhData(1 To totalGrhs) As GrhData
    
    'Open files
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
    
    'Fill Grh List
    
    'Get first Grh Number
    Get handle, , Grh
    
    Do Until Grh <= 0
        'Get highest grh number being used
        If Grh > max Then
            max = Grh
        End If
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            'Resize animation array
            ReDim .Frames(1 To .NumFrames) As Long
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                
                    Get handle, , TempInt
                    
                    'Old format uses integers
                    .Frames(Frame) = TempInt
                    
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > totalGrhs Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , TempInt
                
                'Convert old speed to new one (time based)!
                If IndexMode = "11.X" Then
                    .Speed = CSng(TempInt)
                Else
                    .Speed = CSng(TempInt) * .NumFrames * 1000 / 18
                End If
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , TempInt
                
                'Old format used ints, not longs.
                .FileNum = TempInt
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , .sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                    
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        
        'Get Next Grh Number
        Get handle, , Grh
    Loop
    
    Close handle
    
    'Trim array
    ReDim Preserve GrhData(1 To max) As GrhData
    
    AllGrhData = max
    
    LoadGrhDataOld = True
Exit Function

ErrorHandler:
    LoadGrhDataOld = False
End Function

''
' Finds out the number of grhs for the old file format
'
' @param    path    The path to the folder in which the init file is stored.
'
' @return   The number of grhs that can exist at most.

Private Function NumberOfGrhs(ByVal path As String) As Long
    Dim N As Integer
    Dim GameIni As tGameIni
    Dim MiCabecera As tCabecera
    
    N = FreeFile
    
    Open path & OLD_FORMAT_INIT_FILE For Binary As #N
    
    Get N, , MiCabecera
    
    Get N, , GameIni
    
    Close N
    
    NumberOfGrhs = GameIni.NumeroDeBMPs
End Function

''
' Loads grh data using the new file format.
'
' @param    handle      Handle to the open file containing the grh data.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhDataNew(ByVal handle As Integer) As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim max As Integer
    
    max = -1
    
    'Get file version
    Get handle, , FileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        'Get highest grh number being used
        If Grh > max Then
            max = Grh
        End If
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
            
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close handle
    
    LoadGrhDataNew = True
Exit Function

ErrorHandler:
    AllGrhData = max
    
    LoadGrhDataNew = False
End Function

Public Function SaveGrhDataOld(ByVal path As String) As Boolean
    Dim handle
    Dim Frame As Long
    Dim i As Long
    Dim TempInt As Integer
    Dim MiCabecera As tCabecera
    Dim Contador As Long
    Dim DirectPath As String
    
    DirectPath = path
    
    If IndexMode = "11.X" Then
        'Make sure path is properly set
        If Right$(path, 1) <> "\" Then path = path & "\"
        
        path = path & GRH_DAT_FILE
        
        handle = FreeFile()
        
        If FileExists(path) Then
            Call Kill(path)
        End If
    
        Open path For Binary Access Write As handle
    
        MiCabecera.Desc = OLD_FORMAT_HEADER
        
        'Write headers
        Put handle, , MiCabecera
        Put handle, , TempInt
        Put handle, , TempInt
        Put handle, , TempInt
        Put handle, , TempInt
        Put handle, , TempInt
    Else
        'Call SaveGrhDataNew(DirectPath)
        SaveGrhDataOld = True
        If SaveGrhDataNew(DirectPath) = True Then
            MsgBox "Los Indices se han convertido de 11.X a 12.1"
        End If
        Exit Function
    End If
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            'Index too big for this file format?
            If i > &H7FFF& Then
                Close handle
                Kill path
                Exit Function
            End If
            Contador = i
            Put handle, , CInt(i)
            
            With GrhData(i)
                'Set number of frames
                Put handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Put handle, , CInt(.Frames(Frame))
                    Next Frame
                    
                    If IndexMode = "11.X" Then
                        Put handle, , CInt(.Speed)
                    Else
                        Put handle, , .Speed
                    End If
                Else
                    'Write in normal GRH data
                    Put handle, , CInt(.FileNum)
                    
                    Put handle, , .sX
                    
                    Put handle, , .sY
                        
                    Put handle, , .pixelWidth
                    
                    Put handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close handle
    
    SaveGrhDataOld = True
End Function

''
' Saves grh data using the old (and obsolete) file format. Shouldn't be used if possible.
' New format is valid with the new engine, included in Argentum Online 0.12.1
'
' @param    path    The complete path of the folde rin which to write the grh data file.
'                   If it existed it's deleted first.
'
' @return   True if the file was properly saved, False otherwise.

Public Function SaveGrhDataNew(ByVal path As String) As Boolean
    Dim handle
    Dim Frame As Long
    Dim i As Long
    Dim TempInt As Integer
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & GRH_DAT_FILE
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    'Increment file version
    FileVersion = FileVersion + 1
    
    Put handle, , FileVersion
    
    Put handle, , CLng(UBound(GrhData()))
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            Put handle, , i
            
            With GrhData(i)
                'Set number of frames
                Put handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Put handle, , .Frames(Frame)
                    Next Frame
                    
                    Put handle, , .Speed
                Else
                    'Write in normal GRH data
                    Put handle, , .FileNum
                    
                    Put handle, , .sX
                    
                    Put handle, , .sY
                        
                    Put handle, , .pixelWidth
                    
                    Put handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close handle
    
    SaveGrhDataNew = True
End Function

Private Function ReLoadGrhDataOld(ByVal handle As Integer, ByVal totalGrhs As Long) As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Integer
    Dim Frame As Long
    Dim TempInt As Integer
    Dim max As Integer
    
    max = -1
    
    'Resize array
    ReDim Preserve GrhData(1 To AllGrhData) As GrhData
    
    'Open files
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
    
    'Fill Grh List
    
    'Get first Grh Number
    Get handle, , Grh
    
    Do Until Grh <= 0
        'Get highest grh number being used
        If Grh > max Then
            max = Grh
        End If
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            'Resize animation array
            ReDim .Frames(1 To .NumFrames) As Long
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                
                    Get handle, , TempInt
                    
                    'Old format uses integers
                    .Frames(Frame) = TempInt
                    
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > totalGrhs Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , TempInt
                
                'Convert old speed to new one (time based)!
                .Speed = CSng(TempInt) * .NumFrames * 1000 / 18
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , TempInt
                
                'Old format used ints, not longs.
                .FileNum = TempInt
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , .sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                    
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        
        'Get Next Grh Number
        Get handle, , Grh
    Loop
    
    Close handle
    
    'Trim array
    ReDim Preserve GrhData(1 To max) As GrhData
    
    AllGrhData = max
    
    ReLoadGrhDataOld = True
Exit Function

ErrorHandler:
    ReLoadGrhDataOld = False
End Function
