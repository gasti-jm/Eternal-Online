Attribute VB_Name = "modCompresion"
Option Explicit

Public Const BMP_SOURCE_FILE_EXT As String = ".bmp"
Public Const PNG_SOURCE_FILE_EXT As String = ".png"
Public Const MAP_SOURCE_FILE_EXT As String = ".map"
Public Const INTERFACE_SOURCE_FILE_EXT As String = ".bmp"

Public datContra() As Byte ' Contraseña
Public usaContra As Boolean ' Usa Contraseña?

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER
    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Private Enum PatchInstruction
    Delete_File
    Create_File
    Modify_File
End Enum

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)

'BitMaps Strucures
Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Const BI_RGB As Long = 0
Private Const BI_RLE8 As Long = 1
Private Const BI_RLE4 As Long = 2
Private Const BI_BITFIELDS As Long = 3
Private Const BI_JPG As Long = 4
Private Const BI_PNG As Long = 5


'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim retval As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

''
' Sorts the info headers by their file name. Uses QuickSort.
'
' @param    InfoHead() The array of headers to be ordered.
' @param    first The first index in the list.
' @param    last The last index in the list.

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, ByVal first As Long, ByVal last As Long)
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/20/2007
'Sorts the info headers by their file name using QuickSort.
'*****************************************************************
    Dim aux As INFOHEADER
    Dim min As Long
    Dim max As Long
    Dim comp As String
    
    min = first
    max = last
    
    comp = InfoHead((min + max) \ 2).strFileName
    
    Do While min <= max
        Do While InfoHead(min).strFileName < comp And min < last
            min = min + 1
        Loop
        Do While InfoHead(max).strFileName > comp And max > first
            max = max - 1
        Loop
        If min <= max Then
            aux = InfoHead(min)
            InfoHead(min) = InfoHead(max)
            InfoHead(max) = aux
            min = min + 1
            max = max - 1
        End If
    Loop
    
    If first < max Then Call Sort_Info_Headers(InfoHead, first, max)
    If min < last Then Call Sort_Info_Headers(InfoHead, min, last)
End Sub

''
' Searches for the specified InfoHeader.
'
' @param    ResourceFile A handler to the data file.
' @param    InfoHead The header searched.
' @param    FirstHead The first head to look.
' @param    LastHead The last head to look.
' @param    FileHeaderSize The bytes size of a FileHeader.
' @param    InfoHeaderSize The bytes size of a InfoHeader.
'
' @return   True if found.
'
' @remark   File must be already open.
' @remark   InfoHead must have set its file name to perform the search.

Private Function BinarySearch(ByRef ResourceFile As Integer, ByRef InfoHead As INFOHEADER, ByVal FirstHead As Long, ByVal LastHead As Long, ByVal FileHeaderSize As Long, ByVal InfoHeaderSize As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/21/2007
'Searches for the specified InfoHeader
'*****************************************************************
    Dim ReadingHead As Long
    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2

        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead

        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else
            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1
            End If
        End If
    Loop
End Function

''
' Retrieves the InfoHead of the specified graphic file.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    InfoHead The InfoHead where data is returned.
'
' @return   True if found.

Private Function Get_InfoHeader(ByRef ResourcePath As String, ByRef Filename As String, ByRef InfoHead As INFOHEADER) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/21/2007
'Retrieves the InfoHead of the specified graphic file
'*****************************************************************
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim FileHead As FILEHEADER
    
On Local Error GoTo ErrHandler

    ResourceFilePath = ResourcePath
    
    'Set InfoHeader we are looking for
    InfoHead.strFileName = UCase$(Filename)
    
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            MsgBox "Archivo de recursos dañado. " & ResourceFilePath, , "Error"
            Close ResourceFile
            Exit Function
        End If
        
        'Search for it!
        If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
            Get_InfoHeader = True
        End If
        
    Close ResourceFile
Exit Function

ErrHandler:
    Close ResourceFile
    
    Call MsgBox("Error al intentar leer el archivo " & ResourceFilePath & ". Razón: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Compresses binary data avoiding data loses.
'
' @param    data() The data array.

Private Sub Compress_Data(ByRef data() As Byte)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 16/07/2012 - ^[GS]^
'Compresses binary data avoiding data loses
'*****************************************************************
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim loopc As Long
    
    Dimensions = UBound(data) + 1
    
    ' The worst case scenario, compressed info is 1.06 times the original - see zlib's doc for more info.
    DimBuffer = Dimensions * 1.06
    
    ReDim BufTemp(DimBuffer)
    
    Call compress(BufTemp(0), DimBuffer, data(0), Dimensions)
    
    Erase data
    
    ReDim data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    data = BufTemp
    
    Erase BufTemp
    
    ' GSZAO - Seguridad
    If usaContra Then
        If UBound(datContra) <= UBound(data) And UBound(datContra) <> 0 Then
            For loopc = 0 To UBound(datContra)
                data(loopc) = data(loopc) Xor datContra(loopc)
            Next loopc
        End If
    End If
    ' GSZAO - Seguridad
End Sub

''
' Decompresses binary data.
'
' @param    data() The data array.
' @param    OrigSize The original data size.

Private Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 16/07/2012 - ^[GS]^
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    Dim loopc As Integer
    
    ReDim BufTemp(OrigSize - 1)
    
    ' GSZAO - Seguridad
    If usaContra Then
        If UBound(datContra) <= UBound(data) And UBound(datContra) <> 0 Then
            For loopc = 0 To UBound(datContra)
                data(loopc) = data(loopc) Xor datContra(loopc)
            Next loopc
        End If
    End If
    ' GSZAO - Seguridad
    
    Call uncompress(BufTemp(0), OrigSize, data(0), UBound(data) + 1)
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

''
' Compresses all graphic files to a resource file.
'
' @param    SourcePath The graphic files folder.
' @param    OutputPath The resource file folder.
' @param    version The resource file version.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Compress_Files(ByRef SourcePath As String, ByRef OutputPath As String, ByVal version As Long, ByRef prgBar As ProgressBar, Optional Mode As Byte = 0) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 31/03/2013 - ^[GS]^
'Compresses all graphic files to a resource file
'*****************************************************************
    Dim SourceFileName As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim loopc As Long

On Local Error GoTo ErrHandler

    OutputFilePath = OutputPath
    If Mode = 0 Then
        SourceFileName = Dir(SourcePath & "*" & BMP_SOURCE_FILE_EXT, vbNormal)
    ElseIf Mode = 1 Then
        SourceFileName = Dir(SourcePath & "*" & MAP_SOURCE_FILE_EXT, vbNormal)
    ElseIf Mode = 2 Then
        SourceFileName = Dir(SourcePath & "*" & MAP_SOURCE_FILE_EXT, vbNormal)
    End If
    
    ' Create list of all files to be compressed
    While LenB(SourceFileName) <> 0
        FileHead.lngNumFiles = FileHead.lngNumFiles + 1
        
        ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
        InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
        
        'Search new file
        SourceFileName = Dir()
    Wend
    
    If Mode = 0 And frmMain.cmdGrhPNG.Value = 1 Then ' Comprimimos tambien los Graficos .PNG
        SourceFileName = Dir(SourcePath & "*" & PNG_SOURCE_FILE_EXT, vbNormal)
        ' Create list of all files to be compressed
        While LenB(SourceFileName) <> 0
            FileHead.lngNumFiles = FileHead.lngNumFiles + 1
            
            ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
            InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
            
            'Search new file
            SourceFileName = Dir()
        Wend
    End If
    
    If Mode = 1 And frmMain.cmdMiniMap.Value = 1 Then ' agregamos tambien los BMP junto a los mapas
        SourceFileName = Dir(SourcePath & "*" & BMP_SOURCE_FILE_EXT, vbNormal)  ' GSZAO
        ' Create list of all files to be compressed
        While LenB(SourceFileName) <> 0
            FileHead.lngNumFiles = FileHead.lngNumFiles + 1
            
            ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
            InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
            
            'Search new file
            SourceFileName = Dir()
        Wend
    End If
    
    If Mode = 2 Then ' agregamos tambien los BMP junto a los mapas
        SourceFileName = Dir(SourcePath & "*" & BMP_SOURCE_FILE_EXT, vbNormal)  ' GSZAO
        ' Create list of all files to be compressed
        While LenB(SourceFileName) <> 0
            FileHead.lngNumFiles = FileHead.lngNumFiles + 1
            
            ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
            InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
            
            'Search new file
            SourceFileName = Dir()
        Wend
    End If
    
    If FileHead.lngNumFiles = 0 Then
        If Mode = 0 Then
            If frmMain.cmdGrhPNG.Value = 1 Then
                MsgBox "No se encontraron archivos de extención " & BMP_SOURCE_FILE_EXT & " o " & PNG_SOURCE_FILE_EXT & " en " & SourcePath & ".", vbOKOnly + vbCritical, "Error"
            Else
                MsgBox "No se encontraron archivos de extención " & BMP_SOURCE_FILE_EXT & " en " & SourcePath & ".", vbOKOnly + vbCritical, "Error"
            End If
        ElseIf Mode = 1 Then
            MsgBox "No se encontraron archivos de extención " & MAP_SOURCE_FILE_EXT & " en " & SourcePath & ".", vbOKOnly + vbCritical, "Error"
        ElseIf Mode = 2 Then
            MsgBox "No se encontraron archivos de extención " & INTERFACE_SOURCE_FILE_EXT & " en " & SourcePath & ".", vbOKOnly + vbCritical, "Error"
        End If
        Exit Function
    End If
    
    If Not prgBar Is Nothing Then
        prgBar.Value = 0
        prgBar.max = FileHead.lngNumFiles + 1
    End If
    
    'Destroy file if it previuosly existed
    If LenB(Dir(OutputFilePath, vbNormal)) <> 0 Then
        Kill OutputFilePath
    End If
    
    'Finish setting the FileHeader data
    FileHead.lngFileVersion = version
    FileHead.lngFileSize = Len(FileHead) + FileHead.lngNumFiles * Len(InfoHead(0))
    
    'Order the InfoHeads
    Call Sort_Info_Headers(InfoHead(), 0, FileHead.lngNumFiles - 1)
    
    'Open a new file
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
        ' Move to the end of the headers, where the file data will actually start
        Seek OutputFile, FileHead.lngFileSize + 1
        
        ' Process every file!
        For loopc = 0 To FileHead.lngNumFiles - 1
            
            SourceFile = FreeFile()
            Open SourcePath & InfoHead(loopc).strFileName For Binary Access Read Lock Write As SourceFile
                
                'Find out how large the file is and resize the data array appropriately
                InfoHead(loopc).lngFileSizeUncompressed = LOF(SourceFile)
                ReDim SourceData(LOF(SourceFile) - 1)
                
                'Get the data from the file
                Get SourceFile, , SourceData
                
                'Compress it
                Call Compress_Data(SourceData)
                
                'Store it in the resource file
                Put OutputFile, , SourceData
                
                With InfoHead(loopc)
                    'Set up the info headers
                    .lngFileSize = UBound(SourceData) + 1
                    .lngFileStart = FileHead.lngFileSize + 1
                    
                    'Update the file header
                    FileHead.lngFileSize = FileHead.lngFileSize + .lngFileSize
                End With
                 
                Erase SourceData
            
            Close SourceFile
        
            'Update progress bar
            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
            DoEvents
        Next loopc
        
        'Store the headers in the file
        Seek OutputFile, 1
        Put OutputFile, , FileHead
        Put OutputFile, , InfoHead
        
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
    
    Compress_Files = True
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    Close OutputFile
    
    Dim Filename As String
    If loopc > 0 And loopc <= FileHead.lngNumFiles - 1 Then
        Filename = InfoHead(loopc).strFileName
    Else
        Filename = "(ninguno)"
    End If
    
    Call MsgBox("No se pudo crear el archivo binario. Razón: " & Err.Number & " : " & Err.Description & vbCrLf & _
                "Archivo en el proceso: " & Filename, vbOKOnly, "Error")
End Function

''
' Retrieves a byte array with the compressed data from the specified file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   InfoHead must not be encrypted.
' @remark   Data is not desencrypted.

Public Function Get_File_RawData(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Retrieves a byte array with the compressed data from the specified file
'*****************************************************************
    Dim ResourceFilePath As String
    Dim ResourceFile As Integer
    
On Local Error GoTo ErrHandler
    ResourceFilePath = ResourcePath
    
    'Size the Data array
    ReDim data(InfoHead.lngFileSize - 1)
    
    'Open the binary file
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Get the data
        Get ResourceFile, InfoHead.lngFileStart, data
    'Close the binary file
    Close ResourceFile
    
    Get_File_RawData = True
Exit Function

ErrHandler:
    Close ResourceFile
End Function

''
' Extract the specific file from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Extract_File(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 14/09/2012 - ^[GS]^
'Extract the specific file from a resource file
'*****************************************************************
On Local Error GoTo ErrHandler
    
    If Get_File_RawData(ResourcePath, InfoHead, data) Then
        'Decompress all data
        'If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then ' GSZAO
            Call Decompress_Data(data, InfoHead.lngFileSizeUncompressed)
        'End If
        
        Extract_File = True
    End If
Exit Function

ErrHandler:
    Call MsgBox("Error al intentar decodificar recursos. Razón: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Extracts all files from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    OutputPath The folder where graphic files will be extracted.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Extract_Files(ByRef ResourcePath As String, ByRef OutputPath As String, ByRef prgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 16/07/2012 - ^[GS]^
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim OutputFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    
On Local Error GoTo ErrHandler
    ResourceFilePath = ResourcePath
    
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox("Archivo de recursos dañado. " & ResourceFilePath, , "Error")
            Close ResourceFile
            Exit Function
        End If
        
        'Size the InfoHead array
        ReDim InfoHead(FileHead.lngNumFiles - 1)
        
        'Extract the INFOHEADER
        Get ResourceFile, , InfoHead
        
        'Check if there is enough hard drive space to extract all files
        For loopc = 0 To UBound(InfoHead)
            RequiredSpace = RequiredSpace + InfoHead(loopc).lngFileSizeUncompressed
        Next loopc
        
        If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
            Erase InfoHead
            Close ResourceFile
            Call MsgBox("No hay suficiente espacio en el disco para extraer los archivos.", , "Error")
            Exit Function
        End If
    Close ResourceFile
    
    'Update progress bar
    If Not prgBar Is Nothing Then
        prgBar.Value = 0
        prgBar.max = FileHead.lngNumFiles + 1
    End If
    
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        'Extract this file
        If Extract_File(ResourcePath, InfoHead(loopc), SourceData) Then
            'Destroy file if it previuosly existed
            If FileExist(OutputPath & InfoHead(loopc).strFileName, vbNormal) Then
                Call Kill(OutputPath & InfoHead(loopc).strFileName)
            End If
            
            'Save it!
            OutputFile = FreeFile()
            Open OutputPath & InfoHead(loopc).strFileName For Binary As OutputFile
                Put OutputFile, , SourceData
            Close OutputFile
            
            Erase SourceData
        Else
            Erase SourceData
            Erase InfoHead
            
            Call MsgBox("No se pudo extraer el archivo " & InfoHead(loopc).strFileName, vbOKOnly, "Error")
            Exit Function
        End If
            
        'Update progress bar
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Next loopc
    
    Erase InfoHead
    Extract_Files = True
Exit Function

ErrHandler:
    Close ResourceFile
    Erase SourceData
    Erase InfoHead
    
    Call MsgBox("No se pudo extraer el archivo binario correctamente. Razón: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Retrieves a byte array with the specified file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Get_File_Data(ByRef ResourcePath As String, ByRef Filename As String, ByRef data() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/21/2007
'Retrieves a byte array with the specified file data
'*****************************************************************
    Dim InfoHead As INFOHEADER
    
    If Get_InfoHeader(ResourcePath, Filename, InfoHead) Then
        'Extract!
        Get_File_Data = Extract_File(ResourcePath, InfoHead, data)
    Else
        Call MsgBox("No se se encontro el recurso " & Filename)
    End If
End Function

''
' Retrieves image file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.

Public Function Get_Image(ByRef ResourcePath As String, ByRef Filename As String, ByRef data() As Byte, Optional SoloBMP As Boolean = False) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 09/10/2012 - ^[GS]^
'Retrieves image file data
'*****************************************************************
    Dim InfoHead As INFOHEADER
    Dim ExistFile As Boolean
    
    ExistFile = False
    
    If SoloBMP = True Then
        If Get_InfoHeader(ResourcePath, Filename & ".BMP", InfoHead) Then ' ¿BMP?
            Filename = Filename & ".BMP"
            ExistFile = True
        End If
    Else
        If Get_InfoHeader(ResourcePath, Filename & ".BMP", InfoHead) Then ' ¿BMP?
            Filename = Filename & ".BMP"
            ExistFile = True
        ElseIf Get_InfoHeader(ResourcePath, Filename & ".PNG", InfoHead) Then ' Existe PNG?
            Filename = Filename & ".PNG" ' usamos el PNG
            ExistFile = True
        End If
    End If
    
    If ExistFile = True Then
        If Extract_File(ResourcePath, InfoHead, data) Then Get_Image = True
    Else
        Call MsgBox("Get_Image::No se encontro el recurso " & Filename)
    End If
End Function

''
' Compare two byte arrays to detect any difference.
'
' @param    data1() Byte array.
' @param    data2() Byte array.
'
' @return   True if are equals.

Private Function Compare_Datas(ByRef data1() As Byte, ByRef data2() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 02/11/2007
'Compare two byte arrays to detect any difference
'*****************************************************************
    Dim Length As Long
    Dim act As Long
    
    Length = UBound(data1) + 1
    
    If (UBound(data2) + 1) = Length Then
        While act < Length
            If data1(act) Xor data2(act) Then Exit Function
            
            act = act + 1
        Wend
        
        Compare_Datas = True
    End If
End Function

''
' Retrieves the next InfoHeader.
'
' @param    ResourceFile A handler to the resource file.
' @param    FileHead The reource file header.
' @param    InfoHead The returned header.
' @param    ReadFiles The number of headers that have already been read.
'
' @return   False if there are no more headers tu read.
'
' @remark   File must be already open.
' @remark   Used to walk through the resource file info headers.
' @remark   The number of read files will increase although there is nothing else to read.
' @remark   InfoHead is encrypted.

Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, ByRef FileHead As FILEHEADER, ByRef InfoHead As INFOHEADER, ByRef ReadFiles As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Reads the next InfoHeader
'*****************************************************************

    If ReadFiles < FileHead.lngNumFiles Then
        'Read header
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
        
        'Update
        ReadNext_InfoHead = True
    End If
    
    ReadFiles = ReadFiles + 1
End Function

''
' Retrieves the next bitmap.
'
' @param    ResourcePath The resource file folder.
' @param    ReadFiles The number of bitmaps that have already been read.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   False if there are no more bitmaps tu get.
'
' @remark   Used to walk through the resource file bitmaps.

Public Function GetNext_Bitmap(ByRef ResourcePath As String, ByRef ReadFiles As Long, ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte, ByRef fileIndex As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 09/10/2012 - ^[GS]^
'Reads the next InfoHeader
'*****************************************************************
On Error Resume Next

    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    Dim InfoHead As INFOHEADER
    Dim Filename As String
    
    ResourceFile = FreeFile
    Open ResourcePath For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead
    
    If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then
        Call Get_Image(ResourcePath, InfoHead.strFileName, data())
        Filename = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(Filename, Len(Filename) - 4))
        
        GetNext_Bitmap = True
    End If
    
    Close ResourceFile
End Function

''
' Compares two resource versions and makes a patch file.
'
' @param    NewResourcePath The actual reource file folder.
' @param    OldResourcePath The previous reource file folder.
' @param    OutputPath The patchs file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Make_Patch(ByRef NewResourcePath As String, ByRef OldResourcePath As String, ByRef OutputPath As String, ByRef prgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Compares two resource versions and make a patch file
'*****************************************************************
    Dim NewResourceFile As Integer
    Dim NewResourceFilePath As String
    Dim NewFileHead As FILEHEADER
    Dim NewInfoHead As INFOHEADER
    Dim NewReadFiles As Long
    Dim NewReadNext As Boolean
    
    Dim OldResourceFile As Integer
    Dim OldResourceFilePath As String
    Dim OldFileHead As FILEHEADER
    Dim OldInfoHead As INFOHEADER
    Dim OldReadFiles As Long
    Dim OldReadNext As Boolean
    
    Dim OutputFile As Integer
    Dim OutputFilePath As String
    Dim data() As Byte
    Dim auxData() As Byte
    Dim Instruction As Byte
    
'Set up the error handler
On Local Error GoTo ErrHandler

    NewResourceFilePath = NewResourcePath
    OldResourceFilePath = OldResourcePath
    OutputFilePath = OutputPath
    
    'Open the old binary file
    OldResourceFile = FreeFile
    Open OldResourceFilePath For Binary Access Read Lock Write As OldResourceFile
        
        'Get the old FileHeader
        Get OldResourceFile, 1, OldFileHead
        
        'Check the file for validity
        If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
            Call MsgBox("Archivo de recursos anterior dañado. " & OldResourceFilePath, , "Error")
            Close OldResourceFile
            Exit Function
        End If
        
        'Open the new binary file
        NewResourceFile = FreeFile()
        Open NewResourceFilePath For Binary Access Read Lock Write As NewResourceFile
            
            'Get the new FileHeader
            Get NewResourceFile, 1, NewFileHead
            
            'Check the file for validity
            If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
                Call MsgBox("Archivo de recursos anterior dañado. " & NewResourceFilePath, , "Error")
                Close NewResourceFile
                Close OldResourceFile
                Exit Function
            End If
            
            'Destroy file if it previuosly existed
            If LenB(Dir(OutputFilePath, vbNormal)) <> 0 Then Kill OutputFilePath
            
            'Open the patch file
            OutputFile = FreeFile()
            Open OutputFilePath For Binary Access Read Write As OutputFile
                
                If Not prgBar Is Nothing Then
                    prgBar.Value = 0
                    prgBar.max = (OldFileHead.lngNumFiles + NewFileHead.lngNumFiles) + 1
                End If
                
                'put previous file version (unencrypted)
                Put OutputFile, , OldFileHead.lngFileVersion
                
                'Put the new file header
                Put OutputFile, , NewFileHead
                 
                'Try to read old and new first files
                If ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) _
                  And ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    
                    'Update
                    prgBar.Value = prgBar.Value + 2
                    
                    Do 'Main loop
                        'Comparisons are between encrypted names, for ordering issues
                        If OldInfoHead.strFileName = NewInfoHead.strFileName Then

                            'Get old file data
                            Call Get_File_RawData(OldResourcePath, OldInfoHead, auxData)
                            
                            'Get new file data
                            Call Get_File_RawData(NewResourcePath, NewInfoHead, data)
                            
                            If Not Compare_Datas(data, auxData) Then
                                'File was modified
                                Instruction = PatchInstruction.Modify_File
                                Put OutputFile, , Instruction
                                
                                'Write header
                                Put OutputFile, , NewInfoHead
                                
                                'Write data
                                Put OutputFile, , data
                            End If
                            
                            'Read next OldResource
                            If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                                Exit Do
                            End If
                            
                            'Read next NewResource
                            If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                                'Reread last OldInfoHead
                                OldReadFiles = OldReadFiles - 1
                                Exit Do
                            End If
                            
                            'Update
                            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 2
                        
                        ElseIf OldInfoHead.strFileName < NewInfoHead.strFileName Then
                            
                            'File was deleted
                            Instruction = PatchInstruction.Delete_File
                            Put OutputFile, , Instruction
                            Put OutputFile, , OldInfoHead
                            
                            'Read next OldResource
                            If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                                'Reread last NewInfoHead
                                NewReadFiles = NewReadFiles - 1
                                Exit Do
                            End If
                            
                            'Update
                            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                        
                        Else
                            
                            'New file
                            Instruction = PatchInstruction.Create_File
                            Put OutputFile, , Instruction
                            Put OutputFile, , NewInfoHead
                             
                            'Get file data
                            Call Get_File_RawData(NewResourcePath, NewInfoHead, data)
                            
                            'Write data
                            Put OutputFile, , data
                            
                            'Read next NewResource
                            If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                                'Reread last OldInfoHead
                                OldReadFiles = OldReadFiles - 1
                                Exit Do
                            End If
                            
                            'Update
                            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                        End If
                        
                        DoEvents
                    Loop
                
                Else
                    'if at least one is empty
                    OldReadFiles = 0
                    NewReadFiles = 0
                End If
                
                'Read everything?
                While ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles)
                    'Delete file
                    Instruction = PatchInstruction.Delete_File
                    Put OutputFile, , Instruction
                    Put OutputFile, , OldInfoHead
                    
                    'Update
                    If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                    DoEvents
                Wend
                
                'Read everything?
                While ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles)
                    'Create file
                    Instruction = PatchInstruction.Create_File
                    Put OutputFile, , Instruction
                    Put OutputFile, , NewInfoHead
                    
                    Call Get_File_RawData(NewResourcePath, NewInfoHead, data)
                    'Write data
                    Put OutputFile, , data
                    
                    'Update
                    If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                    DoEvents
                Wend
            
            'Close the patch file
            Close OutputFile
        
        'Close the new binary file
        Close NewResourceFile
    
    'Close the old binary file
    Close OldResourceFile
    
    Make_Patch = True
Exit Function

ErrHandler:
    Close OutputFile
    Close NewResourceFile
    Close OldResourceFile
    
    Call MsgBox("No se pudo terminar de crear el parche. Razón: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Follows patches instructions to update a resource file.
'
' @param    ResourcePath The reource file folder.
' @param    PatchPath The patch file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.
Public Function Apply_Patch(ByRef ResourcePath As String, ByRef OldPath As String, ByRef PatchPath As String, ByRef prgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Follows patches instructions to update a resource file
'*****************************************************************
    Dim ResourceFile As Integer
    Dim sOldResourceFile As String
    Dim FileHead As FILEHEADER
    Dim InfoHead As INFOHEADER
    Dim ResourceReadFiles As Long
    Dim EOResource As Boolean

    Dim PatchFile As Integer
    Dim sPatchFile As String
    Dim PatchFileHead As FILEHEADER
    Dim PatchInfoHead As INFOHEADER
    Dim Instruction As Byte
    Dim OldResourceVersion As Long

    Dim OutputFile As Integer
    Dim sNewResourceFile As String
    Dim data() As Byte
    Dim WrittenFiles As Long
    Dim DataOutputPos As Long

On Local Error GoTo ErrHandler

    sOldResourceFile = OldPath
    sPatchFile = PatchPath
    sNewResourceFile = ResourcePath
    
    'Open the old binary file
    ResourceFile = FreeFile()
    Open sOldResourceFile For Binary Access Read Lock Write As ResourceFile
        
        'Read the old FileHeader
        Get ResourceFile, , FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox("Archivo de recursos anterior dañado. " & sOldResourceFile, , "Error")
            Close ResourceFile
            Exit Function
        End If
        
        'Open the patch file
        PatchFile = FreeFile()
        Open sPatchFile For Binary Access Read Lock Write As PatchFile
            
            'Get previous file version
            Get PatchFile, , OldResourceVersion
            
            'Check the file version
            If OldResourceVersion <> FileHead.lngFileVersion Then
                Call MsgBox("Incongruencia en versiones.", , "Error")
                Close ResourceFile
                Close PatchFile
                Exit Function
            End If
            
            'Read the new FileHeader
            Get PatchFile, , PatchFileHead
            
            'Destroy file if it previuosly existed
            If FileExist(sNewResourceFile, vbNormal) Then Call Kill(sNewResourceFile)
            
            'Open the patch file
            OutputFile = FreeFile()
            Open sNewResourceFile For Binary Access Read Write As OutputFile
                
                'Save the file header
                Put OutputFile, , PatchFileHead

                If Not prgBar Is Nothing Then
                    prgBar.Value = 0
                    prgBar.max = PatchFileHead.lngNumFiles + 1
                End If
                
                'Update
                DataOutputPos = Len(FileHead) + Len(InfoHead) * PatchFileHead.lngNumFiles + 1
                
                'Process loop
                While Loc(PatchFile) < LOF(PatchFile)
                    
                    'Get the instruction
                    Get PatchFile, , Instruction
                    'Get the InfoHead
                    Get PatchFile, , PatchInfoHead
                    
                    Do
                        EOResource = Not ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                        
                        'Comparison is performed among encrypted names for ordering issues
                        If Not EOResource And InfoHead.strFileName < PatchInfoHead.strFileName Then

                            'GetData and update InfoHead
                            Call Get_File_RawData(sOldResourceFile, InfoHead, data)
                            InfoHead.lngFileStart = DataOutputPos
                            
                            'Save file!
                            Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                            Put OutputFile, DataOutputPos, data
                            
                            'Update
                            DataOutputPos = DataOutputPos + UBound(data) + 1
                            WrittenFiles = WrittenFiles + 1
                            If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    Select Case Instruction
                        'Delete
                        Case PatchInstruction.Delete_File
                            If InfoHead.strFileName <> PatchInfoHead.strFileName Then
                                Err.Description = "Incongruencia en archivos de recurso"
                                GoTo ErrHandler
                            End If
                        
                        'Create
                        Case PatchInstruction.Create_File
                            If (InfoHead.strFileName > PatchInfoHead.strFileName) Or EOResource Then

                                'Get file data
                                ReDim data(PatchInfoHead.lngFileSize - 1)
                                Get PatchFile, , data

                                'Save it
                                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                                Put OutputFile, DataOutputPos, data
                                
                                'Reanalize last Resource InfoHead
                                EOResource = False
                                ResourceReadFiles = ResourceReadFiles - 1
                                
                                'Update
                                DataOutputPos = DataOutputPos + UBound(data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                            Else
                                Err.Description = "Incongruencia en archivos de recurso"
                                GoTo ErrHandler
                            End If
                        
                        'Modify
                        Case PatchInstruction.Modify_File
                            If InfoHead.strFileName = PatchInfoHead.strFileName Then
                            
                                'Get file data
                                ReDim data(PatchInfoHead.lngFileSize - 1)
                                Get PatchFile, , data
                                
                                'Save it
                                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                                Put OutputFile, DataOutputPos, data
                                
                                'Update
                                DataOutputPos = DataOutputPos + UBound(data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                            Else
                                Err.Description = "Incongruencia en archivos de recurso"
                                GoTo ErrHandler
                            End If
                    End Select
                    
                    DoEvents
                Wend
                
                'Read everything?
                While ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)

                    'GetData and update InfoHeader
                    Call Get_File_RawData(sOldResourceFile, InfoHead, data)
                    InfoHead.lngFileStart = DataOutputPos
                    
                    'Save file!
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                    Put OutputFile, DataOutputPos, data
                    
                    'Update
                    DataOutputPos = DataOutputPos + UBound(data) + 1
                    WrittenFiles = WrittenFiles + 1
                    If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                    DoEvents
                Wend
            
            'Close the patch file
            Close OutputFile
        
        'Close the new binary file
        Close PatchFile
    
    'Close the old binary file
    Close ResourceFile
    
    'Check integrity
    If (PatchFileHead.lngNumFiles = WrittenFiles) Then
        'Replace File
        If MsgBox("El proceso ha concluido con exito." & vbCrLf & "¿Desea reemplazar el archivo original con la versión parcheada?", vbYesNo + vbQuestion) = vbYes Then
            Call Kill(sOldResourceFile)
            Name sNewResourceFile As sOldResourceFile
        End If
    Else
        Err.Description = "Falla al procesar parche."
        GoTo ErrHandler
    End If
    
    Apply_Patch = True
Exit Function

ErrHandler:
    Close OutputFile
    Close PatchFile
    Close ResourceFile
    'Destroy file if created
    If FileExist(sNewResourceFile, vbNormal) Then Call Kill(sNewResourceFile)
    
    Call MsgBox("No se pudo parchear. Razón: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")
End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
'*****************************************************************
'Author: Unknown
'Last Modify Date: Unknown
'*****************************************************************
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
End Function

''
' Retrieves the version number of a given resource file.
'
' @param    ResourceFilePath The resource file complete path.
'
' @return   The version number of the given file.

Public Function GetVersion(ByVal ResourceFilePath As String) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/23/2008
'
'*****************************************************************
    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
    Close ResourceFile
    
    GetVersion = FileHead.lngFileVersion
End Function

Public Function TestZLib() As Boolean ' GSZAO
'*****************************************************************
'Author: ^[GS]^
'Last Modify Date: 19/06/2011
'*****************************************************************
On Error GoTo ErrHandler
    
    Dim data() As Byte
    Dim lnD As Integer
    
    data = "prueba"
    lnD = UBound(data) + 1
    Call Decompress_Data(data, lnD)
    
    TestZLib = True
    
    Exit Function
ErrHandler:

    TestZLib = False
End Function
