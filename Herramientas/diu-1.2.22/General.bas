Attribute VB_Name = "General"
Option Explicit
Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const OFN_EXPLORER = &H80000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BFFM_INITIALIZED = &H1
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)
Private Const cSingleSelFlags As Long = OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT

Public Declare Function SHGetIDListFromPath Lib "Shell32" Alias "#162" (ByVal pszPath As String) As Long
Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public NoUsarCabecera As Boolean
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer
Public IniPath As String
Public GraphicsFile As String
Public IniPathD As String
Public GraphicsFileD As String
Public DirGraphics As String

Public Const SW_NORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Sub Main()

TilePixelHeight = 32
TilePixelWidth = 32
DIU.Show
' modo de uso
' dindex c:/ao/init/graficos.ind        -> desindexa
' dindex c:/ao/init/graficos.ind.txt -i -> indexa
'Dim command As String
'command = Chr(34) & "D:\AO\Des-Indexador 0.13.0\Graficos3.ind" & Chr(34)
'command = Right(Left(command, Len(command) - 1), Len(command) - 2)
'IniPath = ParsePath(command, vbDirectory)
'GraphicsFile = ParsePath(command, vbArchive)
'MsgBox LoadGrhData
'MsgBox "okay"
End Sub

Public Function GuardarDesindex() As Boolean
On Error Resume Next
    Dim OFName As OPENFILENAME
    Dim sT As String
    GuardarDesindex = False
    
    On Local Error Resume Next

    With OFName
        .lStructSize = Len(OFName)
        .hWndOwner = DIU.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = "Indices desindexados (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*"
        .lpstrTitle = "Guardar el Indice Desindexado"
        .Flags = cSingleSelFlags
        .lpstrFile = Space$(1023)
        .nMaxFile = 1024
    End With

    If Not GetSaveFileName(OFName) = 0 Then
        sT = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
        If LCase$(Right$(sT, 4)) <> ".ini" Then sT = sT & ".ini"
        IniPathD = ParsePath(sT, vbDirectory)
        GraphicsFileD = ParsePath(sT, vbArchive)
        GuardarDesindex = True
    Else
        Exit Function
    End If

End Function

Public Function GuardarIndex() As Boolean
On Error Resume Next
    Dim OFName As OPENFILENAME
    Dim sT As String
    GuardarIndex = False
    
    On Local Error Resume Next

    With OFName
        .lStructSize = Len(OFName)
        .hWndOwner = DIU.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = "Indices (*.ind)" & Chr$(0) & "*.ind" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*"
        .lpstrTitle = "Guardar el Indice Indexado"
        .Flags = cSingleSelFlags
        .lpstrFile = Space$(1023)
        .nMaxFile = 1024
    End With

    If Not GetSaveFileName(OFName) = 0 Then
        sT = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
        If LCase$(Right$(sT, 4)) <> ".ind" Then sT = sT & ".ind"
        IniPathD = ParsePath(sT, vbDirectory)
        GraphicsFileD = ParsePath(sT, vbArchive)
        GuardarIndex = True
    Else
        Exit Function
    End If

End Function

Public Function ExplorarIndex() As Boolean
On Error Resume Next
    Dim OFName As OPENFILENAME
    Dim sT As String
    ExplorarIndex = False
    
    On Local Error Resume Next

    With OFName
        .lStructSize = Len(OFName)
        .hWndOwner = DIU.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = "Indices desindexados (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*"
        .lpstrTitle = "Seleccione un Indice para Indexar"
        .Flags = cSingleSelFlags
        .lpstrFile = Space$(1023)
        .nMaxFile = 1024
    End With

    If Not GetOpenFileName(OFName) = 0 Then
        sT = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
        IniPath = ParsePath(sT, vbDirectory)
        GraphicsFile = ParsePath(sT, vbArchive)
        ExplorarIndex = True
    Else

        Exit Function
    End If

End Function

Public Function ExplorarDesindex() As Boolean
On Error Resume Next
    Dim OFName As OPENFILENAME
    Dim sT As String
    ExplorarDesindex = False
    
    On Local Error Resume Next

    With OFName
        .lStructSize = Len(OFName)
        .hWndOwner = DIU.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = "Indices (*.ind)" & Chr$(0) & "*.ind" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*"
        .lpstrTitle = "Seleccione un Indice para Desindexar"
        .Flags = cSingleSelFlags
        .lpstrFile = Space$(1023)
        .nMaxFile = 1024
    End With

    If Not GetOpenFileName(OFName) = 0 Then
        sT = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
        IniPath = ParsePath(sT, vbDirectory)
        GraphicsFile = ParsePath(sT, vbArchive)
        ExplorarDesindex = True
    Else

        Exit Function
    End If

End Function

Public Function ParsePath(strFullPathName As String, ReturnType As Byte) As String
    Dim strTemp As String, intX As Integer, strPathName As String, strFileName As String
    If Len(strFullPathName) > 0 Then
        strTemp = ""
        intX = Len(strFullPathName)
        Do While strTemp <> "\"
            strTemp = mid(strFullPathName, intX, 1)
            If strTemp = "\" Then
                strPathName = Left(strFullPathName, intX)
                strFileName = Right(strFullPathName, Len(strFullPathName) - intX)
            End If
            intX = intX - 1
        Loop
        Select Case ReturnType
        Case vbDirectory
            ParsePath = strPathName
        Case vbArchive
            ParsePath = strFileName
        Case Else
            ParsePath = strFullPathName
        End Select
    Else
        ParsePath = ""
    End If
End Function

Public Function SelDirGraficos(Optional ByVal Mensaje As String = "Seleccione el Directorio en donde se encuentran los Graficos...") As Boolean

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    DirGraphics = ""
    
    lpIDList = SHGetFolderLocation(DIU.hWnd, 6, SHGetIDListFromPath(App.Path), 0, tBrowseInfo.pIDLRoot)

    With tBrowseInfo
        .hWndOwner = DIU.hWnd
        .lpfnCallback = adr(AddressOf BrowseCallbackProc)
        .lpszTitle = lstrcat(Mensaje, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_NEWDIALOGSTYLE
        .lParam = SHGetIDListFromPath(StrConv(App.Path, vbUnicode))
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        DirGraphics = sBuffer & "\"
        SelDirGraficos = True
    Else
        SelDirGraficos = False
    End If
    
End Function

Function adr(n As Long) As Long
    adr = n
End Function
 
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
  If uMsg = BFFM_INITIALIZED Then
      Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
  End If
End Function

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

