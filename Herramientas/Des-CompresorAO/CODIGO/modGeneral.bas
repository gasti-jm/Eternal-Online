Attribute VB_Name = "modGeneral"
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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
    hwndOwner      As Long
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

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHGetIDListFromPath Lib "Shell32" Alias "#162" (ByVal pszPath As String) As Long
Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function GetTickCount Lib "kernel32" () As Long

Public IniPath As String
Public CompressFile As String
Public IniPathD As String
Public CompressFileD As String

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean

    FileExist = (Dir$(file, FileType) <> "")
    
End Function


Public Function ParsePath(strFullPathName As String, ReturnType As Byte) As String

    Dim strTemp As String, intX As Integer, strPathName As String, strFileName As String
    If Len(strFullPathName) > 0 Then
        strTemp = ""
        intX = Len(strFullPathName)
        Do While strTemp <> "\"
            strTemp = Mid(strFullPathName, intX, 1)
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

Public Function AbrirComprimido(Optional ByVal Mensaje As String = "Seleccione el archivo Graficos.EO", Optional ByVal Filtro As String = "Graficos (*.EO)" & vbNullChar & "*.EO" & vbNullChar & "Todos los archivos" & vbNullChar & "*.*") As Boolean
On Error Resume Next

    Dim OFName As OPENFILENAME
    Dim sT As String
    
    AbrirComprimido = False
    
    On Local Error Resume Next

    With OFName
        .lStructSize = Len(OFName)
        .hwndOwner = frmMain.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filtro
        .lpstrTitle = Mensaje
        .Flags = cSingleSelFlags
        .lpstrFile = Space$(1023)
        .nMaxFile = 1024
    End With

    If Not GetOpenFileName(OFName) = 0 Then
        sT = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
        IniPath = ParsePath(sT, vbDirectory)
        CompressFile = ParsePath(sT, vbArchive)
        AbrirComprimido = True
    Else
        CompressFile = vbNullString
        Exit Function
    End If

End Function

Public Function GuardarComprimido(Optional ByVal Mensaje As String = "Indique donde será guardado Graficos.EO", Optional ByVal Filtro As String = "Graficos (*.EO)" & vbNullChar & "*.EO" & vbNullChar & "Todos los archivos" & vbNullChar & "*.*") As Boolean
On Error Resume Next

    Dim OFName As OPENFILENAME
    Dim sT As String
    
    GuardarComprimido = False
    
    On Local Error Resume Next

    With OFName
        .lStructSize = Len(OFName)
        .hwndOwner = frmMain.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filtro
        .lpstrTitle = Mensaje
        .Flags = cSingleSelFlags
        .lpstrFile = Space$(1023)
        .nMaxFile = 1024
    End With

    If Not GetSaveFileName(OFName) = 0 Then
        sT = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
        If InStr(1, Filtro, ".EO") <> 0 Then
            If UCase$(Right$(sT, 3)) <> ".EO" Then sT = sT & ".EO"
        ElseIf InStr(1, Filtro, ".PATCH") <> 0 Then
            If UCase$(Right$(sT, 6)) <> ".PATCH" Then sT = sT & ".PATCH"
        End If
        IniPathD = ParsePath(sT, vbDirectory)
        CompressFileD = ParsePath(sT, vbArchive)
        GuardarComprimido = True
    Else
        CompressFileD = vbNullString
        Exit Function
    End If

End Function

Function SeleccionarDirectorio(Optional ByVal Mensaje As String = "Seleccione un Directorio") As String

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    SeleccionarDirectorio = ""
    
    lpIDList = SHGetFolderLocation(frmMain.hWnd, 6, SHGetIDListFromPath(App.Path), 0, tBrowseInfo.pIDLRoot)

    With tBrowseInfo
        .hwndOwner = frmMain.hWnd
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
        SeleccionarDirectorio = sBuffer & "\"
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

