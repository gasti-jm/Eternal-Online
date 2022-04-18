Attribute VB_Name = "FindFolder"
Option Explicit

Private Const BIF_RETURNONLYFSDIRS As Long = 1
Private Const MAX_PATH As Long = 260


Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type


Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(ByVal hWnd As Long, ByVal title As String) As String
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    
    With udtBI
        'Set the owner window
        .hWndOwner = hWnd
        'Set the window's title
        .lpszTitle = StrPtr(StrConv(title, vbFromUnicode))
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    
    If lpIDList Then
        sPath = String$(MAX_PATH, vbNullChar)
        
        'Get the path from the IDList
        Call SHGetPathFromIDList(lpIDList, sPath)
        
        'free the block of memory
        Call CoTaskMemFree(lpIDList)
        
        iNull = InStr(sPath, vbNullChar)
        
        If iNull Then
            'We got it!
            BrowseForFolder = Left$(sPath, iNull - 1)
        End If
    End If
End Function
