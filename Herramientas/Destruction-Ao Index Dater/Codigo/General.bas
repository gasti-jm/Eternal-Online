Attribute VB_Name = "General"
Option Explicit

Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    Call writeprivateprofilestring(Main, Var, value, File)
End Sub

Public Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    Call getprivateprofilestring(Main, Var, vbNullString, sSpaces, Len(sSpaces), File)
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Function FileExists(ByVal File As String) As Boolean
    FileExists = Dir$(File, vbArchive) <> ""
End Function

Public Function DirExists(ByVal path As String) As Boolean
    DirExists = Dir$(path, vbDirectory) <> ""
End Function

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean
If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function
