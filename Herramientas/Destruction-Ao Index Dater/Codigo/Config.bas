Attribute VB_Name = "Config"
Option Explicit

Private Const CONFIG_FILE As String = "/DAOIndexDater.dao"

'Configuration variables are publicly accessed
Public BmpPath As String
Public BmpSavePath As String
Public InitPath As String
Public DatPath As String
Public SaveInitPath As String
Public SaveDatPath As String
Public WavPath As String


Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

Public Function LoadConfig() As Boolean
    
    Dim configPath As String
    
    configPath = App.path & CONFIG_FILE
    
    'Make sure the file exists
    If Not FileExists(configPath) Then
        Exit Function
    End If
    
    BmpPath = GetVar(configPath, "Rutas", "Graficos")
    BmpSavePath = GetVar(configPath, "Rutas", "SaveGraficos")
    InitPath = GetVar(configPath, "Rutas", "Inits")
    DatPath = GetVar(configPath, "Rutas", "Dats")
    SaveInitPath = GetVar(configPath, "Rutas", "SaveInit")
    SaveDatPath = GetVar(configPath, "Rutas", "SaveDat")
    WavPath = GetVar(configPath, "Rutas", "WavPath")
    TilePixelHeight = Val(GetVar(configPath, "Constantes", "TileHeight"))
    TilePixelWidth = Val(GetVar(configPath, "Constantes", "TileWidth"))
    
    'Make usre they are valid
    If BmpPath = "" Or Not DirExists(BmpPath) Or InitPath = "" Or Not DirExists(InitPath) _
        Or DatPath = "" Or Not DirExists(DatPath) Or Not DirExists(SaveInitPath) Or SaveInitPath = "" _
        Or Not DirExists(SaveDatPath) Or SaveDatPath = "" Or Not DirExists(WavPath) Or WavPath = "" _
        Or Not DirExists(BmpSavePath) Or BmpSavePath = "" Or TilePixelHeight = 0 Or TilePixelWidth = 0 Then
        Exit Function
    End If
    
    LoadConfig = True
End Function

Public Sub SaveConfig()
    
    Dim configPath As String
    
    configPath = App.path & CONFIG_FILE
    
    Call WriteVar(configPath, "Rutas", "Graficos", BmpPath)
    Call WriteVar(configPath, "Rutas", "SaveGraficos", BmpSavePath)
    Call WriteVar(configPath, "Rutas", "Inits", InitPath)
    Call WriteVar(configPath, "Rutas", "Dats", DatPath)
    Call WriteVar(configPath, "Rutas", "SaveInit", SaveInitPath)
    Call WriteVar(configPath, "Rutas", "SaveDat", SaveDatPath)
    Call WriteVar(configPath, "Rutas", "WavPath", WavPath)
    
    Call WriteVar(configPath, "Constantes", "TileHeight", CStr(TilePixelHeight))
    Call WriteVar(configPath, "Constantes", "TileWidth", CStr(TilePixelWidth))
End Sub
