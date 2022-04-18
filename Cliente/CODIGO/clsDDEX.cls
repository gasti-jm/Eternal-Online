VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================================
' clsDDEX - ENGINE GRAPHIC DDEX v4
' AUTOR OF ENGINE: LOOPZER
'
' Comments: DDEX es una DLL que carga multiples APIS graficas
' tales como DX8, DX9 & OpenGL. Fue creada y liberada por
' el usuario "Loopzer" de GS-ZONE.
'
' For more information: www.gs-zone.org
'                       www.ddex.com.ar
'                       Loopzer@gmail.com
'
' Last modification: 25/06/2020 - by ZenitraM
'====================================================================

Option Explicit

Private Declare Function IniciarPrueba Lib "DDEXv4.dll" () As Long
Private Declare Function DibujarVertices Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef v As Any, ByVal cantidad As Long, ByVal alphaf As Long) As Long
Private Declare Function DibujarCajaEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef R As Any, ByRef color As Any, ByVal angulo As Long) As Long
Private Declare Function DibujarCaja Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef R As Any, ByRef color As Any, ByVal angulo As Long) As Long
Private Declare Function DibujarLineaEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef R As Any, ByRef color As Any) As Long
Private Declare Function PreCargarGrafico Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long) As Long
Private Declare Function ZoomPantalla Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal zoom As Long) As Long
Private Declare Function IniciarMotor Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal ventana As Long, ByRef carga As Any, ByRef Configuracion As Any) As Long
Private Declare Function RotarPantalla Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal angulo As Long) As Long
Private Declare Function TerminarMotor Lib "DDEXv4.dll" (ByRef puntero As Any) As Long
Private Declare Function DibujarRGrafico Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long, ByRef R As Any, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DibujarRGraficoAlpha Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long, ByRef R As Any, ByVal x As Long, ByVal y As Long, ByRef color As Any) As Long
Private Declare Function DibujarRGraficoEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long, ByRef R As Any, ByVal x As Long, ByVal y As Long, ByRef color As Any, ByVal angulo As Long) As Long
Private Declare Function DibujarTexto Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef Texto As Any, ByVal x As Long, ByVal y As Long, ByRef color As Any, ByVal fuente As Long) As Long
Private Declare Function CargarFuenteX Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef fuente As Any, ByVal tam As Long, ByVal negrita As Byte, ByVal cursiva As Byte) As Long
Private Declare Function LimpiarMotor Lib "DDEXv4.dll" (ByRef puntero As Any) As Long
Private Declare Function DibujarMotor Lib "DDEXv4.dll" (ByRef puntero As Any) As Long
Private Declare Function DibujarPantallaEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal ventana As Long, ByRef entrada As Any, ByRef salida As Any) As Long


'Manejo de memoria.
Public Enum MODO_MEMORIA
    DX9_MEM_DEAFAULT = 0 'manejo de memoria por defecto
    DX9_MEM_ADMIN = 1 'manejo de memoria administrador
    DX9_MEM_SYSTEM = 2 'manejo de memoria de sistema
End Enum

'API GRAFICA (DX8, DX9 & OpenGL)
Public Enum API_grafica
    DX9 = 0
    DX9_EXPERIMENTAL = 1
    OGL = 2
    DX8 = 3
End Enum

Public Enum MODO2
    DX9_VH = 0 'usar procesamiento de vertices por hardware
    DX9_VS = 1 'usar procesamiento de vertices por software
    OGL_V = 2 'solo es una guia
End Enum

Public Enum MODO
    DX9_HARD = 0
    DX9_REF = 1
    DX9_SOF = 2
    OGL = 3 'solo es una guia
End Enum

Private Declare Function VB6_ArrayPointer Lib "msvbvm60" Alias "VarPtr" (arr As Any) As Long
Private puntero As Long
Public Sub PreCargar(archivo As Long)
    On Error Resume Next
    PreCargarGrafico ByVal VB6_ArrayPointer(puntero), archivo
End Sub
'vertices x,y,z,rhw,color
Public Sub DibujarVerticesEx(ByRef v, ByVal cantidad As Long, ByVal alphaf As Long)
    DibujarVertices puntero, v, cantidad, alphaf
End Sub

'carpeta = la carpeta donde se encuentren los graficos
'soporta BMP y PNG  automaticamente se carga el BMP numero y si no existe lo busca en PNG
Friend Function Iniciar(ByVal hwnd As Long, carpeta As String, cfg As DDEXCFG) As Boolean
    On Error Resume Next
    puntero = 0
    Iniciar = True
    If IniciarMotor(puntero, hwnd, ByVal StrPtr(StrConv(App.Path & carpeta, vbFromUnicode)), cfg) = -1 Then Iniciar = False
    'IniciarMotor ByVal VB6_ArrayPointer(puntero), hwnd, ByVal StrPtr(StrConv(carpeta, vbFromUnicode)), cc
End Function
Public Sub SetearZoom(zoom_ As Long)
On Error Resume Next
   ZoomPantalla ByVal VB6_ArrayPointer(puntero), zoom_
End Sub
Public Sub Termina()
    TerminarMotor ByVal VB6_ArrayPointer(puntero)
End Sub

Public Function CrearFuente(Nombre As String, tam As Long, Optional negrita As Boolean = False, Optional cursiva As Boolean = False)
    Dim t As Integer
    t = 0
    If negrita Then t = 255
    CrearFuente = CargarFuenteX(ByVal VB6_ArrayPointer(puntero), ByVal StrPtr(StrConv(Nombre, vbFromUnicode)), tam, t, cursiva)
End Function
Friend Sub DBTexto(ByVal Str As String, ByVal x As Integer, ByVal y As Integer, lColor As tDDEXRGBA, Optional fuente As Long = 1)
    If LenB(Str) = 0 Then Exit Sub
    DibujarTexto puntero, ByVal StrPtr(StrConv(Str, vbFromUnicode)), x, y, lColor, fuente
    'DibujarTexto puntero, ByVal StrPtr(StrConv(Str, vbFromUnicode)), X, Y, lColor, fuente
    'DibujarTexto ByVal VB6_ArrayPointer(puntero), ByVal StrPtr(StrConv(Str, vbFromUnicode)), X, Y, lColor, fuente
End Sub
Friend Sub DbCaja(x As Long, y As Long, ancho As Long, alto As Long, lColor As tDDEXRGBA, Optional angulo As Long = 0)
    Dim R As rect
    R.Left = x
    R.Top = y
    R.Bottom = y + alto
    R.Right = x + ancho
    DibujarCaja puntero, R, lColor, angulo
    'DibujarCajaEx ByVal VB6_ArrayPointer(puntero), r, lColor, angulo
End Sub
Friend Sub DbCajaEx(x As Long, y As Long, ancho As Long, alto As Long, lColor As tDDEXRGBA, Optional angulo As Long = 0)
    Dim R As rect
    R.Left = x
    R.Top = y
    R.Bottom = y + alto
    R.Right = x + ancho
    DibujarCajaEx puntero, R, lColor, angulo
    'DibujarCajaEx ByVal VB6_ArrayPointer(puntero), r, lColor, angulo
End Sub

Friend Sub DbLinea(x As Long, y As Long, x2 As Long, y2 As Long, lColor As tDDEXRGBA)
    Dim R As rect
    R.Left = x
    R.Top = y
    R.Bottom = y2
    R.Right = x2
    DibujarLineaEx ByVal VB6_ArrayPointer(puntero), R, lColor
End Sub

Friend Sub MostrarPantallaEx(entrada As rect, salida As rect, ventana As Long)
    DibujarPantallaEx ByVal VB6_ArrayPointer(puntero), ventana, entrada, salida
End Sub
Public Sub MostrarPantalla()
    DibujarMotor puntero
    'DibujarMotor ByVal VB6_ArrayPointer(puntero)
End Sub
Public Sub GirarPantalla(angulo As Long)
    RotarPantalla ByVal VB6_ArrayPointer(puntero), angulo
End Sub
Public Sub LimpiarPantalla()
    LimpiarMotor puntero
    'LimpiarMotor ByVal VB6_ArrayPointer(puntero)
End Sub
Friend Sub DBEx(archivo As Long, Rec As rect, ByVal x As Integer, ByVal y As Integer, c As tDDEXRGBA, Optional angulo As Long = 0)
    DibujarRGraficoEx puntero, archivo, Rec, x, y, c, angulo
    'DibujarRGraficoEx ByVal VB6_ArrayPointer(puntero), archivo, Rec, X, Y, c, angulo
End Sub
Friend Sub DBAlfa(archivo As Long, Rec As rect, ByVal x As Integer, ByVal y As Integer, c As tDDEXRGBA)
    DibujarRGraficoAlpha puntero, archivo, Rec, x, y, c
    'DibujarRGraficoAlpha ByVal VB6_ArrayPointer(puntero), archivo, Rec, X, Y, c
End Sub
Friend Sub DBGrafico(archivo As Long, ByVal x As Long, ByVal y As Long, RRect As rect, Optional ColorKey As Boolean = True)
    DibujarRGrafico puntero, archivo, RRect, x, y
    'DibujarRGrafico ByVal VB6_ArrayPointer(puntero), archivo, RRect, X, Y
End Sub




