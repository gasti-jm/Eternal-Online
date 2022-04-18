Attribute VB_Name = "Declaraciones"
Option Explicit

Public NombreGrafico As String
Public Ruta As String

Type GrhAnim
    Animate() As Long
End Type

Public GrhAnim(1 To 4) As GrhAnim

Public SelectedImg() As Long

Public Animaciones(1 To 4) As Long

Public UnFrame As Boolean

Public IndexDaterIni As String

Public IndexMode As String

Public LoadNews As Boolean

Type IDNews
    Titulo As String
    Noticia As String
End Type

Public DAONews() As IDNews
