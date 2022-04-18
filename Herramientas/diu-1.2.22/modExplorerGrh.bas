Attribute VB_Name = "modExplorerGrh"
Option Explicit

Private Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    NumFrames As Integer
    Frames() As Long
    Speed As Single
End Type

'Private Type tIndiceCabeza
'    Head(1 To 4) As Long
'End Type
'Private Type tIndiceCuerpo
'    Body(1 To 4) As Long
'    HeadOffsetX As Integer
'    HeadOffsetY As Integer
'End Type
'Private Type tIndiceFx
'    Animacion As Long
'    OffsetX As Integer
'    OffsetY As Integer
'End Type

Public exGrhData() As GrhData
Public exGrhCount As Long
Public exfileVersion As Long

Public Function ExExplorerGrh()
On Error Resume Next
    ReDim exGrhData(1) As GrhData
    exGrhCount = 0
    exfileVersion = 0
End Function

Public Function ExplorerGrh() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Leer As New clsIniReader
    Dim Datos As String
    Dim DatoR() As String
    Dim tF As Integer
    
    ExplorerGrh = False

    Call Leer.Initialize(IniPath & GraphicsFile)
    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function
    End If
    
    exGrhCount = Val(Leer.GetValue("INIT", "NumGrh"))
    exfileVersion = Val(Leer.GetValue("INIT", "Version"))
    If (exfileVersion = 0) Then
        exfileVersion = 1
    ElseIf (exGrhCount > 200000 Or exGrhCount <= 0) Then
        MsgBox "La valor de 'NumGrh' es invalido!", vbCritical
        Exit Function
    End If

    ReDim exGrhData(1 To exGrhCount) As GrhData
    
    For Grh = 1 To exGrhCount
        exGrhData(Grh).sX = 0
        exGrhData(Grh).sY = 0
        exGrhData(Grh).pixelWidth = 0
        exGrhData(Grh).pixelHeight = 0
        exGrhData(Grh).FileNum = 0
        exGrhData(Grh).NumFrames = 0
        exGrhData(Grh).Speed = 0
        Datos = Leer.GetValue("Graphics", "Grh" & Grh)
        If LenB(Datos) <> 0 Then
            DatoR() = Split(Datos, "-")
            If DatoR(0) > 1 Then
                exGrhData(Grh).NumFrames = Val(DatoR(0))
                ReDim exGrhData(Grh).Frames(1 To exGrhData(Grh).NumFrames)
                tF = 1
                While Not exGrhData(Grh).NumFrames < tF
                    exGrhData(Grh).Frames(tF) = Val(DatoR(tF))
                    tF = tF + 1
                Wend
                exGrhData(Grh).Speed = Val(DatoR(tF))
            ElseIf DatoR(0) = 1 Then
                exGrhData(Grh).NumFrames = Val(DatoR(0))
                exGrhData(Grh).FileNum = Val(DatoR(1))
                exGrhData(Grh).sX = Val(DatoR(2))
                exGrhData(Grh).sY = Val(DatoR(3))
                exGrhData(Grh).pixelWidth = Val(DatoR(4))
                exGrhData(Grh).pixelHeight = Val(DatoR(5))
            End If
        End If
    Next
    
    ExplorerGrh = True
Exit Function

ErrorHandler:
End Function
    
