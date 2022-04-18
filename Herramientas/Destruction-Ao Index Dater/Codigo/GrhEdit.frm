VERSION 5.00
Begin VB.Form GrhEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direction Grh Editor"
   ClientHeight    =   3195
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   6450
   Icon            =   "GrhEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GrhEdit.frx":08CA
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Draw 
      Interval        =   1
      Left            =   5520
      Top             =   0
   End
   Begin VB.Timer TAnimacion 
      Enabled         =   0   'False
      Left            =   6000
      Top             =   0
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox LstAnim 
      Height          =   1815
      ItemData        =   "GrhEdit.frx":5A026
      Left            =   2160
      List            =   "GrhEdit.frx":5A028
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox TxtSpeed 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox PbShow 
      BackColor       =   &H00FF00FF&
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox TxtPixelHeight 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox TxtPixelWidth 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox TxtSy 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox TxtSx 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxtFileNum 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox TxtNumFrames 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label CmdCancel 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label CmdCreateNew 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label CmdAccept 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image CmdEjes 
      Height          =   435
      Left            =   0
      Picture         =   "GrhEdit.frx":5A02A
      Top             =   2025
      Width           =   1860
   End
   Begin VB.Image CmdContinuar 
      Height          =   420
      Left            =   -15
      Picture         =   "GrhEdit.frx":5D89E
      Top             =   1620
      Width           =   1860
   End
   Begin VB.Image ImgAnim 
      Height          =   660
      Left            =   0
      Picture         =   "GrhEdit.frx":60F22
      Top             =   360
      Width           =   1755
   End
   Begin VB.Image ImgSpeed 
      Height          =   405
      Left            =   0
      Picture         =   "GrhEdit.frx":65FD6
      Top             =   2460
      Width           =   1755
   End
   Begin VB.Image CmdAplicar 
      Height          =   435
      Left            =   0
      Picture         =   "GrhEdit.frx":69176
      Top             =   1200
      Width           =   1890
   End
   Begin VB.Image ImgTodo 
      Height          =   1935
      Left            =   -45
      Picture         =   "GrhEdit.frx":6CAD2
      Top             =   360
      Width           =   2070
   End
End
Attribute VB_Name = "GrhEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GrhToEdit As Long
Dim FrameActual As Long
Dim CantFrames As Byte
Dim AnimChange() As Long
Dim GraficoPath As String
Dim EditEjes As Integer
Public Sub GrhChange(EditType As String, GrhIndex As Integer)
GrhToEdit = GrhIndex
TxtEdit.Text = ""
Draw.Enabled = False
TAnimacion.Enabled = False
If Val(EditType) > 0 Then
    CmdEjes.Visible = True
    EditEjes = Val(EditType)
Else
    EditEjes = 0
    CmdEjes.Visible = False
End If

LstAnim.Clear 'Borro la lista antes de seguir
With GrhData(GrhToEdit)
    If .NumFrames > 1 Then
        Call Visibles("Animacion")
        Me.Caption = "Grh Editor ---> Grh: " & GrhIndex
        TxtNumFrames.Text = .NumFrames
        TxtNumFrames.Enabled = False
        
        Dim T As Byte
        ReDim AnimChange(1 To .NumFrames) As Long
        For T = 1 To .NumFrames
            LstAnim.AddItem .Frames(T)
            AnimChange(T) = .Frames(T)
        Next T
        
        TxtSpeed.Text = .Speed
        
        TAnimacion.Enabled = True
        If IndexMode = "12.1" Then
            TAnimacion.Interval = Round(.Speed / .NumFrames)
        Else
            TAnimacion.Interval = 100
        End If
    Else
        Call Visibles("GrhUnico")
        Me.Caption = "Grh Editor ---> Grh: " & GrhIndex
        TxtNumFrames.Text = .NumFrames
        TxtNumFrames.Enabled = False
        TxtFileNum.Text = .FileNum
        TxtSx.Text = .sX
        TxtSy.Text = .sY
        TxtPixelWidth.Text = .pixelWidth
        TxtPixelHeight.Text = .pixelHeight
        
        Draw.Enabled = True
    End If
End With

End Sub

Private Sub CmdAccept_Click()
Dim CharPos As Integer
Dim NewTxt As String

With GrhData(GrhToEdit)
    If .NumFrames > 1 Then
        .NumFrames = Val(TxtNumFrames.Text)
        Dim T As Integer
        For T = 1 To .NumFrames
            .Frames(T) = AnimChange(T)
        Next T

        .Speed = Val(TxtSpeed.Text)
    Else
        .NumFrames = Val(TxtNumFrames.Text)
        .FileNum = Val(TxtFileNum.Text)
        .sX = Val(TxtSx.Text)
        .sY = Val(TxtSy.Text)
        .pixelWidth = Val(TxtPixelWidth.Text)
        .pixelHeight = Val(TxtPixelHeight.Text)
    End If
End With

Me.Visible = False
TAnimacion.Enabled = False
End Sub

Private Sub CmdAplicar_Click()
If LstAnim.Text = "" Then Exit Sub

Dim Y As Byte
For Y = 1 To GrhData(GrhToEdit).NumFrames
    If AnimChange(Y) = LstAnim.Text Then
        AnimChange(Y) = TxtEdit.Text
    End If
Next Y

LstAnim.Clear
For Y = 1 To GrhData(GrhToEdit).NumFrames
    LstAnim.AddItem AnimChange(Y)
Next Y

TAnimacion.Enabled = True
FrameActual = 0
End Sub

Private Sub CmdCancel_Click()
    Me.Visible = False
    TAnimacion.Enabled = False
    FrameActual = 0
End Sub

Private Sub Visibles(VisibleType As String)
If VisibleType = "Animacion" Then
    'No Mostrar
    'Label2.Visible = False
    'Label3.Visible = False
    'Label4.Visible = False
    'Label5.Visible = False
    'Label6.Visible = False
    ImgTodo.Visible = False
    TxtFileNum.Visible = False
    TxtSx.Visible = False
    TxtSy.Visible = False
    TxtPixelWidth.Visible = False
    TxtPixelHeight.Visible = False
    
    'Mostrar
    TxtSpeed.Visible = True
    TxtEdit.Visible = True
    'Label7.Visible = True
    'Label8.Visible = True
    'Label9.Visible = True
    ImgSpeed.Visible = True
    ImgAnim.Visible = True
    LstAnim.Visible = True
    CmdAplicar.Visible = True
    CmdContinuar.Visible = True
Else
    'No Mostrar
    TxtSpeed.Visible = False
    TxtEdit.Visible = False
    'Label7.Visible = False
    'Label8.Visible = False
    'Label9.Visible = False
    ImgSpeed.Visible = False
    ImgAnim.Visible = False
    LstAnim.Visible = False
    CmdAplicar.Visible = False
    CmdContinuar.Visible = False
    
    'Mostrar
    'Label2.Visible = True
    'Label3.Visible = True
    'Label4.Visible = True
    'Label5.Visible = True
    'Label6.Visible = True
    ImgTodo.Visible = True
    TxtNumFrames.Visible = True
    TxtFileNum.Visible = True
    TxtSx.Visible = True
    TxtSy.Visible = True
    TxtPixelWidth.Visible = True
    TxtPixelHeight.Visible = True
End If
End Sub

Private Sub CmdContinuar_Click()
FrameActual = 0
TAnimacion.Enabled = True
End Sub

Private Sub CmdCreateNew_Click()
Dim CreateNewGrh As Integer
CreateNewGrh = Val(InputBox("Cual va a ser el nuevo Grh?" & vbCrLf & "El ultimo utilizado es el: " & AllGrhData, "Creacion de nuevo Grh a base de otro"))
If CreateNewGrh = 0 Then Exit Sub

If CreateNewGrh > AllGrhData Then
    ReDim Preserve GrhData(1 To CreateNewGrh) As GrhData
    AllGrhData = CreateNewGrh
End If

If GrhData(CreateNewGrh).NumFrames > 0 Then
    MsgBox "Ese Grh ya esta ocupado elige otro"
    Exit Sub
End If

With GrhData(CreateNewGrh)
    If Val(TxtNumFrames.Text) > 1 Then
        .NumFrames = Val(TxtNumFrames.Text)
        Dim T As Integer
        ReDim .Frames(1 To .NumFrames) As Long
        For T = 1 To Val(TxtNumFrames.Text)
            .Frames(T) = AnimChange(T)
        Next T

        .Speed = Val(TxtSpeed.Text)
    Else
        .NumFrames = Val(TxtNumFrames.Text)
        .FileNum = Val(TxtFileNum.Text)
        .sX = Val(TxtSx.Text)
        .sY = Val(TxtSy.Text)
        .pixelWidth = Val(TxtPixelWidth.Text)
        .pixelHeight = Val(TxtPixelHeight.Text)
    End If
End With
End Sub

Private Sub CmdEjes_Click()
Dim HX As Byte
Dim HY As Byte

HX = Val(InputBox("Elige la posicion en el Eje X para la Animacion", "Posicion sobre el Eje X"))
HY = Val(InputBox("Elige la posicion en el Eje Y para la Animacion", "Posicion sobre el Eje Y"))

Fx(EditEjes).OffsetX = HX
Fx(EditEjes).OffsetY = HY

End Sub

Private Sub Draw_Timer()
On Error GoTo ErrHandler
If TxtFileNum.Text <> "" And Val(TxtPixelWidth.Text) > 0 And Val(TxtPixelHeight.Text) > 0 Then
    GraficoPath = Config.BmpPath & "\" & TxtFileNum.Text & ".bmp"
    If FileExist(GraficoPath, vbNormal) = True Then
        PbShow.Cls
        PbShow.PaintPicture LoadPicture(GraficoPath), 0, 0, Val(TxtPixelWidth.Text), Val(TxtPixelHeight.Text), Val(TxtSx.Text), Val(TxtSy.Text), Val(TxtPixelWidth.Text), Val(TxtPixelHeight.Text)
    End If
End If
Draw.Enabled = False
Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    Draw.Enabled = False
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
TAnimacion.Enabled = False
FrameActual = 0
End Sub

Private Sub Label1_Click()

End Sub

Private Sub LstAnim_Click()
TxtEdit.Text = LstAnim.Text
TAnimacion.Enabled = False
FrameActual = 0

Dim MuestraSX As Integer
Dim MuestraSY As Integer
Dim MuestraPixelW As Integer
Dim MuestraPixelH As Integer

With GrhData(Val(LstAnim.Text))
    GraficoPath = Config.BmpPath & "\" & .FileNum & ".bmp"
    If FileExist(GraficoPath, vbNormal) = True Then
        PbShow.PaintPicture LoadPicture(GraficoPath), 0, 0, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight
    End If
End With
End Sub

Private Sub LstAnim_DblClick()
TAnimacion.Enabled = False
Call GrhEdit.GrhChange("General", Val(LstAnim.Text))
End Sub

Private Sub TAnimacion_Timer()
On Error GoTo ErrHandler
Dim AnimacionPosX As Integer
Dim AnimacionPosY As Integer
Dim AnimacionLongX As Integer
Dim AnimacionLongY As Integer
Dim GrhIndexAnim As Long

FrameActual = FrameActual + 1

If AnimChange(FrameActual) = 1 Then Exit Sub

GrhIndexAnim = AnimChange(FrameActual)

If GrhIndexAnim < 1 Then Exit Sub

If GrhIndexAnim > 0 Then
    GraficoPath = Config.BmpPath & "\" & GrhData(GrhIndexAnim).FileNum & ".bmp"
    AnimacionPosX = GrhData(GrhIndexAnim).sX
    AnimacionPosY = GrhData(GrhIndexAnim).sY
    AnimacionLongX = GrhData(GrhIndexAnim).pixelWidth
    AnimacionLongY = GrhData(GrhIndexAnim).pixelHeight
    
    If FileExist(GraficoPath, vbNormal) = True Then
        PbShow.PaintPicture LoadPicture(GraficoPath), 0, 0, AnimacionLongX, AnimacionLongY, AnimacionPosX, AnimacionPosY, AnimacionLongX, AnimacionLongY
    End If
End If

If FrameActual = GrhData(GrhToEdit).NumFrames Then
    FrameActual = 0
End If

Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    TAnimacion.Enabled = False
    Exit Sub
End Sub

Private Sub TxtFileNum_Change()
Draw.Enabled = True
End Sub

Private Sub TxtPixelHeight_Change()
Draw.Enabled = True
End Sub

Private Sub TxtPixelWidth_Change()
Draw.Enabled = True
End Sub

Private Sub TxtSx_Change()
Draw.Enabled = True
End Sub

Private Sub TxtSy_Change()
Draw.Enabled = True
End Sub
