VERSION 5.00
Begin VB.Form DirGrhEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direction Grh Editor"
   ClientHeight    =   2730
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   4920
   Icon            =   "DirGrhEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DirGrhEdit.frx":08CA
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtSetY 
      Height          =   285
      Left            =   1260
      TabIndex        =   9
      Top             =   1380
      Width           =   375
   End
   Begin VB.TextBox TxtSetX 
      Height          =   285
      Left            =   555
      TabIndex        =   8
      Top             =   1380
      Width           =   375
   End
   Begin VB.Timer THeads 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   0
   End
   Begin VB.Timer TAllDirections 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4560
      Top             =   0
   End
   Begin VB.PictureBox PbDir 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   1215
      Index           =   3
      Left            =   3480
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox PbDir 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   1215
      Index           =   2
      Left            =   2040
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox PbDir 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   1215
      Index           =   1
      Left            =   3480
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox PbDir 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   1215
      Index           =   0
      Left            =   2040
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox TxtDir 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   3
      Top             =   1020
      Width           =   1095
   End
   Begin VB.TextBox TxtDir 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   690
      Width           =   1095
   End
   Begin VB.TextBox TxtDir 
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtDir 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   30
      Width           =   1095
   End
   Begin VB.Image ImgHxHy 
      Height          =   330
      Left            =   0
      Picture         =   "DirGrhEdit.frx":3ADCE
      Top             =   1380
      Width           =   1395
   End
   Begin VB.Label CmdCancelar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label CmdAgregar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label CmdAplicar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "DirGrhEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FrameActual As Long
Dim FrameActualAnim(1 To 4) As Long
Dim CantFramesAnim(1 To 4) As Byte
Dim AnimType As Byte
Dim Index As Integer
Dim IndexHead As Long
Dim LongXHead As Integer
Dim LongYHead As Integer
Dim PosXHead As Integer
Dim PosYHead As Integer
Dim HeadModo As Boolean
Dim AnimMode As String

Public Sub DirGrhChange(Modo As String, Indice As Integer)
Dim T As Byte
Dim GraficosPath(1 To 4) As String

Index = Indice
AnimMode = Modo

ImgHxHy.Visible = False
TxtSetY.Visible = False
TxtSetX.Visible = False

Select Case Modo
    Case "Cuerpos"
        For T = 1 To 4
            CantFramesAnim(T) = GrhData(Bodys(Index).Body(T)).NumFrames
            FrameActualAnim(T) = 0
            TxtDir(T).Text = Bodys(Index).Body(T)
            PbDir(T - 1).Cls
        Next T
        
        TxtSetX.Text = Bodys(Index).HeadOffsetX
        TxtSetY.Text = Bodys(Index).HeadOffsetY
            
        ImgHxHy.Visible = True
        TxtSetY.Visible = True
        TxtSetX.Visible = True
        
        AnimType = 0
        TAllDirections.Enabled = True
    Case "Head"
        HeadModo = True
        THeads.Enabled = True
    Case "Cascos"
        HeadModo = False
        THeads.Enabled = True
    Case "Armas"
        For T = 1 To 4
            CantFramesAnim(T) = GrhData(Armas(Index).Arma(T)).NumFrames
            FrameActualAnim(T) = 0
            TxtDir(T).Text = Armas(Index).Arma(T)
            PbDir(T - 1).Cls
        Next T
        
        AnimType = 1
        TAllDirections.Enabled = True
    Case "Escudos"
       For T = 1 To 4
            CantFramesAnim(T) = GrhData(Escudos(Index).Escudo(T)).NumFrames
            FrameActualAnim(T) = 0
            TxtDir(T).Text = Escudos(Index).Escudo(T)
            PbDir(T - 1).Cls
        Next T
    
        AnimType = 2
        TAllDirections.Enabled = True
End Select
End Sub

Private Sub CmdAceptar_Click()

End Sub

Private Sub CmdAgregar_Click()
Dim T As Byte
Select Case AnimMode
    Case "Cuerpos"
        BodysCountNew = BodysCountNew + 1
        ReDim Preserve Bodys(0 To BodysCountNew) As tIndiceCuerpo
        For T = 1 To 4
            Bodys(BodysCountNew).Body(T) = TxtDir(T).Text
        Next T
        
        Bodys(BodysCountNew).HeadOffsetX = TxtSetX.Text
        Bodys(BodysCountNew).HeadOffsetY = TxtSetY.Text
    Case "Head"
        HeadsCountNew = HeadsCountNew + 1
        ReDim Preserve Heads(0 To HeadsCountNew) As tIndiceCabeza
        For T = 1 To 4
            Heads(HeadsCountNew).Head(T) = TxtDir(T).Text
        Next T
    Case "Cascos"
        CascosCountNew = CascosCountNew + 1
        ReDim Preserve Cascos(0 To CascosCountNew) As tIndiceCasco
        For T = 1 To 4
            Cascos(CascosCountNew).Casco(T) = TxtDir(T).Text
        Next T
    Case "Armas"
        ArmasCountNew = ArmasCountNew + 1
        ReDim Preserve Armas(1 To ArmasCountNew) As tIndiceArmas
        For T = 1 To 4
            Armas(ArmasCountNew).Arma(T) = TxtDir(T).Text
        Next T
    Case "Escudos"
        EscudosCountNew = EscudosCountNew + 1
        ReDim Preserve Escudos(1 To EscudosCountNew) As tIndiceEscudos
        For T = 1 To 4
            Escudos(EscudosCountNew).Escudo(T) = TxtDir(T).Text
        Next T
End Select

MsgBox "Movimiento Agregado"
End Sub

Private Sub CmdAplicar_Click()
Dim T As Byte
Select Case AnimMode
    Case "Cuerpos"
        For T = 1 To 4
            Bodys(Index).Body(T) = TxtDir(T).Text
        Next T
        
        Bodys(Index).HeadOffsetX = TxtSetX.Text
        Bodys(Index).HeadOffsetY = TxtSetY.Text
    Case "Head"
        For T = 1 To 4
            Heads(Index).Head(T) = TxtDir(T).Text
        Next T
    Case "Cascos"
        For T = 1 To 4
            Cascos(Index).Casco(T) = TxtDir(T).Text
        Next T
    Case "Armas"
        For T = 1 To 4
            Armas(Index).Arma(T) = TxtDir(T).Text
        Next T
    Case "Escudos"
        For T = 1 To 4
            Escudos(Index).Escudo(T) = TxtDir(T).Text
        Next T
End Select
End Sub

Private Sub CmdCancelar_Click()
Me.Visible = False
TAllDirections.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
TAllDirections.Enabled = False
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub TAllDirections_Timer()
On Error GoTo ErrHandler

Dim AnimacionPosX(1 To 4) As Integer
Dim AnimacionPosY(1 To 4) As Integer
Dim AnimacionLongX(1 To 4) As Integer
Dim AnimacionLongY(1 To 4) As Integer
Dim GraficoPath(1 To 4) As String
Dim GrhIndexAnim(1 To 4) As Long
Dim Anim(1 To 4) As Integer
Dim T As Byte

For T = 1 To 4
    FrameActualAnim(T) = FrameActualAnim(T) + 1
    
    Anim(T) = Val(TxtDir(T).Text)
    If Anim(T) = 0 Then
        FrameActualAnim(T) = 0
        Exit Sub
    End If
    If GrhData(Anim(T)).NumFrames < 1 Or GrhData(Anim(T)).NumFrames < 2 Then
        FrameActualAnim(T) = 0
        Exit Sub
    End If
    
    If FrameActualAnim(T) > UBound(GrhData(Anim(T)).Frames()) Then
        FrameActualAnim(T) = 1
    End If

    GrhIndexAnim(T) = GrhData(Anim(T)).Frames(FrameActualAnim(T))
    
    If GrhIndexAnim(T) = 0 Then 'Por si hay error y no existe.
        TAllDirections.Enabled = False
        Exit Sub
    End If
    
    GraficoPath(T) = Config.BmpPath & "\" & GrhData(GrhIndexAnim(T)).FileNum & ".bmp" 'Busco la Imagen
    
    If Not FileExist(GraficoPath(T), vbNormal) = True Then 'Me fijo si Existe
        TAllDirections.Enabled = False
        Exit Sub
    End If
    
    AnimacionPosX(T) = GrhData(GrhIndexAnim(T)).sX 'Coordenada X
    AnimacionPosY(T) = GrhData(GrhIndexAnim(T)).sY 'Coordenada Y
    AnimacionLongX(T) = GrhData(GrhIndexAnim(T)).pixelWidth 'Longitud sobre X
    AnimacionLongY(T) = GrhData(GrhIndexAnim(T)).pixelHeight 'Longitud sobre Y
    
    PbDir(T - 1).PaintPicture LoadPicture(GraficoPath(T)), 0, 0, AnimacionLongX(T), AnimacionLongY(T), AnimacionPosX(T), AnimacionPosY(T), AnimacionLongX(T), AnimacionLongY(T)
    
    If FrameActualAnim(T) = CantFramesAnim(T) Then
        FrameActualAnim(T) = 0
    End If
    If FrameActualAnim(T) > UBound(GrhData(Anim(T)).Frames()) Then
        FrameActualAnim(T) = 0
    End If
Next T

Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    TAllDirections.Enabled = False
    Exit Sub

End Sub

Private Sub THeads_Timer()
On Error GoTo ErrHandler

Dim T As Byte
Dim GraficosPath(1 To 4) As String

For T = 1 To 4
    If HeadModo = True Then
        IndexHead = Heads(Index).Head(T)
    Else
        IndexHead = Cascos(Index).Casco(T)
    End If
    
    TxtDir(T).Text = IndexHead
    
    LongXHead = GrhData(IndexHead).pixelWidth
    LongYHead = GrhData(IndexHead).pixelHeight
    PosXHead = GrhData(IndexHead).sX
    PosYHead = GrhData(IndexHead).sY

    GraficosPath(T) = Config.BmpPath & "\" & GrhData(IndexHead).FileNum & ".bmp"
    
    If FileExist(GraficosPath(T), vbNormal) = True Then
        PbDir(T - 1).Cls
        PbDir(T - 1).PaintPicture LoadPicture(GraficosPath(T)), 0, 0, LongXHead, LongYHead, PosXHead, PosYHead, LongXHead, LongYHead
    End If
Next T

THeads.Enabled = False

Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    THeads.Enabled = False
    Exit Sub
End Sub

Private Sub TxtDir_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    If AnimMode = "Cuerpos" Or AnimMode = "Escudos" Or AnimMode = "Armas" Then
        TAllDirections.Enabled = False
        Select Case Index
            Case 1
                SelectAnim.SelectAnims ("Norte")
            Case 2
                SelectAnim.SelectAnims ("Este")
            Case 3
                SelectAnim.SelectAnims ("Sur")
            Case 4
                SelectAnim.SelectAnims ("Oeste")
        End Select
    
        Call SelectAnim.Show(vbModal, Me)
    Else
        MsgBox "Solo puedes elegir animaciones para Cuerpos, Escudos y Armas"
    End If
End If

If KeyCode = vbKeyF5 Then
    Call GrhEdit.GrhChange("General", Val(TxtDir(Index).Text))
    Call GrhEdit.Show(vbModal, Me)
End If
End Sub

Public Sub ReplaceDir(Directions As String, IndexDir As Long)
Dim T As Byte

For T = 1 To 4
    FrameActualAnim(T) = 0
Next T

Select Case Directions
    Case "Norte"
        TxtDir(1).Text = IndexDir
        PbDir(0).Cls
    Case "Este"
        TxtDir(3).Text = IndexDir
        PbDir(2).Cls
    Case "Sur"
        TxtDir(2).Text = IndexDir
        PbDir(1).Cls
    Case "Oeste"
        TxtDir(4).Text = IndexDir
        PbDir(3).Cls
End Select
TAllDirections.Enabled = True
End Sub
