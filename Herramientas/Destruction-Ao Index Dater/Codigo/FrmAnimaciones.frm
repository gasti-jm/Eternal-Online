VERSION 5.00
Begin VB.Form FrmAnimaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vision General de Indexacion"
   ClientHeight    =   6660
   ClientLeft      =   3510
   ClientTop       =   1020
   ClientWidth     =   10815
   Icon            =   "FrmAnimaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAnimaciones.frx":08CA
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   721
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CbSearch 
      Height          =   315
      Left            =   8160
      TabIndex        =   15
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox SearchText 
      Height          =   285
      Left            =   6840
      TabIndex        =   14
      Top             =   180
      Width           =   1215
   End
   Begin VB.PictureBox PbCargando 
      Height          =   2055
      Left            =   2700
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   8
      Top             =   1950
      Width           =   5835
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image ImgCargando 
         Height          =   495
         Left            =   0
         Picture         =   "FrmAnimaciones.frx":1392FE
         Top             =   1500
         Width           =   7155
      End
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   0
         Picture         =   "FrmAnimaciones.frx":148936
         Top             =   0
         Width           =   5820
      End
   End
   Begin VB.PictureBox PbShowAnim 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2850
      Index           =   1
      Left            =   4800
      ScaleHeight     =   186
      ScaleMode       =   0  'User
      ScaleWidth      =   191
      TabIndex        =   11
      Top             =   480
      Width           =   2925
   End
   Begin VB.PictureBox PbShowAnim 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2850
      Index           =   3
      Left            =   7800
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   13
      Top             =   3360
      Width           =   2925
   End
   Begin VB.PictureBox PbShowAnim 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2850
      Index           =   2
      Left            =   4800
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   12
      Top             =   3360
      Width           =   2925
   End
   Begin VB.PictureBox PbShowAnim 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2850
      Index           =   0
      Left            =   7800
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   10
      Top             =   480
      Width           =   2925
   End
   Begin VB.Timer TAllDirections 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9840
      Top             =   0
   End
   Begin VB.Timer TAnimacion 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10320
      Top             =   0
   End
   Begin VB.PictureBox PbShow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   5775
      Left            =   4800
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   7
      Top             =   480
      Width           =   6015
   End
   Begin VB.ListBox LstGeneral 
      Height          =   5715
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox LstFx 
      Height          =   1620
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LstEscudos 
      Height          =   1620
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LstArmas 
      Height          =   1620
      Left            =   1680
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LstCabezas 
      Height          =   1620
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LstCascos 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LstCuerpos 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label CmdVolver 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9720
      TabIndex        =   19
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label CmdReload 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label CmdGuardar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label CmdIndex 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Menu MRapido 
      Caption         =   "Menu Rapido"
      Begin VB.Menu MInicio 
         Caption         =   "Inicio"
      End
      Begin VB.Menu MIndexacion 
         Caption         =   "Indexacion"
         Begin VB.Menu MIndexar 
            Caption         =   "Indexar"
         End
      End
      Begin VB.Menu MDateo 
         Caption         =   "Dateo"
         Begin VB.Menu MObjetos 
            Caption         =   "Objetos"
         End
         Begin VB.Menu MNpcs 
            Caption         =   "Npcs"
         End
         Begin VB.Menu MHechizos 
            Caption         =   "Hechizos"
         End
      End
      Begin VB.Menu MConversor 
         Caption         =   "Conversor de Imagenes"
      End
      Begin VB.Menu MRutas 
         Caption         =   "Configurar Rutas"
      End
      Begin VB.Menu MCreditos 
         Caption         =   "Creditos"
      End
      Begin VB.Menu MSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FrmAnimaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LB_FINDSTRING = &H18F
Private Declare Function sendmessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Dim CantFrames As Byte
Dim CantFramesAnim(1 To 4) As Byte
Dim GrhIndex As Long
Dim FrameActual As Long
Dim FrameActualAnim(1 To 4) As Long
Dim BodyIndex As Integer
Dim FxIndex As Integer
Dim HeadIndex As Integer
Dim CascoIndex As Integer
Dim EscudoIndex As Integer
Dim ArmasIndex As Integer
Dim AnimType As Byte
Dim Index As Integer
Dim FirstCharge As Boolean

Private Sub CmdGuardar_Click()

If Grh.FileVersion = -1 Then
    If Not Grh.SaveGrhDataOld(Config.SaveInitPath) Then
        Call MsgBox("Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias.")
        Exit Sub
    Else
        Call MsgBox("Se han guardado los Indices de forma correcta.")
    End If
Else
    If Not Grh.SaveGrhDataNew(Config.SaveInitPath) Then
        Call MsgBox("Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias.")
    Else
        Call MsgBox("Se han guardado los Indices de forma correcta.")
    End If
End If

Call SaveFX
Call SaveEscudos
Call SaveArmas
Call SaveCascos
Call SaveCabezas
Call SaveCuerpos

MsgBox "Se guardaron todos los datos con Exito"

TAllDirections.Enabled = False
TAnimacion.Enabled = False
FrameActual = 0
End Sub

Private Sub CmdIndex_Click()
FrmIndex.Visible = True
End Sub

Private Sub CmdReload_Click()
LstGeneral.Clear
LstCuerpos.Clear
LstFx.Clear
LstCabezas.Clear
LstCascos.Clear
LstEscudos.Clear
LstArmas.Clear

ImgCargando.Width = 1
PbCargando.Visible = True

Dim Cuerpos As Integer
Dim Efectos As Integer
Dim Cabeza As Integer
Dim Casco As Integer
Dim Grhs As Integer
Dim Shields As Integer
Dim Weapon As Integer
    
    For Cuerpos = 0 To BodysCountNew
        If Bodys(Cuerpos).Body(1) > 0 Then
            LstCuerpos.AddItem Cuerpos
        End If
    Next Cuerpos
    
    For Efectos = 1 To FxCountNew
        If Fx(Efectos).Animacion > 0 Then
            LstFx.AddItem Efectos
        End If
    Next Efectos
    
    For Cabeza = 0 To HeadsCountNew
        If Heads(Cabeza).Head(1) > 0 Then
            LstCabezas.AddItem Cabeza
        End If
    Next Cabeza
    
    For Casco = 0 To CascosCountNew
        If Cascos(Casco).Casco(1) > 0 Then
            LstCascos.AddItem Casco
        End If
    Next Casco
    
    For Shields = 1 To EscudosCountNew
        If Escudos(Shields).Escudo(1) > 0 Then
            LstEscudos.AddItem Shields
        End If
    Next Shields
    
    For Weapon = 1 To ArmasCountNew
        If Armas(Weapon).Arma(1) > 0 Then
            LstArmas.AddItem Weapon
        End If
    Next Weapon

Dim Cargando As Integer
Dim EstadoCarga As Byte
Dim ImgCarga As Integer

ImgCarga = CInt(477 / 10)
Cargando = CInt(AllGrhData / 10)
EstadoCarga = 1
For Grhs = 1 To AllGrhData
    If GrhData(Grhs).NumFrames > 1 Then
        LstGeneral.AddItem Grhs & "(ANIMACION)"
    Else
        If GrhData(Grhs).NumFrames <> 0 Then
            LstGeneral.AddItem Grhs
        End If
    End If
    If CInt(Grhs / Cargando) >= EstadoCarga Then
        ImgCargando.Width = ImgCarga * EstadoCarga
        EstadoCarga = EstadoCarga + 1
    End If
Next Grhs

PbCargando.Visible = False
End Sub

Private Sub CmdVolver_Click()
Me.Visible = False
FrmIndexMenu.Visible = True
End Sub

Private Sub Form_Activate()
If FirstCharge = False Then
    'Call LoadCuerpos
    'Call LoadFxs
    'Call LoadCabezas
    'Call LoadCascos
    'Call LoadGrhData(Config.InitPath)
    'Call LoadArmas
    'Call LoadEscudos
    
    Dim Cuerpos As Integer
    Dim Efectos As Integer
    Dim Cabeza As Integer
    Dim Casco As Integer
    Dim Grhs As Integer
    Dim Shields As Integer
    Dim Weapon As Integer
    
    For Cuerpos = 1 To BodysCountOld
        If Bodys(Cuerpos).Body(1) > 0 Then
            LstCuerpos.AddItem Cuerpos
        End If
    Next Cuerpos
    
    For Efectos = 1 To FxCountOld
        If Fx(Efectos).Animacion > 0 Then
            LstFx.AddItem Efectos
        End If
    Next Efectos
    
    For Cabeza = 1 To HeadsCountOld
        If Heads(Cabeza).Head(1) > 0 Then
            LstCabezas.AddItem Cabeza
        End If
    Next Cabeza
    
    For Casco = 1 To CascosCountOld
        If Cascos(Casco).Casco(1) > 0 Then
            LstCascos.AddItem Casco
        End If
    Next Casco
    
    For Shields = 1 To EscudosCountOld
        If Escudos(Shields).Escudo(1) > 0 Then
            LstEscudos.AddItem Shields
        End If
    Next Shields
    
    For Weapon = 1 To ArmasCountOld
        If Armas(Weapon).Arma(1) > 0 Then
            LstArmas.AddItem Weapon
        End If
    Next Weapon
    
    Dim Cargando As Integer
    Dim EstadoCarga As Byte
    Dim ImgCarga As Integer
    
    ImgCarga = CInt(389 / 10)
    Cargando = CInt(AllGrhData / 10)
    EstadoCarga = 1
    
    LstGeneral.Clear
    
    For Grhs = 1 To AllGrhData
        If GrhData(Grhs).NumFrames > 1 Then
            LstGeneral.AddItem Grhs & "(ANIMACION)"
        Else
            If GrhData(Grhs).NumFrames <> 0 Then
                LstGeneral.AddItem Grhs
            End If
        End If
        If CInt(Grhs / Cargando) >= EstadoCarga Then
            ImgCargando.Width = ImgCarga * EstadoCarga
            EstadoCarga = EstadoCarga + 1
        End If
    Next Grhs
    
    PbCargando.Visible = False
    LstCuerpos.Visible = True
    LstFx.Visible = True
    LstCabezas.Visible = True
    LstCascos.Visible = True
    LstEscudos.Visible = True
    LstArmas.Visible = True
    LstGeneral.Visible = True
    
    FirstCharge = True
End If
End Sub

Private Sub Form_Load()
If Not LoadConfig() Then
    Call frmConfig.Show(vbModal, Me)
End If

FirstCharge = False

Dim T As Byte

For T = 1 To 4
    PbShowAnim(T - 1).Visible = False
Next T

CbSearch.AddItem "General", 0
CbSearch.AddItem "Cuerpos", 1
CbSearch.AddItem "Fx", 2
CbSearch.AddItem "Cascos", 3
CbSearch.AddItem "Cabezas", 4
CbSearch.AddItem "Escudos", 5
CbSearch.AddItem "Armas", 6

CbSearch.ListIndex = 0

ImgCargando.Width = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmIndexMenu.Visible = True
End Sub
Private Sub LstArmas_Click()
On Error GoTo ErrHandler
Dim ArmaFrames As Byte
Dim T As Byte
ArmasIndex = Val(LstArmas.Text)

If ArmasIndex = 0 Then Exit Sub

For T = 1 To 4
    CantFramesAnim(T) = GrhData(Armas(ArmasIndex).Arma(T)).NumFrames
    FrameActualAnim(T) = 0
    PbShowAnim(T - 1).Cls
    PbShowAnim(T - 1).Visible = True
Next T

PbShow.Visible = False
Index = ArmasIndex
AnimType = 1
TAllDirections.Enabled = True
TAnimacion.Enabled = False

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub LstArmas_DblClick()
Call DirGrhEdit.DirGrhChange("Armas", Val(LstArmas.Text))
TAllDirections.Enabled = False
'DirGrhEdit.Visible = True
Call DirGrhEdit.Show(vbModal, Me)
End Sub

Private Sub LstCabezas_Click()
On Error GoTo ErrHandler
Dim GraficosPath(1 To 4) As String
Dim IndexHead As Long
Dim LongXHead As Integer
Dim LongYHead As Integer
Dim PosXHead As Integer
Dim PosYHead As Integer
HeadIndex = Val(LstCabezas.Text)

If HeadIndex = 0 Then Exit Sub

Dim T As Byte
For T = 1 To 4
    IndexHead = Heads(HeadIndex).Head(T)
    LongXHead = GrhData(IndexHead).pixelWidth
    LongYHead = GrhData(IndexHead).pixelHeight
    PosXHead = GrhData(IndexHead).sX
    PosYHead = GrhData(IndexHead).sY

    GraficosPath(T) = Config.BmpPath & "\" & GrhData(IndexHead).FileNum & ".bmp"
    If FileExist(GraficosPath(T), vbNormal) = True Then
        PbShowAnim(T - 1).Visible = True
        PbShowAnim(T - 1).Cls
        PbShowAnim(T - 1).PaintPicture LoadPicture(GraficosPath(T)), 0, 0, LongXHead, LongYHead, PosXHead, PosYHead, LongXHead, LongYHead
    End If
Next T

TAllDirections.Enabled = False
TAnimacion.Enabled = False
PbShow.Visible = False

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub

End Sub

Private Sub LstCabezas_DblClick()
Call DirGrhEdit.DirGrhChange("Head", Val(LstCabezas.Text))
TAllDirections.Enabled = False
'DirGrhEdit.Visible = True
Call DirGrhEdit.Show(vbModal, Me)
End Sub

Private Sub LstCascos_Click()
On Error GoTo ErrHandler
Dim GraficosPath(1 To 4) As String
Dim IndexCasco As Long
Dim LongXCasco As Integer
Dim LongYCasco As Integer
Dim PosXCasco As Integer
Dim PosYCasco As Integer
CascoIndex = Val(LstCascos.Text)

If CascoIndex = 0 Then Exit Sub

Dim T As Byte
For T = 1 To 4
    IndexCasco = Cascos(CascoIndex).Casco(T)
    LongXCasco = GrhData(IndexCasco).pixelWidth
    LongYCasco = GrhData(IndexCasco).pixelHeight
    PosXCasco = GrhData(IndexCasco).sX
    PosYCasco = GrhData(IndexCasco).sY
    
    GraficosPath(T) = Config.BmpPath & "\" & GrhData(IndexCasco).FileNum & ".bmp"
    If FileExist(GraficosPath(T), vbNormal) = True Then
        PbShowAnim(T - 1).Visible = True
        PbShowAnim(T - 1).Cls
        PbShowAnim(T - 1).PaintPicture LoadPicture(GraficosPath(T)), 0, 0, LongXCasco, LongYCasco, PosXCasco, PosYCasco, LongXCasco, LongYCasco
    End If
Next T

TAllDirections.Enabled = False
TAnimacion.Enabled = False
PbShow.Visible = False

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub LstCascos_DblClick()
Call DirGrhEdit.DirGrhChange("Cascos", Val(LstCascos.Text))
TAllDirections.Enabled = False
'DirGrhEdit.Visible = True
Call DirGrhEdit.Show(vbModal, Me)
End Sub

Private Sub LstCuerpos_Click()
On Error GoTo ErrHandler
Dim BodyFrames As Byte
Dim T As Byte
BodyIndex = Val(LstCuerpos.Text)

If BodyIndex = 0 Then Exit Sub

For T = 1 To 4
    CantFramesAnim(T) = GrhData(Bodys(BodyIndex).Body(T)).NumFrames
    FrameActualAnim(T) = 0
    PbShowAnim(T - 1).Cls
    PbShowAnim(T - 1).Visible = True
Next T

PbShow.Visible = False
Index = BodyIndex
AnimType = 0
TAllDirections.Enabled = True
TAnimacion.Enabled = False

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub LstCuerpos_DblClick()
Call DirGrhEdit.DirGrhChange("Cuerpos", Val(LstCuerpos.Text))
TAllDirections.Enabled = False
'DirGrhEdit.Visible = True
Call DirGrhEdit.Show(vbModal, Me)
End Sub

Private Sub LstEscudos_Click()
On Error GoTo ErrHandler
Dim EscudoFrames As Byte
Dim T As Byte
EscudoIndex = Val(LstEscudos.Text)

If EscudoIndex = 0 Then Exit Sub

For T = 1 To 4
    CantFramesAnim(T) = GrhData(Escudos(EscudoIndex).Escudo(T)).NumFrames
    FrameActualAnim(T) = 0
    PbShowAnim(T - 1).Cls
    PbShowAnim(T - 1).Visible = True
Next T

PbShow.Visible = False
Index = EscudoIndex
AnimType = 2
TAllDirections.Enabled = True
TAnimacion.Enabled = False

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub LstEscudos_DblClick()
Call DirGrhEdit.DirGrhChange("Escudos", Val(LstEscudos.Text))
TAllDirections.Enabled = False
'DirGrhEdit.Visible = True
Call DirGrhEdit.Show(vbModal, Me)
End Sub

Private Sub LstFx_Click()
On Error GoTo ErrHandler
FxIndex = Val(LstFx.Text)

If FxIndex = 0 Then Exit Sub

GrhIndex = Fx(FxIndex).Animacion
CantFrames = GrhData(GrhIndex).NumFrames

If GrhData(GrhIndex).Speed <> 0 And GrhData(GrhIndex).NumFrames <> 0 Then
    If IndexMode = "12.1" Then
        TAnimacion.Interval = Round(GrhData(GrhIndex).Speed / GrhData(GrhIndex).NumFrames)
    Else
        TAnimacion.Interval = 100
    End If
    TAnimacion.Enabled = True
Else
    TAnimacion.Enabled = False
End If
FrameActual = 0
TAllDirections.Enabled = False
PbShow.Cls
PbShow.Visible = True

Dim T As Byte

For T = 1 To 4
    PbShowAnim(T - 1).Visible = False
Next T
Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub

End Sub

Private Sub LstFx_DblClick()
Dim GrhFxIndex As Integer
GrhFxIndex = Val(Fx(Val(LstFx.Text)).Animacion)
Call GrhEdit.GrhChange(LstFx.Text, GrhFxIndex)
TAnimacion.Enabled = False
FrameActual = 0
'GrhEdit.Visible = True
Call GrhEdit.Show(vbModal, Me)
End Sub

Private Sub LstGeneral_Click()
On Error GoTo ErrHandler
Dim GraficoPath As String
Dim PosicionX As Integer
Dim PosicionY As Integer
Dim LongitudX As Integer
Dim LongitudY As Integer
Dim Frames As Byte

Dim T As Byte

For T = 1 To 4
    PbShowAnim(T - 1).Visible = False
Next T

PbShow.Visible = True

GrhIndex = Val(LstGeneral.Text)

If GrhIndex = 0 Then Exit Sub

If GrhData(GrhIndex).NumFrames > 1 Then
    CantFrames = GrhData(GrhIndex).NumFrames
    If IndexMode = "12.1" Then
        TAnimacion.Interval = Round(GrhData(GrhIndex).Speed / GrhData(GrhIndex).NumFrames)
    Else
        TAnimacion.Interval = 100
    End If
    FrameActual = 0
    TAnimacion.Enabled = True
    TAllDirections.Enabled = False
    PbShow.Cls
Else
    TAnimacion.Enabled = False
    If GrhData(GrhIndex).FileNum > 0 Then
        GraficoPath = Config.BmpPath & "\" & GrhData(GrhIndex).FileNum & ".bmp"
        PosicionX = GrhData(GrhIndex).sX
        PosicionY = GrhData(GrhIndex).sY
        LongitudX = GrhData(GrhIndex).pixelWidth
        LongitudY = GrhData(GrhIndex).pixelHeight
        If FileExist(GraficoPath, vbNormal) = True Then
            PbShow.Cls
            PbShow.PaintPicture LoadPicture(GraficoPath), 0, 0, LongitudX, LongitudY, PosicionX, PosicionY, LongitudX, LongitudY
        End If
    End If
End If

TAllDirections.Enabled = False

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub LstGeneral_DblClick()
Call GrhEdit.GrhChange("General", Val(LstGeneral.Text))
TAnimacion.Enabled = False
FrameActual = 0
'GrhEdit.Visible = True
Call GrhEdit.Show(vbModal, Me)
End Sub

Private Sub MConversor_Click()
FrmConversor.Visible = True
End Sub

Private Sub MRutas_Click()
frmConfig.Visible = True
End Sub

Private Sub MSalir_Click()
End
End Sub

Private Sub TAllDirections_Timer()
On Error GoTo ErrorAnim
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
    
    If AnimType = 0 Then 'Chequeo Tipo de Animacion
        Anim(T) = Bodys(Index).Body(T)
    ElseIf AnimType = 1 Then
        Anim(T) = Armas(Index).Arma(T)
    Else
        Anim(T) = Escudos(Index).Escudo(T)
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
    
    PbShowAnim(T - 1).PaintPicture LoadPicture(GraficoPath(T)), 0, 0, AnimacionLongX(T), AnimacionLongY(T), AnimacionPosX(T), AnimacionPosY(T), AnimacionLongX(T), AnimacionLongY(T)
    
    If FrameActualAnim(T) = CantFramesAnim(T) Then
        FrameActualAnim(T) = 0
    End If
Next T

Exit Sub

ErrorAnim:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    TAllDirections.Enabled = False
    Exit Sub
'For T = 1 To 4
'    FrameActualAnim(T) = 0
'Next T

End Sub

Private Sub TAnimacion_Timer()
On Error GoTo ErrHandler
Dim AnimacionPosX As Integer
Dim AnimacionPosY As Integer
Dim AnimacionLongX As Integer
Dim AnimacionLongY As Integer
Dim GraficoPath As String
Dim GrhIndexAnim As Long

FrameActual = FrameActual + 1

GrhIndexAnim = GrhData(GrhIndex).Frames(FrameActual)

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

If FrameActual = CantFrames Then
    FrameActual = 0
End If
Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    TAnimacion.Enabled = False
    Exit Sub

End Sub

Private Sub SearchText_Change()
Select Case CbSearch.ListIndex
    Case 0
        LstGeneral.ListIndex = sendmessage(LstGeneral.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
    Case 1
        LstCuerpos.ListIndex = sendmessage(LstCuerpos.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
    Case 2
        LstFx.ListIndex = sendmessage(LstFx.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
    Case 3
        LstCascos.ListIndex = sendmessage(LstCascos.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
    Case 4
        LstCabezas.ListIndex = sendmessage(LstCabezas.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
    Case 5
        LstEscudos.ListIndex = sendmessage(LstEscudos.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
    Case 6
        LstArmas.ListIndex = sendmessage(LstArmas.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
End Select
End Sub

Private Sub MCreditos_Click()
FrmCreditos.Visible = True
Me.Visible = False
End Sub

Private Sub MHechizos_Click()
FrmHechizosCreator.Visible = True
Me.Visible = False
End Sub

Private Sub MIndexar_Click()
FrmIndex.Visible = True
Me.Visible = False
End Sub

Private Sub MInicio_Click()
FrmInicio.Visible = True
Me.Visible = False
End Sub

Private Sub MNpcs_Click()
FrmNpcCreator.Visible = True
Me.Visible = False
End Sub

Private Sub MObjetos_Click()
FrmObjSelector.Visible = True
Me.Visible = False
End Sub
