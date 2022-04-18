VERSION 5.00
Begin VB.Form FrmCarga 
   BorderStyle     =   0  'None
   Caption         =   "Cargando"
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   135
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PbCargando 
      Height          =   2055
      Left            =   0
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   0
         Picture         =   "FrmCarga.frx":0000
         Top             =   0
         Width           =   5820
      End
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
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image ImgCargando 
         Height          =   495
         Left            =   0
         Picture         =   "FrmCarga.frx":26494
         Top             =   1500
         Width           =   7155
      End
   End
End
Attribute VB_Name = "FrmCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call LoadCuerpos
    Call LoadFxs
    Call LoadCabezas
    Call LoadCascos
    Call LoadGrhData(Config.InitPath)
    Call LoadArmas
    Call LoadEscudos
    
    Dim Cuerpos As Integer
    Dim Efectos As Integer
    Dim Cabeza As Integer
    Dim Casco As Integer
    Dim Grhs As Integer
    Dim Shields As Integer
    Dim Weapon As Integer
    
    For Cuerpos = 1 To BodysCountOld
        If Bodys(Cuerpos).Body(1) > 0 Then
            FrmAnimaciones.LstCuerpos.AddItem Cuerpos
        End If
    Next Cuerpos
    
    For Efectos = 1 To FxCountOld
        If Fx(Efectos).Animacion > 0 Then
            FrmAnimaciones.LstFx.AddItem Efectos
        End If
    Next Efectos
    
    For Cabeza = 1 To HeadsCountOld
        If Heads(Cabeza).Head(1) > 0 Then
            FrmAnimaciones.LstCabezas.AddItem Cabeza
        End If
    Next Cabeza
    
    For Casco = 1 To CascosCountOld
        If Cascos(Casco).Casco(1) > 0 Then
            FrmAnimaciones.LstCascos.AddItem Casco
        End If
    Next Casco
    
    For Shields = 1 To EscudosCountOld
        If Escudos(Shields).Escudo(1) > 0 Then
            FrmAnimaciones.LstEscudos.AddItem Shields
        End If
    Next Shields
    
    For Weapon = 1 To ArmasCountOld
        If Armas(Weapon).Arma(1) > 0 Then
            FrmAnimaciones.LstArmas.AddItem Weapon
        End If
    Next Weapon
    
    Dim Cargando As Integer
    Dim EstadoCarga As Byte
    Dim ImgCarga As Integer
    
    ImgCarga = CInt(389 / 10)
    Cargando = CInt(AllGrhData / 10)
    EstadoCarga = 1
    
    FrmAnimaciones.LstGeneral.Clear
    
    For Grhs = 1 To AllGrhData
        If GrhData(Grhs).NumFrames > 1 Then
            FrmAnimaciones.LstGeneral.AddItem Grhs & "(ANIMACION)"
        Else
            If GrhData(Grhs).NumFrames <> 0 Then
                FrmAnimaciones.LstGeneral.AddItem Grhs
            End If
        End If
        If CInt(Grhs / Cargando) >= EstadoCarga Then
            ImgCargando.Width = ImgCarga * EstadoCarga
            EstadoCarga = EstadoCarga + 1
        End If
    Next Grhs
    
    Me.Visible = False
    Call Unload(Me)
    FrmInicio.Visible = True
End Sub

Private Sub Form_Load()
ImgCargando.Width = 1
End Sub

