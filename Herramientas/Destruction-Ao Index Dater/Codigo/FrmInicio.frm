VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form FrmInicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Inicio"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9225
   Icon            =   "FrmInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmInicio.frx":08CA
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   615
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ListBox LstNews 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      ItemData        =   "FrmInicio.frx":DD56E
      Left            =   2280
      List            =   "FrmInicio.frx":DD570
      TabIndex        =   6
      Top             =   2280
      Width           =   4695
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   1455
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"FrmInicio.frx":DD572
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton Op121 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5520
      TabIndex        =   4
      Top             =   1920
      Width           =   195
   End
   Begin VB.OptionButton Op112 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4320
      MaskColor       =   &H80000002&
      TabIndex        =   3
      Top             =   1920
      Width           =   195
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   5
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   8880
      TabIndex        =   7
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image CmdActualizar 
      Height          =   465
      Left            =   3510
      Picture         =   "FrmInicio.frx":DD5F7
      Top             =   4275
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Label LblCreditos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label LblDatear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label LblIndexar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Menu MRapido 
      Caption         =   "Menu Rapido"
      Begin VB.Menu MIndexacion 
         Caption         =   "Indexacion"
         Begin VB.Menu MIndexar 
            Caption         =   "Indexar"
         End
         Begin VB.Menu MDatosIndex 
            Caption         =   "Ver Datos de Indexacion"
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
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1
Dim ErrorFound As Boolean

Private Sub CmdActualizar_Click()
Dim ActVersion As String
Dim x

ActVersion = Inet1.OpenURL("www.aodestruction.com.ar/Noticias/VersionPath.txt")
x = ShellExecute(Me.hWnd, "Open", ActVersion, &O0, &O0, SW_NORMAL)
End Sub

Private Sub Form_Activate()
On Error GoTo ErrHandler
If LoadNews = False Then
    Dim CantidadNoticias As String
    Dim T As Integer
    CantidadNoticias = Inet1.OpenURL("www.aodestruction.com.ar/Noticias/IndexNoticias.txt")
    ReDim DAONews(0 To CantidadNoticias) As IDNews
    For T = 1 To CantidadNoticias
        DAONews(T - 1).Titulo = Inet1.OpenURL("www.aodestruction.com.ar/Noticias/Titulo" & T & ".txt")
        DAONews(T - 1).Noticia = Inet1.OpenURL("www.aodestruction.com.ar/Noticias/Noticia" & T & ".txt")
        LstNews.AddItem DAONews(T - 1).Titulo
    Next T
    
    Dim Version As String
    Dim ProgVersion As String
    Version = Inet1.OpenURL("www.aodestruction.com.ar/Noticias/Version.txt")
    ProgVersion = GetVar(App.path & "\DAOIndexDater.dao", "Version", "V")
    If Version <> ProgVersion Then
        CmdActualizar.Visible = True
    Else
        CmdActualizar.Visible = False
    End If
End If
LoadNews = True

Exit Sub
ErrHandler:
    If ErrorFound = False Then
        'MsgBox "El servidor de noticias y actualizaciones ah caido, pero puede seguir utilizando el Index Dater sin problemas."
        ErrorFound = True
    End If
    Inet1.Cancel
    Exit Sub
End Sub

Private Sub Form_Load()
If Not LoadConfig() Then
    Call frmConfig.Show(vbModal, Me)
End If
Op112.value = True
IndexDaterIni = App.path & "\IndexerDats.dao"
LoadNews = False
IndexMode = "11.X"
FrmCarga.Visible = True
ErrorFound = False
'Call LoadGrhData(Config.InitPath)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()
FrmInicio.Inet1.Cancel
End
End Sub

Private Sub LblCreditos_Click()
FrmCreditos.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub LblDatear_Click()
FrmDatMenu.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub LblIndexar_Click()
FrmIndexMenu.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub LstNews_Click()
On Error GoTo ErrHandler
Dim NewsIndex As Byte
NewsIndex = LstNews.ListIndex

Rtb.Text = DAONews(NewsIndex).Titulo
Rtb.SelStart = 0
Rtb.SelLength = Len(DAONews(NewsIndex).Titulo)
With Rtb
    .SelFontSize = 8
    .SelBold = True
    .SelFontName = "Verdana"
    .SelColor = vbBlue
    .SelBold = True
End With

Rtb.Text = Rtb.Text & vbCrLf & DAONews(NewsIndex).Noticia
Rtb.SelStart = Len(Rtb.Text) - Len(DAONews(NewsIndex).Noticia)
Rtb.SelLength = Len(DAONews(NewsIndex).Noticia)
With Rtb
    .SelFontSize = 7
    .SelBold = True
    .SelFontName = "Verdana"
    .SelColor = vbRed
    .SelBold = True
End With

Rtb.SelStart = 0
Rtb.SelLength = 0 'Len(Rtb.Text)
'Rtb.SelAlignment = 2
Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub MConversor_Click()
FrmConversor.Visible = True
End Sub

Private Sub MCreditos_Click()
FrmCreditos.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub MDatosIndex_Click()
FrmAnimaciones.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub MHechizos_Click()
FrmHechizosCreator.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub MIndexar_Click()
FrmIndex.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub MNpcs_Click()
FrmNpcCreator.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub MObjetos_Click()
FrmObjSelector.Visible = True
Me.Visible = False
Inet1.Cancel
End Sub

Private Sub MRutas_Click()
frmConfig.Visible = True
End Sub

Private Sub MSalir_Click()
End
End Sub

Private Sub Op112_Click()
Op121.value = False
IndexMode = "11.X"
Me.Visible = False
FrmCarga.Visible = True
End Sub

Private Sub Op121_Click()
Op112.value = False
IndexMode = "12.1"
Me.Visible = False
FrmCarga.Visible = True
End Sub

