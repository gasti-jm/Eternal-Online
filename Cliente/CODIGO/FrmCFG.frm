VERSION 5.00
Begin VB.Form FrmCFG 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion DDEX"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3615
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00A6A6A6&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmCFG.frx":0000
   ScaleHeight     =   3855
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_memoria 
      BackColor       =   &H00000000&
      ForeColor       =   &H00A6A6A6&
      Height          =   315
      ItemData        =   "FrmCFG.frx":86CA
      Left            =   1440
      List            =   "FrmCFG.frx":86D7
      TabIndex        =   8
      Text            =   "Defecto"
      Top             =   2370
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2130
      MaskColor       =   &H00009CFF&
      TabIndex        =   7
      Top             =   3420
      UseMaskColor    =   -1  'True
      Width           =   1425
   End
   Begin VB.CheckBox chk_vs 
      BackColor       =   &H00000000&
      Caption         =   "sincronizado vertical(VSync)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009CFF&
      Height          =   375
      Left            =   150
      TabIndex        =   6
      Top             =   3000
      Width           =   3525
   End
   Begin VB.ComboBox cmb_vertices 
      BackColor       =   &H00000000&
      ForeColor       =   &H00A6A6A6&
      Height          =   315
      ItemData        =   "FrmCFG.frx":86FB
      Left            =   1440
      List            =   "FrmCFG.frx":8705
      TabIndex        =   4
      Text            =   "HardWare"
      Top             =   1980
      Width           =   1935
   End
   Begin VB.ComboBox cmb_modo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00A6A6A6&
      Height          =   315
      ItemData        =   "FrmCFG.frx":871D
      Left            =   1440
      List            =   "FrmCFG.frx":872A
      TabIndex        =   2
      Text            =   "HardWare"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cmb_api 
      BackColor       =   &H00000000&
      ForeColor       =   &H00A6A6A6&
      Height          =   315
      ItemData        =   "FrmCFG.frx":874E
      Left            =   1440
      List            =   "FrmCFG.frx":875E
      TabIndex        =   1
      Text            =   "DirectX 9"
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Memoria:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009CFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Vertices:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009CFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dispositivo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009CFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1590
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Api grafica:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009CFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1170
      Width           =   1365
   End
End
Attribute VB_Name = "FrmCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cfg As DDEXCFG
Friend Function Configuracion() As DDEXCFG

    Me.Show vbModal
    
    CargarCfg
    Configuracion = cfg
End Function
Private Sub Actualizar()
    If cmb_api.ListIndex = -1 Then Exit Sub
    Select Case cmb_api.ListIndex
        Case 0 Or 1
        
            cmb_modo.Clear
            cmb_modo.AddItem "HardWare"
            cmb_modo.AddItem "Referencia"
            cmb_modo.AddItem "SoftWare"
            
            cmb_vertices.Clear
            cmb_vertices.AddItem "HardWare"
            cmb_vertices.AddItem "SoftWare"
            
            cmb_memoria.Clear
            cmb_memoria.AddItem "Defecto"
            cmb_memoria.AddItem "Administrada"
            cmb_memoria.AddItem "Sistema"
        Case 2
            cmb_modo.Clear
            cmb_modo.AddItem "No soportado"
            cmb_vertices.Clear
            cmb_vertices.AddItem "No soportado"
            
            cmb_memoria.Clear
            cmb_memoria.AddItem "No soportado"
        Case 3
            cmb_modo.Clear
            cmb_modo.AddItem "HardWare"
            cmb_modo.AddItem "Referencia"
            cmb_modo.AddItem "SoftWare"
            
            cmb_vertices.Clear
            cmb_vertices.AddItem "HardWare"
            cmb_vertices.AddItem "SoftWare"
            
            
            cmb_memoria.Clear
            cmb_memoria.AddItem "No soportado"
    End Select
    cmb_vertices.ListIndex = 0
    cmb_memoria.ListIndex = 0
    cmb_modo.ListIndex = 0
End Sub

Private Sub cmb_api_Click()
    Actualizar
End Sub

Private Sub cmb_api_Change()
     Actualizar
End Sub

Private Sub cmb_api_DblClick()
    Actualizar
End Sub

Private Sub cmb_api_KeyDown(KeyCode As Integer, Shift As Integer)
     Actualizar
End Sub

Private Sub cmb_api_Scroll()
    Actualizar
End Sub

Private Sub cmb_api_Validate(Cancel As Boolean)
    Actualizar
End Sub
Private Sub CargarCfg()
    
    cfg.api = IIf(cmb_api.ListIndex = -1, 0, cmb_api.ListIndex)
    cfg.MODO = IIf(cmb_modo.ListIndex = -1, 0, cmb_modo.ListIndex)
    cfg.MODO2 = IIf(cmb_vertices.ListIndex = -1, 0, cmb_vertices.ListIndex)
    cfg.memoria = IIf(cmb_memoria.ListIndex = -1, 0, cmb_memoria.ListIndex)
    cfg.vsync = IIf(chk_vs.value = vbChecked, 1, 0)
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmb_vertices.ListIndex = 0
    cmb_memoria.ListIndex = 0
    cmb_modo.ListIndex = 0
End Sub
