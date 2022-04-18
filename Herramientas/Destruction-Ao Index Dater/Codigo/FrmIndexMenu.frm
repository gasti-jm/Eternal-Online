VERSION 5.00
Begin VB.Form FrmIndexMenu 
   Caption         =   "Menu de Indexacion"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5220
   Icon            =   "FrmIndexMenu.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmIndexMenu.frx":08CA
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label CmdCerrar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label LblVolver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label LblVGeneral 
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
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
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
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Menu MRapido 
      Caption         =   "Menu Rapido"
      Begin VB.Menu MInicio 
         Caption         =   "Inicio"
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
Attribute VB_Name = "FrmIndexMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCerrar_Click()
FrmInicio.Inet1.Cancel
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmInicio.Visible = True
End Sub

Private Sub LblIndexar_Click()
FrmIndex.Visible = True
Me.Visible = False
End Sub

Private Sub LblVGeneral_Click()
FrmAnimaciones.Visible = True
Me.Visible = False
End Sub

Private Sub LblVolver_Click()
FrmInicio.Visible = True
Me.Visible = False
End Sub

Private Sub MConversor_Click()
FrmConversor.Visible = True
End Sub

Private Sub MCreditos_Click()
FrmCreditos.Visible = True
Me.Visible = False
End Sub

Private Sub MHechizos_Click()
FrmHechizosCreator.Visible = True
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

Private Sub MRutas_Click()
frmConfig.Visible = True
End Sub

Private Sub MSalir_Click()
End
End Sub
