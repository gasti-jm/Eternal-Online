VERSION 5.00
Begin VB.Form FrmDatMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu de Dateo"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5220
   Icon            =   "FrmDatMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmDatMenu.frx":08CA
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label CmdCerrar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4800
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label LblDH 
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
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label LblDOBJ 
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label LblDNPC 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
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
         Begin VB.Menu MDatosIndex 
            Caption         =   "Ver Datos de Indexacion"
         End
         Begin VB.Menu MImagenes 
            Caption         =   "Ver Imagenes"
            Visible         =   0   'False
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
Attribute VB_Name = "FrmDatMenu"
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

Private Sub LblDH_Click()
FrmHechizosCreator.Visible = True
Me.Visible = False
End Sub

Private Sub LblDNPC_Click()
FrmNpcCreator.Visible = True
Me.Visible = False
End Sub

Private Sub LblDOBJ_Click()
FrmObjSelector.Visible = True
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

Private Sub MDatosIndex_Click()
FrmAnimaciones.Visible = True
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

Private Sub MRutas_Click()
frmConfig.Visible = True
End Sub

Private Sub MSalir_Click()
End
End Sub
