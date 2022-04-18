VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H001E1E1E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Des-Compresor INIT - Eternal Online"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCompress 
      BackColor       =   &H008080FF&
      Caption         =   "Comprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.OptionButton SelectCFG 
      BackColor       =   &H001E1E1E&
      Caption         =   "CFG CLIENT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton SelectShields 
      BackColor       =   &H001E1E1E&
      Caption         =   "Escudos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton SelectFXs 
      BackColor       =   &H001E1E1E&
      Caption         =   "FXs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton SelectParticles 
      BackColor       =   &H001E1E1E&
      Caption         =   "Particulas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton SelectHeads 
      BackColor       =   &H001E1E1E&
      Caption         =   "Cabezas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton SelectBodys 
      BackColor       =   &H001E1E1E&
      Caption         =   "Cuerpos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdDescompress 
      BackColor       =   &H008080FF&
      Caption         =   "Descomprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.OptionButton SelectHelmets 
      BackColor       =   &H001E1E1E&
      Caption         =   "Cascos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton SelectWeapons 
      BackColor       =   &H001E1E1E&
      Caption         =   "Armas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton SelectGRH 
      BackColor       =   &H001E1E1E&
      Caption         =   "Graficos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
' frmMain - Formulario Principal
' Aca mostramos los tipos de elementos que se pueden
' comprimir y descomprimir
' Autor: ZenitraM
' Last modification: 09/07/2020
' Comments: Este programa cumple con la licencia GPL.
'================================================================

Option Explicit
Private Sub Form_Load()
    '// esto por que la mierda del formulario me selecciona el primero automaticamente :c
    PROCESS_SELECTED = 1
End Sub
Private Sub cmdCompress_Click()
    Call Comprimir(PROCESS_SELECTED)
End Sub
Private Sub cmdDescompress_Click()
    Call Descomprimir(PROCESS_SELECTED)
End Sub
Private Sub SelectGRH_Click()
    If SelectGRH.value Then PROCESS_SELECTED = 1
End Sub
Private Sub SelectBodys_Click()
    If SelectBodys.value Then PROCESS_SELECTED = 2
End Sub
Private Sub SelectWeapons_Click()
    If SelectWeapons.value Then PROCESS_SELECTED = 3
End Sub
Private Sub SelectHelmets_Click()
    If SelectHelmets.value Then PROCESS_SELECTED = 4
End Sub
Private Sub SelectHeads_Click()
    If SelectHeads.value Then PROCESS_SELECTED = 5
End Sub
Private Sub SelectParticles_Click()
    If SelectParticles.value Then PROCESS_SELECTED = 6
End Sub
Private Sub SelectFXs_Click()
    If SelectFXs.value Then PROCESS_SELECTED = 7
End Sub
Private Sub SelectShields_Click()
    If SelectShields.value Then PROCESS_SELECTED = 8
End Sub
Private Sub SelectCFG_Click()
    If SelectCFG.value Then PROCESS_SELECTED = 9
End Sub
