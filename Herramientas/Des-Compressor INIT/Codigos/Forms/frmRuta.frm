VERSION 5.00
Begin VB.Form frmRuta 
   AutoRedraw      =   -1  'True
   BackColor       =   &H001E1E1E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione la ruta de los inits."
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5580
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExaminar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Examinar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox StrRuta 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "D:\Proyectos\Eternal Online\Cliente\Resources\Init"
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton cmdSig 
      BackColor       =   &H008080FF&
      Caption         =   "Siguente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H001E1E1E&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label lblRuta 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la ruta de la carpeta Init:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExaminar_Click()
StrRuta.Text = BrowseForFolder(Me.hWnd, "Carpeta Inits")
End Sub

Private Sub cmdSig_Click()
    If StrRuta.Text = "" Then
        MsgBox "Seleccione una ruta especifica.", vbInformation, "Error"
        Exit Sub
    End If
    
    RUTA_INIT = StrRuta.Text
    frmMain.Show
    Unload Me
End Sub
