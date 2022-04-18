VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Creación de cuenta"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearCuenta.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox nameTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      Left            =   1220
      TabIndex        =   4
      Top             =   2170
      Width           =   3250
   End
   Begin VB.TextBox passTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   1220
      PasswordChar    =   "x"
      TabIndex        =   3
      Top             =   3400
      Width           =   3250
   End
   Begin VB.TextBox pass1Txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   1220
      PasswordChar    =   "x"
      TabIndex        =   2
      Top             =   4670
      Width           =   3250
   End
   Begin VB.TextBox mailTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   1220
      TabIndex        =   1
      Top             =   5910
      Width           =   3250
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Crear"
      Height          =   5535
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Clos 
      Height          =   315
      Left            =   4905
      Top             =   480
      Width           =   345
   End
   Begin VB.Image CrearCuenta 
      Height          =   540
      Left            =   1920
      Top             =   6960
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de cuenta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   10
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electronico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pregunta secreta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label CrearCuentax 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Cuenta."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6660
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5160
      Width           =   2295
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Clos_Click()
Unload Me
End Sub


Private Sub CrearCuenta_Click()

 UserAccount = NameTxt.Text
  
  UserPassword = passTxt.Text
   
   UserEmail = mailTxt.Text
    
    If Not UserPassword = pass1Txt.Text Then
        
        MsgBox "Las contraseñas no coinciden."
        
        Exit Sub
    
    End If
    
    If Not CheckMailString(UserEmail) Then
        
        MsgBox "Direccion de mail invalida."
        
        Exit Sub
    
    End If
    
     UserAnswer = "Tengo que sacar esto por dios!, tengo paja asi que hago esto xD!"
    
     UserQuestion = 2
    
    If Len(UserAnswer) < 11 Then
        
        MsgBox "Respuesta muy corta"
        
        Exit Sub
    
    End If

    EstadoLogin = CrearNuevaCuenta
    
    If frmMain.Socket1.Connected Then
        
        frmMain.Socket1.Disconnect
        
        frmMain.Socket1.Cleanup
        
        DoEvents
    
    End If
    
    frmMain.Socket1.HostName = CurServerIp
    
      frmMain.Socket1.RemotePort = CurServerPort
    
        frmMain.Socket1.Connect
    
    Unload Me

End Sub
Private Sub Form_Load()

If Curper = True Then
  Call FormParser.Parse_Form(Me)
End If

Me.Icon = frmMain.Icon

'Me.Picture = LoadPicture(App.Path & "\Recursos\Interfaces\CrearCuenta\CrearCuenta.bmp")
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Set CrearCuenta.Picture = Nothing
Set Clos.Picture = Nothing
End Sub

Private Sub mailTxt_Click()
If pass1Txt.Text = vbNullString Then
 Else
 If pass1Txt.Text = passTxt.Text Then
   pass1Txt.BackColor = vbGreen
 Else
   pass1Txt.BackColor = vbRed
 End If
End If
End Sub


