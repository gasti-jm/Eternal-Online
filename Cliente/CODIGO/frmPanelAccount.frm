VERSION 5.00
Begin VB.Form frmPanelAccount 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Panel de Cuenta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmPanelAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   3360
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   2445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   10
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   3945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   5445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   6945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   8445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   2445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   3945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   5445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   6945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   8445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   5730
      Width           =   1140
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   8340
      TabIndex        =   23
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   6840
      TabIndex        =   22
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   5340
      TabIndex        =   21
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   3840
      TabIndex        =   20
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   2340
      TabIndex        =   19
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   8340
      TabIndex        =   18
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   6840
      TabIndex        =   17
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   5340
      TabIndex        =   16
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   3840
      TabIndex        =   15
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2340
      TabIndex        =   14
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   6210
      TabIndex        =   13
      Top             =   7620
      Width           =   1605
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6210
      TabIndex        =   12
      Top             =   7770
      Width           =   675
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6210
      TabIndex        =   11
      Top             =   7920
      Width           =   390
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   0
      Left            =   2280
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   1
      Left            =   3780
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   2
      Left            =   5280
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   3
      Left            =   6780
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   4
      Left            =   8280
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   5
      Left            =   2280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   6
      Left            =   3780
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   7
      Left            =   5280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   8
      Left            =   6780
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   9
      Left            =   8280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Label lblAccData 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   2370
      Width           =   3705
   End
   Begin VB.Image ImgCrear 
      Height          =   495
      Left            =   2160
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Image ImgBorrar 
      Height          =   495
      Left            =   4200
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Image ImgConectar 
      Height          =   495
      Left            =   8040
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Image ImgSalir 
      Height          =   735
      Left            =   8040
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmPanelAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Seleccionado As Byte

Private Sub cmdCerrar_Click()
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub

Private Sub cmdConnt_Click()
UserName = lblAccData(1 + Seleccionado).Caption
Call WriteLoginExistingChar
End Sub

Private Sub cmdCrear_Click()

For i = 0 To 7
  If lblAccData(i + 1).Caption = "" Then
     frmCrearPersonaje.Show
     Exit Sub
  End If
Next i

End Sub


Private Sub Form_Load()

On Error Resume Next
    Unload frmConnect

    Me.Icon = frmMain.Icon
    
    Dim CharIndex As Integer
    
    Dim i As Byte
    
    For i = 1 To 10
    
        lblAccData(i).Caption = ""
        
    Next i

Me.Picture = LoadPicture(App.Path & "\Recursos\Interfaces\Cuenta.jpg")

If Curper = True Then
   Call FormParser.Parse_Form(Me)
End If

End Sub

Private Sub Image1_Click()
Dim i As Byte
    For i = 0 To 7
        If lblAccData(i + 1).Caption = "" Then
            frmCrearPersonaje.Show
            Exit Sub
        End If
    Next i
End Sub

Private Sub Image2_Click()
    MsgBox "No habilitado"
End Sub

Private Sub Image3_Click()
    frmMain.Socket1.Disconnect
    Unload Me
    frmConnect.Show
End Sub

Private Sub Image4_Click()
MsgBox "No habilitado"
End Sub

Private Sub Image5_Click()
    If Not lblAccData(Index + 1).Caption = "" Then
        UserName = lblAccData(Index + 1).Caption
        WriteLoginExistingChar
    End If
End Sub

Private Sub lblName_Click(Index As Integer)
    Seleccionado = Index
End Sub

Private Sub imgAcc_Click(Index As Integer)
 Dim i As Byte ' // PA LA RESET MY FRIEND
    '// RESET PAPI :D
    For i = 0 To 9
        imgAcc(i).Picture = Nothing
    Next i

    imgAcc(Index).Picture = LoadPicture(App.Path & "\Recursos\Interfaces\slot" & Index + 1 & ".jpg")
End Sub

Private Sub ImgBorrar_Click()
Call Audio.PlayWave(SND_CLICK)
Call MsgBox("No habilitado.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1)
End Sub

Private Sub ImgConectar_Click()
Call Audio.PlayWave(SND_CLICK)
UserName = lblAccData(Seleccionado).Caption
Call WriteLoginExistingChar
End Sub

Private Sub ImgCrear_Click()
Call Audio.PlayWave(SND_CLICK)

For i = 0 To 7
  If lblAccData(i + 1).Caption = "" Then
     frmCrearPersonaje.Show
     Exit Sub
  End If
Next i
End Sub

Private Sub ImgSalir_Click()
Call Audio.PlayWave(SND_CLICK)
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub


Private Sub picChar_Click(Index As Integer)
    Dim i As Byte ' // PA LA RESET MY FRIEND

    Seleccionado = Index
    If cPJ(Seleccionado).Nombre <> "" Then
        lblCharData(0) = "Nivel: " & cPJ(Seleccionado).Nivel
        lblCharData(1) = cPJ(Seleccionado).Mapa
        lblCharData(2) = "Clase: " & ListaClases(cPJ(Seleccionado).Clase)
    Else
        lblCharData(0) = ""
        lblCharData(1) = ""
        lblCharData(2) = ""
    End If

    '// RESET PAPI :D
    For i = 0 To 9
        imgAcc(i).Picture = Nothing
    Next i

    imgAcc(Index).Picture = LoadPicture(App.Path & "\Recursos\Interfaces\slot" & Index + 1 & ".jpg")
End Sub

Private Sub picChar_DblClick(Index As Integer)
    Seleccionado = Index
    If Not lblAccData(Index + 1).Caption = "" Then
        UserName = lblAccData(1 + Index).Caption
        If Not frmMain.Socket1.Connected Then
            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
            EstadoLogin = Normal
            'Call Login
        Else
            WriteLoginExistingChar
        End If
    Else
        frmCrearPersonaje.Show
    End If
End Sub

Private Sub Timer1_Timer()
Dim i As Byte
For i = 1 To 10
        If Not frmPanelAccount.lblAccData(i).Caption = "" Then
            Call engine.DrawPJSAccount(i)
        End If
    Next i
End Sub
