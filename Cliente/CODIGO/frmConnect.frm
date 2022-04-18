VERSION 5.00
Begin VB.Form frmConnect 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Eternal Online"
   ClientHeight    =   16200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   28800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1080
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1920
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picRender 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   16200
      Left            =   0
      ScaleHeight     =   1080
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1920
      TabIndex        =   4
      Top             =   0
      Width           =   28800
      Begin VB.ComboBox lst_servers 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Timer EfectoIntro 
         Interval        =   1
         Left            =   5280
         Top             =   960
      End
      Begin VB.Timer TimerIntro 
         Interval        =   5800
         Left            =   4680
         Top             =   960
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.TextBox txtPasswd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         IMEMode         =   3  'DISABLE
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3120
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Image imgSalir 
         Height          =   405
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Image imgConectarse 
         Height          =   405
         Left            =   0
         Top             =   960
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Image imgCrearPj 
         Height          =   405
         Left            =   0
         Top             =   480
         Visible         =   0   'False
         Width           =   2460
      End
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4290
      TabIndex        =   0
      Text            =   "7666"
      Top             =   240
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5160
      TabIndex        =   3
      Text            =   "localhost"
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private cBotonCrearPj As clsGraphicalButton
Private cBotonRecuperarPass As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonReglamento As clsGraphicalButton
Private cBotonCodigoFuente As clsGraphicalButton
Private cBotonBorrarPj As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton
Private cBotonLeerMas As clsGraphicalButton
Private cBotonForo As clsGraphicalButton
Private cBotonConectarse As clsGraphicalButton
Private cBotonTeclas As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub EfectoIntro_Timer()

If TypeIntroEffect Then
    If EffectIntro > 0 Then
        EffectIntro = EffectIntro - 1
    Else
        EffectIntro = 0
    End If
Else
    If EffectIntro < 255 Then
        EffectIntro = EffectIntro + 1
    Else
        EffectIntro = 255
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me, E_NORMAL)

    If Not INTRO Then Call LoadOBJFormConnect

    txtNombre.Text = Config_Inicio.AccountName

    'Aplicate in forms
    frmConnect.Width = frmScaleWidth
    frmConnect.Height = frmScaleHeight

    'Aplicate in Renders
    frmConnect.picRender.Width = Config_Inicio.ResolutionX
    frmConnect.picRender.Height = Config_Inicio.ResolutionY
    
    txtNombre.Left = Config_Inicio.ResolutionX / 2 - 61
    txtNombre.Top = Config_Inicio.ResolutionY / 2 - 111
    txtPasswd.Left = Config_Inicio.ResolutionX / 2 - 61
    txtPasswd.Top = Config_Inicio.ResolutionY / 2 - 72
    
    imgConectarse.Left = Config_Inicio.ResolutionX / 2 - 65
    imgConectarse.Top = Config_Inicio.ResolutionY / 2 + 48
    imgCrearPj.Left = Config_Inicio.ResolutionX / 2 - 65
    imgCrearPj.Top = Config_Inicio.ResolutionY / 2 + 82
    
    imgSalir.Left = Config_Inicio.ResolutionX / 2 - 65
    imgSalir.Top = Config_Inicio.ResolutionY / 2 + 116
    
    lst_servers.Left = Config_Inicio.ResolutionX / 2 - 72
    lst_servers.Top = Config_Inicio.ResolutionY / 2 - 18
    
    
    Call LoadListServers
    'Call LoadButtons
        
End Sub
Private Sub LoadListServers()

    lServer(1).port = 7666
    lServer(1).Ip = "127.0.0.1"
    lServer(1).name = "RPG - [LAS] (0/500)"
    
    lServer(2).port = 7666
    lServer(2).Ip = "127.0.0.1"
    lServer(2).name = "PvP - [LAS] (0/200)"
    
    lst_servers.AddItem lServer(1).name
    lst_servers.AddItem lServer(2).name
    lst_servers.ListIndex = 0
End Sub
Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = DirInterface
    
    Set cBotonCrearPj = New clsGraphicalButton
    Set cBotonRecuperarPass = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonReglamento = New clsGraphicalButton
    Set cBotonCodigoFuente = New clsGraphicalButton
    Set cBotonBorrarPj = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonLeerMas = New clsGraphicalButton
    Set cBotonForo = New clsGraphicalButton
    Set cBotonConectarse = New clsGraphicalButton
    Set cBotonTeclas = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgConectarse_Click()
Call FormParser.Parse_Form(Me, E_WAIT)
Call Audio.PlayWave(SND_CLICK)
    
    If frmMain.Client.State <> (sckClosed Or sckConnecting) Then
        frmMain.Client.CloseSck
        DoEvents
    End If
    
    'update user info
    UserName = txtNombre.Text
    
    Dim aux As String
    aux = txtPasswd.Text
    
    UserPassword = aux

    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        frmMain.Client.Connect CurServerIp, CurServerPort
    End If
    
End Sub

Private Sub imgCrearPj_Click()
Call Audio.PlayWave(SND_CLICK)
    
    EstadoLogin = E_MODO.Dados

    If frmMain.Client.State <> (sckClosed Or sckConnecting) Then
        frmMain.Client.CloseSck
        DoEvents
    End If
    frmMain.Client.Connect CurServerIp, CurServerPort
    
End Sub

Private Sub imgSalir_Click()
If MsgBox("�Estas seguro que desea salir?", vbYesNo + vbInformation, "Salir") = vbYes Then prgRun = False
End Sub

Private Sub lst_servers_Click()
    CurServerIp = lServer(lst_servers.ListIndex + 1).Ip
    CurServerPort = lServer(lst_servers.ListIndex + 1).port
End Sub

Private Sub picRender_Click()
    If INTRO Then
        INTRO = False
        TimerIntro.Enabled = False
        EfectoIntro.Enabled = False
        Call LoadOBJFormConnect
    End If
End Sub

Private Sub TimerIntro_Timer()

If BastaINTRO = 3 Then ' si esto llega a 3 finaliza la intro
    INTRO = False
    TimerIntro.Enabled = False
    EfectoIntro.Enabled = False
    Call LoadOBJFormConnect
End If

If TypeIntroEffect Then
    TypeIntroEffect = 0
    BastaINTRO = BastaINTRO + 1
Else
    TypeIntroEffect = 1
    BastaINTRO = BastaINTRO + 1
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not INTRO Then
            imgConectarse_Click
        Else
            INTRO = False
            TimerIntro.Enabled = False
            EfectoIntro.Enabled = False
            Call LoadOBJFormConnect
        End If
    End If
End Sub
