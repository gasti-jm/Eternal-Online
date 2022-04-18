VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Eternal Online"
   ClientHeight    =   16200
   ClientLeft      =   390
   ClientTop       =   675
   ClientWidth     =   28800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1080
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1920
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   16200
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   1080
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   28800
      Begin VB.Timer tTrueno 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   1200
         Top             =   0
      End
      Begin VB.Timer macrotrabajo 
         Enabled         =   0   'False
         Left            =   600
         Top             =   0
      End
      Begin VB.Timer Second 
         Enabled         =   0   'False
         Interval        =   1050
         Left            =   120
         Top             =   0
      End
      Begin VB.PictureBox Minimap 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   27360
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   3
         Top             =   0
         Width           =   1500
         Begin VB.Shape UserP 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   45
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   45
         End
      End
      Begin VB.ListBox hlst 
         BackColor       =   &H00000000&
         ForeColor       =   &H000080FF&
         Height          =   2790
         Left            =   2880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   5880
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   1440
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   1
         Top             =   2160
         Width           =   2400
      End
      Begin RichTextLib.RichTextBox RecTxt 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Mensajes del servidor"
         Top             =   0
         Visible         =   0   'False
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   503
         _Version        =   393217
         BackColor       =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         TextRTF         =   $"frmMain.frx":424A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgDropGLD 
         Height          =   300
         Left            =   1320
         Top             =   0
         Width           =   300
      End
      Begin VB.Image imgMeditate 
         Height          =   330
         Left            =   330
         Top             =   1920
         Width           =   330
      End
      Begin VB.Image imgRuna 
         Height          =   330
         Left            =   6960
         Top             =   1920
         Width           =   330
      End
      Begin VB.Image Macros 
         Height          =   480
         Index           =   6
         Left            =   11400
         Top             =   5640
         Width           =   480
      End
      Begin VB.Image Macros 
         Height          =   480
         Index           =   5
         Left            =   10800
         Top             =   5640
         Width           =   480
      End
      Begin VB.Image Macros 
         Height          =   480
         Index           =   4
         Left            =   10200
         Top             =   5640
         Width           =   480
      End
      Begin VB.Image Macros 
         Height          =   480
         Index           =   3
         Left            =   9600
         Top             =   5640
         Width           =   480
      End
      Begin VB.Image Macros 
         Height          =   480
         Index           =   2
         Left            =   9000
         Top             =   5640
         Width           =   480
      End
      Begin VB.Image Macros 
         Height          =   480
         Index           =   1
         Left            =   8400
         Top             =   5640
         Width           =   480
      End
   End
   Begin VB.Image CmdLanzar 
      Height          =   375
      Left            =   9360
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgClanes 
      Height          =   390
      Left            =   10245
      Top             =   7980
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image imgEstadisticas 
      Height          =   360
      Left            =   10215
      Top             =   7605
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image imgGrupo 
      Height          =   315
      Left            =   10215
      Top             =   6990
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   450
      Left            =   10680
      MousePointer    =   99  'Custom
      Top             =   8520
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   10680
      MouseIcon       =   "frmMain.frx":42C7
      MousePointer    =   99  'Custom
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
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

Option Explicit
Public WithEvents Client As clsSocket
Attribute Client.VB_VarHelpID = -1

Public TX As Byte
Public TY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Private cBotonDiamArriba As clsGraphicalButton
Private cBotonDiamAbajo As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonGrupo As clsGraphicalButton
Private cBotonOpciones As clsGraphicalButton
Private cBotonEstadisticas As clsGraphicalButton
Private cBotonClanes As clsGraphicalButton
Private cBotonAsignarSkill As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public picSkillStar As Picture


Private Sub Form_Load()

Call FormParser.Parse_Form(Me, E_NORMAL)

    Dim i As Byte

    If Config_Inicio.ResolutionX = 800 And Config_Inicio.ResolutionY = 600 Then
        PosHUDX = 298
        PosHUDY = Config_Inicio.ResolutionY - 32
    Else
        PosHUDX = Config_Inicio.ResolutionX / 2
        PosHUDY = Config_Inicio.ResolutionY - 32
    End If

    'Aplicate in forms
    frmMain.Width = frmScaleWidth
    frmMain.Height = frmScaleHeight
    
    'Aplicate in Renders
    frmMain.MainViewPic.Width = Config_Inicio.ResolutionX
    frmMain.MainViewPic.Height = Config_Inicio.ResolutionY
    
    frmMain.Macros(1).Left = PosHUDX - 152
    frmMain.Macros(2).Left = frmMain.Macros(1).Left + 32 + 5
    frmMain.Macros(3).Left = frmMain.Macros(2).Left + 32 + 5
    frmMain.Macros(4).Left = frmMain.Macros(3).Left + 32 + 5
    frmMain.Macros(5).Left = frmMain.Macros(4).Left + 32 + 5
    frmMain.Macros(6).Left = frmMain.Macros(5).Left + 32 + 5
    
    frmMain.imgRuna.Top = PosHUDY - 58
    frmMain.imgRuna.Left = PosHUDX + 174
    frmMain.imgMeditate.Top = PosHUDY - 58
    frmMain.imgMeditate.Left = PosHUDX + 151
    
    

    For i = 1 To 6
        frmMain.Macros(i).Top = PosHUDY - 25
    Next i
    
    PicInv.Left = Config_Inicio.ResolutionX - 160
    PicInv.Top = Config_Inicio.ResolutionY - 160
    
    Minimap.Left = Config_Inicio.ResolutionX - 100
    Minimap.Top = 0
    
    imgDropGLD.Left = 86
    imgDropGLD.Top = 74

    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If
    
    Me.Left = 0
    Me.Top = 0

End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)
    If bTurnOn Then
        imgAsignarSkill.Picture = picSkillStar
    Else
        Set imgAsignarSkill.Picture = Nothing
    End If
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case Index
            Case 1 'subir
                If hlst.ListIndex = 0 Then Exit Sub
            Case 0 'bajar
                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
        
        Select Case Index
            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1
            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Not GetAsyncKeyState(KeyCode) < 0 Then
    Es_Real = False
    Exit Sub
End If
Es_Real = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If UserWritting Then
        Chat_KeyPress (KeyAscii)
        Chat_Change (KeyAscii)
    End If
    KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
'***************************************************

If Es_Real = False Then
    Exit Sub
    End If
Es_Real = True
    
    If Not UserWritting Then
            Select Case KeyCode
                Case vbKeyP
                    'General_Particle_Create 34, 50, 50, -1
                    General_Particle_Create_Render 31, 0, -1 '18
                    'General_Char_Particle_Create 34, 1, -1
                    
                Case vbKeyMultiply
                    View_FPS_OR_MS = Not View_FPS_OR_MS
            
                Case vbKeyF12
                    Call ScreenCapture
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                    
                Case vbKeyF1 To vbKeyF6
                    Call UsarMacro(KeyCode - 111)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            If frmOpciones.Visible = False Then
                Call frmOpciones.Show(vbModeless, frmMain)
            Else
                frmOpciones.Visible = False
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            Call WriteMeditate
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            
            If UserWritting = True Then 'si estaba escribiendo y apretamos el keytalk manda el buffer
                Chat_KeyUP KeyCode
            Else
                If (Not Comerciando) And (Not MirandoAsignarSkills) And _
                (Not frmMSG.Visible) And (Not MirandoForo) And _
                (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                    UserWritting = Not UserWritting
                    'SendTxt.Visible = True
                    'SendTxt.SetFocus
                End If
            End If
            
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Image1_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
End Sub

Private Sub imgAsignarSkill_Click()
    Dim i As Integer
    
    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer
    
    Do While Not LlegaronSkills
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    LlegaronSkills = False
    
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

End Sub

Private Sub imgClanes_Click()
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub imgDropGLD_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
End Sub

Private Sub imgEstadisticas_Click()
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer
    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
End Sub

Private Sub imgGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub lblScroll_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub lblCerrar_Click()
    prgRun = False
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = 1
End Sub
Private Sub imgMeditate_Click()
    If Not MainTimer.Check(TimersIndex.HUD_CLICK) Then Exit Sub
        If UserMinMAN = UserMaxMAN Then Exit Sub
            If UserEstado = 1 Then 'Muerto
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
        Call WriteMeditate
End Sub

Private Sub imgRuna_Click()
'Pongamos freno, no queremos saturar el sv.
    If Not MainTimer.Check(TimersIndex.HUD_CLICK) Then Exit Sub
        Call WriteRegresarHogar
End Sub

Private Sub Macros_Click(Index As Integer)

    If Not FrmMacro.Visible = True Then

    If MacroList(Index).mTipe = 0 Then
        MacroIndex = Index
        FrmMacro.MacroLbl = "Tecla F" & Index
        FrmMacro.Show , Me
    Else
        'Accion!
        Call UsarMacro(CByte(Index))
    End If
    End If
   
End Sub
Private Sub Macros_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not UserEstado = 1 Then
        If FrmMacro.Visible = True Then Exit Sub
        If Button = vbKeyRButton Then
            MacroIndex = Index
            FrmMacro.MacroLbl = "Tecla F" & Index
            FrmMacro.Show , Me
        End If
    Else
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    End If
End Sub

Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not Application.IsAppActive() Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(TX, TY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not (frmCarp.Visible = True) Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
End Sub


Private Sub MainViewPic_Click()
    Form_Click
End Sub

Private Sub MainViewPic_DblClick()
    Form_DblClick
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    'LastPressed.ToggleToNormal
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(TX, TY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(TX, TY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)
End Sub
Private Sub Minimap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call WriteWarpChar("YO", UserMap, IIf(X < 1, 1, X), Y)
Call SmallMap_UserPOS
End Sub
Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseBoton = Button
End Sub

Private Sub picSM_DblClick(Index As Integer)
Select Case Index
    Case eSMType.sResucitation
        Call WriteResuscitationToggle
        
    Case eSMType.sSafemode
        Call WriteSafeToggle
        
    Case eSMType.mSpells
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
    Case eSMType.mWork
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If macrotrabajo.Enabled Then
            Call DesactivarMacroTrabajo
        Else
            Call ActivarMacroTrabajo
        End If
End Select
End Sub


Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else
                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()
    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub

Private Sub hlst_DblClick()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
        End If
    End If
End Sub
Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub
Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub Form_Click()
    'SendTxt.Visible = False
    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, TX, TY)
         
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    Call WriteLeftClick(TX, TY)
                Else
                
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        If Config_Inicio.CursorGraphic Then
                            Call FormParser.Parse_Form(frmMain)
                        Else
                            frmMain.MousePointer = vbDefault
                        End If
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            If Config_Inicio.CursorGraphic Then
                                Call FormParser.Parse_Form(frmMain)
                            Else
                                frmMain.MousePointer = vbDefault
                            End If
                            
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0
                                'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                '    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
                               ' End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0
                                'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    'Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                'End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            If Config_Inicio.CursorGraphic Then
                                Call FormParser.Parse_Form(frmMain)
                            Else
                                frmMain.MousePointer = vbDefault
                            End If
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    If Config_Inicio.CursorGraphic Then
                        Call FormParser.Parse_Form(frmMain)
                    Else
                        frmMain.MousePointer = vbDefault
                    End If
                    
                    Call WriteWorkLeftClick(TX, TY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, TX, TY)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(TX, TY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewPic.Left
    MouseY = Y - MainViewPic.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewPic.Width Then
        MouseX = MainViewPic.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewPic.Height Then
        MouseY = MainViewPic.Height
    End If

    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub lblDropGold_Click()


    
End Sub

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    ' Activo controles de inventario
    PicInv.Visible = True

    ' Desactivo controles de hechizo
    hlst.Visible = False
    cmdINFO.Visible = False
    CmdLanzar.Visible = False
     
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)


    ' Activo controles de hechizos
    hlst.Visible = True
    cmdINFO.Visible = True
    CmdLanzar.Visible = True
    
    
    ' Desactivo controles de inventario
    PicInv.Visible = False

End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
    
    Call UsarItem
    Call EquiparItem
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not Application.IsAppActive() Then Exit Sub
    
    If (Not Comerciando) And (Not MirandoAsignarSkills) And _
        (Not frmMSG.Visible) And (Not MirandoForo) And _
        (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
         
        If PicInv.Visible Then
            PicInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub AbrirMenuViewPort()
Form_DblClick
#If (ConMenuseConextuales = 1) Then

    If MapData(TX, TY).CharIndex > 0 Then
        If charlist(MapData(TX, TY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim M As New frmMenuseFashion
            
            Load M
            M.SetCallback Me
            M.SetMenuId 1
            M.ListaInit 2, False
            
            If charlist(MapData(TX, TY).CharIndex).Nombre <> "" Then
                M.ListaSetItem 0, charlist(MapData(TX, TY).CharIndex).Nombre, True
            Else
                M.ListaSetItem 0, "<NPC>", True
            End If
            M.ListaSetItem 1, "Comerciar"
            
            M.ListaFin
            M.Show , Me

        End If
    End If

#End If
End Sub

''''''''''''''''''''''''''''''''''''''
'     API                            '
''''''''''''''''''''''''''''''''''''''
Private Sub Client_Connect()
 
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
    Second.Enabled = True
    
    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
            Call Login
 
 
        Case E_MODO.Normal
            Call Login
 
        Case E_MODO.Dados
            Call Audio.PlayMIDI("7.mid")
            frmCrearPersonaje.Show vbModal
    End Select
 
End Sub
Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim data() As Byte
    
    Client.GetData RD, vbByte, bytesTotal
        
    data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    HandleIncomingData
    
End Sub
Private Sub Client_CloseSck()
    Dim i As Long
        
    Second.Enabled = False
    Connected = False
    
    If Client.State <> sckClosed Then _
        Client.CloseSck
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name Then
            Unload Forms(i)
        End If
    Loop
    
    'On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        Call Load_Map(2)
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False
 
    pausa = False
    UserMeditar = False
 
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    'UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i
 
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
 
    SkillPoints = 0
    Alocados = 0
 
   ' Dialogos.CantidadDialogos = 0
    
End Sub
Private Sub Client_Error(ByVal number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False
 
    If Client.State <> sckClosed Then _
        Client.CloseSck
 
    frmMain.Visible = False
    frmCrearPersonaje.Visible = False
 
    'If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    'Else
        'frmCrearPersonaje.MousePointer = 0
    'End If
 
End Sub

Private Sub tTrueno_Timer()
    If bRain Then
        TypeTrueno = RandomNumber(0, 1) 'hay 2 efectos de trueno
        HayTrueno = 15
        
        If TypeTrueno Then
            Call Audio.PlayWave(SND_TRUENO)
        Else
            Call Audio.PlayWave(SND_TRUENO2)
        End If
    End If
End Sub

