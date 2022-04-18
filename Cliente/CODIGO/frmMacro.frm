VERSION 5.00
Begin VB.Form frmMacro 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar Macro"
   ClientHeight    =   3330
   ClientLeft      =   14430
   ClientTop       =   5925
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Accion3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Equipar Item"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.OptionButton Accion4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usar Item"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.OptionButton Accion2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lanzar Hechizo"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.OptionButton Accion1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escribir"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Salir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Guardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label MacroLbl 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla F11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   168
      Y1              =   32
      Y2              =   32
   End
End
Attribute VB_Name = "FrmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Accion1_Click()
Text1.Enabled = True
End Sub

Private Sub Accion2_Click()
Text1.Enabled = False
End Sub

Private Sub Accion3_Click()
Text1.Enabled = False
End Sub

Private Sub Accion4_Click()
Text1.Enabled = False
End Sub

Private Sub Guardar_Click()
'Usar Item
    If Accion4.Value Then
        If Inventario.ObjIndex(Inventario.SelectedItem) = 0 Or _
           UsarYequiparObjValido(Inventario.ObjType(Inventario.SelectedItem), True) = False Then
            Call MsgBox("Item Invalido,seleccione otro.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aUsar
            MacroList(MacroIndex).Grh = Inventario.GrhIndex(Inventario.SelectedItem)
            MacroList(MacroIndex).Nombre = Inventario.ItemName(Inventario.SelectedItem)
            MacroList(MacroIndex).ObjIndex = Inventario.ObjIndex(Inventario.SelectedItem)
            MacroList(MacroIndex).slot = Inventario.SelectedItem
            Call SaveMacros(UserName)
            'Call frmMain.RenderMacro(frmMain.Macros(MacroIndex), MacroList(MacroIndex).Grh)
            Unload Me
        End If
    End If

    'Equipar Item
    If Accion3.Value Then
        If Inventario.ObjIndex(Inventario.SelectedItem) = 0 Or _
           UsarYequiparObjValido(Inventario.ObjType(Inventario.SelectedItem), False) = False Then
            Call MsgBox("Item Invalido,seleccione otro.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aEquipar
            MacroList(MacroIndex).Grh = Inventario.GrhIndex(Inventario.SelectedItem)
            MacroList(MacroIndex).Nombre = Inventario.ItemName(Inventario.SelectedItem)
            MacroList(MacroIndex).ObjIndex = Inventario.ObjIndex(Inventario.SelectedItem)
            MacroList(MacroIndex).slot = Inventario.SelectedItem
            Call SaveMacros(UserName)
            'Call frmMain.RenderMacro(frmMain.Macros(MacroIndex), MacroList(MacroIndex).Grh)
            Unload Me
        End If
    End If

    'Usar comandos/Hablar
    If Accion1.Value Then
        If Text1.Text = "" Then
            Call MsgBox("Escriba un comando o una palabra.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aComando
            MacroList(MacroIndex).Grh = 20865
            MacroList(MacroIndex).Nombre = Text1.Text
            Call SaveMacros(UserName)
            'Call frmMain.RenderMacro(frmMain.Macros(MacroIndex), MacroList(MacroIndex).Grh)
            Unload Me
        End If
    End If

    'Usar Hechizo
    If Accion2.Value Then
        If frmMain.hlst.List(frmMain.hlst.ListIndex) = "(None)" Or _
           frmMain.hlst.ListIndex = -1 Then
            Call MsgBox("Hechizo invalido,seleccione otro.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aLanzar
            MacroList(MacroIndex).Grh = 20899
            MacroList(MacroIndex).Nombre = frmMain.hlst.List(frmMain.hlst.ListIndex)
            MacroList(MacroIndex).SpellSlot = frmMain.hlst.ListIndex
            Call SaveMacros(UserName)
            'Call frmMain.RenderMacro(frmMain.Macros(MacroIndex), MacroList(MacroIndex).Grh)
            Unload Me
        End If
    End If
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub


