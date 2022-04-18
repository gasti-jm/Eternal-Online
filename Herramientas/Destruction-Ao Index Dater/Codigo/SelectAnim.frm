VERSION 5.00
Begin VB.Form SelectAnim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Animacion"
   ClientHeight    =   2985
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   5370
   Icon            =   "SelectAnim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SelectAnim.frx":08CA
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox SearchText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer TAnimacion 
      Enabled         =   0   'False
      Left            =   4920
      Top             =   0
   End
   Begin VB.PictureBox PbShow 
      BackColor       =   &H80000007&
      Height          =   3015
      Left            =   2400
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.ListBox LstAnim 
      Height          =   2595
      ItemData        =   "SelectAnim.frx":429BE
      Left            =   0
      List            =   "SelectAnim.frx":429C0
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "SelectAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LB_FINDSTRING = &H18F
Private Declare Function sendmessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Dim FrameActual As Integer
Dim CantFrames As Byte
Dim AnimDir As String
Dim GrhIndex As Long


Private Sub LstAnim_Click()
If Val(LstAnim.Text) = 0 Then
    FrameActual = 0
    Exit Sub
Else
    GrhIndex = Val(LstAnim.Text)
End If
CantFrames = GrhData(GrhIndex).NumFrames
If IndexMode = "12.1" Then
    TAnimacion.Interval = Round(GrhData(GrhIndex).Speed / GrhData(GrhIndex).NumFrames)
Else
    TAnimacion.Interval = 100
End If
FrameActual = 0
TAnimacion.Enabled = True
PbShow.Cls
End Sub

Private Sub LstAnim_DblClick()
Call DirGrhEdit.ReplaceDir(AnimDir, Val(LstAnim.Text))
DirGrhEdit.Visible = True
Me.Visible = False
TAnimacion.Enabled = False
FrameActual = 0
End Sub

Private Sub SearchText_Change()
LstAnim.ListIndex = sendmessage(LstAnim.hWnd, LB_FINDSTRING, -1, ByVal SearchText.Text)
End Sub

Private Sub TAnimacion_Timer()
On Error GoTo ErrHandler
Dim AnimacionPosX As Integer
Dim AnimacionPosY As Integer
Dim AnimacionLongX As Integer
Dim AnimacionLongY As Integer
Dim GraficoPath As String
Dim GrhIndexAnim As Long

FrameActual = FrameActual + 1

GrhIndexAnim = GrhData(GrhIndex).Frames(FrameActual)

If GrhIndexAnim > 0 Then
    GraficoPath = Config.BmpPath & "\" & GrhData(GrhIndexAnim).FileNum & ".bmp"
    AnimacionPosX = GrhData(GrhIndexAnim).sX
    AnimacionPosY = GrhData(GrhIndexAnim).sY
    AnimacionLongX = GrhData(GrhIndexAnim).pixelWidth
    AnimacionLongY = GrhData(GrhIndexAnim).pixelHeight
    
    If FileExist(GraficoPath, vbNormal) = True Then
        PbShow.PaintPicture LoadPicture(GraficoPath), 0, 0, AnimacionLongX, AnimacionLongY, AnimacionPosX, AnimacionPosY, AnimacionLongX, AnimacionLongY
    End If
End If

If FrameActual = CantFrames Then
    FrameActual = 0
End If
Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    TAnimacion.Enabled = False
    Exit Sub
End Sub

Public Sub SelectAnims(Direction As String)
Dim Grhs As Long

LstAnim.Clear

For Grhs = 1 To AllGrhData
    If GrhData(Grhs).NumFrames > 1 Then
        LstAnim.AddItem Grhs & "(ANIMACION)"
    End If
Next Grhs

AnimDir = Direction
End Sub
