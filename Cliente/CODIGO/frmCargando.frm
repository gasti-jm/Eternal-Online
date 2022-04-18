VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmCargando.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Eternal Online"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image picLoading 
      Height          =   810
      Left            =   7920
      Top             =   7920
      Width           =   3855
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    'Aplicate in forms
    Me.Width = frmScaleWidth
    Me.Height = frmScaleHeight

    Call FormParser.Parse_Form(Me, E_WAIT)
    picLoading.Left = Config_Inicio.ResolutionX - 272
    picLoading.Top = Config_Inicio.ResolutionY - 72
    picLoading.Picture = LoadPicture(DirInterface & "VentanaCargando.bmp")
End Sub
