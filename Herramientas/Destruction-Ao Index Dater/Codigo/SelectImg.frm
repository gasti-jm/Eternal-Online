VERSION 5.00
Begin VB.Form SelectImg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   6705
   Icon            =   "SelectImg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SelectImg.frx":08CA
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox FileImag 
      Height          =   4575
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label CmdSelect 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Image ImgShow 
      Height          =   2175
      Left            =   2400
      Picture         =   "SelectImg.frx":92CD2
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "SelectImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdSelect_Click()
Dim T As Long
Dim CantImg As Long

ReDim SelectedImg(1 To 2) As Long
For T = 0 To FileImag.ListCount - 1
    If FileImag.Selected(T) Then
        CantImg = CantImg + 1
        ReDim Preserve SelectedImg(1 To CantImg) As Long
        SelectedImg(CantImg) = Val(FileImag.List(T))
    End If
Next T

Call FrmIndexacion.SpecialIndex
End Sub

Private Sub FileImag_Click()
ImgShow.Picture = LoadPicture(Config.BmpPath & "\" & FileImag.fileName)
End Sub

Private Sub Form_Load()
FileImag.path = Config.BmpPath
MsgBox "Recorda que lo mejor es que las imagenes sean de numeros seguidos para que sea mas facil encontrarlas"
End Sub
