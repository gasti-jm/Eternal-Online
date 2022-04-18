VERSION 5.00
Begin VB.Form FrmIndex 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destruction-Ao Indexador"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   Icon            =   "FrmSeleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSeleccion.frx":08CA
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File 
      Height          =   3600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Command3 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Command2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2370
      Width           =   1455
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image ImagenSelec 
      Height          =   1515
      Left            =   0
      Picture         =   "FrmSeleccion.frx":B0C1E
      Top             =   4125
      Width           =   5820
   End
End
Attribute VB_Name = "FrmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Ruta = File.path & "\" & File.FileName
If Not InStr(UCase(File.FileName), ".BMP") > 0 Then
    MsgBox "Primero elige una imagen con extencion '.bmp'"
    Exit Sub
End If
NombreGrafico = Left(File.FileName, Len(File.FileName) - 4)
Me.Visible = False
FrmIndexacion.Visible = True
End Sub

Private Sub Command2_Click()
Me.Visible = False
FrmAnimaciones.Visible = True
End Sub

Private Sub Command3_Click()
'Me.Visible = False
'FrmCreditos.Visible = True
End Sub

Private Sub File_Click()
On Error GoTo ErrHandler
Ruta = File.path & "\" & File.FileName
If InStr(UCase(File.FileName), ".BMP") > 0 Then
    ImagenSelec.Picture = LoadPicture(Ruta)
End If
Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub Form_Load()
'Call LoadGrhData
'MsgBox "Espero que les guste el programa que hice y que les ayude en la indexacion." & vbCrLf & "Es muy recomendado leer el manual de indexacion que esta en la carpeta en la que descomprimieron este programa." & vbCrLf & "Recuerden poner este programa en la MISMA carpeta que estan los graficos a INDEXAR." & vbCrLf & "Programa hecho por Silver(Director y Programador de Destruction-Ao)" & vbCrLf & "Gracias Cynny por todo, Te Amo" & vbCrLf & "Cualquier problema, error, etc mi Mail y MSN es:                     silverdestruction-ao@hotmail.com"
File.path = Config.BmpPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmIndexMenu.Visible = True
End Sub

