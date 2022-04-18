VERSION 5.00
Begin VB.Form frmRenderer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renderizado"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   532
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   0
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   2
      Top             =   480
      Width           =   7500
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   0
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   480
      Width           =   7500
   End
End
Attribute VB_Name = "frmRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
 
' Recupera la imagen del área del control
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long


Private Sub cmdGuardar_Click()
Call Capturar_Imagen(Picture1, Picture2)
 
'call SavePicture(frmRenderer.Picture2.image, App.Path & "\Mapa1.bmp")
Call SavePicture(frmRenderer.Picture2.Image, App.path & "\Screenshots\" & frmMain.MapPest(4).Caption & ".bmp")
Unload Me
End Sub

Private Sub Capturar_Imagen(Control As Control, Destino As Object)
     
    Dim hdc As Long
    Dim Escala_Anterior As Integer
    Dim Ancho As Long
    Dim Alto As Long
     
    ' Para que se mantenga la imagen por si se repinta la ventana
    Destino.AutoRedraw = True
     
    On Error Resume Next
    ' Si da error es por que el control está dentro de un Frame _
      ya que  los Frame no tiene  dicha propiedad
    Escala_Anterior = Control.Container.ScaleMode
     
    If Err.Number = 438 Then
       ' Si el control está en un Frame, convierte la escala
       Ancho = ScaleX(Control.Width, vbTwips, vbPixels)
       Alto = ScaleY(Control.Height, vbTwips, vbPixels)
    Else
       ' Si no cambia la escala del  contenedor a pixeles
       Control.Container.ScaleMode = vbPixels
       Ancho = Control.Width
       Alto = Control.Height
    End If
     
    ' limpia el error
    On Error GoTo 0
    ' Captura el área de pantalla correspondiente al control
    hdc = GetWindowDC(Control.hwnd)
    ' Copia esa área al picturebox
    BitBlt Destino.hdc, 0, 0, Ancho, Alto, hdc, 0, 0, vbSrcCopy
    ' Convierte la imagen anterior en un Mapa de bits
    Destino.Picture = Destino.Image
    ' Borra la imagen ya que ahora usa el Picture
    Destino.Cls
     
    On Error Resume Next
    If Err.Number = 0 Then
       ' Si el control no está en un  Frame, restaura la escala del contenedor
       Control.Container.ScaleMode = Escala_Anterior
    End If
     
End Sub

