VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmConversor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversor de Gráficos"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   Icon            =   "FrmConversor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmConversor.frx":08CA
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Progreso 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox NewImagen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4080
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Convertir 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   1725
      Width           =   1335
   End
   Begin VB.Label CmdVolver 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label Estado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Empezar Conversion?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmConversor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variables para poder sacar las medidas del bmp
Private Type BITMAPINFOHEADER
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type

Private Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits      As Long
End Type

Dim FileHeaderBMP As BITMAPFILEHEADER
Dim InfoHeaderBMP As BITMAPINFOHEADER
Dim FSO As New FileSystemObject
Dim Fil As File
Dim Fold As Folder
Dim CerroVentana As Boolean
Private Sub CambiarTamaño() 'Completamente cambiado para que sea mas rapido
On Error GoTo ErrorHandler

Dim Numgraficos As Long
Dim ViejaImagen As IPictureDisp
Dim i As Long
Dim hfile As Long
Dim TotalNum As Integer

Set Fold = FSO.GetFolder(Config.BmpPath)

hfile = FreeFile
Numgraficos = 0

Estado.Caption = "Convirtiendo Imagenes."
For Each Fil In Fold.Files
    If CerroVentana = True Then
        Exit For
    End If
    If Right(Fil.Name, 4) = ".bmp" Then
        Numgraficos = Numgraficos + 1
    End If
Next

Progreso.max = Numgraficos
TotalNum = Numgraficos
Numgraficos = 0

For Each Fil In Fold.Files 'selecciona los bmp, los guarda en el array junto a sus medidas viejas y nuevas
    If CerroVentana = True Then
        Exit For
    End If
    If Right(Fil.Name, 4) = ".bmp" Then 'solo enlista a los arhivos .bmp
        Numgraficos = Numgraficos + 1
    
        Open Config.BmpPath & "\" & Fil.Name For Binary Access Read As hfile
            Get hfile, , FileHeaderBMP
            Get hfile, , InfoHeaderBMP
        Close hfile
        
        'Variables para contener informacion del Grafico
        Dim GrafName As String
        Dim GrafHeight As Integer
        Dim GrafHeightNew As Integer
        Dim GrafWidth As Integer
        Dim GrafWidthNew As Integer
        
        GrafName = Fil.Name
        GrafHeight = InfoHeaderBMP.biHeight
        GrafWidth = InfoHeaderBMP.biWidth
        
        If GrafWidth < GrafHeight Then
            GrafHeightNew = ObtenerDimension(GrafHeight)
            GrafWidthNew = ObtenerDimension(GrafHeight)
        ElseIf GrafWidth > GrafHeight Then
            GrafHeightNew = ObtenerDimension(GrafWidth)
            GrafWidthNew = ObtenerDimension(GrafWidth)
        ElseIf GrafWidth = GrafHeight Then
            GrafHeightNew = ObtenerDimension(GrafWidth)
            GrafWidthNew = ObtenerDimension(GrafWidth)
        End If
        Debug.Print GrafHeightNew

        Progreso.value = Numgraficos
        Estado.Caption = "Convirtiendo gráficos: " & Numgraficos & " / " & TotalNum & vbCrLf & "Nombre del Gráfico: " & GrafName

        Set ViejaImagen = LoadPicture(Config.BmpPath & "\" & GrafName)
        NewImagen.Width = GrafWidthNew
        NewImagen.Height = GrafHeightNew
        Call NewImagen.Cls
        Call NewImagen.PaintPicture(ViejaImagen, 0, 0, GrafWidth, GrafHeight)
        Call SavePicture(NewImagen.Image, Config.BmpSavePath & "\" & GrafName)
    End If
    DoEvents
Next

Estado.Caption = "Graficos Convertidos con Exito!!"
Exit Sub
ErrorHandler:
    MsgBox "Al redimencionar un grafico surgio un error, el grafico '" & GrafName & "' no se ha podido redimencionar, generalmente esto surge porque el Grafico tiene errores. Borralo o Modificalo y volve a Empezar. Disculpa las molestias."
End Sub
Private Function ObtenerDimension(ByVal mDimension As Integer) As Integer
Select Case mDimension 'aca esta funcion la hice mas prolija :P
    Case Is <= 32
        ObtenerDimension = 32
        Exit Function
    
    Case Is <= 64
        ObtenerDimension = 64
        Exit Function
    
    Case Is <= 128
        ObtenerDimension = 128
        Exit Function
    
    Case Is <= 256
        ObtenerDimension = 256
        Exit Function
    
    Case Is <= 512
        ObtenerDimension = 512
        Exit Function
    
    Case Is <= 1024
        ObtenerDimension = 1024
        Exit Function
    
    Case Is <= 2048
        ObtenerDimension = 2048
        Exit Function
    
    Case Is <= 4096
        ObtenerDimension = 4096
        Exit Function
End Select
End Function

Private Sub CmdVolver_Click()
CerroVentana = True
Call Unload(Me)
End Sub

Private Sub Convertir_Click()
On Error GoTo ErrorFolder
If FSO.FolderExists(Config.BmpSavePath) Then
    FSO.DeleteFolder (Config.BmpSavePath)
End If
FSO.CreateFolder (Config.BmpSavePath)
CerroVentana = False
CambiarTamaño

Exit Sub
ErrorFolder:
    MsgBox "La carpeta de Graficos fue borrada para volver a generar los graficos desde 0 y el programa no tiene permisos para poder crear la nueva carpeta." & vbCrLf & "Por favor creala a mano y volve a ejecutar el proceso." & vbCrLf & vbCrLf & "Si el error persiste envien un mail a 'soporte@aodestruction.com.ar'. Gracias y disculpen las molestias"
    Call Unload(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CerroVentana = True
Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
CerroVentana = True
Call Unload(Me)
End Sub

Private Sub Label1_Click()

End Sub
