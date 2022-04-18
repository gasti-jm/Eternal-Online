VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1422
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Comprimir interfaces"
      Height          =   435
      Left            =   1680
      TabIndex        =   22
      Top             =   4200
      Width           =   1275
   End
   Begin VB.CheckBox cmdGrhPNG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox cmdMiniMap 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin DesCompresorAO.uAOButton cmdCerrar 
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":29DB2
      PICF            =   "frmMain.frx":2A7DC
      PICH            =   "frmMain.frx":2B49E
      PICV            =   "frmMain.frx":2C430
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox cmdUC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   3840
      Width           =   195
   End
   Begin VB.TextBox txtContra 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000020&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4080
      MaxLength       =   250
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin DesCompresorAO.uAOButton cmdComprimirMaps 
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Comprimir Mapas"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":2D332
      PICF            =   "frmMain.frx":2DD5C
      PICH            =   "frmMain.frx":2EA1E
      PICV            =   "frmMain.frx":2F9B0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4080
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "0"
      Top             =   4200
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar StatusBar 
      Height          =   3975
      Left            =   840
      TabIndex        =   12
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   7011
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin DesCompresorAO.uAOButton cmdDescomprimirMaps 
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Descomprimir Mapas"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":308B2
      PICF            =   "frmMain.frx":312DC
      PICH            =   "frmMain.frx":31F9E
      PICV            =   "frmMain.frx":32F30
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesCompresorAO.uAOButton cmdParcheMaps 
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Crear Parche Mapas"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":33E32
      PICF            =   "frmMain.frx":3485C
      PICH            =   "frmMain.frx":3551E
      PICV            =   "frmMain.frx":364B0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesCompresorAO.uAOButton cmdAplicarMaps 
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Aplicar Parche Mapas"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":373B2
      PICF            =   "frmMain.frx":37DDC
      PICH            =   "frmMain.frx":38A9E
      PICV            =   "frmMain.frx":39A30
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesCompresorAO.uAOButton cmdComprimirGrh 
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Comprimir Graficos"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":3A932
      PICF            =   "frmMain.frx":3B35C
      PICH            =   "frmMain.frx":3C01E
      PICV            =   "frmMain.frx":3CFB0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesCompresorAO.uAOButton cmdDescomprimirGrh 
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Descomprimir Graficos"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":3DEB2
      PICF            =   "frmMain.frx":3E8DC
      PICH            =   "frmMain.frx":3F59E
      PICV            =   "frmMain.frx":40530
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesCompresorAO.uAOButton cmdParcheGrh 
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Crear Parche Graficos"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":41432
      PICF            =   "frmMain.frx":41E5C
      PICH            =   "frmMain.frx":42B1E
      PICV            =   "frmMain.frx":43AB0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesCompresorAO.uAOButton cmdAplicarGrh 
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      TX              =   "Aplicar Parche Graficos"
      ENAB            =   -1  'True
      FCOL            =   12640511
      OCOL            =   16777215
      PICE            =   "frmMain.frx":449B2
      PICF            =   "frmMain.frx":453DC
      PICH            =   "frmMain.frx":4609E
      PICV            =   "frmMain.frx":47030
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizar PNG:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar MiniMapas:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by ^[GS]^ (www.gs-zone.org)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   2520
      TabIndex        =   17
      Top             =   480
      Width           =   2640
   End
   Begin VB.Label lVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.x.x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1800
      TabIndex        =   16
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizar contraseña:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label lEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versión:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112

Private Sub Trabajando(ByVal Estado As Boolean)

On Error Resume Next

    If Estado = True Then
        cmdComprimirGrh.Enabled = False
        cmdDescomprimirGrh.Enabled = False
        cmdParcheGrh.Enabled = False
        cmdAplicarGrh.Enabled = False
        cmdComprimirMaps.Enabled = False
        cmdDescomprimirMaps.Enabled = False
        cmdParcheMaps.Enabled = False
        cmdAplicarMaps.Enabled = False
        txtContra.Enabled = False
        txtVersion.Enabled = False
        cmdUC.Enabled = False
    Else
        cmdComprimirGrh.Enabled = True
        cmdDescomprimirGrh.Enabled = True
        cmdParcheGrh.Enabled = True
        cmdAplicarGrh.Enabled = True
        cmdComprimirMaps.Enabled = True
        cmdDescomprimirMaps.Enabled = True
        cmdParcheMaps.Enabled = True
        cmdAplicarMaps.Enabled = True
        txtContra.Enabled = True
        txtVersion.Enabled = True
        cmdUC.Enabled = True
    End If
End Sub


Private Sub cmdAplicarGrh_Click()

On Error Resume Next

    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim PatchPath As String
    
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    OldVersion = CLng(txtVersion.Text)
    NewVersion = OldVersion + 1
    
    If AbrirComprimido = True Then
        OldResourcePath = IniPath & CompressFile
        If AbrirComprimido("Seleccione el archivo Graficos.PATCH", "Patch de Graficos (*.PATCH)" & Chr$(0) & "*.PATCH" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*") = True Then ' el parche
            PatchPath = IniPath & CompressFile
            If GuardarComprimido = True Then
            
                NewResourcePath = IniPathD & CompressFileD
                
                'Check if the old resource file exists
                If Not FileExist(OldResourcePath, vbNormal) Then
                    MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & OldResourcePath, , "Error"
                    Exit Sub
                End If
                
                'We look if there's a patch to apply to the current version.
                If Not FileExist(PatchPath, vbArchive) Then
                    MsgBox "No existe el archivo .PATCH", vbExclamation, "Error"
                    Exit Sub
                End If
                
                'Look if the new resource file already exists.
                If FileExist(NewResourcePath, vbArchive) Then
                    If (MsgBox("Ya se encuentra el archivo parcheado, ¿Desea reemplazarlo?", vbInformation + vbYesNo, "Error") = vbNo) Then Exit Sub
                End If
                
                lEstado.Caption = "Aplicando Parche de " & OldVersion & " a " & NewVersion

                Call Trabajando(True)
                'Patch!
                If Apply_Patch(NewResourcePath, OldResourcePath, PatchPath, frmMain.StatusBar) Then
                    'Show we finished
                    MsgBox "Operación terminada con éxito!", vbInformation
                Else
                    'Show we finished
                    MsgBox "Operación abortada!", vbCritical
                End If
                Call Trabajando(False)
                
                lEstado.Caption = vbNullString
                
            End If
        End If
    End If
    
End Sub


Private Sub cmdAplicarMaps_Click()

On Error Resume Next

    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim PatchPath As String
    
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    OldVersion = CLng(txtVersion.Text)
    NewVersion = OldVersion + 1
    
    If AbrirComprimido = True Then
        OldResourcePath = IniPath & CompressFile
        If AbrirComprimido("Seleccione el archivo Mapas.PATCH", "Patch de Mapas (*.PATCH)" & Chr$(0) & "*.PATCH" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*") = True Then ' el parche
            PatchPath = IniPath & CompressFile
            If GuardarComprimido = True Then
            
                NewResourcePath = IniPathD & CompressFileD
                
                'Check if the old resource file exists
                If Not FileExist(OldResourcePath, vbNormal) Then
                    MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & OldResourcePath, , "Error"
                    Exit Sub
                End If
                
                'We look if there's a patch to apply to the current version.
                If Not FileExist(PatchPath, vbArchive) Then
                    MsgBox "No existe el archivo .PATCH", vbExclamation, "Error"
                    Exit Sub
                End If
                
                'Look if the new resource file already exists.
                If FileExist(NewResourcePath, vbArchive) Then
                    If (MsgBox("Ya se encuentra el archivo parcheado, ¿Desea reemplazarlo?", vbInformation + vbYesNo, "Error") = vbNo) Then Exit Sub
                End If
                
                lEstado.Caption = "Aplicando Parche de " & OldVersion & " a " & NewVersion
                
                Call Trabajando(True)
                'Patch!
                If Apply_Patch(NewResourcePath, OldResourcePath, PatchPath, frmMain.StatusBar) Then
                    'Show we finished
                    MsgBox "Operación terminada con éxito!", vbInformation
                Else
                    'Show we finished
                    MsgBox "Operación abortada!", vbCritical
                End If
                Call Trabajando(False)
                
                lEstado.Caption = vbNullString
                
            End If
        End If
    End If

End Sub

Private Sub cmdCerrar_Click()
    End
End Sub

Private Sub cmdComprimirGrh_Click()

On Error Resume Next

    Dim SourcePath As String
    Dim OutputPath As String

    SourcePath = SeleccionarDirectorio("Seleccione el Directorio de las Imagenes")
    
    If SourcePath <> vbNullString Then
        If GuardarComprimido = True Then
        
            OutputPath = IniPathD & CompressFileD
            
            'Check if the version already exists
            If FileExist(OutputPath, vbNormal) Then
                If MsgBox("El archivo de Graficos.EO ya se encuentra comprimido. ¿Desea reemplazarlo?", vbYesNo, "Atencion") = vbNo Then _
                    Exit Sub
            End If

            lEstado.Caption = "Comprimiendo..."
            
            Call Trabajando(True)
            'Compress!
            If Compress_Files(SourcePath, OutputPath, CLng(txtVersion.Text), StatusBar, 0) Then
                'Show we finished
                MsgBox "Operación terminada con éxito!", vbInformation
            Else
                'Show we finished
                MsgBox "Operación abortada!", vbCritical
            End If
            Call Trabajando(False)
            
            lEstado.Caption = vbNullString

        End If
    End If

End Sub

Private Sub cmdComprimirMaps_Click()

On Error Resume Next

    Dim SourcePath As String
    Dim OutputPath As String

    SourcePath = SeleccionarDirectorio("Seleccione el Directorio de los Mapas")
    
    If SourcePath <> vbNullString Then
        If GuardarComprimido("Indique donde será guardado Mapas.EO", "Mapas (*.EO)" & vbNullChar & "*.EO" & vbNullChar & "Todos los archivos" & vbNullChar & "*.*") = True Then
        
            OutputPath = IniPathD & CompressFileD
            
            'Check if the version already exists
            If FileExist(OutputPath, vbNormal) Then
                If MsgBox("El archivo de Mapas.EO ya se encuentra comprimido. ¿Desea reemplazarlo?", vbYesNo, "Atencion") = vbNo Then _
                    Exit Sub
            End If

            lEstado.Caption = "Comprimiendo..."
            
            Call Trabajando(True)
            'Compress!
            If Compress_Files(SourcePath, OutputPath, CLng(txtVersion.Text), StatusBar, 1) Then
                'Show we finished
                MsgBox "Operación terminada con éxito!", vbInformation
            Else
                'Show we finished
                MsgBox "Operación abortada!", vbCritical
            End If
            Call Trabajando(False)
            
            lEstado.Caption = vbNullString

        End If
    End If
End Sub

Private Sub cmdDescomprimirMaps_Click()

On Error Resume Next
    
    Dim ResourcePath As String
    Dim OutputPath As String
    
    If AbrirComprimido("Seleccione el archivo Mapas.EO", "Graficos (*.EO)" & vbNullChar & "*.EO" & vbNullChar & "Todos los archivos" & vbNullChar & "*.*") = True Then
    
        ResourcePath = IniPath & CompressFile
        OutputPath = SeleccionarDirectorio("Seleccione el Directorio donde desea descomprimir los Mapas")
        
        If OutputPath <> vbNullString Then

            'Check if the resource file exists
            If Not FileExist(ResourcePath, vbNormal) Then
                MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
                Exit Sub
            End If
            
            'Check if the version is already extracted
            If FileExist(OutputPath, vbDirectory) Then
                If MsgBox("El directorio ya se encuentra creado y puede contener otros archivos." & vbCrLf & "¿Desea utilizarlo de todas formas?", vbYesNo, "Atencion") = vbNo Then _
                    Exit Sub
            Else
                'Create this version folder
                MkDir OutputPath
            End If
            
            lEstado.Caption = "Descomprimiendo..."
            
            Call Trabajando(True)
            'Extract!
            If Extract_Files(ResourcePath, OutputPath, StatusBar) Then
                'Show we finished
                MsgBox "Operación terminada con éxito!", vbInformation
            Else
                'Show we finished
                MsgBox "Operación abortada!", vbCritical
            End If
            Call Trabajando(False)
            
            lEstado.Caption = vbNullString
            
        End If
    End If
End Sub

Private Sub cmdParcheGrh_Click()

On Error Resume Next

    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim OutputPath As String
    
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    NewVersion = CLng(txtVersion.Text)
    OldVersion = NewVersion - 1 'we patch from the last version
    
    If AbrirComprimido("Seleccione el antiguo archivo de Graficos.EO") = True Then ' viejo
        OldResourcePath = IniPath & CompressFile
        If AbrirComprimido("Seleccione el nuevo archivo de Graficos.EO") = True Then ' nuevo
            NewResourcePath = IniPath & CompressFile
            If GuardarComprimido("Seleccione donde será creado el nuevo Graficos.PATCH", "Patch de Graficos (*.PATCH)" & Chr$(0) & "*.PATCH" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*") = True Then ' parche
            
                OutputPath = IniPathD & CompressFileD
                
                'Check if the new resource file exists
                If Not FileExist(NewResourcePath, vbNormal) Then
                    MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & NewResourcePath, , "Error"
                    Exit Sub
                End If
                
                'Check if the old resource file exists
                If Not FileExist(OldResourcePath, vbNormal) Then
                    MsgBox "No se encontraron los recursos de la version anterior." & vbCrLf & OldResourcePath, , "Error"
                    Exit Sub
                End If
                
                'Check if the version is already extracted
                If FileExist(OutputPath, vbNormal) Then
                    If MsgBox("El parche ya se encuentra creado. ¿Desea reemplazarlo?", vbYesNo, "Atencion") = vbNo Then _
                        Exit Sub
                End If
                
                lEstado.Caption = "Creando el parche de " & OldVersion & " a " & NewVersion
                
                Call Trabajando(True)
                'Patch!
                If Make_Patch(NewResourcePath, OldResourcePath, OutputPath, StatusBar) Then
                    'Show we finished
                    MsgBox "Operación terminada con éxito!", vbInformation
                Else
                    'Show we finished
                    MsgBox "Operación abortada!", vbCritical
                End If
                Call Trabajando(False)
                
                lEstado.Caption = vbNullString
                
            End If
        End If
    End If
    
End Sub

Private Sub cmdDescomprimirGrh_Click()

On Error Resume Next

    Dim ResourcePath As String
    Dim OutputPath As String
    
    If AbrirComprimido = True Then
    
        ResourcePath = IniPath & CompressFile
        OutputPath = SeleccionarDirectorio("Seleccione el Directorio donde desea descomprimir los Graficos")
        
        If OutputPath <> vbNullString Then

            'Check if the resource file exists
            If Not FileExist(ResourcePath, vbNormal) Then
                MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
                Exit Sub
            End If
            
            'Check if the version is already extracted
            If FileExist(OutputPath, vbDirectory) Then
                If MsgBox("El directorio ya se encuentra creado y puede contener otros archivos." & vbCrLf & "¿Desea utilizarlo de todas formas?", vbYesNo, "Atencion") = vbNo Then _
                    Exit Sub
            Else
                'Create this version folder
                MkDir OutputPath
            End If
            
            lEstado.Caption = "Descomprimiendo..."
            
            Call Trabajando(True)
            'Extract!
            If Extract_Files(ResourcePath, OutputPath, StatusBar) Then
                'Show we finished
                MsgBox "Operación terminada con éxito!", vbInformation
            Else
                'Show we finished
                MsgBox "Operación abortada!", vbCritical
            End If
            Call Trabajando(False)
            
            lEstado.Caption = vbNullString
            
        End If
    End If

End Sub

Private Sub cmdParcheMaps_Click()

On Error Resume Next

    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim OutputPath As String
    
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    NewVersion = CLng(txtVersion.Text)
    OldVersion = NewVersion - 1 'we patch from the last version
    
    If AbrirComprimido("Seleccione el antiguo archivo de Mapas.EO") = True Then ' viejo
        OldResourcePath = IniPath & CompressFile
        If AbrirComprimido("Seleccione el nuevo archivo de Mapas.EO") = True Then ' nuevo
            NewResourcePath = IniPath & CompressFile
            If GuardarComprimido("Seleccione donde será creado el nuevo Mapas.PATCH", "Patch de Mapas (*.PATCH)" & Chr$(0) & "*.PATCH" & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*") = True Then ' parche
            
                OutputPath = IniPathD & CompressFileD
                
                'Check if the new resource file exists
                If Not FileExist(NewResourcePath, vbNormal) Then
                    MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & NewResourcePath, , "Error"
                    Exit Sub
                End If
                
                'Check if the old resource file exists
                If Not FileExist(OldResourcePath, vbNormal) Then
                    MsgBox "No se encontraron los recursos de la version anterior." & vbCrLf & OldResourcePath, , "Error"
                    Exit Sub
                End If
                
                'Check if the version is already extracted
                If FileExist(OutputPath, vbNormal) Then
                    If MsgBox("El parche ya se encuentra creado. ¿Desea reemplazarlo?", vbYesNo, "Atencion") = vbNo Then _
                        Exit Sub
                End If
                
                lEstado.Caption = "Creando el parche de " & OldVersion & " a " & NewVersion
                
                Call Trabajando(True)
                'Patch!
                If Make_Patch(NewResourcePath, OldResourcePath, OutputPath, StatusBar) Then
                    'Show we finished
                    MsgBox "Operación terminada con éxito!", vbInformation
                Else
                    'Show we finished
                    MsgBox "Operación abortada!", vbCritical
                End If
                Call Trabajando(False)
                
                lEstado.Caption = vbNullString
                
            End If
        End If
    End If
End Sub

Private Sub Command1_Click()

On Error Resume Next

    Dim SourcePath As String
    Dim OutputPath As String

    SourcePath = SeleccionarDirectorio("Seleccione el Directorio de las interfaces")
    
    If SourcePath <> vbNullString Then
        If GuardarComprimido("Indique donde será guardado Interface.EO", "Interface (*.EO)" & vbNullChar & "*.EO" & vbNullChar & "Todos los archivos" & vbNullChar & "*.*") = True Then
        
            OutputPath = IniPathD & CompressFileD
            
            'Check if the version already exists
            If FileExist(OutputPath, vbNormal) Then
                If MsgBox("El archivo de Interface.EO ya se encuentra comprimido. ¿Desea reemplazarlo?", vbYesNo, "Atencion") = vbNo Then _
                    Exit Sub
            End If

            lEstado.Caption = "Comprimiendo..."
            
            Call Trabajando(True)
            'Compress!
            If Compress_Files(SourcePath, OutputPath, CLng(txtVersion.Text), StatusBar, 2) Then
                'Show we finished
                MsgBox "Operación terminada con éxito!", vbInformation
            Else
                'Show we finished
                MsgBox "Operación abortada!", vbCritical
            End If
            Call Trabajando(False)
            
            lEstado.Caption = vbNullString

        End If
    End If
End Sub

Private Sub Form_Load()

On Error Resume Next

    lVer.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    ChDir App.Path
    ChDrive App.Path
    
    ' nos fijamos si tiene registrado zlib.dll :S
    If TestZLib = False Then
        MsgBox "Es necesario disponer de la librería zlib.dll registrada para el correcto funcionamiento del programa.", vbCritical, "Error"
        End ' no seguimos, si es para fallar
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    If Button = 1 Then
        ReleaseCapture
        SendMessage hWnd, WM_NCLBUTTONDOWN, _
            HTCAPTION, 0&
    End If
End Sub

Private Sub lBy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lVer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub txtContra_Change()

On Error Resume Next

    If LenB(txtContra.Text) <> 0 Then
        cmdUC.Value = 1
        Call GenerateContra
    Else
        cmdUC.Value = 0
    End If
End Sub

Private Sub txtContra_LostFocus()

On Error Resume Next
        
    Call GenerateContra
End Sub

Private Sub txtVersion_Change()

On Error Resume Next
    
    If IsNumeric(txtVersion.Text) = False Then
        txtVersion.Text = 0
    End If
End Sub

Private Sub cmdUC_Click()

On Error Resume Next

    If cmdUC.Value = 1 Then
        txtContra.BackColor = &H202020
        usaContra = True
    Else
        txtContra.BackColor = &H20
        usaContra = False
    End If
End Sub

Public Function GenerateContra()

On Error Resume Next

    Dim loopc As Byte
    Erase datContra
    
    If LenB(txtContra.Text) <> 0 Then
        ReDim datContra(Len(txtContra.Text) - 1)
        For loopc = 0 To UBound(datContra)
            datContra(loopc) = Asc(Mid(txtContra.Text, loopc + 1, 1))
        Next loopc
    End If
    
End Function
