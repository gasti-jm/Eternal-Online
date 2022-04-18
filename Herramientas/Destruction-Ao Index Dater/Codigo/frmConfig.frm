VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConfig.frx":08CA
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox BmpSavePath 
      Height          =   285
      Left            =   3360
      TabIndex        =   20
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox TxtWavPath 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox DatPath 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox SaveDatPath 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox SaveInitPath 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox initPathTxt 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox bmpPathTxt 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   150
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiles"
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox HeightTxt 
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Text            =   "32"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox WidthTxt 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "32"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alto del tile:"
         Height          =   195
         Left            =   4080
         TabIndex        =   4
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ancho del tile:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Label SearchBmpSave 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label CancelCmd 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Accept 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label SearchInit 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label SearchInitSave 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label SearchDat 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label SearchDatSave 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label SearchWav 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label SearchBmp 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_HEIGHT As Byte = 32
Private Const DEFAULT_WIDTH As Byte = 32

Private Sub Accept_Click()
    If bmpPathTxt.Text = vbNullString Then
        MsgBox "Por favor ingresa una ruta para los Graficos, antes de continuar."
        Exit Sub
    End If
    
    If BmpSavePath.Text = vbNullString Then
        MsgBox "Por favor ingresa una ruta para grabar los graficos modificados a potencia de 2, antes de continuar"
        Exit Sub
    End If
    
    If initPathTxt.Text = vbNullString Then
        MsgBox "Por favor ingresa una ruta para los Inits, antes de continuar."
        Exit Sub
    End If
    
    If DatPath.Text = vbNullString Then
        MsgBox "Por favor ingresa una ruta para los Dats, antes de continuar."
        Exit Sub
    End If
    
    If SaveInitPath.Text = vbNullString Then
        MsgBox "Por favor ingresa una ruta para Guardar los Inits modificados, antes de continuar."
        Exit Sub
    End If
    
    If SaveDatPath.Text = vbNullString Then
        MsgBox "Por favor ingresa una ruta para guardar los Dats modificados, antes de continuar."
        Exit Sub
    End If
    
    If TxtWavPath.Text = vbNullString Then
        MsgBox "Por favor ingresa una ruta para los WAV, antes de continuar."
        Exit Sub
    End If
    
    If Val(HeightTxt.Text) <= 0 Then
        MsgBox "Tile height must be a positive number and different from zero."
        Exit Sub
    End If
    
    If Val(WidthTxt.Text) <= 0 Then
        MsgBox "Tile width must be a positive number and different from zero."
        Exit Sub
    End If
    
    
    Config.BmpPath = bmpPathTxt.Text
    Config.BmpSavePath = BmpSavePath.Text
    Config.InitPath = initPathTxt.Text
    Config.SaveInitPath = SaveInitPath.Text
    Config.SaveDatPath = SaveDatPath.Text
    Config.TilePixelHeight = Val(HeightTxt.Text)
    Config.TilePixelWidth = Val(WidthTxt.Text)
    Config.DatPath = DatPath.Text
    Config.WavPath = TxtWavPath.Text

    Config.SaveConfig
    
    Call Unload(Me)
End Sub
Private Sub CancelCmd_Click()
    'Is there a saved config, or we are requesting it for the first time?
    If Config.LoadConfig() Then
        Call Unload(Me)
    Else
        'Shut down!
        End
    End If
End Sub

Private Sub SearchBmpSave_Click()
    BmpSavePath.Text = BrowseForFolder(Me.hWnd, "Ruta para guardado de Graficos en potencia de 2")
End Sub

Private Sub SearchDat_Click()
    DatPath.Text = BrowseForFolder(Me.hWnd, "Ruta de los Dats")
End Sub

Private Sub Form_Load()
    bmpPathTxt.Text = Config.BmpPath
    initPathTxt.Text = Config.InitPath
    DatPath.Text = Config.DatPath
    SaveInitPath.Text = Config.SaveInitPath
    SaveDatPath.Text = Config.SaveDatPath
    TxtWavPath.Text = Config.WavPath
    BmpSavePath.Text = Config.BmpSavePath
    
    HeightTxt.Text = CStr(32)
    WidthTxt.Text = CStr(32)
End Sub

Private Sub HeightTxt_Change()
    If Not IsNumeric(HeightTxt.Text) Then
        HeightTxt.Text = DEFAULT_HEIGHT
    End If
End Sub

Private Sub SearchBmp_Click()
    bmpPathTxt.Text = BrowseForFolder(Me.hWnd, "Ruta de las Imagenes")
End Sub

Private Sub SearchDatSave_Click()
    SaveDatPath.Text = BrowseForFolder(Me.hWnd, "Ruta para guardar los Dats Modificados")
End Sub

Private Sub SearchInit_Click()
    initPathTxt.Text = BrowseForFolder(Me.hWnd, "Ruta de los Inits")
End Sub

Private Sub SearchInitSave_Click()
    SaveInitPath.Text = BrowseForFolder(Me.hWnd, "Ruta para guardar los Init modificados")
End Sub

Private Sub SearchWav_Click()
    TxtWavPath.Text = BrowseForFolder(Me.hWnd, "Ruta de los WAV")
End Sub

Private Sub TxtWav_Change()

End Sub

Private Sub WidthTxt_Change()
    If Not IsNumeric(WidthTxt.Text) Then
        WidthTxt.Text = DEFAULT_WIDTH
    End If
End Sub
