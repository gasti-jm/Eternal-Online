VERSION 5.00
Begin VB.Form DIU 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Des-Indexador Universal"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7200
   Icon            =   "DIU.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4245
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox NoCabeceras 
      BackColor       =   &H00000000&
      Caption         =   "No utilizar Cabeceras (Modo IAO Clon, etc)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   2520
      Width           =   4815
   End
   Begin VB.ListBox lFormato 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   510
      ItemData        =   "DIU.frx":030A
      Left            =   720
      List            =   "DIU.frx":0314
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   480
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   7170
      TabIndex        =   5
      Top             =   3750
      Width           =   7200
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Programado por ^[GS]^ - Website: http://www.gs-zone.org"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000E000&
         Height          =   240
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   100
         Width           =   6570
      End
      Begin VB.Image imgCreditos 
         Height          =   495
         Left            =   0
         Picture         =   "DIU.frx":036A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9375
      End
      Begin VB.Label lblCreditos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Programado por ^[GS]^ - Website: http://www.gs-zone.org"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000E000&
         Height          =   240
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   7
         Top             =   120
         Width           =   6570
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   345
         Left            =   7320
         TabIndex        =   6
         Top             =   60
         Width           =   105
      End
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdSalir 
      Height          =   1935
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3413
      Caption         =   "&Salir"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   128
   End
   Begin VB.ListBox lModo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   990
      ItemData        =   "DIU.frx":03C2
      Left            =   720
      List            =   "DIU.frx":03D2
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdDesindexar 
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Caption         =   "&Desindexar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   32768
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdIndexar 
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Caption         =   "&Indexar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   32896
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdVer 
      Height          =   735
      Left            =   4080
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      Caption         =   "&Explorar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   16744576
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   8388608
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formato:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "DIU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDesindexar_Click()
If ExplorarDesindex Then
    If GuardarDesindex Then
        DoEvents
        If lFormato.ListIndex = 0 Then ' original
            Select Case lModo.ListIndex
                Case 0
                If Desindexar0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
                Case 1
                If DesindexarCabezas0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
                Case 2
                If DesindexarCuerpos0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
                Case 3
                If DesindexarFX0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
            End Select
        Else ' 0.13
            Select Case lModo.ListIndex
                Case 0
                If Desindexar0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
                Case 1
                If DesindexarCabezas0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
                Case 2
                If DesindexarCuerpos0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
                Case 3
                If DesindexarFX0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido desindexado con exito!", vbInformation
                End If
            End Select
        End If
    End If
End If
End Sub

Private Sub cmdIndexar_Click()
If ExplorarIndex Then
    If GuardarIndex Then
        DoEvents
        If lFormato.ListIndex = 0 Then ' original
            Select Case lModo.ListIndex
                Case 0
                If Indexar0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
                Case 1
                If IndexarCabezas0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
                Case 2
                If IndexarCuerpos0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
                Case 3
                If IndexarFX0120 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
            End Select
        Else    ' 0.13
            Select Case lModo.ListIndex
                Case 0
                If Indexar0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
                Case 1
                If IndexarCabezas0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
                Case 2
                If IndexarCuerpos0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
                Case 3
                If IndexarFX0130 Then
                    MsgBox IniPathD & GraphicsFileD & " ha sido indexado con exito!", vbInformation
                End If
            End Select
        End If
    End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVer_Click()
If lModo.ListIndex = 0 Then
    If ExplorarIndex Then
        If SelDirGraficos Then
            Me.WindowState = 1
            DoEvents
            If ExplorerGrh Then
                Call frmExplorerGrh.Show
                frmExplorerGrh.Caption = GraphicsFile & " [Explorador]"
            End If
        End If
    End If
Else
    MsgBox "Opción solo disponible para el Indice de Graficos!", vbInformation
End If

End Sub

Private Sub Form_Load()
Me.Caption = "Des-Indexador Universal v" & App.Major & "." & App.Minor & "." & App.Revision
lFormato.ListIndex = 0
lModo.ListIndex = 0
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
MsgBox Data.Files(Data.Files.Count)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Label2_Click()
    Call ShellExecute(Me.hWnd, "Open", "http://www.gs-zone.org", &O0, &O0, SW_NORMAL)
End Sub

Private Sub NoCabeceras_Click()
    NoUsarCabecera = (NoCabeceras.value = 1)
End Sub
