VERSION 5.00
Begin VB.Form frmExplorerGrh 
   BackColor       =   &H00000000&
   Caption         =   "Explorar Indice de Graficos"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExplorerGrh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerFrame 
      Enabled         =   0   'False
      Left            =   6960
      Top             =   3960
   End
   Begin VB.ListBox lZoom 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
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
      Height          =   870
      ItemData        =   "frmExplorerGrh.frx":030A
      Left            =   8280
      List            =   "frmExplorerGrh.frx":0320
      TabIndex        =   5
      ToolTipText     =   "Zoom de Muestra"
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox imgGrh 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   3000
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox tCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
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
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   6255
   End
   Begin DesIndexadorUniversal.LynxGrid lGrh 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10398
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   2105376
      BackColorBkg    =   0
      BackColorEdit   =   14737632
      BackColorSel    =   0
      ForeColor       =   16777215
      ForeColorHdr    =   8421504
      ForeColorSel    =   8438015
      BackColorEvenRows=   3158064
      CustomColorFrom =   4210752
      CustomColorTo   =   8421504
      GridLines       =   2
      ThemeColor      =   5
      ScrollBars      =   1
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      HotHeaderTracking=   0   'False
   End
   Begin DesIndexadorUniversal.LynxGrid lInfo 
      Height          =   2175
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   2105376
      BackColorBkg    =   0
      BackColorEdit   =   14737632
      BackColorSel    =   0
      ForeColor       =   16777215
      ForeColorHdr    =   8421504
      ForeColorSel    =   8438015
      BackColorEvenRows=   3158064
      CustomColorFrom =   4210752
      CustomColorTo   =   8421504
      GridLines       =   2
      ThemeColor      =   5
      ScrollBars      =   1
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      HotHeaderTracking=   0   'False
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdShow 
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Forzar el muestro de Grh"
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Mostrar Grh"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   4210752
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   65280
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdPausa 
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      ToolTipText     =   "Detener animación"
      Top             =   4800
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "| |"
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
      cFHover         =   4210752
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   32896
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdNext 
      Height          =   495
      Left            =   8640
      TabIndex        =   8
      ToolTipText     =   "Siguiente cuadro"
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   ">>"
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
      cFHover         =   4210752
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   16576
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdPrev 
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      ToolTipText     =   "Cuadro previo"
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "<<"
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
      cFHover         =   4210752
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   16744576
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdPlay 
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      ToolTipText     =   "Animar"
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "|>"
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
      cFHover         =   4210752
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   32768
   End
   Begin DesIndexadorUniversal.lvButtons_H cmdRegresar 
      Height          =   735
      Left            =   6960
      TabIndex        =   12
      Top             =   5400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      Caption         =   "&Regresar"
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
   Begin DesIndexadorUniversal.lvButtons_H cmdMantener 
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      ToolTipText     =   "Mantener el cambio mientras continue en el Explorador"
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Mantener"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   4210752
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin VB.Image hT 
      Height          =   495
      Left            =   7560
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lGrhFrame 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lbZoom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmExplorerGrh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ZoomImg As Single

Private showDecode As String
Private showGrhNF As Integer
Private showGrhFile As String
Private showGrhX As Integer
Private showGrhY As Integer
Private showGrhW As Integer
Private showGrhH As Integer
Private showGrhFF() As Long
Private showGrhS As Single

Private SelectedGrh As Long
Private showActualFrame As Integer
Private Transp As Boolean

Private Sub cmdMantener_Click()
    If LenB(tCode.Text) <> 0 Then
        Dim y As Integer
        Dim TT As Integer
        Dim DatoX() As String
        DatoX() = Split(tCode.Text, "-")
        TT = Val(InStr(1, DatoX(0), "="))
        DatoX(0) = Right(DatoX(0), Len(DatoX(0)) - TT)
        showGrhNF = Val(DatoX(0)) ' Numero de Frames
        If (showGrhNF = 1) Then ' no anim
            showGrhFile = DatoX(1) ' Nombre de el BMP (no solo numero así se puede testear con chars)
            showGrhX = Val(DatoX(2)) ' X
            showGrhY = Val(DatoX(3)) ' Y
            showGrhW = Val(DatoX(4)) ' Ancho
            showGrhH = Val(DatoX(5)) ' Alto
            showDecode = tCode.Text
        ElseIf (showGrhNF > 1) Then ' es anim
            ReDim showGrhFF(1 To showGrhNF) As Long
            For y = 1 To showGrhNF
                showGrhFF(y) = Val(DatoX(y))
            Next
            showGrhS = Val(DatoX(showGrhNF + 1))
            showDecode = tCode.Text
        Else ' error!
            imgGrh.Cls
            showDecode = tCode.Text
            Exit Sub
        End If
        ' pasamos al grh los datos
        Dim ST As String
        exGrhData(SelectedGrh).NumFrames = showGrhNF
        ST = exGrhData(SelectedGrh).NumFrames & "-"
        If exGrhData(SelectedGrh).NumFrames = 1 Then
            exGrhData(SelectedGrh).FileNum = Val(showGrhFile)
            ST = ST & exGrhData(SelectedGrh).FileNum & "-"
            exGrhData(SelectedGrh).sX = showGrhX
            ST = ST & exGrhData(SelectedGrh).sX & "-"
            exGrhData(SelectedGrh).sY = showGrhY
            ST = ST & exGrhData(SelectedGrh).sY & "-"
            exGrhData(SelectedGrh).pixelWidth = showGrhW
            ST = ST & exGrhData(SelectedGrh).pixelWidth & "-"
            exGrhData(SelectedGrh).pixelHeight = showGrhH
            ST = ST & exGrhData(SelectedGrh).pixelHeight & "-"
        ElseIf exGrhData(SelectedGrh).NumFrames > 1 Then
            ReDim exGrhData(SelectedGrh).Frames(1 To exGrhData(SelectedGrh).NumFrames)
            For y = 1 To showGrhNF
                exGrhData(SelectedGrh).Frames(y) = showGrhFF(y)
                ST = ST & exGrhData(SelectedGrh).Frames(y) & "-"
            Next
            exGrhData(SelectedGrh).Speed = showGrhS
            ST = ST & exGrhData(SelectedGrh).Speed
        Else
            MsgBox "Numero de frames invalido.", vbCritical
            Exit Sub
        End If
        ' guardar en ini
        MsgBox "El cambio se mantendrá mientras se encuentre usando el Explorador." & vbCrLf & "No se ha guardado el valor en " & GraphicsFile & ".", vbInformation + vbOKOnly
        'Dim cIni As New clsIniReader
        'Call cIni.Initialize(IniPath & GraphicsFile)
        'Call cIni.ChangeValue("Graphics", "Grh" & SelectedGrh, ST)
    End If
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
    showActualFrame = showActualFrame + 1
    If showActualFrame > showGrhNF Then showActualFrame = 1
    Call ShowGrhN(exGrhData(showGrhFF(showActualFrame)).FileNum, exGrhData(showGrhFF(showActualFrame)).sX, exGrhData(showGrhFF(showActualFrame)).sY, exGrhData(showGrhFF(showActualFrame)).pixelWidth, exGrhData(showGrhFF(showActualFrame)).pixelHeight)
    lGrhFrame.Caption = "GrhFrame: " & showGrhFF(showActualFrame)
End Sub

Private Sub cmdPausa_Click()
On Error Resume Next
    TimerFrame.Enabled = False
    cmdPlay.Enabled = True
    cmdPausa.Enabled = False
End Sub

Private Sub cmdPlay_Click()
On Error Resume Next
    TimerFrame.Enabled = True
    cmdPausa.Enabled = True
    cmdPlay.Enabled = False
End Sub

Private Sub cmdPrev_Click()
On Error Resume Next
    showActualFrame = showActualFrame - 1
    If showActualFrame <= 0 Then showActualFrame = showGrhNF
    Call ShowGrhN(exGrhData(showGrhFF(showActualFrame)).FileNum, exGrhData(showGrhFF(showActualFrame)).sX, exGrhData(showGrhFF(showActualFrame)).sY, exGrhData(showGrhFF(showActualFrame)).pixelWidth, exGrhData(showGrhFF(showActualFrame)).pixelHeight)
    lGrhFrame.Caption = "GrhFrame: " & showGrhFF(showActualFrame)
End Sub

Private Sub cmdRegresar_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdShow_Click()
On Error Resume Next
    Call ShowGrh
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Dim I As Long
    Dim k As Long
    
    frmExplorerGrh.lGrh.Clear
    frmExplorerGrh.lGrh.Redraw = False
    frmExplorerGrh.lGrh.Visible = False

    frmExplorerGrh.lGrh.AddColumn "ID", 1
    frmExplorerGrh.lGrh.AddColumn "Grafico", 3

    For I = 1 To exGrhCount
        frmExplorerGrh.lGrh.AddItem I
        k = frmExplorerGrh.lGrh.Rows - 1
        If exGrhData(I).NumFrames > 1 Then
            frmExplorerGrh.lGrh.CellText(k, 1) = "[ANIMACIÓN]"
        ElseIf exGrhData(I).FileNum <> 0 Then
            frmExplorerGrh.lGrh.CellText(k, 1) = exGrhData(I).FileNum
            If FileExist(DirGraphics & exGrhData(I).FileNum & ".png", vbArchive) = False And FileExist(DirGraphics & exGrhData(I).FileNum & ".bmp", vbArchive) = False Then
                frmExplorerGrh.lGrh.CellBackColor(k, 1) = vbRed
            'Else
                'frmExplorerGrh.lGrh.CellBackColor(k, 1) = vbWhite
            End If
        Else
            frmExplorerGrh.lGrh.CellText(k, 1) = "[LIBRE]"
        End If
    Next

    frmExplorerGrh.lGrh.Visible = True
    frmExplorerGrh.lGrh.Redraw = True
    frmExplorerGrh.lGrh.ColForceFit
    
    frmExplorerGrh.lInfo.AddColumn "Información", 1
    frmExplorerGrh.lInfo.AddColumn "Valor", 3
    
    lGrhFrame.Caption = vbNullString
    Transp = False
    ZoomImg = 1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    TimerFrame.Enabled = False
    Call ExExplorerGrh
    DIU.Show
    DIU.WindowState = 0
End Sub

Private Sub hT_Click()
On Error Resume Next
    Transp = Not Transp
    If Transp = True Then
        imgGrh.BackColor = vbMagenta
    Else
        imgGrh.BackColor = vbBlack
    End If
    Call ShowGrh
End Sub

Private Sub lGrh_Click()
On Error Resume Next
    Call CargarInfo
End Sub

Private Sub CargarInfo()
On Error Resume Next

    Dim I As Long
    Dim y As Long
    Dim k As Long
    Dim TempFrame As String
    
    I = frmExplorerGrh.lGrh.CellText(-1, 0)
    If I > 0 And I <= exGrhCount Then
        
        frmExplorerGrh.lInfo.Visible = False
        frmExplorerGrh.lInfo.Redraw = False
        frmExplorerGrh.lInfo.Clear
        TimerFrame.Enabled = False
        tCode.Text = vbNullString
        lGrhFrame.Caption = vbNullString
        SelectedGrh = I
        
        If (exGrhData(I).NumFrames = 1) Then ' no anim
            frmExplorerGrh.lInfo.AddItem "Nro.:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = I
            frmExplorerGrh.lInfo.AddItem "Archivo:"
            y = frmExplorerGrh.lInfo.Rows - 1
            If FileExist(DirGraphics & exGrhData(I).FileNum & ".bmp", vbArchive) Then
                TempFrame = DirGraphics & exGrhData(I).FileNum & ".bmp"
            ElseIf FileExist(DirGraphics & exGrhData(I).FileNum & ".png", vbArchive) Then
             TempFrame = DirGraphics & exGrhData(I).FileNum & ".png"
            End If
            If Len(TempFrame) > 39 Then
                TempFrame = "..." & Right(TempFrame, 35)
            End If
            frmExplorerGrh.lInfo.CellText(y, 1) = TempFrame
            frmExplorerGrh.lInfo.AddItem "Posición X:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = exGrhData(I).sX
            frmExplorerGrh.lInfo.AddItem "Posición Y:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = exGrhData(I).sY
            frmExplorerGrh.lInfo.AddItem "Ancho:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = exGrhData(I).pixelWidth
            frmExplorerGrh.lInfo.AddItem "Alto:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = exGrhData(I).pixelHeight
            tCode.Text = "Grh" & I & "=1-" & exGrhData(I).FileNum & "-" & exGrhData(I).sX & "-" & exGrhData(I).sY & "-" & exGrhData(I).pixelWidth & "-" & exGrhData(I).pixelHeight
        ElseIf (exGrhData(I).NumFrames > 1) Then  ' es anim
            frmExplorerGrh.lInfo.AddItem "Nro.:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = I
            frmExplorerGrh.lInfo.AddItem "NumFrames:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = exGrhData(I).NumFrames
            frmExplorerGrh.lInfo.AddItem "Velocidad:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = exGrhData(I).Speed
            TempFrame = vbNullString
            For k = 1 To exGrhData(I).NumFrames
                frmExplorerGrh.lInfo.AddItem "GrhFrame" & k & ":"
                y = frmExplorerGrh.lInfo.Rows - 1
                frmExplorerGrh.lInfo.CellText(y, 1) = exGrhData(I).Frames(k)
                TempFrame = TempFrame & "-" & exGrhData(I).Frames(k)
            Next
            tCode.Text = "Grh" & I & "=" & exGrhData(I).NumFrames & TempFrame & "-" & exGrhData(I).Speed
        Else ' libre
            frmExplorerGrh.lInfo.AddItem "Nro.:"
            y = frmExplorerGrh.lInfo.Rows - 1
            frmExplorerGrh.lInfo.CellText(y, 1) = I
            tCode.Text = "Grh" & I & "="
        End If
        
        frmExplorerGrh.lInfo.Visible = True
        frmExplorerGrh.lInfo.Redraw = True
        frmExplorerGrh.lInfo.ColForceFit
        DoEvents
        
        Call ShowGrh
        DoEvents
    End If
        
    
End Sub

Private Sub lGrh_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Call CargarInfo
End Sub

Private Sub ShowGrh()
On Error GoTo Errores
    ' Read tCode!
    Dim DatoX() As String
    If LenB(tCode.Text) <> 0 Then
        showActualFrame = 1
        TimerFrame.Enabled = False
        cmdPlay.Enabled = False
        cmdPausa.Enabled = False
        cmdNext.Enabled = False
        cmdPrev.Enabled = False
        If (showDecode <> tCode.Text) Then
            DatoX() = Split(tCode.Text, "-")
            Dim TT As Integer
            TT = Val(InStr(1, DatoX(0), "="))
            DatoX(0) = Right(DatoX(0), Len(DatoX(0)) - TT)
            showGrhNF = Val(DatoX(0)) ' Numero de Frames
            If (showGrhNF = 1) Then ' no anim
                showGrhFile = DatoX(1) ' Nombre de el BMP (no solo numero así se puede testear con chars)
                showGrhX = Val(DatoX(2)) ' X
                showGrhY = Val(DatoX(3)) ' Y
                showGrhW = Val(DatoX(4)) ' Ancho
                showGrhH = Val(DatoX(5)) ' Alto
                showDecode = tCode.Text
            ElseIf (showGrhNF > 1) Then ' es anim
                Dim y As Integer
                ReDim showGrhFF(1 To showGrhNF) As Long
                For y = 1 To showGrhNF
                    showGrhFF(y) = Val(DatoX(y))
                Next
                showGrhS = Val(DatoX(showGrhNF + 1))
                showDecode = tCode.Text
            Else ' error!
                imgGrh.Cls
                showDecode = tCode.Text
                Exit Sub
            End If
        End If
                
        If (showGrhNF = 1) Then ' no anim
            Call ShowGrhN(showGrhFile, showGrhX, showGrhY, showGrhW, showGrhH)
        ElseIf (showGrhNF > 1) Then '  anim
            Call ShowGrhN(exGrhData(showGrhFF(showActualFrame)).FileNum, exGrhData(showGrhFF(showActualFrame)).sX, exGrhData(showGrhFF(showActualFrame)).sY, exGrhData(showGrhFF(showActualFrame)).pixelWidth, exGrhData(showGrhFF(showActualFrame)).pixelHeight)
            lGrhFrame.Caption = "GrhFrame: " & showGrhFF(showActualFrame)
            If (showGrhS > 0) Then
                TimerFrame.Interval = showGrhS / 2
                TimerFrame.Enabled = True
                'cmdPlay.Enabled = True
                cmdPausa.Enabled = True
                cmdNext.Enabled = True
                cmdPrev.Enabled = True
            End If
        End If

    Else
        imgGrh.Cls
    End If
    
    Exit Sub
Errores:
    ' Fallo!
    imgGrh.Cls

End Sub

Private Sub ShowGrhN(ByVal sFile As String, ByVal sX As Integer, ByVal sY As Integer, ByVal tW As Integer, ByVal tH As Integer)
On Error Resume Next
    'imgGrh.Visible = False
    imgGrh.Cls
    Dim xT As Integer
    Dim yT As Integer
    Dim lTransp As Long
    If FileExist(DirGraphics & sFile & ".bmp", vbArchive) = True Then
        imgGrh.PaintPicture LoadPicture(DirGraphics & sFile & ".bmp"), 1, 1, tW * ZoomImg, tH * ZoomImg, sX, sY, tW, tH
        If Transp = True Then
            lTransp = imgGrh.Point(1, 1)
            For xT = 0 To imgGrh.ScaleWidth - 1
                For yT = 0 To imgGrh.ScaleWidth - 1
                    If imgGrh.Point(xT, yT) = lTransp Then
                        imgGrh.PSet (xT, yT), imgGrh.BackColor
                    End If
                Next
            Next
        End If
    ElseIf FileExist(DirGraphics & sFile & ".png", vbArchive) = True Then
        'imgGrh.PaintPicture LoadPicture(DirGraphics & sFile & ".png"), 1, 1, tW * ZoomImg, tH * ZoomImg, sX, sY, tW, tH
        imgGrh.PaintPicture StdPictureEx.LoadPicture(DirGraphics & sFile & ".png"), 1, 1, tW * ZoomImg, tH * ZoomImg, sX, sY, tW, tH
        If Transp = True Then
            lTransp = imgGrh.Point(1, 1)
            For xT = 0 To imgGrh.ScaleWidth - 1
                For yT = 0 To imgGrh.ScaleWidth - 1
                    If imgGrh.Point(xT, yT) = lTransp Then
                        imgGrh.PSet (xT, yT), imgGrh.BackColor
                    End If
                Next
            Next
        End If
    End If
    imgGrh.AutoRedraw = True
    imgGrh.Visible = True
End Sub


Private Sub lZoom_Click()
On Error Resume Next
    ZoomImg = Val(lZoom.List(lZoom.ListIndex))
    Call ShowGrh
End Sub

Private Sub TimerFrame_Timer()
On Error Resume Next
    If (showGrhNF > 1) Then
        showActualFrame = showActualFrame + 1
        If showActualFrame > showGrhNF Then showActualFrame = 1
        lGrhFrame.Caption = "GrhFrame: " & showGrhFF(showActualFrame)
        Call ShowGrhN(exGrhData(showGrhFF(showActualFrame)).FileNum, exGrhData(showGrhFF(showActualFrame)).sX, exGrhData(showGrhFF(showActualFrame)).sY, exGrhData(showGrhFF(showActualFrame)).pixelWidth, exGrhData(showGrhFF(showActualFrame)).pixelHeight)
        DoEvents
    End If
End Sub
