VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indexador Alkon"
   ClientHeight    =   9090
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   11115
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   606
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   741
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9720
      TabIndex        =   22
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame grhFrame 
      Caption         =   "Grh"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   8160
      Width           =   8295
      Begin VB.TextBox bmpTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhWidthTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhHeightTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhYTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhXTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bmp:"
         Height          =   195
         Left            =   6960
         TabIndex        =   20
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alto:"
         Height          =   195
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.CheckBox grhOnly 
      Caption         =   "Mostrar solamente el Grh"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Timer animation 
      Enabled         =   0   'False
      Left            =   240
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zoom"
      Height          =   735
      Left            =   8640
      TabIndex        =   5
      Top             =   8160
      Width           =   2055
      Begin VB.CommandButton ZoomOut 
         Caption         =   "-"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton ZoomIn 
         Caption         =   "+"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox ZoomTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         Top             =   285
         Width           =   120
      End
   End
   Begin VB.HScrollBar picScrollH 
      Height          =   255
      LargeChange     =   10
      Left            =   3000
      TabIndex        =   4
      Top             =   7800
      Width           =   7695
   End
   Begin VB.VScrollBar picScrollV 
      Height          =   7695
      LargeChange     =   10
      Left            =   10680
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.ListBox fileList 
      Height          =   2205
      ItemData        =   "frmMain.frx":08CA
      Left            =   120
      List            =   "frmMain.frx":08CC
      TabIndex        =   2
      Top             =   5760
      Width           =   2655
   End
   Begin VB.ListBox grhList 
      Height          =   5130
      ItemData        =   "frmMain.frx":08CE
      Left            =   120
      List            =   "frmMain.frx":08D0
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox previewer 
      AutoRedraw      =   -1  'True
      Height          =   7680
      Left            =   3000
      ScaleHeight     =   508
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   0
      Top             =   120
      Width           =   7680
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu SaveMnu 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveOldMnu 
         Caption         =   "Save in &old format"
      End
      Begin VB.Menu SaveNewMnu 
         Caption         =   "Save in &new format"
      End
   End
   Begin VB.Menu GrhMnu 
      Caption         =   "&Grh"
      Begin VB.Menu AddGrhMnu 
         Caption         =   "&Agregar Grh..."
         Shortcut        =   ^N
      End
      Begin VB.Menu RemoveGrhMnu 
         Caption         =   "&Remover Grh"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Default zoom, 100%
Private Const DEFAULT_ZOOM As Integer = 100

''
' Maximum zoom possible, 10 times bigger.
Private Const MAX_ZOOM As Integer = DEFAULT_ZOOM * 10

''
' Minimum zoom possible, 10 times smaller.
Private Const MIN_ZOOM As Integer = DEFAULT_ZOOM / 10

''
' Step by which zoom is altered.
Private Const ZOOM_STEP As Integer = 10

''
' Means no grh is being rendered.
Private Const NO_GRH As Long = -1


''
' Defines the different points of the selection box that are being edited.
'
' @param    sbpeNone            No coord is being modified.
' @param    sbpeStartX          Starting x coord is being modified.
' @param    sbpeStartY          Starting y coord is being modified.
' @param    sbpeEndX            Ending x coord is being modified.
' @param    sbpeEndY            Ending y coord is being modified.
' @param    sbpeStartXStartY    Starting x coord and starting y coord are being modified.
' @param    sbpeEndXEndY        Ending x coord and ending y coord are being modified.
' @param    sbpeStartXEndY      Starting x coord and ending y coord are being modified.
' @param    sbpeEndXStartY      Ending x coord and starting y coord are being modified.

Private Enum eSelectionBoxPointEdition
    sbpeNone
    sbpeStartX
    sbpeStartY
    sbpeEndX
    sbpeEndY
    sbpeStartXStartY
    sbpeEndXEndY
    sbpeStartXEndY
    sbpeEndXStartY
End Enum

''
' The current zoom, 1 == 100%
Private zoom As Single

''
'Currently loaded picture. Used to render avoiding to reload everytime zoom or scroll happens.
Private currentPic As StdPicture

''
' X coord where a selection started.
Private selectionAreaStartX As Single

''
' Y coord where a selection started.
Private selectionAreaStartY As Single

''
' X coord where a selection ended.
Private selectionAreaEndX As Single

''
' Y coord where a selection ended.
Private selectionAreaEndY As Single

''
' Cord currently being edited.
Private editionCoord As eSelectionBoxPointEdition

''
' The grh currently being displayed
Private currentGrh As Long

''
' The current frame of the grh being displayed
Private currentFrame As Long

''
' Flag used to ignore calls to RenderSelectionBox.
Private ignoreSelectionBoxRender As Boolean

''
' Flag used to ignore update events to grh' data textboxes.
Private ignoreGrhTextUpdate As Boolean



Private Sub AddGrhMnu_Click()
    Call frmAddGrh.Show(vbModal, Me)
End Sub

Private Sub animation_Timer()
    Dim path As String
    
    'If an animated grh is chosen, animate!
    If currentGrh <> NO_GRH Then
        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!
            currentFrame = currentFrame + 1
            
            If currentFrame > GrhData(currentGrh).NumFrames Then
                currentFrame = 1
            End If
            
            'Load new bitmap
            If Right$(Config.BmpPath, 1) <> "\" Then
                path = Config.BmpPath & "\" & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
            Else
                path = Config.BmpPath & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
            End If
            
            'Prevent memory leaks
            Set currentPic = Nothing
            Set currentPic = LoadPicture(path)
            
            Call RedrawPicture(currentGrh, currentFrame)
        End If
    End If
End Sub

Private Sub bmpTxt_Change()
    Dim path As String
    
    'Prevent non numeric characters
    If Not IsNumeric(bmpTxt.Text) Then
        bmpTxt.Text = Val(bmpTxt.Text)
    End If
    
    'Prevent overflow
    If Val(bmpTxt.Text) > &H7FFFFFFF Then
        bmpTxt.Text = &H7FFFFFFF
    End If
    
    'Prevent underrflow
    If Val(bmpTxt.Text) < 1 Then
        bmpTxt.Text = "1"
    End If
    
    
    If Right$(Config.BmpPath, 1) <> "\" Then
        path = Config.BmpPath & "\" & bmpTxt.Text & ".bmp"
    Else
        path = Config.BmpPath & bmpTxt.Text & ".bmp"
    End If
    
    'If file exists, load it
    If FileExists(path) And currentGrh <> NO_GRH Then
        GrhData(currentGrh).FileNum = CLng(bmpTxt.Text)
        
        'Prevent memory leaks
        Set currentPic = Nothing
        Set currentPic = LoadPicture(path)
        
        'Set scrollers!
        Call SetScrollers
        
        'Display the grh!
        Call RedrawPicture(currentGrh, currentFrame)
        
        'Show selection box (if needed)
        ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
        Call RenderSelectionBox
    End If
End Sub

Private Sub Command1_Click()
FrmAnimaciones.Visible = True
End Sub

Private Sub fileList_Click()
    Dim path As String
    
    If Right$(Config.BmpPath, 1) <> "\" Then
        path = Config.BmpPath & "\" & fileList.Text & ".bmp"
    Else
        path = Config.BmpPath & fileList.Text & ".bmp"
    End If
    
    'Prevent memory leaks
    Set currentPic = Nothing
    Set currentPic = LoadPicture(path)
    
    'Reset selection box
    selectionAreaEndX = 0
    selectionAreaEndY = 0
    selectionAreaStartX = 0
    selectionAreaStartY = 0
    
    'Set scrollers!
    Call SetScrollers
    
    currentGrh = NO_GRH
    
    bmpTxt.Text = fileList.Text
    
    'Draw!
    Call RedrawPicture(NO_GRH, 0)
    
    ignoreSelectionBoxRender = False
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim fileName As String
    Dim path As String
    
    If Not LoadConfig() Then
        'Show config form
        Call frmConfig.Show(vbModal, Me)
    End If
    
    'Load Grhs!
    Call LoadGrhData(Config.InitPath)
    
    'Fill the lists
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            If GrhData(i).NumFrames = 1 Then
                Call grhList.AddItem(CStr(i))
            Else
                Call grhList.AddItem(CStr(i) & " (ANIMACIÓN)")
            End If
        End If
    Next i
    
    'Set up bmp search path
    If Right$(Config.BmpPath, 1) <> "\" Then
        path = Config.BmpPath & "\*.bmp"
    Else
        path = Config.BmpPath & "*.bmp"
    End If
    
    fileName = Dir$(path, vbArchive)
    
    While fileName <> ""
        'Add it!
        fileName = Left$(fileName, InStr(1, fileName, ".") - 1)
        
        'Make usre it's numeric
        If IsNumeric(fileName) Then
            Call fileList.AddItem(fileName)
        End If
        
        fileName = Dir()
    Wend
    
    'Set default zoom value
    ZoomTxt.Text = DEFAULT_ZOOM
    
    editionCoord = sbpeNone
    
    currentGrh = NO_GRH
    
    'By default update events are not ignored
    ignoreGrhTextUpdate = False
    
    'Show first grh by default
    If grhList.ListCount > 0 Then
        grhList.ListIndex = 0
    ElseIf fileList.ListCount > 0 Then
        fileList.ListIndex = 0
    End If
End Sub

Private Sub grhHeightTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhHeightTxt.Text) Then
        grhHeightTxt.Text = Val(grhHeightTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhHeightTxt.Text) > &H7FFF Then
        grhHeightTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhHeightTxt.Text) > previewer.ScaleY(currentPic.Height) - Val(grhYTxt.Text) Then
        grhHeightTxt.Text = Round(previewer.ScaleY(currentPic.Height) - Val(grhYTxt.Text))
    End If
    
    'Prevent negative values
    If CInt(grhHeightTxt.Text) < 0 Then
        grhHeightTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).pixelHeight = CInt(grhHeightTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaEndY = selectionAreaStartY + Val(grhHeightTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhList_Click()
    Dim path As String
    
    ' Set current grh and reset frame
    currentGrh = Val(grhList.Text)
    currentFrame = 1
    
    'Should grh controls be enabled?
    Call SetGrhControlsEnabled(grhList.Text = CStr(currentGrh))
    
    If Right$(Config.BmpPath, 1) <> "\" Then
        path = Config.BmpPath & "\" & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
    Else
        path = Config.BmpPath & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
    End If
    
    'Prevent memory leaks
    Set currentPic = Nothing
    Set currentPic = LoadPicture(path)
    
    'Enable animations if necessary
    If GrhData(currentGrh).NumFrames > 1 Then
        animation.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
        animation.Enabled = True
        
        grhOnly.value = vbChecked
        grhOnly.Enabled = False
    Else
        animation.Enabled = False
        
        If Not grhOnly.Enabled Then
            grhOnly.Enabled = True
            
            grhOnly.value = vbChecked
        ElseIf grhOnly.value = vbUnchecked Then
            'Set selection box!
            Call SelectGrhArea(currentGrh)
        End If
        
        'Show bmp
        bmpTxt.Text = GrhData(currentGrh).FileNum
        
        'Filelist will reset the currentGrh, restore it!
        currentGrh = Val(grhList.Text)
        
        'Set selection box!
        Call SelectGrhArea(currentGrh)
        
        'Display grh info
        grhXTxt.Text = GrhData(currentGrh).sX
        grhYTxt.Text = GrhData(currentGrh).sY
        grhWidthTxt.Text = GrhData(currentGrh).pixelWidth
        grhHeightTxt.Text = GrhData(currentGrh).pixelHeight
    End If
    
    'Set scrollers!
    Call SetScrollers
    
    'Display the grh!
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box (if needed)
    ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
    Call RenderSelectionBox
End Sub

Private Sub grhOnly_Click()
    If currentGrh = NO_GRH Then Exit Sub
    
    Call RedrawPicture(currentGrh, currentFrame)
    
    ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
    
    'Set selection box!
    Call SelectGrhArea(currentGrh)
    
    Call RenderSelectionBox
End Sub

Private Sub grhWidthTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhWidthTxt.Text) Then
        grhWidthTxt.Text = Val(grhWidthTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhWidthTxt.Text) > &H7FFF Then
        grhWidthTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhWidthTxt.Text) > previewer.ScaleX(currentPic.Width) - Val(grhXTxt.Text) Then
        grhWidthTxt.Text = Round(previewer.ScaleX(currentPic.Width) - Val(grhXTxt.Text))
    End If
    
    'Prevent negative values
    If CInt(grhWidthTxt.Text) < 0 Then
        grhWidthTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).pixelWidth = CInt(grhWidthTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaEndX = selectionAreaStartX + CInt(grhWidthTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhXTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhXTxt.Text) Then
        grhXTxt.Text = Val(grhXTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhXTxt.Text) > &H7FFF Then
        grhXTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhXTxt.Text) > previewer.ScaleX(currentPic.Width) Then
        grhXTxt.Text = Round(previewer.ScaleX(currentPic.Width))
    End If
    
    'Prevent negative values
    If CInt(grhXTxt.Text) < 0 Then
        grhXTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).sX = CInt(grhXTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaStartX = CInt(grhXTxt.Text)
    selectionAreaEndX = selectionAreaStartX + Val(grhWidthTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhYTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhYTxt.Text) Then
        grhYTxt.Text = Val(grhYTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhYTxt.Text) > &H7FFF Then
        grhYTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhYTxt.Text) > previewer.ScaleY(currentPic.Height) Then
        grhYTxt.Text = Round(previewer.ScaleY(currentPic.Height))
    End If
    
    'Trim height to prevent invalid values
    If CInt(grhYTxt.Text) + Val(grhHeightTxt.Text) > previewer.ScaleY(currentPic.Height) Then
        grhHeightTxt.Text = Round(previewer.ScaleY(currentPic.Height)) - CInt(grhYTxt.Text)
    End If
    
    'Prevent negative values
    If CInt(grhYTxt.Text) < 0 Then
        grhYTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).sY = CInt(grhYTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaStartY = Val(grhYTxt.Text)
    selectionAreaEndY = selectionAreaStartY + Val(grhHeightTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub picScrollH_Change()
    'Redraw
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

Private Sub picScrollV_Change()
    'Redraw
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

Private Sub previewer_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'If no picture is loaded, there is nothing to be done
    If currentPic Is Nothing Then Exit Sub
    
    If Button And vbLeftButton Then
        If currentGrh <> NO_GRH And grhOnly.value = vbChecked Then Exit Sub
        
        Select Case Me.MousePointer
            Case vbDefault
                'A new box is being created, we are fixing start x-y coord and moving end x-y
                editionCoord = sbpeEndXEndY
                
                'Make sure selection box doesn't go beyond bmp
                If ViewPortToBmpPosX(x) > previewer.ScaleX(currentPic.Width) Then
                    x = BmpToViewPortPosX(previewer.ScaleX(currentPic.Width))
                End If
                
                If ViewPortToBmpPosY(Y) > previewer.ScaleY(currentPic.Height) Then
                    Y = BmpToViewPortPosY(previewer.ScaleY(currentPic.Height))
                End If
                
                
                'Convert mouse pos to pixel pos of origin
                selectionAreaStartX = ViewPortToBmpPosX(x)
                selectionAreaStartY = ViewPortToBmpPosY(Y)
                
                'Reset end area, we are starting a new rectangle
                selectionAreaEndX = selectionAreaStartX
                selectionAreaEndY = selectionAreaStartY
                
                'Show selection box!
                Call RenderSelectionBox
            
            Case vbSizeNS
                If Abs(selectionAreaStartY - ViewPortToBmpPosY(Y)) < 2 Then
                    editionCoord = sbpeStartY
                ElseIf Abs(selectionAreaEndY - ViewPortToBmpPosY(Y)) < 2 Then
                    editionCoord = sbpeEndY
                End If
            
            Case vbSizeWE
                If Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 2 Then
                    editionCoord = sbpeStartX
                ElseIf Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 2 Then
                    editionCoord = sbpeEndX
                End If
            
            Case vbSizeNWSE
                If (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(Y)) < 5) Then
                    editionCoord = sbpeStartXStartY
                ElseIf (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(Y)) < 5) Then
                    editionCoord = sbpeEndXEndY
                End If
            
            Case vbSizeNESW
                If (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(Y)) < 5) Then
                    editionCoord = sbpeStartXEndY
                ElseIf (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(Y)) < 5) Then
                    editionCoord = sbpeEndXStartY
                End If
        End Select
    End If
End Sub

Private Sub previewer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button And vbLeftButton Then
        If currentGrh <> NO_GRH And grhOnly.value = vbChecked Then Exit Sub
        
        'If we got past the border, we scroll!!
        If x < 0 Then
            x = 0
            
            If picScrollH.value > 0 And picScrollH.Enabled Then
                picScrollH.value = picScrollH.value - 1
            End If
        ElseIf x > previewer.Width Then
            x = previewer.Width
            
            If picScrollH.value < picScrollH.max And picScrollH.Enabled Then
                picScrollH.value = picScrollH.value + 1
            End If
        End If
        
        If Y < 0 Then
            Y = 0
            
            If picScrollV.value > 0 And picScrollV.Enabled Then
                picScrollV.value = picScrollV.value - 1
            End If
        ElseIf Y > previewer.Height Then
            Y = previewer.Height
            
            If picScrollV.value < picScrollV.max And picScrollV.Enabled Then
                picScrollV.value = picScrollV.value + 1
            End If
        End If
        
        
        'Make sure selection box doesn't go beyond bmp
        If ViewPortToBmpPosX(x) > previewer.ScaleX(currentPic.Width) Then
            x = BmpToViewPortPosX(previewer.ScaleX(currentPic.Width))
        End If
        
        If ViewPortToBmpPosY(Y) > previewer.ScaleY(currentPic.Height) Then
            Y = BmpToViewPortPosY(previewer.ScaleY(currentPic.Height))
        End If
        
        
        'Update coords
        Call UpdateSelectionBox(x, Y)
        
        'Show selection box!
        Call RenderSelectionBox
    ElseIf Not ignoreSelectionBoxRender And selectionAreaStartX <> selectionAreaEndX And selectionAreaStartY <> selectionAreaEndY Then
        'Allow the user to resize the selection box!
        
        'Set mouse pointer appropiately
        If (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(Y)) < 5) _
                Or (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(Y)) < 5) Then
            Me.MousePointer = vbSizeNWSE
        
        ElseIf (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(Y)) < 5) _
                Or (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(Y)) < 5) Then
            Me.MousePointer = vbSizeNESW
        
        ElseIf (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 2 Or Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 2) _
                And ViewPortToBmpPosY(Y) > selectionAreaStartY And ViewPortToBmpPosY(Y) < selectionAreaEndY Then
            Me.MousePointer = vbSizeWE
        
        ElseIf (Abs(selectionAreaStartY - ViewPortToBmpPosY(Y)) < 2 Or Abs(selectionAreaEndY - ViewPortToBmpPosY(Y)) < 2) _
                And ViewPortToBmpPosX(x) > selectionAreaStartX And ViewPortToBmpPosX(x) < selectionAreaEndX Then
            Me.MousePointer = vbSizeNS
        
        Else
            Me.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub previewer_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button And vbLeftButton Then
        If currentGrh <> NO_GRH And grhOnly.value = vbChecked Then Exit Sub
        
        'Make sure selection box doesn't go beyond bmp
        If ViewPortToBmpPosX(x) > previewer.ScaleX(currentPic.Width) Then
            x = BmpToViewPortPosX(previewer.ScaleX(currentPic.Width))
        End If
        
        If ViewPortToBmpPosY(Y) > previewer.ScaleY(currentPic.Height) Then
            Y = BmpToViewPortPosY(previewer.ScaleY(currentPic.Height))
        End If
        
        'Update selection box
        Call UpdateSelectionBox(x, Y)
        
        'Show selection box!
        Call RenderSelectionBox
    End If
End Sub

Private Sub RemoveGrhMnu_Click()
    Dim i As Long
    
    If currentGrh = NO_GRH Then
        MsgBox "There is no grh selected."
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete the grh " & currentGrh & "?" & vbCrLf & "This change can't be undone.", vbOKCancel) = vbOK Then
        'Reset it
        With GrhData(currentGrh)
            .FileNum = 0
            ReDim .Frames(0)
            .NumFrames = 0
            .pixelHeight = 0
            .pixelWidth = 0
            .Speed = 0
            .sX = 0
            .sY = 0
            .TileHeight = 0
            .TileWidth = 0
        End With
        
        'Remove it
        For i = 0 To grhList.ListCount - 1
            If Val(grhList.List(i)) = currentGrh Then
                grhList.RemoveItem (i)
                Exit For
            End If
        Next i
        
        'Select next grh
        If i < grhList.ListCount Then
            grhList.ListIndex = i
        Else
            grhList.ListIndex = grhList.ListCount - 1
        End If
    End If
End Sub

Private Sub SaveMnu_Click()
    'Detect the original file format and save it
    If Grh.FileVersion = -1 Then
        If Not Grh.SaveGrhDataOld(Config.InitPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk, or you are using grh indexes above 32767, which are only supported in the new file format.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    Else
        If Not Grh.SaveGrhDataNew(Config.InitPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    End If
End Sub

Private Sub SaveNewMnu_Click()
    If Not Grh.SaveGrhDataNew(Config.InitPath) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveOldMnu_Click()
    If MsgBox("The old file format speed system is FPS based, animation's speed may be altered. Do you want to proceed?", vbYesNo) = vbYes Then
        If Not Grh.SaveGrhDataOld(Config.InitPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk, or you are using grh indexes above 32767, which are only supported in the new file format.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    End If
End Sub

Private Sub ZoomIn_Click()
    ZoomTxt.Text = Val(ZoomTxt.Text) + ZOOM_STEP
End Sub

Private Sub ZoomOut_Click()
    ZoomTxt.Text = Val(ZoomTxt.Text) - ZOOM_STEP
End Sub

Private Sub ZoomTxt_Change()
    'Validate
    If Not IsNumeric(ZoomTxt.Text) Then
        ZoomTxt.Text = DEFAULT_ZOOM
        Exit Sub
    End If
    
    If Val(ZoomTxt.Text) > MAX_ZOOM Then
        ZoomTxt.Text = MAX_ZOOM
        Exit Sub
    End If
    
    If Val(ZoomTxt.Text) < MIN_ZOOM Then
        ZoomTxt.Text = MIN_ZOOM
        Exit Sub
    End If
    
    'Recompute zoom
    zoom = CSng(ZoomTxt.Text) / DEFAULT_ZOOM
    
    
    'Reset scrollbars
    Call SetScrollers
    
    'Redraw
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

''
' Sets the scrollers' properties appropiately for the current picture loaded, zoom and value.

Private Sub SetScrollers()
    Dim oldMax As Integer
    
    If currentPic Is Nothing Then
        picScrollH.Enabled = False
        picScrollV.Enabled = False
        Exit Sub
    End If
    
    'Set up scrollers
    If previewer.Width < previewer.ScaleX(currentPic.Width) * zoom Then
        oldMax = IIf(picScrollH.max > 0, picScrollH.max, 1)
        
        picScrollH.max = previewer.ScaleX(currentPic.Width) - previewer.Width / zoom
        picScrollH.value = picScrollH.value * picScrollH.max / oldMax
        picScrollH.Enabled = True
    Else
        picScrollH.value = 0
        picScrollH.Enabled = False
    End If
    
    If previewer.Height < previewer.ScaleY(currentPic.Height) * zoom Then
        oldMax = IIf(picScrollV.max > 0, picScrollV.max, 1)
        
        picScrollV.max = previewer.ScaleX(currentPic.Height) - previewer.Height / zoom
        picScrollV.value = picScrollV.value * picScrollV.max / oldMax
        picScrollV.Enabled = True
    Else
        picScrollV.value = 0
        picScrollV.Enabled = False
    End If
End Sub

''
' Renders the last laoded picture.
'
' @param    grh     The grh to be rendered within the loaded picture. Can be @code NO_GRH
' @param    frame   The frame of the grh to be rendered. Only important if grh is not @code NO_GRH

Private Sub RedrawPicture(ByVal Grh As Long, ByVal Frame As Long)
    If currentPic Is Nothing Then Exit Sub
    
    'Clear picturebox
    Set previewer.Picture = Nothing
    previewer.Picture = LoadPicture("")
    
    'Render!
    If Grh <> NO_GRH And grhOnly.value = vbChecked Then
        'Transform grh to actual frame grh.
        Grh = GrhData(Grh).Frames(Frame)
        
        Call previewer.PaintPicture(currentPic, -picScrollH.value * zoom, -picScrollV.value * zoom, _
                                    GrhData(Grh).pixelWidth * zoom, _
                                    GrhData(Grh).pixelHeight * zoom, _
                                    GrhData(Grh).sX, GrhData(Grh).sY, _
                                    GrhData(Grh).pixelWidth, GrhData(Grh).pixelHeight)
    Else
        Call previewer.PaintPicture(currentPic, -picScrollH.value * zoom, -picScrollV.value * zoom, _
                                    previewer.ScaleX(currentPic.Width) * zoom, _
                                    previewer.ScaleY(currentPic.Height) * zoom)
    End If
End Sub

''
' Renders the selection box.

Private Sub RenderSelectionBox()
    Dim startX As Long
    Dim startY As Long
    Dim endX As Long
    Dim endY As Long
    
    If ignoreSelectionBoxRender Then Exit Sub
    
    'Transform origin coord to those in the picturebox
    startX = BmpToViewPortPosX(selectionAreaStartX)
    startY = BmpToViewPortPosY(selectionAreaStartY)
    
    'Transform end coord to those in the picturebox
    endX = BmpToViewPortPosX(selectionAreaEndX)
    endY = BmpToViewPortPosY(selectionAreaEndY)
    
    previewer.AutoRedraw = False
    previewer.Cls
    previewer.Line (startX, startY)-(endX, endY), vbRed, B
    previewer.AutoRedraw = True
End Sub

''
' Converts a bmp absolute pixel pos in the x axis to the picturebox's view area coord.
'
' @param    x   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function BmpToViewPortPosX(ByVal x As Long) As Long
    BmpToViewPortPosX = (x - picScrollH.value) * zoom
End Function

''
' Converts a bmp absolute pixel pos in the y axis to the picturebox's view area coord.
'
' @param    y   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function BmpToViewPortPosY(ByVal Y As Long) As Long
    BmpToViewPortPosY = (Y - picScrollV.value) * zoom
End Function

''
' Converts a picturebox's view area pos in the x axis to the bmp absolute pixel coord.
'
' @param    x   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function ViewPortToBmpPosX(ByVal x As Long) As Long
    ViewPortToBmpPosX = picScrollH.value + Fix(x / zoom)
End Function

''
' Converts a picturebox's view area pos in the y axis to the bmp absolute pixel coord.
'
' @param    y   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function ViewPortToBmpPosY(ByVal Y As Long) As Long
    ViewPortToBmpPosY = picScrollV.value + Fix(Y / zoom)
End Function

''
' Updates the appropiate selection box coords according to the current value of @code editionCoord.
'
' @param    x   The mouse pos in the x coord within the previewer.
' @param    y   The mouse pos in the y coord within the previewer.

Private Sub UpdateSelectionBox(ByVal x As Long, ByVal Y As Long)
    Dim tmp As Long
    
    Select Case editionCoord
        Case sbpeNone
            'Convert mouse pos to pixel pos of end
            selectionAreaEndX = ViewPortToBmpPosX(x)
            selectionAreaEndY = ViewPortToBmpPosY(Y)
        
        Case sbpeStartX
            selectionAreaStartX = ViewPortToBmpPosX(x)
        
        Case sbpeStartY
            selectionAreaStartY = ViewPortToBmpPosY(Y)
        
        Case sbpeEndX
            selectionAreaEndX = ViewPortToBmpPosX(x)
        
        Case sbpeEndY
            selectionAreaEndY = ViewPortToBmpPosY(Y)
        
        Case sbpeStartXStartY
            selectionAreaStartX = ViewPortToBmpPosX(x)
            selectionAreaStartY = ViewPortToBmpPosY(Y)
        
        Case sbpeEndXEndY
            selectionAreaEndX = ViewPortToBmpPosX(x)
            selectionAreaEndY = ViewPortToBmpPosY(Y)
        
        Case sbpeStartXEndY
            selectionAreaStartX = ViewPortToBmpPosX(x)
            selectionAreaEndY = ViewPortToBmpPosY(Y)
        
        Case sbpeEndXStartY
            selectionAreaEndX = ViewPortToBmpPosX(x)
            selectionAreaStartY = ViewPortToBmpPosY(Y)
    End Select
    
    'Invert coordinates if needed to prevent pointer from going crazy on corners.
    If selectionAreaStartX > selectionAreaEndX Then
        tmp = selectionAreaEndX
        selectionAreaEndX = selectionAreaStartX
        selectionAreaStartX = tmp
        
        'Invert edition coord accordingly.
        Select Case editionCoord
            Case sbpeEndX
                editionCoord = sbpeStartX
            
            Case sbpeEndXEndY
                editionCoord = sbpeStartXEndY
            
            Case sbpeEndXStartY
                editionCoord = sbpeStartXStartY
            
            Case sbpeStartX
                editionCoord = sbpeEndX
            
            Case sbpeStartXEndY
                editionCoord = sbpeEndXEndY
            
            Case sbpeStartXStartY
                editionCoord = sbpeEndXStartY
        End Select
    End If
    
    If selectionAreaStartY > selectionAreaEndY Then
        tmp = selectionAreaEndY
        selectionAreaEndY = selectionAreaStartY
        selectionAreaStartY = tmp
        
        'Invert edition coord accordingly.
        Select Case editionCoord
            Case sbpeEndY
                editionCoord = sbpeStartY
            
            Case sbpeEndXEndY
                editionCoord = sbpeEndXStartY
            
            Case sbpeEndXStartY
                editionCoord = sbpeEndXEndY
            
            Case sbpeStartY
                editionCoord = sbpeEndY
            
            Case sbpeStartXEndY
                editionCoord = sbpeStartXStartY
            
            Case sbpeStartXStartY
                editionCoord = sbpeStartXEndY
        End Select
    End If
    
    'Display data at the bottom
    ignoreGrhTextUpdate = True
    
    grhHeightTxt.Text = selectionAreaEndY - selectionAreaStartY
    grhWidthTxt.Text = selectionAreaEndX - selectionAreaStartX
    grhXTxt.Text = selectionAreaStartX
    grhYTxt.Text = selectionAreaStartY
    
    ignoreGrhTextUpdate = False
End Sub

''
' Sets up the selection area around the given grh within it's bmp.
'
' @param    grh     The grh to be selected.

Private Sub SelectGrhArea(ByVal Grh As Long)
    selectionAreaStartX = GrhData(Grh).sX
    selectionAreaStartY = GrhData(Grh).sY
    selectionAreaEndX = selectionAreaStartX + GrhData(Grh).pixelWidth
    selectionAreaEndY = selectionAreaStartY + GrhData(Grh).pixelHeight
End Sub

''
'Enables / disables the grh controls (those within the grhFrame control).
'
' @param    enable  True if controls should be enabled, False otherwise.

Private Sub SetGrhControlsEnabled(ByVal enable As Boolean)
    Dim i As Long
    
    For i = 0 To frmMain.Controls.Count - 1
        If Not TypeOf frmMain.Controls(i) Is Timer And Not TypeOf frmMain.Controls(i) Is Menu Then
            If frmMain.Controls(i).Container Is grhFrame Then
                frmMain.Controls(i).Enabled = enable
            End If
        End If
    Next i
    
    grhFrame.Enabled = enable
End Sub
