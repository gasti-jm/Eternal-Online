Attribute VB_Name = "modMapColor"
Option Explicit

Private lMapColor() As Long
Private lMapColorCount As Long

Public Sub LoadMapColor()
  
    'If FileExist(App.Path & "\GrhIndex\MapColor.bin", vbArchive) = False Then
    '    MsgBox "Falta el archivo 'MapColor.bin'", vbCritical
    '    End
    'End If
    
    Dim iT As Long
    Dim Handle As Long
    Handle = FreeFile
    
    Open App.path & "\small_maps.bin" For Binary Access Read As Handle
    Seek Handle, 1
    Get Handle, , lMapColorCount
    
    ReDim lMapColor(1 To lMapColorCount) As Long
    
    For iT = 1 To lMapColorCount
        Get Handle, , lMapColor(iT)
    Next
    
    Close Handle
    
    'If grh_count <> lMapColorCount Then
    '    MsgBox "ALERTA: El archivo 'MapColor.bin' se encuentra desactualizado." & vbCrLf & "Graficos Indexados: " & grh_count & " <> MapColor Indexados: " & lMapColorCount, vbExclamation
    'End If
    
End Sub


Public Sub DrawMiniMap(Optional ByVal NPCs As Boolean = True)
    frmMain.picRadar.Cls
    If lMapColorCount = 0 Then Exit Sub
    Dim bMapX As Byte, bMapY As Byte, bCapa As Byte, tGrh As Long, tColor As Long, colRed As Byte, colGreen As Byte, colBlue As Byte
    For bMapX = 1 To 100
        For bMapY = 1 To 100
            For bCapa = 1 To 3 ' Dibujamos las 3 primeras capas
                tGrh = MapData(bMapX, bMapY).Graphic(bCapa).grhindex
                If tGrh > 0 And tGrh <= lMapColorCount Then
                    SetPixel frmMain.picRadar.hdc, bMapX - 1, bMapY - 1, lMapColor(tGrh)
                End If
            Next
            tGrh = MapData(bMapX, bMapY).ObjGrh.grhindex ' Dibujamos los objetos
            If tGrh > 0 And tGrh <= lMapColorCount Then
                tColor = GetPixel(frmMain.picRadar.hdc, bMapX - 1, bMapY - 1)
                colRed = (lMapColor(tGrh) And &HFF&)
                colGreen = (lMapColor(tGrh) And &HFF00&) / &H100
                colBlue = (lMapColor(tGrh) And &HFF0000) / &H10000
                tColor = (tColor - RGB(colRed / 200, colGreen / 200, colBlue / 200)) ' solo "un poco" del color del objeto sobre la base
                SetPixel frmMain.picRadar.hdc, bMapX - 1, bMapY - 1, tColor
            End If
            If MapData(bMapX, bMapY).CharIndex > 0 And NPCs = True Then ' Dibujamos los NPC's
                tGrh = charlist(MapData(bMapX, bMapY).CharIndex).Body.Walk(1).grhindex
                If tGrh > 0 And tGrh <= lMapColorCount Then
                    SetPixel frmMain.picRadar.hdc, bMapX - 1, bMapY - 1, lMapColor(tGrh)
                End If
            End If
        Next
    Next
    frmMain.picRadar.Refresh
End Sub

Public Sub MoveUserMiniMap()
   ' With frmMain.UserMiniMap
   '    .left = UserPos.X - 3
   '    .top = UserPos.Y - 3
   ' End With
End Sub
