VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'System text for texture
Private Type tFont
    font_size As Integer
    ascii_code(0 To 255) As Integer 'indice de cada letra
End Type
Private font_types(1 To 2) As tFont

'Engine
Private MouseTileX As Byte
Private MouseTileY As Byte
Public Tile_Buffer_Size As Integer
Public timerElapsedTime As Double
Public timerTicksPerFrame As Double
Public EngineBaseSpeed As Single
Private lFrameTimer As Long
Private lFrameLimiter As Long
Private MainScreenRect As RECT

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Byte
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Private Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
   
    Speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long
 
    frame_counter As Single
    frame_speed As Single
   
    stream_type As Byte
 
    particle_stream() As Particle
    particle_count As Long
   
    grh_index_list() As Long
    grh_index_count As Long
   
    alpha_blend As Boolean
   
    alive_counter As Long
    never_die As Boolean
   
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(3) As Long
   
    'Added by Juan Martín Sotuyo Dodero
    Speed As Single
    life_counter As Long
   
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
'Particle system
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long

'================================================================================================
'GUI
Private Const GRH_HUD As Long = 21632
Private Const GRH_BARRA_EXP As Long = 21633
Private Const GRH_CONNECT As Long = 21813
Private Const GRH_MINIMAP As Long = 21815
Private Const GRH_INVENTARIOS As Long = 21816
Private Const GRH_STATS As Long = 21817
Private Const GRH_E_HP As Long = 21822
Private Const GRH_E_MP As Long = 21823
Private Const GRH_BARRA_HAMBRE As Long = 21819
Private Const GRH_BARRA_SED As Long = 21820
Private Const GRH_BARRA_ENERGIA As Long = 21821
'================================================================================================
Private EndTime As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0)
    Call SpriteBatch.Begin
End Sub
Public Sub Engine_EndScene(ByRef destRect As RECT, Optional ByVal hWndDest As Long = 0)
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
    
    If hWndDest = 0 Then
        Call DirectDevice.Present(destRect, ByVal 0&, ByVal 0&, ByVal 0&)
    Else
        Call DirectDevice.Present(destRect, ByVal 0, hWndDest, ByVal 0)
    End If
End Sub
Public Sub Engine_Start()
    Static lastFlush As Long
    
    Call Load_Map(2) 'for RenderConnect
    
    Do While prgRun
        If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
            ShowNextFrame frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY
            Check_Keys
        Else
            Render_Connect
        End If
        
        Call Engine_Update_FPS
        
        If timeGetTime >= lastFlush Then
            ' If there is anything to be sent, we send it
            Call FlushBuffer
            lastFlush = timeGetTime + 10
        End If
        
        DoEvents
    Loop

    'Close
    frmMain.Client.CloseSck
    'Start Socket
    Set frmMain.Client = Nothing
    Call CloseClient
End Sub
Public Sub Engine_Update_FPS()
Dim i As Integer

    If FPSLastCheck + 1000 < timeGetTime Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        FPSLastCheck = timeGetTime

        For i = 1 To LastChar
            If charlist(i).dialog <> "" Then
                charlist(i).dialog_life = charlist(i).dialog_life - 1
                If charlist(i).dialog_life <= 0 Then
                    charlist(i).dialog = ""
                    charlist(i).dialog_color = 0
                    charlist(i).dialog_life = 0
                    charlist(i).dialog_offset_counter_y = 0
                    charlist(i).dialog_scroll = False
                End If
            End If
        Next i
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If

    'Get timing info
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * EngineBaseSpeed
End Sub

Public Sub DirectX_Init()
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = IIf(Config_Inicio.VSYNC, D3DSWAPEFFECT_COPY_VSYNC, D3DSWAPEFFECT_COPY)
        .BackBufferWidth = Config_Inicio.ResolutionX
        .BackBufferHeight = Config_Inicio.ResolutionY
        .BackBufferFormat = DispMode.Format
    End With
    
    If Not DirectDevice Is Nothing Then
        Set DirectDevice = Nothing
    End If

    Set DirectDevice = DirectD3D.CreateDevice( _
                        D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                        frmMain.MainViewPic.hWnd, _
                        D3DCREATE_HARDWARE_VERTEXPROCESSING, _
                        D3DWindow)
        
    'Seteamos la matriz de proyeccion.
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)

    Engine_Init_RenderStates
    
    'Carga dinamica de texturas por defecto.
    Set SurfaceDB = New clsTextureManager

    'Sprite batching.
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialise(2000)
    
    Call SurfaceDB.Initialize(DirectD3D8)
    
    Engine_Init_PrincipalVariables
    
    EndTime = timeGetTime
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
End Sub
Private Sub Engine_Init_PrincipalVariables()
    
    Tile_Buffer_Size = 1 '// Lo uso en 1, por que manda a dibujar fuera del render al pedo :P
    
    'Tamaño de Tiles = 32x32
    Tile_Pixel_Size = 32
    
    HalfWindowTileHeight = Round(Config_Inicio.ResolutionY / 32, 0) \ 2
    HalfWindowTileWidth = Round(Config_Inicio.ResolutionX / 32, 0) \ 2
    
    'Velocidad del Engine
    EngineBaseSpeed = 0.029
    
    'Limites del mapa
    MinXBorder = MinMapSize + HalfWindowTileWidth: MinYBorder = MinMapSize + HalfWindowTileHeight
    MaxXBorder = MaxMapSize - HalfWindowTileWidth: MaxYBorder = MaxMapSize - HalfWindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(MinMapSize To MaxMapSize, MinMapSize To MaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    With MainScreenRect
        .bottom = Config_Inicio.ResolutionY
        .Right = Config_Inicio.ResolutionX
    End With
End Sub
Private Sub Engine_Init_RenderStates()
    'Set the render states
    With DirectDevice
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    End With
End Sub
Public Sub Engine_DeInit()
On Error Resume Next
    
    '   Clean Particles
    Call Particle_Group_Remove_All
    
    '   Clean Texture
    Call DirectDevice.SetTexture(0, Nothing)
    
    '   Erase Data
    Erase MapData()
    Erase charlist()
    
    Set DirectD3D = Nothing
    Set DirectDevice = Nothing
    Set DirectX = Nothing
    Set SpriteBatch = Nothing
End Sub
Public Sub Draw_GrhIndex(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, Optional ScaleX As Integer = 0, Optional ScaleY As Integer = 0, Optional Opacidad As Byte = 255)
    Dim rgb_list(3) As Long
    rgb_list(0) = D3DColorARGB(Opacidad, 255, 255, 255)
    rgb_list(1) = D3DColorARGB(Opacidad, 255, 255, 255)
    rgb_list(2) = D3DColorARGB(Opacidad, 255, 255, 255)
    rgb_list(3) = D3DColorARGB(Opacidad, 255, 255, 255)
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - Tile_Pixel_Size) \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * Tile_Pixel_Size) + Tile_Pixel_Size
            End If
        End If

        '// BEST IF NIGGA BRO
        'Este if sirve para las barras de vida, mana, etc, en render. Osea para ir cambiando su resolucion.
        If ScaleX = 0 And ScaleY = 0 Then
            Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, rgb_list())
        Else
            Call Device_Textured_Render(X, Y, ScaleX, ScaleY, .sX, .sY, .FileNum, rgb_list())
        End If
    End With
End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, _
                ByVal Y As Integer, ByVal Center As Byte, _
                ByVal Animate As Byte, ByRef Color_List() As Long, Optional movSpeed As Single = 1, Optional ByVal Alpha As Boolean = False, _
                Optional ByVal angle As Single = 0, _
                Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Long
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * movSpeed
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth * ScaleX - Tile_Pixel_Size) \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * Tile_Pixel_Size) + Tile_Pixel_Size
            End If
        End If

        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
    End With
    
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in Draw_Grh, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient
    End If
End Sub
Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    Dim Y               As Long     'Keeps track of where on map we are
    Dim X               As Long     'Keeps track of where on map we are
    Dim screenminY      As Integer  'Start Y pos on current screen
    Dim screenmaxY      As Integer  'End Y pos on current screen
    Dim screenminX      As Integer  'Start X pos on current screen
    Dim screenmaxX      As Integer  'End X pos on current screen
    Dim minY            As Integer  'Start Y pos on current map
    Dim maxY            As Integer  'End Y pos on current map
    Dim minX            As Integer  'Start X pos on current map
    Dim maxX            As Integer  'End X pos on current map
    Dim ScreenX         As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY         As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset      As Integer
    Dim minYOffset      As Integer
    Dim temp_array(3)   As Long
    Dim Techos(3) As Long
    Dim ResumeOffsetX As Integer
    Dim ResumeOffsetY As Integer
            
    'Figure out Ends and Starts of screen
    screenminY = TileY - HalfWindowTileHeight: screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth: screenmaxX = TileX + HalfWindowTileWidth
    minY = screenminY - Tile_Buffer_Size:     maxY = screenmaxY + Tile_Buffer_Size
    minX = screenminX - Tile_Buffer_Size:     maxX = screenmaxX + Tile_Buffer_Size
    
    EffectDay
    
    'Make sure mins and maxs are allways in map bounds
    If minY < MinMapSize Then
        minYOffset = MinMapSize - minY
        minY = MinMapSize
    End If
    
    If maxY > MaxMapSize Then maxY = MaxMapSize
    
    If minX < MinMapSize Then
        minXOffset = MinMapSize - minX
        minX = MinMapSize
    End If
    
    If maxX > MaxMapSize Then maxX = MaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > MinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < MaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > MinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < MaxMapSize Then screenmaxX = screenmaxX + 1
    
    'Draw layer 1
    ScreenY = minYOffset - Tile_Buffer_Size
    For Y = minY To maxY
        ScreenX = minXOffset - Tile_Buffer_Size
        For X = minX To maxX
            With MapData(X, Y)
                If .Layer(1).GrhIndex Then
                    Call Draw_Grh(.Layer(1), _
                                ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                                ScreenY * Tile_Pixel_Size + PixelOffsetY, _
                                1, 1, AmbientClima)
                End If
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw layer 2 and small OBJs
    ScreenY = minYOffset - Tile_Buffer_Size
    For Y = minY To maxY
        ScreenX = minXOffset - Tile_Buffer_Size
        For X = minX To maxX
            With MapData(X, Y)
                If .Layer(2).GrhIndex Then
                    Call Draw_Grh(.Layer(2), _
                            ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                            ScreenY * Tile_Pixel_Size + PixelOffsetY, _
                            1, 1, AmbientClima)
                End If
                
                'Sistema de huellas
                If .Huella.GrhIndex Then
                    Call Draw_Grh(.Huella, ScreenX * Tile_Pixel_Size + PixelOffsetX, ScreenY * Tile_Pixel_Size + PixelOffsetY, 1, 0, AmbientClima)
                End If
                
                If .ObjGrh.GrhIndex Then
                    If GrhData(.ObjGrh.GrhIndex).pixelHeight = 32 And GrhData(.ObjGrh.GrhIndex).pixelWidth = 32 Then
                        Call Draw_Grh(.ObjGrh, _
                            ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                            ScreenY * Tile_Pixel_Size + PixelOffsetY, _
                            1, 1, AmbientClima)
                    End If
                End If
                
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw Transparent Layers
    ScreenY = minYOffset - Tile_Buffer_Size
    For Y = minY To maxY
        ScreenX = minXOffset - Tile_Buffer_Size
        For X = minX To maxX
            With MapData(X, Y)
            
                If .CharIndex Then
                    Call Char_Render(.CharIndex, _
                        ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                        ScreenY * Tile_Pixel_Size + PixelOffsetY)
                End If
                
                If .CharIndex Then
                    Call NickName_Render(.CharIndex, _
                        ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                        ScreenY * Tile_Pixel_Size + PixelOffsetY)
                End If
                
                If .ObjGrh.GrhIndex Then
                    If GrhData(.ObjGrh.GrhIndex).pixelHeight <> 32 Or GrhData(.ObjGrh.GrhIndex).pixelWidth <> 32 Then
                        Call Draw_Grh(.ObjGrh, _
                            ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                            ScreenY * Tile_Pixel_Size + PixelOffsetY, _
                            1, 1, AmbientClima)
                    End If
                End If

                'If Not MapData(UserPos.X, UserPos.Y).Blocked = 15 Or MapData(UserPos.X, UserPos.Y).Blocked = 10 Then
                If .Layer(3).GrhIndex Then
                    Call Draw_Grh(.Layer(3), _
                        ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                        ScreenY * Tile_Pixel_Size + PixelOffsetY, _
                        1, 1, AmbientClima)
                End If
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw Layer 5
    ScreenY = minYOffset - Tile_Buffer_Size
    For Y = minY To maxY
        ScreenX = minXOffset - Tile_Buffer_Size
        For X = minX To maxX
            With MapData(X, Y)
                    If .Layer(5).GrhIndex Then
                        Call Draw_Grh(.Layer(5), _
                            ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                            ScreenY * Tile_Pixel_Size + PixelOffsetY, _
                            1, 1, AmbientClima)
                    End If
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw Dialogs
    ScreenY = minYOffset - Tile_Buffer_Size
    For Y = minY To maxY
        ScreenX = minXOffset - Tile_Buffer_Size
        For X = minX To maxX
            With MapData(X, Y)
                '*** Start Dialogs ***
                If .CharIndex Then
                    If charlist(.CharIndex).dialog <> "" Then
                        'Set colors
                        temp_array(0) = charlist(MapData(X, Y).CharIndex).dialog_color
                        temp_array(1) = charlist(MapData(X, Y).CharIndex).dialog_color
                        temp_array(2) = charlist(MapData(X, Y).CharIndex).dialog_color
                        temp_array(3) = charlist(MapData(X, Y).CharIndex).dialog_color


                        If charlist(.CharIndex).dialog_offset_counter_y < 5 Then
                            charlist(.CharIndex).dialog_offset_counter_y = charlist(.CharIndex).dialog_offset_counter_y + 1
                        End If
                        
                        Dim screen_y As Integer, screen_x As Integer
                        screen_x = (ScreenX * 32) + (CInt(charlist(.CharIndex).MoveOffsetX) + PixelOffsetX)
                        screen_y = (ScreenY * 32) + (CInt(charlist(.CharIndex).MoveOffsetY) + PixelOffsetY)
                        
                        'Render dialog
                        Engine_Text_Render charlist(.CharIndex).dialog, screen_x + 15 - Engine_Text_Width(charlist(.CharIndex).dialog, True) / 2, screen_y + (charlist(.CharIndex).Body.HeadOffset.Y + 6) - Engine_Text_Height(charlist(.CharIndex).dialog, True) - charlist(.CharIndex).dialog_offset_counter_y, temp_array, charlist(.CharIndex).dialog_font_index, True
                    End If
                End If
                '*** End Dialogs ***
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw blocked tiles and grid
    AlphaLayerFour Techos
    
    ScreenY = minYOffset - Tile_Buffer_Size
    For Y = minY To maxY
        ScreenX = minXOffset - Tile_Buffer_Size
        For X = minX To maxX
            With MapData(X, Y)
                If .Layer(4).GrhIndex Then
                    Call Draw_Grh(.Layer(4), _
                        ScreenX * Tile_Pixel_Size + PixelOffsetX, _
                        ScreenY * Tile_Pixel_Size + PixelOffsetY, _
                        1, 1, Techos)
                End If
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw Particle Map
    ScreenY = minYOffset - Tile_Buffer_Size
        For Y = minY To maxY
            ScreenX = minXOffset - Tile_Buffer_Size
            For X = minX To maxX
                With MapData(X, Y)
                    If .particle_group_index Then
                        Particle_Group_Render .particle_group_index, ScreenX * Tile_Pixel_Size + PixelOffsetX, ScreenY * Tile_Pixel_Size + PixelOffsetY
                    End If
                End With
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
        
    'Draw Particle Rain
    If bRain And Not MapInfo.Zone = 1 Then
        If particle_group_index_render(1) Then
                Particle_Group_Render particle_group_index_render(1), 0, 0
        End If
    End If
    
    'Draw Particle Mouse
    If particle_group_index_render(0) Then
        Particle_Group_Render particle_group_index_render(0), frmMain.MouseX - 15, frmMain.MouseY + 20
    End If
    
    Call Render_HUD
    Call DibujarCartel
    
    'FPS, MS & Position MAP.
    If View_FPS_OR_MS Then
        Engine_Text_Render "(" & UserMap & ", " & UserPos.X & ", " & UserPos.Y & ")", Config_Inicio.ResolutionX - 100, 107, LongWhite
        Engine_Text_Render FPS & " FPS", Config_Inicio.ResolutionX - 160, 0, LongWhite
        Engine_Text_Render pingTime & " MS", Config_Inicio.ResolutionX - 160, 10, LongWhite
    End If

End Sub
Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single

    Engine_BeginScene
        If UserMoving Then
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - charlist(UserCharIndex).Scroll_Pixels_Per_Frame * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(Tile_Pixel_Size * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - charlist(UserCharIndex).Scroll_Pixels_Per_Frame * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(Tile_Pixel_Size * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
    Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
    
    Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
    
    Engine_EndScene MainScreenRect, 0
    
    If frmBancoObj.PicBancoInv.Visible Then _
        Call InvBanco(0).DrawInv
         
    If frmBancoObj.PicInv.Visible Then _
        Call InvBanco(1).DrawInv
    
    If frmComerciar.picInvNpc.Visible Then _
        Call InvComNpc.DrawInv
        
    If frmComerciar.picInvUser.Visible Then _
        Call InvComUsu.DrawInv
        
    If frmComerciarUsu.picInvComercio.Visible Then _
        InvComUsu.DrawInv (1)
    
    If frmComerciarUsu.picInvOfertaProp.Visible Then _
        InvOfferComUsu(0).DrawInv (1)
    
    If frmComerciarUsu.picInvOfertaOtro.Visible Then _
        InvOfferComUsu(1).DrawInv (1)
    
    If frmComerciarUsu.picInvOroProp.Visible Then _
        InvOroComUsu(0).DrawInv (1)
    
    If frmComerciarUsu.picInvOroOfertaProp.Visible Then _
        InvOroComUsu(1).DrawInv (1)
    
    If frmComerciarUsu.picInvOroOfertaOtro.Visible Then _
        InvOroComUsu(2).DrawInv (1)

    If frmHerrero.Visible Then
        With frmHerrero
        If .picLingotes0.Visible Or .picMejorar0.Visible Then _
            InvLingosHerreria(1).DrawInv (1)
            
        If .picLingotes1.Visible Or .picMejorar1.Visible Then _
            InvLingosHerreria(2).DrawInv (1)
            
        If .picLingotes2.Visible Or .picMejorar2.Visible Then _
            InvLingosHerreria(3).DrawInv (1)
            
        If .picLingotes3.Visible Or .picMejorar3.Visible Then _
            InvLingosHerreria(4).DrawInv (1)
        End With
    End If

    If frmCarp.Visible Then
        With frmCarp
            If .picMaderas0.Visible Or .imgMejorar0.Visible Then _
                InvMaderasCarpinteria(1).DrawInv (1)
                
            If .picMaderas1.Visible Or .imgMejorar1.Visible Then _
                InvMaderasCarpinteria(2).DrawInv (1)
            
            If .picMaderas2.Visible Or .imgMejorar2.Visible Then _
                InvMaderasCarpinteria(3).DrawInv (1)
            
            If .picMaderas3.Visible Or .imgMejorar3.Visible Then _
                InvMaderasCarpinteria(4).DrawInv (1)
        End With
    End If

    '//Inventario
    If frmMain.Visible Then Call Inventario.DrawInv
End Sub

Private Sub NickName_Render(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
Dim Pos As Integer
Dim line As String
Dim Color As Long
Dim temp_array(3) As Long
Dim i As Integer
'Draw name over head
With charlist(CharIndex)
PixelOffsetX = PixelOffsetX + .MoveOffsetX
PixelOffsetY = PixelOffsetY + .MoveOffsetY

    If LenB(.Nombre) > 0 Then
        If Nombres Then
            Pos = getTagPosition(.Nombre)
            'Pos = InStr(.Nombre, "<")
            'If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
            If .priv = 0 Then
                If .Criminal Then
                    Color = D3DColorXRGB(128, 0, 0)  'CAOS
                Else
                    Color = D3DColorXRGB(0, 128, 192)  'IMPE
                End If
            Else
                Color = D3DColorXRGB(0, 128, 0) 'GM
            End If
            
            temp_array(0) = Color
            temp_array(1) = Color
            temp_array(2) = Color
            temp_array(3) = Color

            'Nick
            line = Left$(.Nombre, Pos - 2)
            Engine_Text_Render line, PixelOffsetX + 16 - Engine_Text_Width(line) / 2, PixelOffsetY + 30, temp_array
            
            'Clan
            line = Mid$(.Nombre, Pos)
            Engine_Text_Render line, PixelOffsetX + 16 - Engine_Text_Width(line) / 2, PixelOffsetY + 45, temp_array
                            
            If Not .priv = 0 Then ' // CLAN ADMIN
                line = "<Eternal Online Staff>"
                Engine_Text_Render line, PixelOffsetX + 16 - Engine_Text_Width(line) / 2, PixelOffsetY + 45, temp_array
            End If
     
        End If
    End If
End With
End Sub
Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, Optional angle As Byte = 0)
    Dim moved As Boolean
    Dim i As Integer
    Dim rgb_list(0 To 3) As Long
    
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + .Scroll_Pixels_Per_Frame * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + .Scroll_Pixels_Per_Frame * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                    .Arma.WeaponWalk(.Heading).Started = 1
                    .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
                .Body.Walk(.Heading).Started = 0
                .Body.Walk(.Heading).FrameCounter = 1
              
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
              
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1
              
                .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If .Envenenado Then
            rgb_list(0) = D3DColorXRGB(0, 255, 0)
            rgb_list(1) = rgb_list(0)
            rgb_list(2) = rgb_list(0)
            rgb_list(3) = rgb_list(0)
        End If
        
        If Not .invisible And Not .muerto Then
            rgb_list(0) = D3DColorXRGB(DayR, DayG, DayB)
            rgb_list(1) = rgb_list(0)
            rgb_list(2) = rgb_list(0)
            rgb_list(3) = rgb_list(0)
            Call EffectAlphaForChars(rgb_list(), CharIndex)
        Else
            Call RoofAlphaCalculateToAlpha(rgb_list(), CharIndex)
        End If

        If .Body.Walk(.Heading).GrhIndex Then _
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, rgb_list(), 0.5)

        If .Head.Head(.Heading).GrhIndex Then _
            Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X / 2, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, rgb_list(), 0.5)

        If .Casco.Head(.Heading).GrhIndex Then _
            Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X / 2, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 0, rgb_list(), 0.5)
        
        If .Arma.WeaponWalk(.Heading).GrhIndex Then _
            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, rgb_list(), 0.5)
                    
        If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
            Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, rgb_list(), 0.5)

        If UserWritting Then
            If CharIndex = UserCharIndex Then
                Call Draw_GrhIndex(20943, PixelOffsetX + .Body.HeadOffset.X + 15, PixelOffsetY + .Body.HeadOffset.Y, 0)
            End If
        End If
        
        If .particle_count > 0 Then
            For i = 1 To .particle_count
                If .particle_group(i) > 0 Then
                    Particle_Group_Render .particle_group(i), PixelOffsetX, PixelOffsetY
                End If
            Next i
        End If

        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, rgb_list(), 1, True)
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
   
 
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
    End If
End Function
 
Private Function Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
 
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
 
    particle_group_list(particle_group_index).active = True
 
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
 
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
 
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
 
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    particle_group_list(particle_group_index).stream_type = stream_type
 
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
   
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    MapData(map_x, map_y).particle_group_index = particle_group_index
End Function
Private Function Particle_Group_Make_Render(ByVal particle_group_index As Long, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional AutomaticScaleResolution As Byte)
 
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
 
    particle_group_list(particle_group_index).active = True
 
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
 
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
 
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    particle_group_list(particle_group_index).stream_type = stream_type
 
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    
    If AutomaticScaleResolution = 1 Then ' LLuvia
        particle_group_list(particle_group_index).x2 = Config_Inicio.ResolutionX
        particle_group_list(particle_group_index).y2 = Config_Inicio.ResolutionY
    Else
        particle_group_list(particle_group_index).x2 = x2
        particle_group_list(particle_group_index).y2 = y2
    End If
    
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
   
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    particle_group_index_render(AutomaticScaleResolution) = particle_group_index
End Function
 
Private Function Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
 
    If no_move = False Then
                If temp_particle.alive_counter = 0 Then
                    InitGrh temp_particle.Grh, grh_index, alpha_blend
                    temp_particle.X = RandomNumber(x1, x2)
                    temp_particle.Y = RandomNumber(y1, y2)
                    temp_particle.vector_x = RandomNumber(vecx1, vecx2)
                    temp_particle.vector_y = RandomNumber(vecy1, vecy2)
                    temp_particle.angle = angle
                    temp_particle.alive_counter = RandomNumber(life1, life2)
                    temp_particle.friction = fric
                Else
 
                    If gravity = True Then
                        temp_particle.vector_y = temp_particle.vector_y + grav_strength
                        If temp_particle.Y > 0 Then
                            temp_particle.vector_y = bounce_strength
                        End If
                    End If
 
                   If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                               
                    If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
                    If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
                End If
 
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
 
        temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
 
    'If grh_resize = True Then
    '    If temp_particle.Grh.GrhIndex Then
    '         Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
    '        Exit Function
    '    End If
    'End If
 
    If temp_particle.Grh.GrhIndex Then
        Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, rgb_list, 1, alpha_blend, temp_particle.angle
    End If
End Function
Private Function Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
 
    Dim loopC As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
 
    If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then 'for renderconnect
        If UserMinHP = 0 Then
            temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
            temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
            temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
            temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        Else
            temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
            temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
            temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
            temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
        End If
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
       
    If particle_group_list(particle_group_index).alive_counter Then
 
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
 
        For loopC = 1 To particle_group_list(particle_group_index).particle_count
 
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopC), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
                           
        Next loopC
       
        If no_move = False Then
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
   
    Else
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
   
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
   
    Particle_Group_Remove_All = True
End Function
 
Private Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopC As Long
   
    loopC = 1
    Do Until particle_group_list(loopC).id = id
        If loopC = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopC = loopC + 1
    Loop
   
    Particle_Group_Find = loopC
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Private Function Particle_Group_Destroy(ByVal particle_group_index As Long)
 
    Dim temp As particle_group
    Dim i As Integer
   
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    Exit For
                End If
            Next i
        End If
    End If
   
    particle_group_list(particle_group_index) = temp
 
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Function
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Function
Private Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function
Public Function Particle_Group_Create_Render(ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional AutomaticScaleResolution As Byte) As Long
   
 

        Particle_Group_Create_Render = Particle_Group_Next_Open
        Particle_Group_Make_Render Particle_Group_Create_Render, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey, AutomaticScaleResolution

End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
 
    Dim char_part_free_index As Integer
   
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
   
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
       
    End If
 
End Function
 Private Function Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                               
 
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
 
    particle_group_list(particle_group_index).active = True
 
    particle_group_list(particle_group_index).char_index = char_index
 
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
 
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
 
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
 
    particle_group_list(particle_group_index).stream_type = stream_type
 
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
 
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
 
    particle_group_list(particle_group_index).id = id
 
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
End Function
Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
 
On Error GoTo ErrorHandler:
    Dim loopC As Long
   
    loopC = 1
    Do Until charlist(char_index).particle_group(loopC) = 0
        If loopC = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopC = loopC + 1
    Loop
   
    Char_Particle_Group_Next_Open = loopC
 
Exit Function
 
ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1
 
End Function
Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopC As Long
   
    loopC = 1
    Do Until particle_group_list(loopC).active = False
        If loopC = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopC = loopC + 1
    Loop
   
    Particle_Group_Next_Open = loopC
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function
Private Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean
 
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
   
End Function
 
Private Function Map_In_Bounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
    Map_In_Bounds = True
End Function
Private Sub EffectAlphaForChars(roofrgb_list() As Long, ByVal CharIndex As Integer)
    Dim Color As D3DCOLORVALUE
    
    With charlist(CharIndex)
        If .StatusAlpha = False Then
            If .AlphaX < 255 Then
                .AlphaX = .AlphaX + (timerTicksPerFrame * 6.7)
                .last_tick = GetTickCount
            Else
                .AlphaX = 255
            End If
        Else
            If .AlphaX > 0 Then
                .AlphaX = .AlphaX - (timerTicksPerFrame * 6.7)
                .last_tick = GetTickCount
            Else
                .AlphaX = 0
                 If .StatusAlpha = True Then Call EraseChar(CharIndex)
            End If
        End If
    End With
    
    Color = AmbientColor
    Color.a = charlist(CharIndex).AlphaX
    D3DColorToRgbList roofrgb_list(), Color
End Sub
Private Sub RoofAlphaCalculateToAlpha(roofrgb_list() As Long, ByVal CharIndex As Integer)
    Dim Color As D3DCOLORVALUE
    Static State As Byte
   
    If Not CharIndex = UserCharIndex Then
        If State Then
            If charlist(CharIndex).AlphaX > 0 Then
                charlist(CharIndex).AlphaX = charlist(CharIndex).AlphaX - (timerTicksPerFrame * 6.7)
                charlist(CharIndex).last_tick = GetTickCount
            Else
                State = 0
                charlist(CharIndex).AlphaX = 0
            End If
        Else
            If charlist(CharIndex).AlphaX < 255 Then
                charlist(CharIndex).AlphaX = charlist(CharIndex).AlphaX + (timerTicksPerFrame * 6.7)
                charlist(CharIndex).last_tick = GetTickCount
            Else
                State = 1
                charlist(CharIndex).AlphaX = 255
            End If
        End If
    Else
       charlist(CharIndex).AlphaX = 120
    End If
       
       
    Color = AmbientColor
    Color.a = charlist(CharIndex).AlphaX
    D3DColorToRgbList roofrgb_list(), Color
End Sub


Sub Engine_Text_Render(Texto As String, X As Integer, Y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

    If (Len(Texto) = 0) Then Exit Sub
    
    d = 0
    If multi_line = False Then
        For a = 1 To Len(Texto)
            b = Asc(Mid$(Texto, a, 1))
            graf.GrhIndex = font_types(font_index).ascii_code(b)
            If b <> 32 Then
                If graf.GrhIndex <> 0 Then
                    graf.GrhIndex = font_types(font_index).ascii_code(b)
                    Draw_Grh graf, (X + d), Y, 0, 0, text_color
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                End If
            Else
                d = d + 4
            End If
        Next a
    Else
        e = 0
        f = 0
        For a = 1 To Len(Texto)
            b = Asc(Mid$(Texto, a, 1))
            graf.GrhIndex = font_types(font_index).ascii_code(b)
            If b = 32 Or b = 13 Then
                If e >= 20 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If b = 32 Then d = d + 4
                End If
            Else
                If graf.GrhIndex > 12 Then
                    graf.GrhIndex = font_types(font_index).ascii_code(b)
                    Draw_Grh graf, (X + d), Y + f * 14, 0, 0, text_color
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                End If
           End If
           e = e + 1
        Next a
    End If
End Sub

Function Engine_Text_Width(Texto As String, Optional multi As Boolean = False) As Integer
Dim a As Integer, b As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

    If multi = False Then
        For a = 1 To Len(Texto)
            b = Asc(Mid$(Texto, a, 1))
            graf.GrhIndex = font_types(1).ascii_code(b)
            If (b <> 32) And (b <> 5) And (b <> 129) And (b <> 9) And (b <> 4) And (b <> 255) And (b <> 2) And (b <> 151) And (b <> 152) Then
                Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
            Else
                Engine_Text_Width = Engine_Text_Width + 4
            End If
        Next a
    Else
        e = 0
        f = 0
        For a = 1 To Len(Texto)
            b = Asc(Mid$(Texto, a, 1))
            graf.GrhIndex = font_types(1).ascii_code(b)
            If b = 32 Or b = 13 Then
                If e >= 20 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If b = 32 Then d = d + 4
                End If
            Else
                If graf.GrhIndex > 12 Then
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                    If d > Engine_Text_Width Then Engine_Text_Width = d
                End If
            End If
            e = e + 1
        Next a
    End If
End Function

Function Engine_Text_Height(Texto As String, Optional multi As Boolean = False) As Integer
Dim a As Integer, b As Integer, c  As Integer, d  As Integer, e As Integer, f As Integer
    If multi = False Then
        Engine_Text_Height = 0
    Else
        e = 0
        f = 0
        For a = 1 To Len(Texto)
            b = Asc(Mid$(Texto, a, 1))
            If b = 32 Or b = 13 Then
                If e >= 20 Then
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If b = 32 Then
                        d = d + 4
                    End If
                End If
            End If
            e = e + 1
        Next a
        Engine_Text_Height = f * 14
    End If
End Function
Public Sub Char_Dialog_Set(ByVal char_index As Integer, ByVal char_dialog As String, ByVal char_dialog_color As Long, ByVal char_dialog_life As Byte, Optional ByVal font_index As Integer = 1)
    charlist(char_index).dialog = char_dialog
    charlist(char_index).dialog_color = char_dialog_color
    charlist(char_index).dialog_life = char_dialog_life
    charlist(char_index).dialog_font_index = font_index
    charlist(char_index).dialog_offset_counter_y = 0
End Sub
Public Sub Font_Initialize()
Dim a As Integer

font_types(1).font_size = 9

'Numeros
font_types(1).ascii_code(48) = 21686
font_types(1).ascii_code(49) = 21687
font_types(1).ascii_code(50) = 21688
font_types(1).ascii_code(51) = 21689
font_types(1).ascii_code(52) = 21690
font_types(1).ascii_code(53) = 21691
font_types(1).ascii_code(54) = 21692
font_types(1).ascii_code(55) = 21693
font_types(1).ascii_code(56) = 21694
font_types(1).ascii_code(57) = 21695

'abc
For a = 0 To 25
    font_types(1).ascii_code(a + 97) = 21634 + a
Next a

'ABC
For a = 0 To 25
    font_types(1).ascii_code(a + 65) = 21660 + a
Next a


font_types(1).ascii_code(33) = 21696
font_types(1).ascii_code(161) = 21697
font_types(1).ascii_code(34) = 21698
font_types(1).ascii_code(36) = 21699
font_types(1).ascii_code(191) = 21700
font_types(1).ascii_code(35) = 21701
font_types(1).ascii_code(36) = 21702
font_types(1).ascii_code(37) = 21703
font_types(1).ascii_code(38) = 21704
font_types(1).ascii_code(47) = 21705
font_types(1).ascii_code(92) = 21706
font_types(1).ascii_code(40) = 21707
font_types(1).ascii_code(41) = 21708
font_types(1).ascii_code(61) = 21709
font_types(1).ascii_code(39) = 21710
font_types(1).ascii_code(123) = 21711
font_types(1).ascii_code(125) = 21712
font_types(1).ascii_code(95) = 21713
font_types(1).ascii_code(45) = 21714

font_types(1).ascii_code(63) = 21699

font_types(1).ascii_code(64) = 21715
font_types(1).ascii_code(94) = 21716
font_types(1).ascii_code(91) = 21717
font_types(1).ascii_code(93) = 21718
font_types(1).ascii_code(60) = 21719
font_types(1).ascii_code(62) = 21720
font_types(1).ascii_code(42) = 21721
font_types(1).ascii_code(43) = 21722
font_types(1).ascii_code(46) = 21723
font_types(1).ascii_code(44) = 21724
font_types(1).ascii_code(58) = 21725
font_types(1).ascii_code(59) = 21726
font_types(1).ascii_code(124) = 21727

font_types(1).ascii_code(252) = 21728
font_types(1).ascii_code(220) = 21729
font_types(1).ascii_code(225) = 21730
font_types(1).ascii_code(233) = 21731
font_types(1).ascii_code(237) = 21732
font_types(1).ascii_code(243) = 21733
font_types(1).ascii_code(250) = 21734
font_types(1).ascii_code(253) = 21735
font_types(1).ascii_code(193) = 21736
font_types(1).ascii_code(201) = 21737
font_types(1).ascii_code(205) = 21738
font_types(1).ascii_code(211) = 21739
font_types(1).ascii_code(218) = 21740
font_types(1).ascii_code(221) = 21741
font_types(1).ascii_code(224) = 21742
font_types(1).ascii_code(232) = 21743
font_types(1).ascii_code(236) = 21744
font_types(1).ascii_code(242) = 21745
font_types(1).ascii_code(249) = 21746
font_types(1).ascii_code(192) = 21747
font_types(1).ascii_code(200) = 21748
font_types(1).ascii_code(204) = 21749
font_types(1).ascii_code(210) = 21750
font_types(1).ascii_code(217) = 21751
font_types(1).ascii_code(241) = 21752
font_types(1).ascii_code(209) = 21753

font_types(1).ascii_code(196) = 21773
font_types(1).ascii_code(194) = 21774
font_types(1).ascii_code(203) = 21775
font_types(1).ascii_code(207) = 21776
font_types(1).ascii_code(214) = 21777
font_types(1).ascii_code(212) = 21778

font_types(2).font_size = 9
font_types(2).ascii_code(97) = 21757
font_types(2).ascii_code(108) = 21758

font_types(2).ascii_code(115) = 21780

font_types(2).ascii_code(70) = 21759
font_types(2).ascii_code(48) = 21760
font_types(2).ascii_code(49) = 21761
font_types(2).ascii_code(50) = 21762
font_types(2).ascii_code(51) = 21763
font_types(2).ascii_code(52) = 21764
font_types(2).ascii_code(53) = 21765
font_types(2).ascii_code(54) = 21766
font_types(2).ascii_code(55) = 21767
font_types(2).ascii_code(56) = 21768
font_types(2).ascii_code(57) = 21769
font_types(2).ascii_code(33) = 21770
font_types(2).ascii_code(161) = 21771
font_types(2).ascii_code(42) = 21772

End Sub
Public Sub Draw_Box(ByVal X As Integer, ByVal Y As Integer, ByVal w As Integer, ByVal h As Integer, BackgroundColor() As Long)
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(X, Y, w, h, BackgroundColor)
End Sub
Public Sub Device_Textured_Render(ByVal X As Integer, ByVal Y As Integer, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef Color() As Long, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!)

        Dim Texture As Direct3DTexture8
        
        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
        
        With SpriteBatch
                Call .SetTexture(Texture)
                Call .SetAlpha(Alpha)
                
                If TextureWidth <> 0 And TextureHeight <> 0 Then
                    Call .Draw(X, Y, Width, Height, Color, (sX / TextureWidth) + 0.0001, (sY / TextureHeight) + 0.0001, ((sX + Width) / TextureWidth) + 0.0001, ((sY + Height) / TextureHeight) + 0.0001, angle)
                Else
                    Call .Draw(X, Y, TextureWidth, TextureHeight, Color, , , , , angle)
                End If
        End With
        
End Sub
Private Sub AlphaLayerFour(roofrgb_list() As Long)
Dim Color As D3DCOLORVALUE
Static AlphaY As Integer
Static last_tick As Long
    
        If bTecho Then
            If AlphaY > 0 Then
                If GetTickCount - last_tick >= 18 Then
                    AlphaY = AlphaY - 5
                    last_tick = GetTickCount
                End If
            End If
        Else
            If AlphaY < 255 Then
                If GetTickCount - last_tick >= 18 Then
                    AlphaY = AlphaY + 5
                    last_tick = GetTickCount
                End If
            End If
        End If
       
    Color = AmbientColor
    Color.a = AlphaY
    D3DColorToRgbList roofrgb_list(), Color
End Sub
Public Sub D3DColorToRgbList(rgb_list() As Long, Color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub
Private Function GetElapsedTime() As Single
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Function Engine_ElapsedTime() As Long
'**************************************************************
'Gets the time that past since the last call
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
'**************************************************************

    Dim Start_Time As Long

    'Get current time
    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime

    'Get next end time
    EndTime = Start_Time

End Function
Private Sub Render_UserStats()
Dim line As String

    'Dibujamos la esfera vacia de vida.
    Call Draw_GrhIndex(GRH_E_HP, PosHUDX - 220, PosHUDY, 1, 85, (86 - (((UserMinHP / 100) / (UserMaxHP / 100)) * 86)))
    
    'Dibujamos la esfera vacia de mana.
    If Not UserMaxMAN = 0 Then '// Si sos guerre que dibuje el "0/0" y no me bugee todo el engine.
        Call Draw_GrhIndex(GRH_E_MP, PosHUDX + 217, PosHUDY - 1, 1, 85, (86 - (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 86)))
    Else
        Call Draw_GrhIndex(GRH_E_MP, PosHUDX + 217, PosHUDY - 1, 1)
    End If
    
    '// HP TEXT
    line = UserMinHP & "/" & UserMaxHP
    Engine_Text_Render line, PosHUDX - 205 - Engine_Text_Width(line) / 2, PosHUDY - 20, LongWhite

    '// MP TEXT
    line = UserMinMAN & "/" & UserMaxMAN
    If Not UserMaxMAN = 0 Then
        Engine_Text_Render line, PosHUDX + 233 - Engine_Text_Width(line) / 2, PosHUDY - 20, LongWhite
    End If

    'LEVEL USER!
    line = UserLvl
    Engine_Text_Render line, PosHUDX + 184 - Engine_Text_Width(line) / 2, PosHUDY + 12, LongWhite

    '// BARRA DE EXPERIENCIA
    If Not UserLvl = 50 Then
        line = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel)) & "%"
        Call Draw_GrhIndex(GRH_BARRA_EXP, PosHUDX - 6, PosHUDY - 9, 1, Round(((UserExp / 100) / (UserPasarNivel / 100)) * 330), 9)
        Engine_Text_Render line, PosHUDX - Engine_Text_Width(line) / 2, PosHUDY + 11, LongWhite
    Else
        line = "¡Nivel maximo!"
        Call Draw_GrhIndex(GRH_BARRA_EXP, PosHUDX - 6, PosHUDY - 9, 1)
        Engine_Text_Render line, PosHUDX - Engine_Text_Width(line) / 2, PosHUDY + 11, LongWhite
    End If

    Call Draw_GrhIndex(GRH_STATS, 105, 65, 1)
    
    Call Draw_GrhIndex(GRH_BARRA_HAMBRE, 135, -11, 1, (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 171), 11)
    Call Draw_GrhIndex(GRH_BARRA_SED, 141, 7, 1, (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 160), 11)
    Call Draw_GrhIndex(GRH_BARRA_ENERGIA, 136, 34, 1, (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 169), 11)
    
    line = UserMinHAM & "/" & UserMaxHAM
    Engine_Text_Render line, 153 - Engine_Text_Width(line) / 2, 8, LongWhite
    
    line = UserMinAGU & "/" & UserMaxAGU
    Engine_Text_Render line, 153 - Engine_Text_Width(line) / 2, 26, LongWhite
    
    line = UserMinSTA & "/" & UserMaxSTA
    Engine_Text_Render line, 153 - Engine_Text_Width(line) / 2, 53, LongWhite
    
    line = UserGLD
    Engine_Text_Render line, 160 - Engine_Text_Width(line) / 2, 74, LongYellow
    
    line = UserFuerza
    Engine_Text_Render line, 28 - Engine_Text_Width(line) / 2, 76, LongGreen
    
    line = UserAgilidad
    Engine_Text_Render line, 51 - Engine_Text_Width(line) / 2, 76, LongYellow


End Sub
Public Sub Render_HUD()
Dim i As Byte, LongRed(3) As Long

    LongRed(0) = D3DColorXRGB(255, 0, 0)
    LongRed(1) = LongRed(0)
    LongRed(2) = LongRed(0)
    LongRed(3) = LongRed(0)
    
    Render_Console
    Render_Chat
    
    Call Draw_GrhIndex(GRH_HUD, PosHUDX, PosHUDY, 1)
    Call Draw_GrhIndex(GRH_MINIMAP, Config_Inicio.ResolutionX - 70, 73, 1)
    Call Draw_GrhIndex(GRH_INVENTARIOS, Config_Inicio.ResolutionX - 99, Config_Inicio.ResolutionY - 34, 1)
    
    Render_UserStats
    
    'Call Draw_GrhIndex(23655, Config_Inicio.ResolutionX - 102, Config_Inicio.ResolutionY - 32, 1, lighthandle) '//Inv
    If Not bRain Then
        Call Draw_GrhIndex(Clima(Hour(Time)).GRH_CLIMA, PosHUDX + 110, PosHUDY - 25, 1)
    Else
        Select Case MapInfo.Terrain
            Case 1 'Nieve
                Call Draw_GrhIndex(22019, PosHUDX + 110, PosHUDY - 25, 1)
            Case 2 'Tormenta de arena
                Call Draw_GrhIndex(22018, PosHUDX + 110, PosHUDY - 25, 1)
            Case Else 'LLuvia'
                Call Draw_GrhIndex(22020, PosHUDX + 110, PosHUDY - 25, 1)
        End Select
    End If
    
    For i = 1 To 6
        If MacroList(i).mTipe = 0 Or MacroList(i).Grh <= 0 Then
            'Call Draw_GrhIndex(1, frmMain.Macros(i).Left, PosHUDY - 25, 1, lighthandle)
        Else
            Call Draw_GrhIndex(MacroList(i).Grh, frmMain.Macros(i).Left, PosHUDY - 25, 1)
        End If
        
        If MacroList(i).mTipe = eMacros.aUsar Or MacroList(i).mTipe = eMacros.aEquipar Then
            Engine_Text_Render CStr(TotalItemAmountGet(MacroList(i).ObjIndex)), frmMain.Macros(i).Left + 1, frmMain.Macros(i).Top + 20, LongWhite
        End If
        
        If ObjIndexEquipped(MacroList(i).ObjIndex) Then
           Engine_Text_Render "+", frmMain.Macros(i).Left + 22, frmMain.Macros(i).Top, LongRed
        End If
    Next i
    
End Sub
Sub Render_Console()
Dim i As Byte
Dim temp_array_console(3) As Long
If OffSetConsola > 0 Then OffSetConsola = OffSetConsola - 1
If OffSetConsola = 0 Then UltimaLineavisible = True
 
For i = 1 To MaxLineas - 1
    temp_array_console(0) = D3DColorARGB(Con(i).a, Con(i).r, Con(i).g, Con(i).b)
    temp_array_console(1) = temp_array_console(0)
    temp_array_console(2) = temp_array_console(0)
    temp_array_console(3) = temp_array_console(0)
    Engine_Text_Render Con(i).T, 0, (Config_Inicio.ResolutionY - 235) + (i * 15) + OffSetConsola, temp_array_console
Next i

If UltimaLineavisible = True Then
    temp_array_console(0) = D3DColorARGB(Con(MaxLineas).a, Con(MaxLineas).r, Con(MaxLineas).g, Con(MaxLineas).b)
    temp_array_console(1) = temp_array_console(0)
    temp_array_console(2) = temp_array_console(0)
    temp_array_console(3) = temp_array_console(0)
    Engine_Text_Render Con(MaxLineas).T, 0, (Config_Inicio.ResolutionY - 235) + (MaxLineas * 15) + OffSetConsola, temp_array_console
End If

End Sub
Public Sub Render_Chat()
    If UserWritting Then
        Engine_Text_Render "Chat: " & ChatBuffer, 0, Config_Inicio.ResolutionY - 115, LongWhite
    End If
End Sub
Public Sub Render_Connect()
'===============================
'Conectar renderizado
'===============================
Dim View_X As Long
Dim View_Y As Long
Dim AmbientClima(3) As Long
Dim MaxViewTilesX As Byte
Dim MaxViewTilesY As Byte
Dim TotalLayers As Byte
Dim GRH_CONNECT_POS_X As Long
Dim GRH_CONNECT_POS_Y As Long
Static Effect As Byte

GRH_CONNECT_POS_X = Config_Inicio.ResolutionX - 326
GRH_CONNECT_POS_Y = 0

'Max view tiles in render.
MaxViewTilesX = (Config_Inicio.ResolutionX / 32) + 1
MaxViewTilesY = (Config_Inicio.ResolutionY / 32) + 1
'Day
AmbientClima(0) = D3DColorXRGB(Effect, Effect, Effect)
AmbientClima(1) = AmbientClima(0)
AmbientClima(2) = AmbientClima(0)
AmbientClima(3) = AmbientClima(0)

    Engine_BeginScene
    
    If Not INTRO Then
    
    If Effect < 255 Then Effect = Effect + 1
    
        'Draw all layers
        For TotalLayers = 1 To 4
            For View_X = 1 To MaxViewTilesX
                For View_Y = 1 To MaxViewTilesY
                    With MapData(View_X + 20, View_Y + 29)
                            If .Layer(TotalLayers).GrhIndex <> 0 Then
                                Call Draw_Grh(.Layer(TotalLayers), _
                                    (View_X - 1) * 32, (View_Y - 1) * 32, _
                                    1, 1, AmbientClima)
                            End If
                    End With
                Next View_Y
            Next View_X
        Next TotalLayers
        
        For View_X = 1 To MaxViewTilesX
            For View_Y = 1 To MaxViewTilesY
                With MapData(View_X + 20, View_Y + 29)
                        If .Layer(5).GrhIndex <> 0 Then
                            Call Draw_Grh(.Layer(5), _
                                (View_X - 1) * 32, (View_Y - 1) * 32, _
                                1, 1, AmbientClima)
                        End If
                End With
            Next View_Y
        Next View_X
        
        For View_X = 1 To MaxViewTilesX
            For View_Y = 1 To MaxViewTilesY
                With MapData(View_X + 20, View_Y + 29)
                    If .particle_group_index Then
                        Particle_Group_Render .particle_group_index, (View_X - 1) * 32, (View_Y - 1) * 32
                    End If
                End With
            Next View_Y
        Next View_X
        
        Call Draw_Box(GRH_CONNECT_POS_X, GRH_CONNECT_POS_Y, 260, Config_Inicio.ResolutionY, LongBlack)
        Call Draw_GrhIndex(21818, Config_Inicio.ResolutionX / 2, GRH_CONNECT_POS_Y + 150, 1, 0, 0, Effect) 'Eternal Online
        
        'Call Draw_GrhIndex(GRH_CONNECT, GRH_CONNECT_POS_X, GRH_CONNECT_POS_Y, 1, 0, 0, Effect)
        Engine_Text_Render FPS & " FPS", Config_Inicio.ResolutionX - 52, 0, LongWhite
    Else
        If Not BastaINTRO >= 2 Then
            Call Draw_GrhIndex(22026, Config_Inicio.ResolutionX / 2 - 15, Config_Inicio.ResolutionY / 2 + 250, 1, 0, 0, EffectIntro)
        Else
            Call Draw_GrhIndex(22027, Config_Inicio.ResolutionX / 2 - 15, Config_Inicio.ResolutionY / 2 + 250, 1, 0, 0, EffectIntro)
        End If
    End If
    
    Engine_EndScene MainScreenRect, frmConnect.picRender.hWnd

End Sub

