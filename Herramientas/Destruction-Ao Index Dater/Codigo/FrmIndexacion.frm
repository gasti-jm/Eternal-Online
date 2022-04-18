VERSION 5.00
Begin VB.Form FrmIndexacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destruction-Ao Indexador"
   ClientHeight    =   2610
   ClientLeft      =   4260
   ClientTop       =   645
   ClientWidth     =   5625
   Icon            =   "FrmIndexacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdIndexacion 
      Caption         =   "Empezar Indexacion"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2115
      Width           =   1935
   End
   Begin VB.Image Imagen 
      Height          =   1920
      Left            =   0
      Picture         =   "FrmIndexacion.frx":08CA
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "FrmIndexacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TipoDeIndexacion As Byte    '--->Para saber que tipo de Indexacion
Dim NumeroDeGrhIndex As Long    '--->Para saber desde donde empezar el GRHINDEX
Dim CantidadAlto As Byte        '--->Define la cantidad de Animaciones a lo Alto
Dim CantidadLargo As Byte       '--->Define la cantidad de Animaciones a lo Largo
Dim AnimacionLargo As Integer   '--->Este nos da las dimenciones de cada cuadro a lo largo y se le suma uno para que de 25 Luego 1 mas para que de 26
Dim AnimacionAlto As Integer    '--->Este nos da las dimenciones de cada cuadro a lo alto y se le suma 1 despues para que de 46
Dim Index As String             '--->En esta variable guardamos todo el renglon de cada Cuadro de animacion que se Printea en la Base de Datos
Dim SumaLado As Integer         '--->Esto es para sumarle al indexador para que tome la Imagen de al lado(el cuadro)
Dim SumaAlto As Integer         '--->Esto es para sumarle al indexador para que tome la Imagen de abajo(el cuadro)
Dim NorteSur As Boolean         '--->Para definir si es la animacion de 6 Cuadros o la de 5
Dim AnimacionFinal() As Long    '--->Para poder poner la Animacion al final
Dim Animacion As String         '--->Es la variable que Almacena toda la ANIMACION para grabarla en la Base de Datos
Dim CuantosLargo As Byte        '--->Para saber cuantas Animaciones hay a lo LARGO(Cada cuadro)
Dim CuantosAlto As Byte         '--->Para saber cuantas Animaciones hay a lo ALTO(Cada cuadro)
Dim SumandoAnim As Byte         '--->Para definir los Grh que componen cada Animacion
Dim CantImagenes As Byte        '--->Para definir cuantas Imagenes son si es una animacion conjunta de varias imagenes
Private Sub cmdIndexacion_Click()
Dim J As Byte
Dim T As Byte
Dim Y As Byte
Dim Imag As Byte

TipoDeIndexacion = Val(InputBox("¿Que indexacion queres hacer? " & vbCrLf & "1=Index de Movimiento: Armaduras, Tunicas, Animaciones que tengan direcion Norte, Sur, Este y Oeste." & vbCrLf & "2=Indexaciones de: Cascos, Sombreros y Cabezas." & vbCrLf & "3=Indexaciones de: NPC's como el Dragon que tienen una direccion por Imagen(en una esta el sur, en otra el norte, etc)." & vbCrLf & "4=Indexaciones de: Hechizos en una MISMA imagen y de Meditaciones." & vbCrLf & "5=Indexaciones de: Graficos De un solo cuadro(incluyendo Objetos de Inventario)." & vbCrLf & "6=Indexaciones de: Animaciones de movimiento completas(Ej: arañas, Npc's de comercio,etc)" & vbCrLf & "7= Animaciones completas pero en varias imagenes(Ej: Apocalipsis)" & vbCrLf & "8= Movimiento de Escudos" & vbCrLf & "9= Movimiento de Armas", "Tipo de Indexacion"))
If TipoDeIndexacion < 1 Then
    Me.Visible = False
    FrmIndex.Visible = True
    Exit Sub
End If

NumeroDeGrhIndex = AllGrhData + 1

'Todos los valores en 0
SumaLado = 0
SumaAlto = 0

ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData

Select Case TipoDeIndexacion
    Case 1 'Armaduras, Tunicas y Ropajes TIPO(6,6,5,5)
        AnimacionAlto = Int(Imagen.Height / 4)
        AnimacionLargo = Int(Imagen.Width / 6) + 1
        For T = 1 To 4
            SumaLado = 0
            If T < 3 Then
                For Y = 1 To 6
                    With GrhData(NumeroDeGrhIndex)
                        .NumFrames = 1
                        .FileNum = NombreGrafico
                        .sX = SumaLado
                        .sY = SumaAlto
                        .pixelWidth = AnimacionLargo + 1
                        .pixelHeight = AnimacionAlto + 1
                        ReDim Preserve GrhAnim(T).Animate(1 To 6) As Long
                        GrhAnim(T).Animate(Y) = NumeroDeGrhIndex
                    End With
                    If Y = 2 Or Y = 3 Then
                        SumaLado = SumaLado + AnimacionLargo - 1
                    Else
                        SumaLado = SumaLado + AnimacionLargo
                    End If
                    NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                    ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
                Next Y
                SumaAlto = SumaAlto + AnimacionAlto
            Else
                For Y = 1 To 5
                    With GrhData(NumeroDeGrhIndex)
                        .NumFrames = 1
                        .FileNum = NombreGrafico
                        .sX = SumaLado
                        .sY = SumaAlto
                        .pixelWidth = AnimacionLargo + 1
                        .pixelHeight = AnimacionAlto + 1
                        ReDim Preserve GrhAnim(T).Animate(1 To 5) As Long
                        GrhAnim(T).Animate(Y) = NumeroDeGrhIndex
                    End With
                    If Y = 2 Or Y = 3 Then
                        SumaLado = SumaLado + AnimacionLargo - 1
                    Else
                        SumaLado = SumaLado + AnimacionLargo
                    End If
                    NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                    ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
                Next Y
                SumaAlto = SumaAlto + AnimacionAlto
            End If
        Next T
        
        For T = 1 To 4
            If T < 3 Then
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 6
                    For Y = 1 To 6
                        ReDim Preserve .Frames(1 To 6) As Long
                        .Frames(Y) = GrhAnim(T).Animate(Y)
                    Next Y
                    .Speed = 6 * 1000 / 18
                    If IndexMode = "12.1" Then
                        .Speed = 6 * 1000 / 18
                    Else
                        .Speed = 1
                    End If
                    Animaciones(T) = NumeroDeGrhIndex
                End With
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            Else
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 5
                    For Y = 1 To 5
                        ReDim Preserve .Frames(1 To 5) As Long
                        .Frames(Y) = GrhAnim(T).Animate(Y)
                    Next Y
                    If IndexMode = "12.1" Then
                        .Speed = 5 * 1000 / 18
                    Else
                        .Speed = 1
                    End If
                    Animaciones(T) = NumeroDeGrhIndex
                End With
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            End If
        Next T
        
    Case 2 'Cascos, Cabezas y Sombreros
        AnimacionAlto = Int(Imagen.Height)
        AnimacionLargo = Int(Imagen.Width / 4)
        For T = 1 To 4
            With GrhData(NumeroDeGrhIndex)
                .NumFrames = 1
                .FileNum = NombreGrafico
                .sX = SumaLado
                .sY = 0
                .pixelHeight = AnimacionAlto
                .pixelWidth = AnimacionLargo
                ReDim GrhAnim(T).Animate(1 To 1) As Long
                GrhAnim(T).Animate(1) = NumeroDeGrhIndex
                Animaciones(T) = NumeroDeGrhIndex
            End With
            Animaciones(T) = NumeroDeGrhIndex
            SumaLado = SumaLado + AnimacionLargo
            NumeroDeGrhIndex = NumeroDeGrhIndex + 1
            ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
        Next T
        
    Case 3 'Animaciones En las cuales cada Direccion esta en una imagen distinta
        SelectImg.Visible = True
        UnFrame = False
        
    Case 4 'Hechizos y Meditaciones en misma Imagen
        CuantosLargo = Val(InputBox("¿Cuantas imagenes tiene a lo Largo?", "Imagenes a lo Largo"))
        CuantosAlto = Val(InputBox("¿Cuantas imagenes tiene a lo Alto?", "Imagenes a lo Alto"))
        
        If CuantosLargo < 1 Or CuantosAlto < 1 Then Exit Sub
        
        AnimacionLargo = Int(Imagen.Width / CuantosLargo)
        AnimacionAlto = Int(Imagen.Height / CuantosAlto)
        
        Dim Cont As Byte
        
        For Y = 1 To CuantosAlto
            SumaLado = 0
            For T = 1 To CuantosLargo
                With GrhData(NumeroDeGrhIndex)
                    Cont = Cont + 1
                    .NumFrames = 1
                    .FileNum = NombreGrafico
                    .sX = SumaLado
                    .sY = SumaAlto
                    .pixelWidth = AnimacionLargo
                    .pixelHeight = AnimacionAlto
                    ReDim Preserve GrhAnim(1).Animate(1 To Cont) As Long
                    GrhAnim(1).Animate(Cont) = NumeroDeGrhIndex
                End With
                SumaLado = SumaLado + AnimacionLargo
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            Next T
            SumaAlto = SumaAlto + AnimacionAlto
        Next Y
        
        With GrhData(NumeroDeGrhIndex)
            .NumFrames = Cont
            ReDim Preserve GrhData(NumeroDeGrhIndex).Frames(1 To .NumFrames) As Long
            For T = 1 To .NumFrames
                .Frames(T) = GrhAnim(1).Animate(T)
            Next T
            If IndexMode = "12.1" Then
                .Speed = .NumFrames * 1000 / 18
            Else
                .Speed = 1
            End If
        End With
        Animaciones(1) = NumeroDeGrhIndex
        
    Case 5 'Objetos de un solo Frame
        With GrhData(NumeroDeGrhIndex)
            .NumFrames = 1
            .FileNum = NombreGrafico
            .sX = 0
            .sY = 0
            .pixelWidth = Imagen.Width
            .pixelHeight = Imagen.Height
        End With
    
        MsgBox "La indexacion se realizo con Exito"
        
    Case 6 'Animaciones en una misma imagen con todas las direcciones
        CuantosLargo = Val(InputBox("¿Cuantas imagenes tiene a lo Largo?", "Imagenes a lo Largo"))
        CuantosAlto = Val(InputBox("¿Cuantas imagenes tiene a lo Alto?", "Imagenes a lo Alto"))
        
        If CuantosLargo < 1 Or CuantosAlto < 1 Then Exit Sub
        
        AnimacionLargo = Int(Imagen.Width / CuantosLargo)
        AnimacionAlto = Int(Imagen.Height / CuantosAlto)
        
        For Y = 1 To CuantosAlto
            SumaLado = 0
            ReDim Preserve GrhAnim(Y).Animate(1 To CuantosLargo) As Long
            For T = 1 To CuantosLargo
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 1
                    .FileNum = NombreGrafico
                    .sX = SumaLado
                    .sY = SumaAlto
                    .pixelWidth = AnimacionLargo
                    .pixelHeight = AnimacionAlto
                    GrhAnim(Y).Animate(T) = NumeroDeGrhIndex
                End With
                SumaLado = SumaLado + AnimacionLargo
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            Next T
            SumaAlto = SumaAlto + AnimacionAlto
        Next Y
        
        For Y = 1 To CuantosAlto
            With GrhData(NumeroDeGrhIndex)
                .NumFrames = CuantosLargo
                ReDim Preserve GrhData(NumeroDeGrhIndex).Frames(1 To .NumFrames) As Long
                For T = 1 To .NumFrames
                    .Frames(T) = GrhAnim(Y).Animate(T)
                Next T
                If IndexMode = "12.1" Then
                    .Speed = .NumFrames * 1000 / 18
                Else
                    .Speed = 1
                End If
            End With
            Animaciones(Y) = NumeroDeGrhIndex
            NumeroDeGrhIndex = NumeroDeGrhIndex + 1
            ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
        Next Y
        
        MsgBox "La indexacion se realizo con Exito"
    
    Case 7
        SelectImg.Visible = True
        UnFrame = True
    
    Case 8 'Escudos
        AnimacionAlto = Int(Imagen.Height / 4)
        AnimacionLargo = Int(Imagen.Width / 6) + 1
        For T = 1 To 4
            SumaLado = 0
            If T < 3 Then
                For Y = 1 To 6
                    With GrhData(NumeroDeGrhIndex)
                        .NumFrames = 1
                        .FileNum = NombreGrafico
                        .sX = SumaLado
                        .sY = SumaAlto
                        .pixelWidth = AnimacionLargo + 1
                        .pixelHeight = AnimacionAlto + 1
                        ReDim Preserve GrhAnim(T).Animate(1 To 6) As Long
                        GrhAnim(T).Animate(Y) = NumeroDeGrhIndex
                    End With
                    If Y = 2 Or Y = 3 Then
                        SumaLado = SumaLado + AnimacionLargo - 1
                    Else
                        SumaLado = SumaLado + AnimacionLargo
                    End If
                    NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                    ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
                Next Y
                SumaAlto = SumaAlto + AnimacionAlto
            Else
                For Y = 1 To 5
                    With GrhData(NumeroDeGrhIndex)
                        .NumFrames = 1
                        .FileNum = NombreGrafico
                        .sX = SumaLado
                        .sY = SumaAlto
                        .pixelWidth = AnimacionLargo + 1
                        .pixelHeight = AnimacionAlto + 1
                        ReDim Preserve GrhAnim(T).Animate(1 To 5) As Long
                        GrhAnim(T).Animate(Y) = NumeroDeGrhIndex
                    End With
                    If Y = 2 Or Y = 3 Then
                        SumaLado = SumaLado + AnimacionLargo - 1
                    Else
                        SumaLado = SumaLado + AnimacionLargo
                    End If
                    NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                    ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
                Next Y
                SumaAlto = SumaAlto + AnimacionAlto
            End If
        Next T
        
        For T = 1 To 4
            If T < 3 Then
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 6
                    For Y = 1 To 6
                        ReDim Preserve .Frames(1 To 6) As Long
                        .Frames(Y) = GrhAnim(T).Animate(Y)
                    Next Y
                    If IndexMode = "12.1" Then
                        .Speed = 6 * 1000 / 18
                    Else
                        .Speed = 1
                    End If
                    Animaciones(T) = NumeroDeGrhIndex
                End With
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            Else
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 5
                    For Y = 1 To 5
                        ReDim Preserve .Frames(1 To 5) As Long
                        .Frames(Y) = GrhAnim(T).Animate(Y)
                    Next Y
                    If IndexMode = "12.1" Then
                        .Speed = 5 * 1000 / 18
                    Else
                        .Speed = 1
                    End If
                    Animaciones(T) = NumeroDeGrhIndex
                End With
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            End If
        Next T
        
    Case 9 'Armas
        AnimacionAlto = Int(Imagen.Height / 4)
        AnimacionLargo = Int(Imagen.Width / 6) + 1
        For T = 1 To 4
            SumaLado = 0
            If T < 3 Then
                For Y = 1 To 6
                    With GrhData(NumeroDeGrhIndex)
                        .NumFrames = 1
                        .FileNum = NombreGrafico
                        .sX = SumaLado
                        .sY = SumaAlto
                        .pixelWidth = AnimacionLargo + 1
                        .pixelHeight = AnimacionAlto + 1
                        ReDim Preserve GrhAnim(T).Animate(1 To 6) As Long
                        GrhAnim(T).Animate(Y) = NumeroDeGrhIndex
                    End With
                    If Y = 2 Or Y = 3 Then
                        SumaLado = SumaLado + AnimacionLargo - 1
                    Else
                        SumaLado = SumaLado + AnimacionLargo
                    End If
                    NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                    ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
                Next Y
                SumaAlto = SumaAlto + AnimacionAlto
            Else
                For Y = 1 To 5
                    With GrhData(NumeroDeGrhIndex)
                        .NumFrames = 1
                        .FileNum = NombreGrafico
                        .sX = SumaLado
                        .sY = SumaAlto
                        .pixelWidth = AnimacionLargo + 1
                        .pixelHeight = AnimacionAlto + 1
                        ReDim Preserve GrhAnim(T).Animate(1 To 5) As Long
                        GrhAnim(T).Animate(Y) = NumeroDeGrhIndex
                    End With
                    If Y = 2 Or Y = 3 Then
                        SumaLado = SumaLado + AnimacionLargo - 1
                    Else
                        SumaLado = SumaLado + AnimacionLargo
                    End If
                    NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                    ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
                Next Y
                SumaAlto = SumaAlto + AnimacionAlto
            End If
        Next T
        
        For T = 1 To 4
            If T < 3 Then
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 6
                    For Y = 1 To 6
                        ReDim Preserve .Frames(1 To 6) As Long
                        .Frames(Y) = GrhAnim(T).Animate(Y)
                    Next Y
                    If IndexMode = "12.1" Then
                        .Speed = 6 * 1000 / 18
                    Else
                        .Speed = 1
                    End If
                    Animaciones(T) = NumeroDeGrhIndex
                End With
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            Else
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 5
                    For Y = 1 To 5
                        ReDim Preserve .Frames(1 To 5) As Long
                        .Frames(Y) = GrhAnim(T).Animate(Y)
                    Next Y
                    If IndexMode = "12.1" Then
                        .Speed = 5 * 1000 / 18
                    Else
                        .Speed = 1
                    End If
                    Animaciones(T) = NumeroDeGrhIndex
                End With
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            End If
        Next T
    
    Case Else
        MsgBox "La opcion elegida no es valida"
        
End Select
AllGrhData = NumeroDeGrhIndex

Select Case TipoDeIndexacion
    Case 1
        BodysCountNew = BodysCountNew + 1
        ReDim Preserve Bodys(0 To BodysCountNew) As tIndiceCuerpo
        With Bodys(BodysCountNew)
            .Body(1) = Animaciones(2)
            .Body(2) = Animaciones(4)
            .Body(3) = Animaciones(1)
            .Body(4) = Animaciones(3)
            
            
            Dim PosHead As Byte
            
            PosHead = Val(InputBox("La ropa indexada es para: " & vbCrLf & "1) Altos" & vbCrLf & "2) Enanos", "Seleccion de posición de cabeza"))
            
            If PosHead > 2 Or PosHead < 1 Then Exit Sub
            
            If PosHead = 1 Then
                .HeadOffsetX = 0
                .HeadOffsetY = -38
            Else
                .HeadOffsetX = 0
                .HeadOffsetY = -28
            End If
        End With

        MsgBox "La indexacion se realizo con Exito"
    
    Case 2
        Dim TipoHead As Byte
        
        TipoHead = Val(InputBox("Eliga tipo de Head:" & vbCrLf & vbCrLf & "1)Cabezas" & vbCrLf & "2)Cascos y Sombreros"))
        
        If TipoHead > 2 Or TipoHead < 1 Then Exit Sub
        
        If TipoHead = 1 Then
            HeadsCountNew = HeadsCountNew + 1
            ReDim Preserve Heads(0 To HeadsCountNew) As tIndiceCabeza
                Heads(HeadsCountNew).Head(1) = Animaciones(4)
                Heads(HeadsCountNew).Head(2) = Animaciones(2)
                Heads(HeadsCountNew).Head(3) = Animaciones(1)
                Heads(HeadsCountNew).Head(4) = Animaciones(3)
        Else
            CascosCountNew = CascosCountNew + 1
            ReDim Preserve Cascos(0 To CascosCountNew) As tIndiceCasco
                Cascos(CascosCountNew).Casco(1) = Animaciones(4)
                Cascos(CascosCountNew).Casco(2) = Animaciones(2)
                Cascos(CascosCountNew).Casco(3) = Animaciones(1)
                Cascos(CascosCountNew).Casco(4) = Animaciones(3)
        End If
        
        MsgBox "La indexacion se realizo con Exito"
        
    Case 4
        FxCountNew = FxCountNew + 1
        ReDim Preserve Fx(1 To FxCountNew) As tIndiceFx
        Fx(FxCountNew).Animacion = Animaciones(1)
        Fx(FxCountNew).OffsetX = Val(InputBox("Asignele posicion sobre X en la que se debe posicionar la animacion." & vbCrLf & "El valor Default(por lo general) es '0'"))
        Fx(FxCountNew).OffsetY = Val(InputBox("Asignele posicion sobre Y en la que se debe posicionar la animacion." & vbCrLf & "El valor Default(por lo general) es '0'"))
        
        MsgBox "La indexacion se realizo con Exito"
        
    Case 6
        BodysCountNew = BodysCountNew + 1
        ReDim Preserve Bodys(0 To BodysCountNew) As tIndiceCuerpo
        With Bodys(BodysCountNew)
            .Body(1) = Animaciones(2)
            .Body(2) = Animaciones(4)
            .Body(3) = Animaciones(1)
            .Body(4) = Animaciones(3)
            
            .HeadOffsetX = 0
            .HeadOffsetY = 0
        End With
        
        MsgBox "La indexacion se realizo con Exito"
        
    Case 8
        EscudosCountNew = EscudosCountNew + 1
        ReDim Preserve Escudos(1 To EscudosCountNew) As tIndiceEscudos
            Escudos(EscudosCountNew).Escudo(1) = Animaciones(2)
            Escudos(EscudosCountNew).Escudo(2) = Animaciones(4)
            Escudos(EscudosCountNew).Escudo(3) = Animaciones(1)
            Escudos(EscudosCountNew).Escudo(4) = Animaciones(3)
                
        MsgBox "La indexacion se realizo con Exito"
        
    Case 9
        ArmasCountNew = ArmasCountNew + 1
        ReDim Preserve Armas(1 To ArmasCountNew) As tIndiceArmas
            Armas(ArmasCountNew).Arma(1) = Animaciones(2)
            Armas(ArmasCountNew).Arma(2) = Animaciones(4)
            Armas(ArmasCountNew).Arma(3) = Animaciones(1)
            Armas(ArmasCountNew).Arma(4) = Animaciones(3)
        
        MsgBox "La indexacion se realizo con Exito"
        
End Select
Me.Visible = False
FrmIndex.Visible = True
End Sub

Private Sub Form_Activate()
Imagen.Picture = LoadPicture(Ruta)
If Imagen.Height = 16 Then
    Me.Width = Imagen.Width * 35
    Me.Height = Imagen.Height * 80
ElseIf Imagen.Height > 587 Or Imagen.Width > 713 Then
    Me.Width = 12690
    Me.Height = 10710
    cmdIndexacion.Left = 0
    cmdIndexacion.Top = Me.ScaleHeight - cmdIndexacion.Height
    MsgBox "La imagen es muy GRANDE... abrila y asegurate de la cantidad de animaciones a lo Largo y a lo Alto"
ElseIf Imagen.Height < 33 Or Imagen.Width < 33 Then
    Me.Width = 2000
    Me.Height = 1500
    cmdIndexacion.Left = 0
    cmdIndexacion.Top = Me.ScaleHeight - cmdIndexacion.Height
Else
    Me.Width = Imagen.Width * 15
    Me.Height = Imagen.Height * 23
End If
cmdIndexacion.Left = 0
cmdIndexacion.Top = Me.ScaleHeight - cmdIndexacion.Height
End Sub

Private Sub Form_Load()
Imagen.Picture = LoadPicture(Ruta)
If Imagen.Height = 16 Then
    Me.Width = Imagen.Width * 35
    Me.Height = Imagen.Height * 80
    cmdIndexacion.Left = 0
    cmdIndexacion.Top = Me.ScaleHeight - cmdIndexacion.Height
Else
    Me.Width = Imagen.Width * 15
    Me.Height = Imagen.Height * 23
    cmdIndexacion.Left = 0
    cmdIndexacion.Top = Me.ScaleHeight - cmdIndexacion.Height
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmIndex.Visible = True
End Sub

Public Sub SpecialIndex()
Dim T As Byte
Dim Y As Byte
Dim ContAnim As Integer
Dim Imag As Byte

    NumeroDeGrhIndex = AllGrhData + 1
    ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
    
If UnFrame = False Then
    'CuantosLargo = InputBox("Cuantas Imagenes a lo Largo tiene la imagen numero: " & SelectedImg(Imag) & "?", "Cantidad de Imagenes a lo Largo")
    'CuantosAlto = InputBox("Cuantas Imagenes a lo Alto tiene la imagen numero: " & SelectedImg(Imag) & "?", "Cantidad de Imagenes a lo Alto")
    
    CuantosLargo = Val(InputBox("Cuantas Imagenes a lo Largo tiene la imagen?", "Cantidad de Imagenes a lo Largo"))
    CuantosAlto = Val(InputBox("Cuantas Imagenes a lo Alto tiene la imagen?", "Cantidad de Imagenes a lo Alto"))
    
    If CuantosLargo < 1 Or CuantosAlto < 1 Then Exit Sub
    
    For Imag = 1 To UBound(SelectedImg())
        Imagen.Picture = LoadPicture(Config.BmpPath & "\" & SelectedImg(Imag) & ".bmp")
        AnimacionLargo = Int(Imagen.Width / CuantosLargo)
        AnimacionAlto = Int(Imagen.Height / CuantosAlto)
        ContAnim = 0
        ReDim Preserve GrhAnim(Imag).Animate(1 To CuantosLargo * CuantosAlto) As Long
        SumaAlto = 0
        For T = 1 To CuantosAlto
            For Y = 1 To CuantosLargo
                ContAnim = ContAnim + 1
                With GrhData(NumeroDeGrhIndex)
                    .NumFrames = 1
                    .FileNum = SelectedImg(Imag)
                    .sX = SumaLado
                    .sY = SumaAlto
                    .pixelWidth = AnimacionLargo
                    .pixelHeight = AnimacionAlto
                    GrhAnim(Imag).Animate(ContAnim) = NumeroDeGrhIndex
                End With
                SumaLado = SumaLado + AnimacionLargo
                NumeroDeGrhIndex = NumeroDeGrhIndex + 1
                ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
            Next Y
            SumaAlto = SumaAlto + AnimacionAlto
            SumaLado = 0
        Next T
    Next Imag
    
    For T = 1 To 4
        With GrhData(NumeroDeGrhIndex)
            .NumFrames = UBound(GrhAnim(T).Animate())
            ReDim Preserve GrhData(NumeroDeGrhIndex).Frames(1 To .NumFrames) As Long
            For Y = 1 To .NumFrames
                .Frames(Y) = GrhAnim(T).Animate(Y)
            Next Y
            If IndexMode = "12.1" Then
                .Speed = .NumFrames * 1000 / 18
            Else
                .Speed = 1
            End If
        End With
        Animaciones(T) = NumeroDeGrhIndex
        NumeroDeGrhIndex = NumeroDeGrhIndex + 1
        ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
    Next T
    
    BodysCountNew = BodysCountNew + 1
    ReDim Preserve Bodys(0 To BodysCountNew) As tIndiceCuerpo
    Bodys(BodysCountNew).Body(1) = Animaciones(2)
    Bodys(BodysCountNew).Body(2) = Animaciones(4)
    Bodys(BodysCountNew).Body(3) = Animaciones(1)
    Bodys(BodysCountNew).Body(4) = Animaciones(3)
    Bodys(BodysCountNew).HeadOffsetX = 0
    Bodys(BodysCountNew).HeadOffsetY = 0
    
    MsgBox "La indexacion se realizo con Exito"
    
Else
    For Imag = 1 To UBound(SelectedImg())
        ReDim Preserve GrhAnim(1).Animate(1 To UBound(SelectedImg())) As Long
        Imagen.Picture = LoadPicture(Config.BmpPath & "\" & SelectedImg(Imag) & ".bmp")
        With GrhData(NumeroDeGrhIndex)
            .NumFrames = 1
            .FileNum = SelectedImg(Imag)
            .sX = 0
            .sY = 0
            .pixelWidth = Imagen.Width
            .pixelHeight = Imagen.Height
            GrhAnim(1).Animate(Imag) = NumeroDeGrhIndex
        End With
        NumeroDeGrhIndex = NumeroDeGrhIndex + 1
        ReDim Preserve GrhData(1 To NumeroDeGrhIndex) As GrhData
    Next Imag
    
    With GrhData(NumeroDeGrhIndex)
        .NumFrames = UBound(SelectedImg())
        ReDim Preserve GrhData(NumeroDeGrhIndex).Frames(1 To .NumFrames) As Long
        For T = 1 To .NumFrames
            .Frames(T) = GrhAnim(1).Animate(T)
        Next T
        If IndexMode = "12.1" Then
            .Speed = .NumFrames * 1000 / 18
        Else
            .Speed = 1
        End If
    End With
    Animaciones(1) = NumeroDeGrhIndex
    
    FxCountNew = FxCountNew + 1
    ReDim Preserve Fx(1 To FxCountNew) As tIndiceFx
    Fx(FxCountNew).Animacion = Animaciones(1)
    Fx(FxCountNew).OffsetX = Val(InputBox("Asignele posicion sobre X en la que se debe posicionar la animacion." & vbCrLf & "El valor Default(por lo general) es '0'"))
    Fx(FxCountNew).OffsetY = Val(InputBox("Asignele posicion sobre Y en la que se debe posicionar la animacion." & vbCrLf & "El valor Default(por lo general) es '0'"))
    
    MsgBox "La indexacion se realizo con Exito"
End If
AllGrhData = NumeroDeGrhIndex
    
End Sub
