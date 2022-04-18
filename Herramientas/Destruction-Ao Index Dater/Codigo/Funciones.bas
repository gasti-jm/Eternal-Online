Attribute VB_Name = "Funciones"
'Option Explicit
'Public Sub AllCharge()

'Call LoadCuerpos
'Call LoadFxs
'Call LoadCabezas
'Call LoadCascos
'Call LoadGrhData(Config.InitPath)
'Call LoadArmas
'Call LoadEscudos

'Dim Cuerpos As Integer
'Dim Efectos As Integer
'Dim Cabeza As Integer
'Dim Casco As Integer
'Dim Grhs As Integer
'Dim Shields As Integer
'Dim Weapon As Integer

'For Cuerpos = 1 To BodysCountOld
'    If Bodys(Cuerpos).Body(1) > 0 Then
'        LstCuerpos.AddItem Cuerpos
'    End If
'Next Cuerpos

'For Efectos = 1 To FxCountOld
'    If Fx(Efectos).Animacion > 0 Then
'        LstFx.AddItem Efectos
'    End If
'Next Efectos

'For Cabeza = 1 To HeadsCountOld
'    If Heads(Cabeza).Head(1) > 0 Then
'        LstCabezas.AddItem Cabeza
'    End If
'Next Cabeza

'For Casco = 1 To CascosCountOld
'    If Cascos(Casco).Casco(1) > 0 Then
'        LstCascos.AddItem Casco
'    End If
'Next Casco

'For Shields = 1 To EscudosCountOld
'    If Escudos(Shields).Escudo(1) > 0 Then
'        LstEscudos.AddItem Shields
'    End If
'Next Shields

'For Weapon = 1 To ArmasCountOld
'    If Armas(Weapon).Arma(1) > 0 Then
'        LstArmas.AddItem Weapon
'    End If
'Next Weapon

'Dim Cargando As Integer
'Dim EstadoCarga As Byte
'Dim ImgCarga As Integer

'ImgCarga = CInt(477 / 10)
'Cargando = CInt(AllGrhData / 10)
'EstadoCarga = 1
'For Grhs = 1 To AllGrhData
'    If GrhData(Grhs).NumFrames > 1 Then
'        LstGeneral.AddItem Grhs & "(ANIMACION)"
'    Else
'        If GrhData(Grhs).NumFrames <> 0 Then
'            LstGeneral.AddItem Grhs
'        End If
'    End If
'    If CInt(Grhs / Cargando) >= EstadoCarga Then
'        ImgCargando.Width = ImgCarga * EstadoCarga
'        EstadoCarga = EstadoCarga + 1
'    End If
'Next Grhs

'PbCargando.Visible = False
'LstCuerpos.Visible = True
'LstFx.Visible = True
'LstCabezas.Visible = True
'LstCascos.Visible = True
'LstEscudos.Visible = True
'LstArmas.Visible = True
'LstGeneral.Visible = True

'End Sub
