Attribute VB_Name = "Indexacion"
Option Explicit

Public Type tCabecera '---> Datos antes del Index(Creo)
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Type tIndiceCabeza '--->Datos de Cabezas
    Head(1 To 4) As Integer
End Type

Public Heads() As tIndiceCabeza
Public HeadsCountOld As Integer
Public HeadsCountNew As Integer
 
Type tIndiceCasco '--->Datos de Cascos
    Casco(1 To 4) As Integer
End Type

Public Cascos() As tIndiceCasco
Public CascosCountOld As Integer
Public CascosCountNew As Integer
 
Type tIndiceCuerpo '--->Datos de Cuerpos
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type
 
Public Bodys() As tIndiceCuerpo
Public BodysCountOld As Integer
Public BodysCountNew As Integer
 
Type tIndiceFx '--->Datos de FX
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public Fx() As tIndiceFx
Public FxCountOld As Integer
Public FxCountNew As Integer

Type tIndiceArmas '--->Datos de Cuerpos
    Arma(1 To 4) As Integer
End Type
 
Public Armas() As tIndiceArmas
Public ArmasCountOld As Integer
Public ArmasCountNew As Integer

Type tIndiceEscudos '--->Datos de Cuerpos
    Escudo(1 To 4) As Integer
End Type
 
Public Escudos() As tIndiceEscudos
Public EscudosCountOld As Integer
Public EscudosCountNew As Integer

Public Sub LoadCabezas()
    Dim N As Integer
    Dim Anim As Long
   
    N = FreeFile()
    Open Config.InitPath & "\Heads.ind" For Binary Access Read As #N
   
    'cabecera
    Get #N, , MiCabecera
   
    'num de cabezas
    Get #N, , HeadsCountOld
   
    ReDim Heads(0 To HeadsCountOld) As tIndiceCabeza
   
    For Anim = 1 To HeadsCountOld
        Get #N, , Heads(Anim)
    Next Anim
   
    HeadsCountNew = HeadsCountOld
   
    Close #N
End Sub
   
Public Sub SaveCabezas()
    Dim N As Integer
    Dim Anim As Long
   
    N = FreeFile()
    Open Config.SaveInitPath & "\Heads.ind" For Binary Access Write As #N
        'BORRA LA CABECERA
        Put #N, , MiCabecera
        
        Put #N, , HeadsCountNew
       
        ReDim Preserve Heads(0 To HeadsCountNew) As tIndiceCabeza
       
        For Anim = 1 To HeadsCountNew
            Put #N, , Heads(Anim)
        Next Anim
   
    Close #N
End Sub

Public Sub LoadFxs()
    Dim N As Integer
    Dim Anim As Long
   
    N = FreeFile()
    Open Config.InitPath & "\Fxs.ind" For Binary Access Read As #N
   
    'cabecera
    Get #N, , MiCabecera
   
    'num de FX
    Get #N, , FxCountOld
   
    'Resize array
    ReDim Fx(1 To FxCountOld) As tIndiceFx
   
    For Anim = 1 To FxCountOld
        Get #N, , Fx(Anim)
    Next Anim
   
    FxCountNew = FxCountOld
    Close #N
End Sub
   
Public Sub SaveFX()
    Dim N As Integer
    Dim Anim As Long
    
    N = FreeFile()
    Open Config.SaveInitPath & "\Fxs.ind" For Binary Access Write As #N
        'BORRA LA CABECERA
        Put #N, , MiCabecera
        
        Put #N, , FxCountNew
       
        ReDim Preserve Fx(1 To FxCountNew) As tIndiceFx
       
        For Anim = 1 To FxCountNew
            Put #N, , Fx(Anim)
        Next Anim
   
    Close #N
End Sub

Public Sub LoadCuerpos()
    Dim N As Integer
    Dim Anim As Long
   
    N = FreeFile()
    Open Config.InitPath & "\Bodys.ind" For Binary Access Read As #N
        
         'cabecera
         Get #N, , MiCabecera
        
         'num de cuerpos
         Get #N, , BodysCountOld
        
         'Resize array
         ReDim Bodys(0 To BodysCountOld) As tIndiceCuerpo
        
         For Anim = 1 To BodysCountOld
             Get #N, , Bodys(Anim)
         Next Anim
    Close #N
    
    BodysCountNew = BodysCountOld
End Sub

Public Sub SaveCuerpos()
    Dim N As Integer
    Dim Anims As Long
    
    N = FreeFile()
    Open Config.SaveInitPath & "\Bodys.ind" For Binary Access Write As #N
        Put #N, , MiCabecera
        
        Put #N, , BodysCountNew
       
        ReDim Preserve Bodys(0 To BodysCountNew) As tIndiceCuerpo
       
        For Anims = 1 To BodysCountNew
            Put #N, , Bodys(Anims)
        Next Anims
    Close #N
End Sub

Public Sub LoadCascos()
    Dim N As Integer
    Dim Anim As Long
   
    N = FreeFile()
    Open Config.InitPath & "\Helmets.ind " For Binary Access Read As #N
   
    'cabecera
    Get #N, , MiCabecera
   
    'num de cabezas
    Get #N, , CascosCountOld
   
    ReDim Cascos(0 To CascosCountOld) As tIndiceCasco
   
    For Anim = 1 To CascosCountOld
        Get #N, , Cascos(Anim)
    Next Anim
   
    CascosCountNew = CascosCountOld
   
    Close #N
End Sub
   
Public Sub SaveCascos()
    Dim N As Integer
    Dim Anim As Long
   
    N = FreeFile()
    Open Config.SaveInitPath & "\Helmets.ind" For Binary Access Write As #N
        
        Put #N, , MiCabecera
        
        Put #N, , CascosCountNew
       
        ReDim Preserve Cascos(0 To CascosCountNew) As tIndiceCasco
       
        For Anim = 1 To CascosCountNew
            Put #N, , Cascos(Anim)
        Next Anim
        
    Close #N
End Sub

Public Sub LoadArmas()
Dim Anims As Integer
Dim Moves As Byte

ArmasCountOld = Val(GetVar(Config.InitPath & "\Weapons.dat", "INIT", "NumArmas"))

ReDim Armas(1 To ArmasCountOld) As tIndiceArmas

For Anims = 1 To ArmasCountOld
    For Moves = 1 To 4
        Armas(Anims).Arma(Moves) = Val(GetVar(Config.InitPath & "\Weapons.dat", "Arma" & Anims, "Dir" & Moves))
    Next Moves
Next Anims

ArmasCountNew = ArmasCountOld

End Sub

Public Sub SaveArmas()
Dim Anims As Integer
Dim Moves As Byte

Call WriteVar(Config.SaveInitPath & "\Weapons.dat", "INIT", "NumArmas", ArmasCountNew)

ReDim Preserve Armas(1 To ArmasCountNew) As tIndiceArmas

For Anims = 1 To ArmasCountNew
    For Moves = 1 To 4
        Call WriteVar(Config.SaveInitPath & "\Weapons.dat", "Arma" & Anims, "Dir" & Moves, Armas(Anims).Arma(Moves))
    Next Moves
Next Anims
End Sub

Public Sub LoadEscudos()
Dim Anims As Integer
Dim Moves As Byte

EscudosCountOld = Val(GetVar(Config.InitPath & "\Shields.dat", "INIT", "NumEscudos"))

ReDim Escudos(1 To EscudosCountOld) As tIndiceEscudos

For Anims = 1 To EscudosCountOld
    For Moves = 1 To 4
        Escudos(Anims).Escudo(Moves) = Val(GetVar(Config.InitPath & "\Shields.dat", "ESC" & Anims, "Dir" & Moves))
    Next Moves
Next Anims

EscudosCountNew = EscudosCountOld

End Sub

Public Sub SaveEscudos()
Dim Anims As Integer
Dim Moves As Byte

Call WriteVar(Config.SaveInitPath & "\Shields.dat", "INIT", "NumEscudos", EscudosCountNew)

ReDim Preserve Escudos(1 To EscudosCountNew) As tIndiceEscudos

For Anims = 1 To EscudosCountNew
    For Moves = 1 To 4
        Call WriteVar(Config.SaveInitPath & "\Shields.dat", "ESC" & Anims, "Dir" & Moves, Escudos(Anims).Escudo(Moves))
    Next Moves
Next Anims
End Sub
