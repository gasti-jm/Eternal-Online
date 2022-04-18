Attribute VB_Name = "modSeguridadClones"
'[Rezniaq]
'El módulo 'modSeguridadClones' se encarga de limitar la cantidad de personajes
'que puede crear un mismo jugador en un determinado plazo de tiempo.

Option Explicit

Private Const limite_de_personajes_k    As Integer = 10

Private Type jugador_t

    ip_v                                As String
    personajes_creados_v                As Long

End Type

Private jugadores_m()                   As jugador_t

Public Sub seguridad_clones_construir()

    ReDim jugadores_m(0 To 0)

End Sub

Public Sub seguridad_clones_destruir()

    Erase jugadores_m()

End Sub

Public Function seguridad_clones_validar(ByVal ip_p As String) As Boolean

    Dim iterador_v As Long
  
    ip_p = UCase$(ip_p)
  
    For iterador_v = LBound(jugadores_m) To UBound(jugadores_m)
  
        With jugadores_m(iterador_v)
      
            If .ip_v = ip_p Then
          
                If .personajes_creados_v >= limite_de_personajes_k Then
              
                    seguridad_clones_validar = False
                    Exit Function
                  
                Else
              
                    .personajes_creados_v = .personajes_creados_v + 1
                  
                    seguridad_clones_validar = True
                    Exit Function
                  
                End If
          
            End If
      
        End With
      
    Next
  
    ReDim Preserve jugadores_m(LBound(jugadores_m) To UBound(jugadores_m) + 1)
  
    With jugadores_m(UBound(jugadores_m))
  
        .ip_v = ip_p
        .personajes_creados_v = 1
  
    End With

    seguridad_clones_validar = True

End Function

Public Sub seguridad_clones_limpiar()

    Erase jugadores_m()
    ReDim jugadores_m(0 To 0)

End Sub
