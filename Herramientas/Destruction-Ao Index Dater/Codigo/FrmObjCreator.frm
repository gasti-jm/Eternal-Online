VERSION 5.00
Begin VB.Form FrmObjCreator 
   Caption         =   "Creacion de Objetos"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9270
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmObjCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
FrmIndex.Visible = True
End Sub
