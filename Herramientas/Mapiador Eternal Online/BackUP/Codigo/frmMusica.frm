VERSION 5.00
Begin VB.Form frmMusica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Musica"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmMusica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
      _extentx        =   4260
      _extenty        =   873
      caption         =   "&Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":628A
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdAplicarYCerrar 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   2415
      _extentx        =   4260
      _extenty        =   873
      caption         =   "&Aplicar y Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":62B6
      mode            =   0
      value           =   0   'False
      enabled         =   0   'False
      cback           =   12648447
   End
   Begin WorldEditor.lvButtons_H cmdDetener 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      caption         =   "&Detener"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":62E2
      mode            =   0
      value           =   0   'False
      enabled         =   0   'False
      cback           =   12632319
   End
   Begin WorldEditor.lvButtons_H cmdEscuchar 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      caption         =   "&Escuchar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":630E
      mode            =   0
      value           =   0   'False
      enabled         =   0   'False
      cback           =   12648384
   End
   Begin VB.FileListBox fleMusicas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   120
      Pattern         =   "*.mid"
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMusica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit

Private MidiActual As String

''
' Aplica la Musica seleccionada y oculta la ventana
'

Private Sub cmdAplicarYCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If Len(MidiActual) >= 5 Then
    MapInfo.Music = Left(MidiActual, Len(MidiActual) - 4)
    frmMapInfo.txtMapMusica.text = MapInfo.Music
    frmMain.lblMapMusica = MapInfo.Music
    MidiActual = Empty
End If
Me.Hide
End Sub


''
' Oculta la ventana
'

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.Hide
End Sub

''
' Detiene la Musica que se encuentra Reproduciendo
'

Private Sub cmdDetener_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Stop_Midi
cmdEscuchar.Enabled = True
cmdDetener.Enabled = False
Play = False
End Sub

''
' Inicia la reproduccion de la Musica Seleccionada
'

Private Sub cmdEscuchar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Play_Midi
cmdDetener.Enabled = True
cmdEscuchar.Enabled = False
Play = True
End Sub

''
' Selecciona una nueva Musica del listado
'

Private Sub fleMusicas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MidiActual = fleMusicas.List(fleMusicas.ListIndex)
'CargarMIDI fleMusicas.Path & "\" & fleMusicas.List(fleMusicas.ListIndex)
cmdAplicarYCerrar.Enabled = True
If Play = False Then cmdEscuchar.Enabled = True
End Sub

