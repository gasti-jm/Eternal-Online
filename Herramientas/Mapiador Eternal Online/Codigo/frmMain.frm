VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H001E1E1E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WorldEditor"
   ClientHeight    =   9735
   ClientLeft      =   390
   ClientTop       =   840
   ClientWidth     =   14850
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   990
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picRadar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   120
      ScaleHeight     =   98.99
      ScaleMode       =   0  'User
      ScaleWidth      =   98.99
      TabIndex        =   99
      Top             =   120
      Width           =   1500
      Begin VB.Shape ApuntadorRadar 
         BorderColor     =   &H80000004&
         Height          =   225
         Left            =   480
         Top             =   480
         Width           =   300
      End
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   15360
      Top             =   3000
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   120
      Picture         =   "frmMain.frx":424A
      ScaleHeight     =   4365
      ScaleWidth      =   4365
      TabIndex        =   34
      Top             =   1800
      Width           =   4395
      Begin VB.ComboBox nBloqueos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":168AE
         Left            =   2880
         List            =   "frmMain.frx":168DF
         TabIndex        =   100
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   60
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   59
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   58
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":16916
         Left            =   1080
         List            =   "frmMain.frx":16926
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   56
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   55
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":16936
         Left            =   120
         List            =   "frmMain.frx":16938
         Sorted          =   -1  'True
         TabIndex        =   54
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":1693A
         Left            =   3360
         List            =   "frmMain.frx":1693C
         TabIndex        =   53
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":1693E
         Left            =   840
         List            =   "frmMain.frx":16940
         TabIndex        =   52
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":16942
         Left            =   120
         List            =   "frmMain.frx":16944
         TabIndex        =   51
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   50
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":16946
         Left            =   840
         List            =   "frmMain.frx":16948
         TabIndex        =   49
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":1694A
         Left            =   3360
         List            =   "frmMain.frx":1694C
         TabIndex        =   48
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   47
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":1694E
         Left            =   120
         List            =   "frmMain.frx":16950
         TabIndex        =   46
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":16952
         Left            =   120
         List            =   "frmMain.frx":16954
         TabIndex        =   45
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   44
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   43
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   42
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   41
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   40
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   39
         Top             =   0
         Width           =   0
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":16956
         Left            =   840
         List            =   "frmMain.frx":16958
         TabIndex        =   38
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":1695A
         Left            =   120
         List            =   "frmMain.frx":1695C
         TabIndex        =   36
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":1695E
         Left            =   3360
         List            =   "frmMain.frx":16960
         TabIndex        =   35
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   68
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   735
         Left            =   2400
         TabIndex        =   71
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   74
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   76
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   77
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   78
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   79
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   80
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   2400
         TabIndex        =   81
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   83
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   98
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   97
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   96
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   94
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   93
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   92
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   91
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   90
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   89
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   88
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   87
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   85
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   84
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.PictureBox PreviewGrh 
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      Height          =   3420
      Left            =   120
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   33
      Top             =   6240
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3420
      Left            =   120
      ScaleHeight     =   228
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6240
      Width           =   4455
   End
   Begin VB.TextBox StatTxt 
      Alignment       =   2  'Center
      BackColor       =   &H001E1E1E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2115
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "frmMain.frx":16962
      Top             =   7560
      Width           =   10080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H001E1E1E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000B&
      Height          =   1770
      Left            =   1560
      TabIndex        =   22
      Top             =   0
      Width           =   3105
      Begin WorldEditor.lvButtons_H cmdInformacionDelMapa 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Caption         =   "&Información del Mapa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cmdQuitarFunciones 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Quitar Todas las Funciones Activadas"
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         Caption         =   "&Quitar Funciones (F4)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632319
      End
      Begin VB.Label lblMapVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1680
         TabIndex        =   30
         Top             =   600
         Width           =   105
      End
      Begin VB.Label lblMapMusica 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1440
         TabIndex        =   29
         Top             =   352
         Width           =   90
      End
      Begin VB.Label lblMapNombre 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Mapa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1440
         TabIndex        =   28
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblFMusica 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Musica:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   105
         TabIndex        =   27
         Top             =   315
         Width           =   2895
      End
      Begin VB.Label lblFVersion 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Versión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblFNombreMapa 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre del Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   105
         TabIndex        =   25
         Top             =   60
         Width           =   2895
      End
   End
   Begin VB.PictureBox Renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   6000
      Left            =   4680
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   672
      TabIndex        =   0
      Top             =   1440
      Width           =   10080
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   15000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   11475
      TabIndex        =   14
      Top             =   0
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Particulas"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":169A2
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   12960
      TabIndex        =   15
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":175F4
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   9960
      TabIndex        =   16
      Top             =   0
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":17BBA
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8595
      TabIndex        =   17
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":180BB
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   7080
      TabIndex        =   18
      Top             =   0
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":1846F
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   5565
      TabIndex        =   19
      Top             =   0
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Translados (F6)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":187F0
      ImgSize         =   24
      cBack           =   16777215
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   4800
      TabIndex        =   20
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   12632319
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":1BE50
      ImgSize         =   24
      cBack           =   16777215
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   9480
      TabIndex        =   21
      Top             =   330
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":1F396
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   13980
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   13215
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5565
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6330
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7095
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7860
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8625
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9390
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   10155
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   10920
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   11685
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   12450
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarcomoBMP 
         Caption         =   "Guardar Render en &BMP"
      End
      Begin VB.Menu mnuGuardarcomoJPG 
         Caption         =   "Guardar Render en &JPG"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuComo 
         Caption         =   "¿ Como seleccionar ? ---- Mantener SHIFT y arrastrar el cursor."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBloquearS 
         Caption         =   "&Bloquear Selección"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRealizarOperacion 
         Caption         =   "&Realizar Operación en Selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeshacerPegado 
         Caption         =   "Deshacer P&egado de Selección"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuParticle 
         Caption         =   "...&Particle"
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPanelesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFuncionesLine1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "&Actualizar Indices de..."
         Begin VB.Menu mnuActualizarSuperficies 
            Caption         =   "&Superficies"
         End
         Begin VB.Menu mnuActualizarNPCs 
            Caption         =   "&NPC's"
         End
         Begin VB.Menu mnuActualizarObjs 
            Caption         =   "&Objetos"
         End
         Begin VB.Menu mnuActualizarTriggers 
            Caption         =   "&Trigger's"
         End
         Begin VB.Menu mnuActualizarCabezas 
            Caption         =   "C&abezas"
         End
         Begin VB.Menu mnuActualizarCuerpos 
            Caption         =   "C&uerpos"
         End
         Begin VB.Menu mnuActualizarGraficos 
            Caption         =   "&Graficos"
         End
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "&GRH => BMP"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarUltimaConfig 
         Caption         =   "&Guardar Ultima Configuración"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuManual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLineAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
End
Attribute VB_Name = "frmMain"
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
Public MouseX As Integer
Public MouseY As Integer

Public Sub PonerAlAzar(ByVal n As Integer, T As Byte)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 by GS
'*************************************************
Dim OBJIndex As Long
Dim NPCIndex As Long
Dim X, Y, i
Dim Head As Integer
Dim Body As Integer
Dim Heading As Byte
Dim Leer As New clsIniReader
i = n

modEdicion.Deshacer_Add "Aplicar " & IIf(T = 0, "Objetos", "NPCs") & " al Azar" ' Hago deshacer

Do While i > 0
    X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
    Y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
    Select Case T
        Case 0
            If MapData(X, Y).OBJInfo.OBJIndex = 0 Then
                  i = i - 1
                  If frmMain.cInsertarBloqueo.value = True Then
                    MapData(X, Y).Blocked = frmMain.nBloqueos.text
                  Else
                    MapData(X, Y).Blocked = 0
                  End If
                  If frmMain.cNumFunc(2).text > 0 Then
                      OBJIndex = frmMain.cNumFunc(2).text
                      InitGrh MapData(X, Y).ObjGrh, ObjData(OBJIndex).grhindex
                      MapData(X, Y).OBJInfo.OBJIndex = OBJIndex
                      MapData(X, Y).OBJInfo.Amount = Val(frmMain.cCantFunc(2).text)
                      Select Case ObjData(OBJIndex).ObjType ' GS
                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                      End Select
                  End If
            End If
        Case 1
           If MapData(X, Y).Blocked = 0 Then
                  i = i - 1
                  If frmMain.cNumFunc(T - 1).text > 0 Then
                        NPCIndex = frmMain.cNumFunc(T - 1).text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex
                  End If
            End If
        Case 2
           If MapData(X, Y).Blocked = 0 Then
                  i = i - 1
                  If frmMain.cNumFunc(T - 1).text >= 0 Then
                        NPCIndex = frmMain.cNumFunc(T - 1).text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex
                  End If
           End If
        End Select
        DoEvents
Loop
End Sub







Private Sub Command1_Click()
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(0, 255, 255)
rgb_list(1) = RGB(0, 255, 255)
rgb_list(2) = RGB(0, 255, 255)
rgb_list(3) = RGB(0, 255, 255)
 
Dim grh_list(1) As Long
grh_list(1) = 15
Engine.Particle_Group_Create 50, 50, grh_list(), rgb_list(), 50, 4, 1, -1, 0, , 0, 0, 0, -20, 20, -10, -50, 10, 50, 8, 1, 1, 2, -15, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0
 
End Sub


Private Sub Form_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.SetFocus

End Sub

Private Sub Form_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
Dim tX As Byte
Dim tY As Byte

If Not MapaCargado Then Exit Sub

If SobreX > 0 And SobreY > 0 Then
    DobleClick Val(SobreX), Val(SobreY)
End If
End Sub






Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Seleccionando Then CopiarSeleccion
End Sub

Private Sub mnuActualizarCabezas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modDX8FIFO.CargarCabezas
End Sub

Private Sub mnuActualizarCuerpos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modDX8FIFO.CargarCuerpos
End Sub

Private Sub mnuActualizarGraficos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modDX8FIFO.LoadGrhData
End Sub

Private Sub mnuActualizarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesSuperficie
End Sub

Private Sub mnuAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dialog.CancelError = True
On Error GoTo ErrHandler

DeseaGuardarMapa Dialog.FileName

ObtenerNombreArchivo False

If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If
    
    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    Call DrawMiniMap
Exit Sub
ErrHandler:
End Sub

Private Sub mnuacercade_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAbout.Show
End Sub



Private Sub mnuActualizarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesNPC
End Sub

Private Sub mnuActualizarObjs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesOBJ
End Sub

Private Sub mnuActualizarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesTriggers
End Sub

Private Sub mnuAutoCapturarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)
End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

End Sub

Private Sub mnuAutoGuardarMapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAutoGuardarMapa.Show
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

End Sub

Private Sub mnuBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 2 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next

modPaneles.VerFuncion 2, True
End Sub

Private Sub mnuBloquearBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloquear_Bordes
End Sub

Private Sub mnuBloquearMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(1)
End Sub

Private Sub mnuBloquearS_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Bloquear Selección")
Call BlockearSeleccion
End Sub

Private Sub mnuConfigAvanzada_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmConfigSup.Show
End Sub

Private Sub mnuConfigObjTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Cfg_TrOBJ = frmMain.cNumFunc(2).text
End Sub

Private Sub mnuCopiar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call CopiarSeleccion
End Sub

Private Sub mnuCortar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Cortar Selección")
Call CortarSeleccion
End Sub

Private Sub mnuDeshacer_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Call modEdicion.Deshacer_Recover
End Sub

Private Sub mnuDeshacerPegado_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Deshacer Pegado de Selección")
Call DePegar
End Sub

Private Sub mnuGRHaBMP_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
frmGRHaBMP.Show
End Sub

Private Sub mnuGuardarcomoBMP_Click()
Radio = Val(InputBox("Escriba la escala de 1 a 5 en la que generemos su mapa", "la escala se multiplica x 32"))
If Radio = 0 Then Radio = 1
 
frmRenderer.Picture1.Width = (Radio * 100)
frmRenderer.Picture1.Height = (Radio * 100)
 
frmRenderer.Picture2.Width = (Radio * 100)
frmRenderer.Picture2.Height = (Radio * 100)
 
frmRenderer.Show
End Sub

Private Sub mnuGuardarcomoJPG_Click()
'*************************************************
'Author: Salvito
'Last modified: 01/05/2008 - ^[GS]^
'*************************************************
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))
    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then
        Engine.RenderToPicture
    End If
End Sub

Private Sub mnuGuardarMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa Dialog.FileName
End Sub

Private Sub mnuGuardarMapaComo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa
End Sub

Private Sub mnuGuardarUltimaConfig_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 23/05/06
'*************************************************
mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)
End Sub

Private Sub mnuInfoMap_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMapInfo.Show
frmMapInfo.Visible = True
End Sub

Private Sub mnuInformes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmInformes.Show
End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Azar
End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Bordes
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Todo
End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub

Private Sub mnuManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
If LenB(Dir(App.path & "\manual\index.html", vbArchive)) <> 0 Then
    Call Shell("explorer " & App.path & "\manual\index.html")
    DoEvents
End If
End Sub

Private Sub mnuModoCaminata_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
ToggleWalkMode
End Sub

Private Sub mnuNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 3 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 3, True
End Sub



'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer
DeseaGuardarMapa Dialog.FileName

For loopc = 0 To frmMain.MapPest.Count
    frmMain.MapPest(loopc).Visible = False
Next

frmMain.Dialog.FileName = Empty

If WalkMode = True Then
    Call modGeneral.ToggleWalkMode
End If

Call modMapIO.NuevoMapa

Call frmMain.cmdInformacionDelMapa_Click
Call DrawMiniMap

End Sub

Private Sub mnuObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 5 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 5, True
End Sub


Private Sub mnuOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
frmOptimizar.Show
End Sub

Private Sub mnuParticle_Click()
    mnuParticle.Checked = Not mnuParticle.Checked
End Sub

Private Sub mnuPegar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Pegar Selección")
Call PegarSeleccion
End Sub

Private Sub mnuQBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 2, False
End Sub

Private Sub mnuQNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 3, False
End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 5, False
End Sub

Private Sub mnuQSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 0, False
End Sub

Private Sub mnuQTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 1, False
End Sub

Private Sub mnuQTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 6, False
End Sub


Private Sub mnuQuitarBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(0)
End Sub

Public Sub mnuQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Superficies
frmMain.cSeleccionarSuperficie.value = False
Call frmMain.cSeleccionarSuperficie_Click
frmMain.cQuitarEnEstaCapa.value = False
Call frmMain.cQuitarEnEstaCapa_Click
frmMain.cQuitarEnTodasLasCapas.value = False
Call frmMain.cQuitarEnTodasLasCapas_Click
' Translados
frmMain.cQuitarTrans.value = False
Call frmMain.cQuitarTrans_Click
frmMain.cInsertarTrans.value = False
Call frmMain.cInsertarTrans_Click
' Bloqueos
frmMain.cQuitarBloqueo.value = False
Call frmMain.cQuitarBloqueo_Click
frmMain.cInsertarBloqueo.value = False
Call frmMain.cInsertarBloqueo_Click
' Otras funciones
frmMain.cInsertarFunc(0).value = False
Call frmMain.cInsertarFunc_Click(0)
frmMain.cInsertarFunc(1).value = False
Call frmMain.cInsertarFunc_Click(1)
frmMain.cInsertarFunc(2).value = False
Call frmMain.cInsertarFunc_Click(2)
frmMain.cQuitarFunc(0).value = False
Call frmMain.cQuitarFunc_Click(0)
frmMain.cQuitarFunc(1).value = False
Call frmMain.cQuitarFunc_Click(1)
frmMain.cQuitarFunc(2).value = False
Call frmMain.cQuitarFunc_Click(2)
' Triggers
frmMain.cInsertarTrigger.value = False
Call frmMain.cInsertarTrigger_Click
frmMain.cQuitarTrigger.value = False
Call frmMain.cQuitarTrigger_Click
End Sub

Private Sub mnuQuitarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_NPCs(False)
End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Objetos
End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Bordes
End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Capa(frmMain.cCapas.text)
End Sub

Private Sub mnuQuitarTODO_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Borrar_Mapa
End Sub

Private Sub mnuQuitarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Call modEdicion.Quitar_Translados
End Sub

Private Sub mnuQuitarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Triggers
End Sub

Private Sub mnuReAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error GoTo ErrHandler
    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName
        End If
    End If
    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
Exit Sub
ErrHandler:
End Sub

Private Sub mnuRealizarOperacion_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Realizar Operación en Selección")
Call AccionSeleccion
End Sub

Private Sub mnuSalir_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

Private Sub mnuSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 0 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 0, True
End Sub

Private Sub mnuTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 1 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 1, True
End Sub

Private Sub mnuTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 6 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 6, True
End Sub

Private Sub mnuUtilizarDeshacer_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)
End Sub


Private Sub mnuVerAutomatico_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)
End Sub

Private Sub mnuVerBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMain.cVerBloqueos.value = (frmMain.cVerBloqueos.value = False)
mnuVerBloqueos.Checked = frmMain.cVerBloqueos.value

End Sub

Private Sub mnuVerCapa1_Click()
mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
End Sub

Private Sub mnuVerCapa2_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
End Sub

Private Sub mnuVerCapa3_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
End Sub

Private Sub mnuVerCapa4_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
End Sub


Private Sub mnuVerGrilla_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
VerGrilla = (VerGrilla = False)
mnuVerGrilla.Checked = VerGrilla
End Sub

Private Sub mnuVerNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

End Sub

Private Sub mnuVerObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)

End Sub

Private Sub mnuVerTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

End Sub

Private Sub mnuVerTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMain.cVerTriggers.value = (frmMain.cVerTriggers.value = False)
mnuVerTriggers.Checked = frmMain.cVerTriggers.value
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'Last modified: 20/11/07 - Loopzer
'*************************************************

Dim tX As Integer
Dim tY As Integer

If Not MapaCargado Then Exit Sub

ConvertCPtoTP X, Y, tX, tY


If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionIX = tX '+ UserPos.X
    SeleccionIY = tY '+ UserPos.Y
Else
    ClickEdit Button, tX, tY
End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'*************************************************

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If Not MapaCargado Then Exit Sub
HotKeysAllow = True

'Make sure click is in view window
'If X <= Renderer.Left Or X >= Renderer.Left + MainViewWidth Or Y <= Renderer.Top Or Y >= Renderer.Top + MainViewHeight Then
'    Exit Sub
'End If

ConvertCPtoTP X, Y, tX, tY


frmMain.Caption = "Eternal Online Map-Editor v1.0 DX8 - X: " & tX & " - Y: " & tY

 If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionFX = tX '+ TileX
    SeleccionFY = tY '+ TileY
Else
    ClickEdit Button, tX, tY
End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************

' Guardar configuración
WriteVar IniPath & "MapEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")
If frmMain.mnuGuardarUltimaConfig.Checked = True Then
    WriteVar IniPath & "MapEditor.ini", "PATH", "UltimoMapa", Dialog.FileName
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.Y
    WriteVar IniPath & "MapEditor.ini", "CONFIGURACION", "UtilizarDeshacer", IIf(frmMain.mnuUtilizarDeshacer.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
    WriteVar IniPath & "MapEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)
End If

'Allow MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub

Private Sub Renderer_Click()
    Me.SetFocus
End Sub
Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If X < 11 Then X = 11
If X > 189 Then X = 189
If Y < 10 Then Y = 10
If Y > 192 Then Y = 192
UserPos.X = X
UserPos.Y = Y
bRefreshRadar = True
End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
MiRadarX = X
MiRadarY = Y
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
' HotKeys
If HotKeysAllow = False Then Exit Sub

Select Case UCase(Chr(KeyAscii))
    Case "S" ' Activa/Desactiva Insertar Superficie
        frmMain.cSeleccionarSuperficie.value = (frmMain.cSeleccionarSuperficie.value = False)
        Call frmMain.cSeleccionarSuperficie_Click
    Case "T" ' Activa/Desactiva Insertar Translados
        frmMain.cInsertarTrans.value = (frmMain.cInsertarTrans.value = False)
        Call frmMain.cInsertarTrans_Click
    Case "B" ' Activa/Desactiva Insertar Bloqueos
        frmMain.cInsertarBloqueo.value = (frmMain.cInsertarBloqueo.value = False)
        Call frmMain.cInsertarBloqueo_Click
    Case "N" ' Activa/Desactiva Insertar NPCs
        frmMain.cInsertarFunc(0).value = (frmMain.cInsertarFunc(0).value = False)
        Call frmMain.cInsertarFunc_Click(0)
   ' Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
   '     cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
   '     Call cInsertarFunc_Click(1)
    Case "O" ' Activa/Desactiva Insertar Objetos
        frmMain.cInsertarFunc(2).value = (frmMain.cInsertarFunc(2).value = False)
        Call frmMain.cInsertarFunc_Click(2)
    Case "G" ' Activa/Desactiva Insertar Triggers
        frmMain.cInsertarTrigger.value = (frmMain.cInsertarTrigger.value = False)
        Call frmMain.cInsertarTrigger_Click
    Case "Q" ' Quitar Funciones
        Call frmMain.mnuQuitarFunciones_Click
End Select
End Sub
Private Sub Renderer_DblClick()
    Call Form_DblClick
End Sub

Private Sub Renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X, Y)
    Call DrawMiniMap(True)
End Sub

Private Sub Renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, X, Y)
    MouseX = X
    MouseY = Y
End Sub

Private Sub Renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseUp(Button, Shift, X, Y)
End Sub


Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

With Dialog
    .Filter = "Mapas de Argentum Online (*.map)|*.map"
    If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
    Else
        .DialogTitle = "Cargar"
        .FileName = vbNullString
        .flags = cdlOFNFileMustExist
        .ShowOpen
    End If
End With
End Sub

Public Sub cQuitarFunc_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarFunc(index).value = True Then
    cInsertarFunc(index).Enabled = False
    cAgregarFuncalAzar(index).Enabled = False
    cCantFunc(index).Enabled = False
    cNumFunc(index).Enabled = False
    cFiltro((index) + 1).Enabled = False
    lListado((index) + 1).Enabled = False
    Call modPaneles.EstSelectPanel((index) + 3, True)
Else
    cInsertarFunc(index).Enabled = True
    cAgregarFuncalAzar(index).Enabled = True
    cCantFunc(index).Enabled = True
    cNumFunc(index).Enabled = True
    cFiltro((index) + 1).Enabled = True
    lListado((index) + 1).Enabled = True
    Call modPaneles.EstSelectPanel((index) + 3, False)
End If
End Sub
Public Sub cQuitarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrigger.value = True Then
    lListado(4).Enabled = False
    cInsertarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(6, True)
Else
    lListado(4).Enabled = True
    cInsertarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(6, False)
End If
End Sub
Public Sub cQuitarEnTodasLasCapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarEnTodasLasCapas.value = True Then
    cCapas.Enabled = False
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    cCapas.Enabled = True
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub
Private Sub cAgregarFuncalAzar_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If IsNumeric(cCantFunc(index).text) = False Or cCantFunc(index).text > 100 Then
    MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 100.", vbCritical
    Exit Sub
End If
cAgregarFuncalAzar(index).Enabled = False
Call frmMain.PonerAlAzar(CInt(cCantFunc(index).text), 1 + (IIf(index = 2, -1, index)))
cAgregarFuncalAzar(index).Enabled = True
End Sub
Public Sub cQuitarEnEstaCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarEnEstaCapa.value = True Then
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnTodasLasCapas.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnTodasLasCapas.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub
Private Sub cNumFunc_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next

If KeyAscii = 13 Then
    Dim Cont As String
    Cont = frmMain.cNumFunc(index).text
    Call frmMain.cNumFunc_LostFocus(index)
    If Cont <> frmMain.cNumFunc(index).text Then Exit Sub
    If frmMain.cNumFunc(index).ListCount > 5 Then
        frmMain.cNumFunc(index).RemoveItem 0
    End If
    frmMain.cNumFunc(index).AddItem frmMain.cNumFunc(index).text
    Exit Sub
ElseIf KeyAscii = 8 Then
    
ElseIf IsNumeric(Chr(KeyAscii)) = False Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub cNumFunc_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If cNumFunc(index).text = vbNullString Then
    frmMain.cNumFunc(index).text = IIf(index = 1, 500, 1)
End If
End Sub

Public Sub cNumFunc_LostFocus(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If index = 0 Then
        If frmMain.cNumFunc(index).text > 499 Or frmMain.cNumFunc(index).text < 1 Then
            frmMain.cNumFunc(index).text = 1
        End If
    ElseIf index = 1 Then
        If frmMain.cNumFunc(index).text < 100 Or frmMain.cNumFunc(index).text > 32000 Then
            frmMain.cNumFunc(index).text = 100
        End If
    ElseIf index = 2 Then
        If frmMain.cNumFunc(index).text < 1 Or frmMain.cNumFunc(index).text > 32000 Then
            frmMain.cNumFunc(index).text = 1
        End If
    End If
End Sub
Private Sub cGrh_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
If KeyAscii = 13 Then
    Call fPreviewGrh(cGrh.text)
    If frmMain.cGrh.ListCount > 5 Then
        frmMain.cGrh.RemoveItem 0
    End If
    frmMain.cGrh.AddItem frmMain.cGrh.text
End If
Exit Sub
Fallo:
    cGrh.text = 1

End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub cFiltro_GotFocus(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
HotKeysAllow = False
End Sub

Private Sub cFiltro_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If KeyAscii = 13 Then
    Call Filtrar(index)
End If
End Sub

Private Sub cFiltro_LostFocus(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
HotKeysAllow = True
End Sub
Public Sub cInsertarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cInsertarBloqueo.value = True Then
    cQuitarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cQuitarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub

Public Sub cQuitarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cQuitarBloqueo.value = True Then
    cInsertarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cInsertarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub
Private Sub cverBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMain.mnuVerBloqueos.Checked = cVerBloqueos.value
End Sub

Private Sub cVerTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMain.mnuVerTriggers.Checked = cVerTriggers.value
End Sub

Private Sub lListado_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
On Error Resume Next
If HotKeysAllow = False Then
    lListado(index).Tag = lListado(index).ListIndex
    Select Case index
        Case 0
            cGrh.text = DameGrhIndex(ReadField(2, lListado(index).text, Asc("#")))
            If SupData(ReadField(2, lListado(index).text, Asc("#"))).Capa <> 0 Then
                If LenB(ReadField(2, lListado(index).text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.text
                cCapas.text = SupData(ReadField(2, lListado(index).text, Asc("#"))).Capa
            Else
                If LenB(cCapas.Tag) <> 0 Then
                    cCapas.text = cCapas.Tag
                    cCapas.Tag = vbNullString
                End If
            End If
            If SupData(ReadField(2, lListado(index).text, Asc("#"))).Block = True Then
                If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.value = True, 1, 0)
                cInsertarBloqueo.value = True
                Call cInsertarBloqueo_Click
            Else
                If LenB(cInsertarBloqueo.Tag) <> 0 Then
                    cInsertarBloqueo.value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                    cInsertarBloqueo.Tag = vbNullString
                    Call cInsertarBloqueo_Click
                End If
            End If
            Call fPreviewGrh(cGrh.text)
            Call modPaneles.VistaPreviaDeSup
        Case 1
            cNumFunc(0).text = ReadField(2, lListado(index).text, Asc("#"))
        Case 2
            cNumFunc(1).text = ReadField(2, lListado(index).text, Asc("#"))
        Case 3
            cNumFunc(2).text = ReadField(2, lListado(index).text, Asc("#"))
    End Select
Else
    lListado(index).ListIndex = lListado(index).Tag
End If

End Sub

Private Sub lListado_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If index = 3 And Button = 2 Then
    If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc
End If
End Sub

Private Sub lListado_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
On Error Resume Next
HotKeysAllow = False
End Sub
Public Sub cInsertarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
If cInsertarTrans.value = True Then
    cQuitarTrans.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    cQuitarTrans.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub
Private Sub cUnionManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarTrans.value = (cUnionManual.value = True)
Call cInsertarTrans_Click
End Sub
Private Sub cUnionAuto_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub
Private Sub cCantFunc_Change(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If Val(cCantFunc(index)) < 1 Then
      cCantFunc(index).text = 1
    End If
    If Val(cCantFunc(index)) > 10000 Then
      cCantFunc(index).text = 10000
    End If
End Sub
Private Sub cCapas_Change()
'*************************************************
'Author: ^[GS]^
'Last modified: 31/05/06
'*************************************************
    If Val(cCapas.text) < 1 Then
      cCapas.text = 1
    End If
    If Val(cCapas.text) > 4 Then
      cCapas.text = 4
    End If
    cCapas.Tag = vbNullString
End Sub

Public Sub cQuitarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrans.value = True Then
    cInsertarTransOBJ.Enabled = False
    cInsertarTrans.Enabled = False
    cUnionManual.Enabled = False
    cUnionAuto.Enabled = False
    tTMapa.Enabled = False
    tTX.Enabled = False
    tTY.Enabled = False
    frmMain.mnuInsertarTransladosAdyasentes.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    tTMapa.Enabled = True
    tTX.Enabled = True
    tTY.Enabled = True
    cUnionAuto.Enabled = True
    cUnionManual.Enabled = True
    cInsertarTrans.Enabled = True
    cInsertarTransOBJ.Enabled = True
    frmMain.mnuInsertarTransladosAdyasentes.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub
Public Sub cmdQuitarFunciones_Click()
Call frmMain.mnuQuitarFunciones_Click
End Sub
Public Sub cmdInformacionDelMapa_Click()
frmMapInfo.Show
frmMapInfo.Visible = True
End Sub
Public Sub cInsertarFunc_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarFunc(index).value = True Then
    cQuitarFunc(index).Enabled = False
    cAgregarFuncalAzar(index).Enabled = False
    If index <> 2 Then cCantFunc(index).Enabled = False
    Call modPaneles.EstSelectPanel((index) + 3, True)
Else
    cQuitarFunc(index).Enabled = True
    cAgregarFuncalAzar(index).Enabled = True
    If index <> 2 Then cCantFunc(index).Enabled = True
    Call modPaneles.EstSelectPanel((index) + 3, False)
End If
End Sub
Public Sub cInsertarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarTrigger.value = True Then
    cQuitarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(6, True)
Else
    cQuitarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(6, False)
End If
End Sub
Private Sub TimAutoGuardarMapa_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If mnuAutoGuardarMapas.Checked = True Then
    bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1
    If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
        If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
            modMapIO.GuardarMapa Dialog.FileName
        End If
        bAutoGuardarMapaCount = 0
    End If
End If
End Sub
Public Sub cSeleccionarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cSeleccionarSuperficie.value = True Then
    cQuitarEnTodasLasCapas.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    cQuitarEnTodasLasCapas.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub
Private Sub MapPest_Click(index As Integer)
If (index + NumMap_Save - 4) <> NumMap_Save Then
    Dialog.CancelError = True
    On Error GoTo ErrHandler
    Dialog.FileName = PATH_Save & NameMap_Save & (index + NumMap_Save - 4) & ".map"
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName
        End If
    End If
        Call modMapIO.NuevoMapa
        DoEvents
        modMapIO.AbrirMapa Dialog.FileName
        EngineRun = True
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description

End If
End Sub



Private Sub SelectPanel_Click(index As Integer)
Dim i As Byte
For i = 0 To 7
    If i <> index Then
        SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
If frmMain.mnuAutoQuitarFunciones.Checked = True Then Call frmMain.mnuQuitarFunciones_Click
Call VerFuncion(index, SelectPanel(index).value)
End Sub

