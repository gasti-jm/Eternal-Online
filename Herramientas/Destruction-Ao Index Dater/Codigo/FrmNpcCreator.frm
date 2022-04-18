VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmNpcCreator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creacion de Npc"
   ClientHeight    =   8220
   ClientLeft      =   105
   ClientTop       =   705
   ClientWidth     =   9660
   Icon            =   "FrmNpcCreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmNpcCreator.frx":08CA
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtDomable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   40
      Top             =   1500
      Width           =   975
   End
   Begin VB.TextBox TxtInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   39
      Top             =   480
      Width           =   5655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   2760
      TabIndex        =   24
      Top             =   3360
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Cabeza/Cuerpo"
      TabPicture(0)   =   "FrmNpcCreator.frx":15ACFE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label24"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtBody"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtHead"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "PbHB(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "PbHB(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "PbHB(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "PbHB(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtSearchHB"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LstHB"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "THB"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CbHeading"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Objetos"
      TabPicture(1)   =   "FrmNpcCreator.frx":15AD1A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label26"
      Tab(1).Control(1)=   "Label27"
      Tab(1).Control(2)=   "Label28"
      Tab(1).Control(3)=   "Label29"
      Tab(1).Control(4)=   "OpName2"
      Tab(1).Control(5)=   "OpNum2"
      Tab(1).Control(6)=   "LstObjName"
      Tab(1).Control(7)=   "TxtObjSearch"
      Tab(1).Control(8)=   "LstObj"
      Tab(1).Control(9)=   "TxtCantObj"
      Tab(1).Control(10)=   "CmdAddObj"
      Tab(1).Control(11)=   "LstObjAdd"
      Tab(1).Control(12)=   "CmdRemoverObj"
      Tab(1).Control(13)=   "CmdListClear"
      Tab(1).Control(14)=   "TxtObjCant"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Hechizos"
      TabPicture(2)   =   "FrmNpcCreator.frx":15AD36
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label31"
      Tab(2).Control(1)=   "Label33"
      Tab(2).Control(2)=   "Label1"
      Tab(2).Control(3)=   "TxtCantSpell"
      Tab(2).Control(4)=   "CmdListClearSpell"
      Tab(2).Control(5)=   "CmdRemoverSpell"
      Tab(2).Control(6)=   "LstSpellAdd"
      Tab(2).Control(7)=   "CmdAddSpell"
      Tab(2).Control(8)=   "TxtSpSearch"
      Tab(2).Control(9)=   "OpNum3"
      Tab(2).Control(10)=   "OpName3"
      Tab(2).Control(11)=   "LstSpellName"
      Tab(2).Control(12)=   "LstSpell"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Agregados"
      TabPicture(3)   =   "FrmNpcCreator.frx":15AD52
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LstAdd"
      Tab(3).Control(1)=   "TxtAdd"
      Tab(3).Control(2)=   "CmdCargar"
      Tab(3).Control(3)=   "CmdAgregar"
      Tab(3).Control(4)=   "CmdChange"
      Tab(3).Control(5)=   "LblAdd"
      Tab(3).Control(6)=   "Label30"
      Tab(3).ControlCount=   7
      Begin VB.ListBox LstSpell 
         Height          =   3180
         ItemData        =   "FrmNpcCreator.frx":15AD6E
         Left            =   -74760
         List            =   "FrmNpcCreator.frx":15AD70
         TabIndex        =   68
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ListBox LstSpellName 
         Height          =   3180
         ItemData        =   "FrmNpcCreator.frx":15AD72
         Left            =   -74760
         List            =   "FrmNpcCreator.frx":15AD74
         TabIndex        =   74
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton OpName3 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   -72960
         TabIndex        =   71
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OpNum3 
         Caption         =   "Numero"
         Height          =   195
         Left            =   -73920
         TabIndex        =   70
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxtSpSearch 
         Height          =   285
         Left            =   -73920
         TabIndex        =   69
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton CmdAddSpell 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   -72000
         TabIndex        =   67
         Top             =   1920
         Width           =   975
      End
      Begin VB.ListBox LstSpellAdd 
         Height          =   3180
         ItemData        =   "FrmNpcCreator.frx":15AD76
         Left            =   -70920
         List            =   "FrmNpcCreator.frx":15AD78
         TabIndex        =   66
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton CmdRemoverSpell 
         Caption         =   "Remover"
         Height          =   255
         Left            =   -72000
         TabIndex        =   65
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton CmdListClearSpell 
         Caption         =   "Limpiar Lista"
         Height          =   255
         Left            =   -72120
         TabIndex        =   64
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TxtCantSpell 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -68880
         TabIndex        =   63
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox LstAdd 
         Height          =   2985
         ItemData        =   "FrmNpcCreator.frx":15AD7A
         Left            =   -72960
         List            =   "FrmNpcCreator.frx":15AD7C
         TabIndex        =   62
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox TxtAdd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74640
         TabIndex        =   61
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar Parametros"
         Height          =   375
         Left            =   -74640
         TabIndex        =   58
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Parametros"
         Height          =   375
         Left            =   -74640
         TabIndex        =   57
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton CmdChange 
         Caption         =   "Cambiar"
         Height          =   315
         Left            =   -74640
         TabIndex        =   56
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox TxtObjCant 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -68880
         TabIndex        =   54
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton CmdListClear 
         Caption         =   "Limpiar Lista"
         Height          =   255
         Left            =   -72120
         TabIndex        =   53
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton CmdRemoverObj 
         Caption         =   "Remover"
         Height          =   255
         Left            =   -72000
         TabIndex        =   52
         Top             =   2280
         Width           =   975
      End
      Begin VB.ListBox LstObjAdd 
         Height          =   3180
         ItemData        =   "FrmNpcCreator.frx":15AD7E
         Left            =   -70920
         List            =   "FrmNpcCreator.frx":15AD80
         TabIndex        =   51
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton CmdAddObj 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   -72000
         TabIndex        =   50
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox TxtCantObj 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72000
         TabIndex        =   48
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox LstObj 
         Height          =   3180
         ItemData        =   "FrmNpcCreator.frx":15AD82
         Left            =   -74760
         List            =   "FrmNpcCreator.frx":15AD84
         TabIndex        =   43
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TxtObjSearch 
         Height          =   285
         Left            =   -73920
         TabIndex        =   45
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox LstObjName 
         Height          =   2595
         ItemData        =   "FrmNpcCreator.frx":15AD86
         Left            =   -74760
         List            =   "FrmNpcCreator.frx":15AD88
         TabIndex        =   44
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton OpNum2 
         Caption         =   "Numero"
         Height          =   195
         Left            =   -73920
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OpName2 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   -72960
         TabIndex        =   41
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox CbHeading 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.Timer THB 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6360
         Top             =   360
      End
      Begin VB.ListBox LstHB 
         Height          =   2790
         ItemData        =   "FrmNpcCreator.frx":15AD8A
         Left            =   2280
         List            =   "FrmNpcCreator.frx":15AD8C
         TabIndex        =   37
         Top             =   1620
         Width           =   2295
      End
      Begin VB.TextBox TxtSearchHB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   34
         Top             =   1020
         Width           =   1455
      End
      Begin VB.PictureBox PbHB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1695
         Index           =   4
         Left            =   4680
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   33
         Top             =   2700
         Width           =   2055
      End
      Begin VB.PictureBox PbHB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1695
         Index           =   3
         Left            =   4680
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   32
         Top             =   1020
         Width           =   2055
      End
      Begin VB.PictureBox PbHB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1695
         Index           =   2
         Left            =   120
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   31
         Top             =   2700
         Width           =   2055
      End
      Begin VB.PictureBox PbHB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1695
         Index           =   1
         Left            =   120
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   30
         Top             =   1020
         Width           =   2055
      End
      Begin VB.TextBox TxtHead 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   28
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtBody 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   25
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   76
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   73
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Hechizos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   72
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label LblAdd 
         Caption         =   "Bloque para ingresar parametros personalizados."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74640
         TabIndex        =   60
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Cambiar Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   59
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Objetos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71160
         TabIndex        =   55
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   49
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   46
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LblList 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   1380
         Width           =   2535
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cabeza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Head"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuerpo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.ComboBox CbTipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   23
      Top             =   1830
      Width           =   1455
   End
   Begin VB.TextBox TxtDefensa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8520
      TabIndex        =   22
      Top             =   1860
      Width           =   1095
   End
   Begin VB.TextBox TxtPAtaque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   21
      Top             =   2550
      Width           =   855
   End
   Begin VB.TextBox TxtPEvacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   20
      Top             =   2550
      Width           =   855
   End
   Begin VB.TextBox TxtHpMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   19
      Top             =   2220
      Width           =   615
   End
   Begin VB.TextBox TxtHitMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   18
      Top             =   2220
      Width           =   615
   End
   Begin VB.TextBox TxtHpMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   17
      Top             =   2220
      Width           =   615
   End
   Begin VB.TextBox TxtHitMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   16
      Top             =   2220
      Width           =   615
   End
   Begin VB.TextBox TxtOro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      TabIndex        =   15
      Top             =   2925
      Width           =   1455
   End
   Begin VB.TextBox TxtExp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   14
      Top             =   2925
      Width           =   1455
   End
   Begin VB.ComboBox CbRespawn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      TabIndex        =   13
      Top             =   1500
      Width           =   1095
   End
   Begin VB.ComboBox CbAlineacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      Top             =   1830
      Width           =   1335
   End
   Begin VB.ComboBox CbHostil 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      TabIndex        =   11
      Top             =   1155
      Width           =   1095
   End
   Begin VB.ComboBox CbComercia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   1500
      Width           =   975
   End
   Begin VB.ComboBox CbAtacable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   9
      Top             =   1155
      Width           =   1215
   End
   Begin VB.ComboBox CbMovimiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   8
      Top             =   1155
      Width           =   1455
   End
   Begin VB.TextBox TxtDesc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox TxtNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   5655
   End
   Begin VB.ListBox LstNpc 
      Height          =   6885
      ItemData        =   "FrmNpcCreator.frx":15AD8E
      Left            =   0
      List            =   "FrmNpcCreator.frx":15AD90
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.ListBox LstNpcName 
      Height          =   6300
      ItemData        =   "FrmNpcCreator.frx":15AD92
      Left            =   0
      List            =   "FrmNpcCreator.frx":15AD94
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox TxtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton OpName 
      BackColor       =   &H00C00000&
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   540
      Width           =   180
   End
   Begin VB.OptionButton OpNum 
      BackColor       =   &H00C00000&
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   540
      Width           =   180
   End
   Begin VB.Label CmdVolver 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   79
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label CmdCrear 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   8640
      TabIndex        =   78
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label CmdModificar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3960
      TabIndex        =   77
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   4440
      TabIndex        =   75
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblCarga 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   750
      Width           =   2775
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu IParametros 
         Caption         =   "Insertar Parametros"
         Begin VB.Menu InsertType 
            Caption         =   "Tipo de NPC"
         End
         Begin VB.Menu InsertMovement 
            Caption         =   "Tipo de Movimiento"
         End
         Begin VB.Menu InserAlin 
            Caption         =   "Alineacion"
         End
      End
   End
   Begin VB.Menu Reloads 
      Caption         =   "Reloads"
      Begin VB.Menu ReloadNpc 
         Caption         =   "Recargar NPC"
      End
      Begin VB.Menu ReloadHeads 
         Caption         =   "Recargar Cabezas"
      End
      Begin VB.Menu ReloadBodys 
         Caption         =   "Recargar Cuerpos"
      End
      Begin VB.Menu ReloadObj 
         Caption         =   "Recargar Objetos"
      End
      Begin VB.Menu ReloadAll 
         Caption         =   "Recargar Todo"
      End
   End
   Begin VB.Menu MRapido 
      Caption         =   "Menu Rapido"
      Begin VB.Menu MInicio 
         Caption         =   "Inicio"
      End
      Begin VB.Menu MIndexacion 
         Caption         =   "Indexacion"
         Begin VB.Menu MIndexar 
            Caption         =   "Indexar"
         End
         Begin VB.Menu MDatosIndex 
            Caption         =   "Ver Datos de Indexacion"
         End
      End
      Begin VB.Menu MDateo 
         Caption         =   "Dateo"
         Begin VB.Menu MObjetos 
            Caption         =   "Objetos"
         End
         Begin VB.Menu MHechizos 
            Caption         =   "Hechizos"
         End
      End
      Begin VB.Menu MConversor 
         Caption         =   "Conversor de Imagenes"
      End
      Begin VB.Menu MRutas 
         Caption         =   "Configurar Rutas"
      End
      Begin VB.Menu MCreditos 
         Caption         =   "Creditos"
      End
      Begin VB.Menu MSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FrmNpcCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LB_FINDSTRING = &H18F
Private Declare Function sendmessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Dim TypeNpc As String
Dim NpcPath As String
Dim NpcHPath As String
Dim ListHB As String
Dim CantFrames As Byte
Dim CantFramesAnim(1 To 4) As Byte
Dim FrameActual As Long
Dim FrameActualAnim(1 To 4) As Long
Dim Index As Integer
Dim ComboTipoValue As Integer
Dim ComboAlineacionValue As Integer
Dim ComboMovimientoValue As Integer
Dim NpcSavePath As String
Dim NumNpc
Dim NpcIndex As Integer
Dim Reload As Boolean
Dim CantNpc As Integer
Dim Mensaje As String

Private Sub CmdAddObj_Click()
Dim NumeroObj As Integer
NumeroObj = Mid(LstObj.Text, 1, InStr(1, LstObj.Text, "-") - 1)
If TxtCantObj.Text = "" Then
    MsgBox "Ingrese una cantidad valida"
    Exit Sub
End If
If LstObj.Text = "" Then
    MsgBox "Seleccione un Objeto antes de agregar"
    Exit Sub
End If
LstObjAdd.AddItem NumeroObj & "-" & TxtCantObj.Text
TxtObjCant.Text = Val(TxtObjCant.Text) + 1
End Sub

Private Sub CmdAddSpell_Click()
Dim NumeroHz As Integer
NumeroHz = Mid(LstSpell.Text, 1, InStr(1, LstSpell.Text, "-") - 1)
If LstSpell.Text = "" Then
    MsgBox "Seleccione un Hechizo antes de agregar"
    Exit Sub
End If
LstSpellAdd.AddItem LstSpell.Text
TxtCantSpell.Text = Val(TxtCantSpell.Text) + 1
End Sub

Private Sub CmdCrear_Click()
NpcIndex = CantNpc + 1
If TypeNpc = "HOSTILES" Then
    NpcSavePath = Config.SaveDatPath & "\NPCs.dat"
    If FileExist(Config.SaveDatPath & "\NPCs.dat", vbNormal) Then
        Call CrearNPC(False)
    Else
        Call FileCopy(Config.DatPath & "\NPCs.dat", Config.SaveDatPath & "\NPCs.dat")
        Call CrearNPC(False)
    End If
Else
    NpcSavePath = Config.SaveDatPath & "\NPCs-HOSTILES.dat"
    If FileExist(Config.SaveDatPath & "\NPCs-HOSTILES.dat", vbNormal) Then
        Call CrearNPC(False)
    Else
        Call FileCopy(Config.DatPath & "\NPCs-HOSTILES.dat", Config.SaveDatPath & "\NPCs-HOSTILES.dat")
        Call CrearNPC(False)
    End If
End If
If Mensaje <> "" Then
    MsgBox "El npc no se ha podido crear ya que se genero un error al momento de guardarlo"
    Mensaje = ""
Else
    MsgBox "El Npc se ha creado con exito."
End If
End Sub

Private Sub CmdListClearSpell_Click()
LstSpellAdd.Clear
TxtCantSpell.Text = 0
End Sub

Private Sub CmdModificar_Click()
If TypeNpc = "HOSTILES" Then
    NpcSavePath = Config.SaveDatPath & "\NPCs.dat"
    If FileExist(Config.SaveDatPath & "\NPCs.dat", vbNormal) Then
        Call CrearNPC(True)
    Else
        Call FileCopy(Config.DatPath & "\NPCs.dat", Config.SaveDatPath & "\NPCs.dat")
        Call CrearNPC(True)
    End If
Else
    NpcSavePath = Config.SaveDatPath & "\NPCs-HOSTILES.dat"
    If FileExist(Config.SaveDatPath & "\NPCs-HOSTILES.dat", vbNormal) Then
        Call CrearNPC(True)
    Else
        Call FileCopy(Config.DatPath & "\NPCs-HOSTILES.dat", Config.SaveDatPath & "\NPCs-HOSTILES.dat")
        Call CrearNPC(True)
    End If
End If

If Mensaje <> "" Then
    MsgBox "El npc no se ha podido modificar ya que se genero un error al momento de guardarlo"
    Mensaje = ""
Else
    MsgBox "El Npc se ha modificado con exito."
End If
End Sub

Private Sub CmdRemoverObj_Click()
If LstObjAdd.ListCount <> 0 Then
    LstObjAdd.RemoveItem (LstObjAdd.ListIndex)
    TxtObjCant.Text = Val(TxtObjCant.Text) - 1
End If
End Sub

Private Sub CmdListClear_Click()
LstObjAdd.Clear
TxtObjCant.Text = 0
End Sub

Private Sub CmdRemoverSpell_Click()
If LstSpellAdd.ListCount <> 0 Then
    LstSpellAdd.RemoveItem (LstSpellAdd.ListIndex)
    TxtCantSpell.Text = Val(TxtCantSpell.Text) - 1
End If
End Sub

Private Sub CmdVolver_Click()
FrmDatMenu.Visible = True
Me.Visible = False
End Sub

Private Sub Form_Activate()
Call Parameters
If IndexMode = "12.1" Then
    Call LoadNpcH
    TypeNpc = "HOSTILES"
    LblCarga.Caption = "Cargar 'NPCs-HOSTILES.dat'"
End If
End Sub

Private Sub Form_Load()
Call Parameters
LblCarga.Caption = "Cargar 'NPCs.dat'"
LblList.Caption = "Cargar Cabezas"
TypeNpc = "NORMALES"
OpNum.value = True
OpName2.value = True
OpName3.value = True
TxtHead.Enabled = False
TxtBody.Enabled = False
TxtObjCant.Enabled = False
TxtCantSpell.Enabled = False

Call LoadDats
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmDatMenu.Visible = True
End Sub

Private Sub InserAlin_Click()
Dim Cant As Integer
Dim Dato As String
Cant = GetVar(IndexDaterIni, "CANTS", "NpcAlineacion")
Dato = InputBox("Ingrese una Alineacion", "Alineacion del NPC")
If Dato <> "" Then
    Call WriteVar(IndexDaterIni, "CANTS", "NpcAlineacion", Cant + 1)
    Call WriteVar(IndexDaterIni, "NPCALINEACION", "NpcAlineacion" & Cant + 1, Dato)
End If

Dim Cont As Integer
CbAlineacion.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "NpcAlineacion"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "NPCALINEACION", "NpcAlineacion" & Cont)
    CbAlineacion.AddItem Cont & "-" & Dato
Next Cont
CbAlineacion.AddItem "(Elegir)"
CbAlineacion.ListIndex = Cant
ComboAlineacionValue = Cant
End Sub

Private Sub InsertMovement_Click()
Dim Cant As Integer
Dim Dato As String
Dim Dato2 As Integer
Cant = GetVar(IndexDaterIni, "CANTS", "NpcMovement")
Dato2 = Val(InputBox("Ingrese el numero de Movimiento", "Movimientos de NPC"))
Dato = InputBox("Ingrese un Nombre para el movimiento", "Movimientos de NPC")
If Dato <> "" Then
    Call WriteVar(IndexDaterIni, "CANTS", "NpcMovement", Cant + 1)
    Call WriteVar(IndexDaterIni, "NPCMOVNUM", "NpcMovNum" & Cant + 1, Dato2)
    Call WriteVar(IndexDaterIni, "NPCMOVEMENT", "NpcMovement" & Cant + 1, Dato)
End If

Dim CbCont As Integer
Dim DatoS As String
Dim Cont As Integer
CbMovimiento.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "NpcMovement"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "NPCMOVEMENT", "NpcMovement" & Cont)
    DatoS = Str(GetVar(IndexDaterIni, "NPCMOVNUM", "NpcMovNum" & Cont))
    If Dato <> "" Then
        CbMovimiento.AddItem DatoS & "-" & Dato
        CbCont = CbCont + 1
    End If
Next Cont
CbMovimiento.AddItem "(Elegir)"
CbMovimiento.ListIndex = CbCont
ComboMovimientoValue = CbCont

End Sub

Private Sub InsertType_Click()
Dim Cant As Integer
Dim Dato As String
Cant = GetVar(IndexDaterIni, "CANTS", "NpcTypes")
Dato = InputBox("Ingrese un Tipo de Npc", "Tipo de NPC")
If Dato <> "" Then
    Call WriteVar(IndexDaterIni, "CANTS", "NpcTypes", Cant + 1)
    Call WriteVar(IndexDaterIni, "NPCTYPES", "NpcType" & Cant + 1, Dato)
End If

Dim Cont As Integer
CbTipo.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "NpcTypes"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "NPCTYPES", "NpcType" & Cont)
    CbTipo.AddItem Cont & "-" & Dato
Next Cont
CbTipo.AddItem "(Elegir)"
CbTipo.ListIndex = Cant
ComboTipoValue = Cant
End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub LblCarga_Click()
LstNpc.Clear
LstNpcName.Clear
If TypeNpc = "HOSTILES" Then
    LblCarga.Caption = "Cargar 'NPCs.dat'"
    TypeNpc = "NORMALES"
    Call LoadNpc
Else
    LblCarga.Caption = "Cargar 'NPCs-HOSTILES.dat'"
    TypeNpc = "HOSTILES"
    Call LoadNpcH
End If
End Sub

Private Sub LoadDats()
ListHB = "BODY"
Dim T As Integer
For T = 1 To BodysCountNew
    If Bodys(T).Body(1) <> 0 Then
        LstHB.AddItem T
    End If
Next T

'Valores "SI" o "NO"
CbAtacable.Clear
CbAtacable.AddItem "No", 0
CbAtacable.AddItem "Si", 1
CbAtacable.AddItem "(Elegir)", 2

CbHostil.Clear
CbHostil.AddItem "No", 0
CbHostil.AddItem "Si", 1
CbHostil.AddItem "(Elegir)", 2

CbComercia.Clear
CbComercia.AddItem "No", 0
CbComercia.AddItem "Si", 1
CbComercia.AddItem "(Elegir)", 2

CbRespawn.Clear
CbRespawn.AddItem "No", 0
CbRespawn.AddItem "Si", 1
CbRespawn.AddItem "(Elegir)", 2

CbHeading.Clear
CbHeading.AddItem "(Elegir)", 0
CbHeading.AddItem "Sur", 1
CbHeading.AddItem "Este", 2
CbHeading.AddItem "Norte", 3
CbHeading.AddItem "Oeste", 4

'Valores Predefinidos
Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

CbTipo.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "NpcTypes"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "NPCTYPES", "NpcType" & Cont)
    CbTipo.AddItem Cont & "-" & Dato
Next Cont
CbTipo.AddItem "(Elegir)"
CbTipo.ListIndex = Cant
ComboTipoValue = Cant

CbAlineacion.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "NpcAlineacion"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "NPCALINEACION", "NpcAlineacion" & Cont)
    CbAlineacion.AddItem Cont & "-" & Dato
Next Cont
CbAlineacion.AddItem "(Elegir)"
CbAlineacion.ListIndex = Cant
ComboAlineacionValue = Cant

Dim CbCont As Integer
Dim Dato2 As String
CbMovimiento.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "NpcMovement"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "NPCMOVEMENT", "NpcMovement" & Cont)
    Dato2 = GetVar(IndexDaterIni, "NPCMOVNUM", "NpcMovNum" & Cont)
    If Dato <> "" Then
        CbMovimiento.AddItem Dato2 & "-" & Dato
        CbCont = CbCont + 1
    End If
Next Cont
CbMovimiento.AddItem "(Elegir)"
CbMovimiento.ListIndex = CbCont
ComboMovimientoValue = CbCont

LstNpc.Clear
LstNpcName.Clear
Cant = Val(GetVar(NpcHPath, "INIT", "NumNPCs"))
CantNpc = Cant
Dim Inicio As Integer
If IndexMode = "12.1" Then
    Inicio = 1
Else
    Inicio = 500
End If
For Cont = Inicio To Cant
    Dato = GetVar(NpcHPath, "NPC" & Cont, "Name")
    If Dato <> "" Then
        LstNpc.AddItem Cont & "-" & Dato
        LstNpcName.AddItem Dato
    End If
Next Cont

Cant = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosNpc"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROSNPC", "Parametro" & Cont)
    LstAdd.AddItem Dato & "=0"
Next Cont

LstSpell.Clear
LstSpellName.Clear
Dim Hechizo As Integer
Dim HechizoName As String
Dim CantHechizos As Integer
CantHechizos = GetVar(Config.DatPath & "\Hechizos.dat", "INIT", "NumeroHechizos")
For Hechizo = 1 To CantHechizos
    HechizoName = GetVar(Config.DatPath & "\Hechizos.dat", "HECHIZO" & Hechizo, "Nombre")
    If HechizoName <> "" Then
        LstSpell.AddItem Hechizo & "-" & HechizoName
        LstSpellName.AddItem HechizoName
    End If
Next Hechizo

Call LoadObj

Call SetComboValue
End Sub

Private Sub LblList_Click()

LstHB.Clear
If ListHB = "HEAD" Then
    LblList.Caption = "Cargar Cabezas"
    ListHB = "BODY"
    Call LoadBodys
Else
    LblList.Caption = "Cargar Cuerpos"
    ListHB = "HEAD"
    Call LoadHeads
End If
End Sub

Private Sub LstHB_Click()
Dim T As Byte
If ListHB = "HEAD" Then
    Dim GraficosPath(1 To 4) As String
    Dim HeadIndex As Integer
    Dim IndexHead As Long
    Dim LongXHead As Integer
    Dim LongYHead As Integer
    Dim PosXHead As Integer
    Dim PosYHead As Integer
    HeadIndex = Val(LstHB.Text)
    
    If HeadIndex = 0 Then Exit Sub
    THB.Enabled = False
    TxtHead.Text = HeadIndex
    
    For T = 1 To 4
        IndexHead = Heads(HeadIndex).Head(T)
        LongXHead = GrhData(IndexHead).pixelWidth
        LongYHead = GrhData(IndexHead).pixelHeight
        PosXHead = GrhData(IndexHead).sX
        PosYHead = GrhData(IndexHead).sY
    
        GraficosPath(T) = Config.BmpPath & "\" & GrhData(IndexHead).FileNum & ".bmp"
        If FileExist(GraficosPath(T), vbNormal) = True Then
            PbHB(T).Visible = True
            PbHB(T).Cls
            PbHB(T).PaintPicture LoadPicture(GraficosPath(T)), 0, 0, LongXHead, LongYHead, PosXHead, PosYHead, LongXHead, LongYHead
        End If
    Next T
Else
    Dim BodyIndex As Integer
    TxtBody.Text = LstHB.Text
    If LstHB.Text <> "" Then
        BodyIndex = Val(LstHB.Text)
        For T = 1 To 4
            CantFramesAnim(T) = GrhData(Bodys(BodyIndex).Body(T)).NumFrames
            FrameActualAnim(T) = 0
            PbHB(T).Cls
        Next T
        
        THB.Enabled = True
        Index = BodyIndex
    End If
End If
End Sub

Private Sub LstNpc_DblClick()
On Error GoTo ErrHandler
Dim NpcRuta As String
Call ClearData

NpcIndex = Mid(LstNpc.Text, 1, InStr(1, LstNpc.Text, "-") - 1)

If TypeNpc = "HOSTILES" Then
    NpcRuta = NpcPath
Else
    NpcRuta = NpcHPath
End If

TxtNombre.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "Name")
TxtDesc.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "Desc")
TxtInfo.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "Info")
CbAtacable.ListIndex = Val(GetVar(NpcRuta, "NPC" & NpcIndex, "Attackable"))
CbHostil.ListIndex = Val(GetVar(NpcRuta, "NPC" & NpcIndex, "Hostile"))
CbComercia.ListIndex = Val(GetVar(NpcRuta, "NPC" & NpcIndex, "Comercia"))
TxtDomable.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "Domable")
CbRespawn.ListIndex = Val(GetVar(NpcRuta, "NPC" & NpcIndex, "ReSpawn"))
CbAlineacion.ListIndex = Val(GetVar(NpcRuta, "NPC" & NpcIndex, "Alineacion")) - 1
CbTipo.ListIndex = Val(GetVar(NpcRuta, "NPC" & NpcIndex, "NpcType"))
CbHeading.ListIndex = Val(GetVar(NpcRuta, "NPC" & NpcIndex, "Heading"))
TxtDefensa.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "DEF")
TxtHpMin.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "MinHP")
TxtHpMax.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "MaxHP")
TxtHitMin.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "MinHIT")
TxtHitMax.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "MaxHIT")
TxtPAtaque.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "PoderAtaque")
TxtPEvacion.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "PoderEvasion")
TxtExp.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "GiveEXP")
TxtOro.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "GiveGLD")
TxtHead.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "Head")
TxtBody.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "Body")
TxtObjCant.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "NROITEMS")

Dim Combo As Integer
For Combo = 0 To CbMovimiento.ListCount - 1
    If Combo <> (CbMovimiento.ListCount - 1) Then
        If Mid(CbMovimiento.List(Combo), 1, InStr(1, CbMovimiento.List(Combo), "-") - 1) = GetVar(NpcRuta, "NPC" & NpcIndex, "Movement") Then
            CbMovimiento.ListIndex = Combo
        End If
    End If
Next Combo

If Val(TxtObjCant.Text) <> 0 Then
    Dim T As Integer
    Dim Dato As String
    For T = 1 To Val(TxtObjCant.Text)
        Dato = GetVar(NpcRuta, "NPC" & NpcIndex, "Obj" & T)
        LstObjAdd.AddItem Dato
    Next T
End If

Dim CantParm As Integer
CantParm = LstAdd.ListCount
LstAdd.Clear
For T = 1 To CantParm
    Dim ValorParm As String
    Dim Parametro As String
    Parametro = GetVar(IndexDaterIni, "PARAMETROSNPC", "Parametro" & T)
    ValorParm = GetVar(NpcRuta, "NPC" & NpcIndex, Parametro)
    If ValorParm = "" Then ValorParm = 0
    LstAdd.AddItem Parametro & "=" & ValorParm
Next T

TxtCantSpell.Text = GetVar(NpcRuta, "NPC" & NpcIndex, "LanzaSpells")
LstSpellAdd.Clear
For T = 1 To Val(TxtCantSpell.Text)
    Dim NombreH As String
    Dim NpcSpell As String
    NpcSpell = GetVar(NpcRuta, "NPC" & NpcIndex, "SP" & T)
    NombreH = GetVar(Config.DatPath & "\Hechizos.dat", "HECHIZO" & NpcSpell, "Nombre")
    LstSpellAdd.AddItem NpcSpell & "-" & NombreH
Next T
Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub MConversor_Click()
FrmConversor.Visible = True
End Sub

Private Sub MRutas_Click()
frmConfig.Visible = True
End Sub

Private Sub OpName_Click()
If OpName.value = True Then OpNum.value = False
End Sub

Private Sub OpName2_Click()
OpNum2.value = False
End Sub

Private Sub OpName3_Click()
OpNum3.value = False
End Sub

Private Sub OpNum_Click()
If OpNum.value = True Then OpName.value = False
End Sub

Private Sub OpNum2_Click()
OpName2.value = False
End Sub

Private Sub OpNum3_Click()
OpName3.value = False
End Sub

Private Sub ReloadAll_Click()
Call LoadDats
End Sub

Private Sub ReloadBodys_Click()
If ListHB = "BODY" Then
    Call LoadBodys
End If
End Sub

Private Sub ReloadHeads_Click()
If ListHB = "HEAD" Then
    Call LoadHeads
End If
End Sub

Private Sub ReloadNpc_Click()
If TypeNpc = "HOSTILES" Then
    Call LoadNpcH
Else
    Call LoadNpc
End If
End Sub

Private Sub ReloadObj_Click()
Call LoadObj
End Sub

Private Sub THB_Timer()
On Error GoTo ErrorAnim
Dim AnimacionPosX(1 To 4) As Integer
Dim AnimacionPosY(1 To 4) As Integer
Dim AnimacionLongX(1 To 4) As Integer
Dim AnimacionLongY(1 To 4) As Integer
Dim GraficoPath(1 To 4) As String
Dim GrhIndexAnim(1 To 4) As Long
Dim Anim(1 To 4) As Integer
Dim T As Byte

For T = 1 To 4
    FrameActualAnim(T) = FrameActualAnim(T) + 1
    
    Anim(T) = Bodys(Index).Body(T)
    
    GrhIndexAnim(T) = GrhData(Anim(T)).Frames(FrameActualAnim(T))
    
    If GrhIndexAnim(T) = 0 Then 'Por si hay error y no existe.
        THB.Enabled = False
        Exit Sub
    End If
    
    GraficoPath(T) = Config.BmpPath & "\" & GrhData(GrhIndexAnim(T)).FileNum & ".bmp" 'Busco la Imagen
    
    If Not FileExist(GraficoPath(T), vbNormal) = True Then 'Me fijo si Existe
        THB.Enabled = False
        Exit Sub
    End If
    
    AnimacionPosX(T) = GrhData(GrhIndexAnim(T)).sX 'Coordenada X
    AnimacionPosY(T) = GrhData(GrhIndexAnim(T)).sY 'Coordenada Y
    AnimacionLongX(T) = GrhData(GrhIndexAnim(T)).pixelWidth 'Longitud sobre X
    AnimacionLongY(T) = GrhData(GrhIndexAnim(T)).pixelHeight 'Longitud sobre Y
    
    PbHB(T).PaintPicture LoadPicture(GraficoPath(T)), 0, 0, AnimacionLongX(T), AnimacionLongY(T), AnimacionPosX(T), AnimacionPosY(T), AnimacionLongX(T), AnimacionLongY(T)
    
    If FrameActualAnim(T) = CantFramesAnim(T) Then
        FrameActualAnim(T) = 0
    End If
Next T

Exit Sub

ErrorAnim:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    THB.Enabled = False
    Exit Sub
'For T = 1 To 4
'    FrameActualAnim(T) = 0
'Next T
End Sub

Private Sub TxtObjSearch_Change()
If OpNum2.value = True Then
    LstObj.ListIndex = sendmessage(LstObj.hWnd, LB_FINDSTRING, -1, ByVal TxtObjSearch.Text)
Else
    LstObjName.ListIndex = sendmessage(LstObjName.hWnd, LB_FINDSTRING, -1, ByVal TxtObjSearch.Text)
    LstObj.ListIndex = LstObjName.ListIndex
End If
End Sub

Private Sub TxtSearch_Change()
If OpNum.value = True Then
        LstNpc.ListIndex = sendmessage(LstNpc.hWnd, LB_FINDSTRING, -1, ByVal TxtSearch.Text)
    Else
        LstNpcName.ListIndex = sendmessage(LstNpcName.hWnd, LB_FINDSTRING, -1, ByVal TxtSearch.Text)
        LstNpc.ListIndex = LstNpcName.ListIndex
    End If
End Sub

Private Sub TxtSearchHB_Change()
LstHB.ListIndex = sendmessage(LstHB.hWnd, LB_FINDSTRING, -1, ByVal TxtSearchHB.Text)
End Sub

Private Sub LoadHeads()
Dim T As Integer
For T = 1 To HeadsCountNew
    If Heads(T).Head(1) <> 0 Then
        LstHB.AddItem T
    End If
Next T
End Sub

Private Sub LoadBodys()
Dim T As Integer
For T = 1 To BodysCountNew
    If Bodys(T).Body(1) <> 0 Then
        LstHB.AddItem T
    End If
Next T
End Sub

Private Sub LoadNpc()
Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

LstNpc.Clear
LstNpcName.Clear
Cant = GetVar(NpcHPath, "INIT", "NumNPCs")
NumNpc = Cant
For Cont = 500 To Cant
    Dato = GetVar(NpcHPath, "NPC" & Cont, "Name")
    If Dato <> "" Then
        LstNpc.AddItem Cont & "-" & Dato
        LstNpcName.AddItem Dato
    End If
Next Cont
End Sub

Private Sub LoadNpcH()
Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

LstNpc.Clear
LstNpcName.Clear
Cant = GetVar(NpcPath, "INIT", "NumNPCs")
NumNpc = Cant
For Cont = 1 To Cant
    Dato = GetVar(NpcPath, "NPC" & Cont, "Name")
    If Dato <> "" Then
        LstNpc.AddItem Cont & "-" & Dato
        LstNpcName.AddItem Dato
    End If
Next Cont
End Sub

Private Sub SetComboValue()
CbAtacable.ListIndex = 2
CbHostil.ListIndex = 2
CbRespawn.ListIndex = 2
CbComercia.ListIndex = 2
CbHeading.ListIndex = 1
End Sub

Private Sub ClearData()
Call SetComboValue
TxtNombre.Text = ""
TxtInfo.Text = ""
TxtDesc.Text = ""
TxtDomable.Text = ""
TxtHpMin.Text = ""
TxtHpMax.Text = ""
TxtHitMin.Text = ""
TxtHitMax.Text = ""
TxtPAtaque.Text = ""
TxtPEvacion.Text = ""
TxtExp.Text = ""
TxtOro.Text = ""
TxtHead.Text = ""
TxtBody.Text = ""
CbTipo.ListIndex = ComboTipoValue
CbMovimiento.ListIndex = ComboMovimientoValue
CbAlineacion.ListIndex = ComboAlineacionValue
LstObjAdd.Clear
End Sub

Private Sub LoadObj()
LstObj.Clear
LstObjName.Clear
Dim CantItems As Integer
Dim Obj As Integer
Dim ObjName As String
CantItems = GetVar(Config.DatPath & "\Obj.dat", "INIT", "NumOBJs")
For Obj = 1 To CantItems
    ObjName = GetVar(Config.DatPath & "\Obj.dat", "OBJ" & Obj, "Name")
    If ObjName <> "" Then
        LstObj.AddItem Obj & "-" & ObjName
        LstObjName.AddItem ObjName
    End If
Next Obj
End Sub

Private Sub CrearNPC(Modifica As Boolean)
If NpcError = False Then
    MsgBox Mensaje
    Exit Sub
End If
If Modifica = False Then
    Call WriteVar(NpcSavePath, "INIT", "NumNPCs", NpcIndex)
End If
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Name", TxtNombre.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Info", TxtInfo.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Desc", TxtDesc.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Head", TxtHead.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Body", TxtBody.Text)
If CbHeading.ListIndex <> 0 Then
    Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Heading", CbHeading.ListIndex)
End If
If CbMovimiento.Text <> "(Elegir)" Then
    Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Movement", Mid(CbMovimiento.Text, 1, InStr(1, CbMovimiento.Text, "-") - 1))
Else
    MsgBox "Ingrese un tipo de movimiento"
End If
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Comercia", CbComercia.ListIndex)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Hostile", CbHostil.ListIndex)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Attackable", CbAtacable.ListIndex)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "GiveEXP", TxtExp.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "GiveGLD", TxtOro.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "MinHP", TxtHpMin.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "MaxHP", TxtHpMax.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "MinHIT", TxtHitMin.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "MaxHit", TxtHitMax.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "DEF", TxtDefensa.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "PoderAtaque", TxtPAtaque.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "PoderEvacion", TxtPEvacion.Text)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Alineacion", CbAlineacion.ListIndex)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Domable", TxtDomable.Text)
If TypeNpc = "NORMALES" Then
    Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "BackUp", 1)
End If
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "ReSpawn", CbRespawn.ListIndex)
Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "NROITEMS", TxtObjCant.Text)
Dim T As Integer
For T = 0 To LstObjAdd.ListCount
    Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "Obj" & T + 1, LstObjAdd.List(T))
Next T

Dim Parms As Integer
For Parms = 0 To LstAdd.ListCount - 1
    Dim NombreParm As String
    Dim ValorParm As String
    NombreParm = Left(LstAdd.List(Parms), InStr(1, LstAdd.List(Parms), "=") - 1)
    ValorParm = Mid(LstAdd.List(Parms), InStr(1, LstAdd.List(Parms), "=") + 1, Trim(Len(LstAdd.List(Parms))))
    Call WriteVar(NpcSavePath, "NPC" & NpcIndex, NombreParm, ValorParm)
Next Parms

Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "LanzaSpells", TxtCantSpell.Text)

For T = 0 To LstSpellAdd.ListCount - 1
    Call WriteVar(NpcSavePath, "NPC" & NpcIndex, "SP" & T + 1, Mid(LstSpellAdd.List(T), 1, InStr(1, LstSpellAdd.List(T), "-") - 1))
Next T

End Sub

Private Sub MCreditos_Click()
FrmCreditos.Visible = True
Me.Visible = False
End Sub

Private Sub MDatosIndex_Click()
FrmAnimaciones.Visible = True
Me.Visible = False
End Sub

Private Sub MHechizos_Click()
FrmHechizosCreator.Visible = True
Me.Visible = False
End Sub

Private Sub MIndexar_Click()
FrmIndex.Visible = True
Me.Visible = False
End Sub

Private Sub MInicio_Click()
FrmInicio.Visible = True
Me.Visible = False
End Sub

Private Sub MObjetos_Click()
FrmObjSelector.Visible = True
Me.Visible = False
End Sub

Private Sub MSalir_Click()
End
End Sub

Private Sub CmdAgregar_Click()
Dim NewParm As String
Dim CantParm As Integer
Dim TotalParm As Integer
NewParm = InputBox("Agregue el nuevo parametro, SIN valor inicia y SIN el '='.", "Agregar Parametros")

If NewParm = "" Then Exit Sub

CantParm = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosNpc"))
TotalParm = CantParm + 1
Call WriteVar(IndexDaterIni, "CANTS", "ParametrosNpc", TotalParm)
Call WriteVar(IndexDaterIni, "PARAMETROSNPC", "Parametro" & TotalParm, NewParm)

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
LstAdd.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosNpc"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROSNPC", "Parametro" & Cont)
    LstAdd.AddItem Dato & "=0"
Next Cont

MsgBox "Parametro Ingresado"
End Sub

Private Sub CmdCargar_Click()
Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
LstAdd.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosNpc"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROSNPC", "Parametro" & Cont)
    LstAdd.AddItem Dato & "=0"
Next Cont
End Sub

Private Sub CmdChange_Click()
If LstAdd.Text = "" Then Exit Sub

Dim Cont As Integer
For Cont = 0 To LstAdd.ListCount - 1
    If LstAdd.Selected(Cont) = True Then
        LstAdd.List(Cont) = Left(LstAdd.List(Cont), InStr(1, LstAdd.List(Cont), "=") - 1) & "=" & TxtAdd.Text
    End If
Next Cont
End Sub

Private Sub TxtSpSearch_Change()
If OpNum3.value = True Then
    LstSpell.ListIndex = sendmessage(LstSpell.hWnd, LB_FINDSTRING, -1, ByVal TxtSpSearch.Text)
Else
    LstSpellName.ListIndex = sendmessage(LstSpellName.hWnd, LB_FINDSTRING, -1, ByVal TxtSpSearch.Text)
    LstSpell.ListIndex = LstSpellName.ListIndex
End If
End Sub

Private Sub Parameters()
If IndexMode = "12.1" Then
    NpcPath = Config.DatPath & "\NPCs.dat"
    NpcHPath = Config.DatPath & "\NPCs.dat"
    LblCarga.Visible = False
Else
    NpcPath = Config.DatPath & "\NPCs.dat"
    NpcHPath = Config.DatPath & "\NPCs-HOSTILES.dat"
    LblCarga.Visible = True
End If
End Sub

Private Function NpcError() As Boolean
NpcError = True
If CbHeading.Text = "(Elegir)" Then
    Mensaje = "Eliga una direccion para el NPC, 'Heading'"
    NpcError = False
End If
If CbMovimiento.Text = "(Elegir)" Then
    Mensaje = "Ingrese un tipo de movimiento"
    NpcError = False
End If
If CbComercia.Text = "(Elegir)" Then
    Mensaje = "Eliga si el Npc comercia o no"
    NpcError = False
End If
If CbHostil.Text = "(Elegir)" Then
    Mensaje = "Eliga si el Npc es Hostil o no"
    NpcError = False
End If
If CbAtacable.Text = "(Elegir)" Then
    Mensaje = "Eliga si el Npc es atacable o no"
    NpcError = False
End If
If CbRespawn.Text = "(Elegir)" Then
    Mensaje = "Eliga si el Npc va a respawnear o no"
    NpcError = False
End If
If CbTipo.Text = "(Elegir)" Then
    Mensaje = "Eliga un tipo de NPC"
    NpcError = False
End If

End Function
