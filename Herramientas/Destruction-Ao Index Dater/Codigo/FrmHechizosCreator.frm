VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmHechizosCreator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creacion de Hechizos"
   ClientHeight    =   7335
   ClientLeft      =   105
   ClientTop       =   705
   ClientWidth     =   9450
   Icon            =   "FrmHechizosCreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmHechizosCreator.frx":08CA
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   630
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox LstHechizos 
      Height          =   6495
      ItemData        =   "FrmHechizosCreator.frx":12D6A6
      Left            =   0
      List            =   "FrmHechizosCreator.frx":12D6A8
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.ListBox LstHName 
      Height          =   6105
      ItemData        =   "FrmHechizosCreator.frx":12D6AA
      Left            =   0
      List            =   "FrmHechizosCreator.frx":12D6AC
      TabIndex        =   123
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.OptionButton OpNum 
      BackColor       =   &H00C00000&
      Height          =   195
      Left            =   645
      TabIndex        =   122
      Top             =   480
      Width           =   180
   End
   Begin VB.OptionButton OpName 
      BackColor       =   &H00C00000&
      Height          =   195
      Left            =   1680
      TabIndex        =   121
      Top             =   480
      Width           =   180
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4575
      Left            =   2880
      TabIndex        =   8
      Top             =   2280
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Generales"
      TabPicture(0)   =   "FrmHechizosCreator.frx":12D6AE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblTipo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblTarget"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblSkills"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblMana"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblStamina"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblNStaff"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblSAffected"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LblResis"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CbTipo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CbTarget"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtSkills"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtMana"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtStamina"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtNStaff"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CbResis"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CbSAffected"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Daños"
      TabPicture(1)   =   "FrmHechizosCreator.frx":12D6CA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblSMana"
      Tab(1).Control(1)=   "LblSVida"
      Tab(1).Control(2)=   "LblSStamina"
      Tab(1).Control(3)=   "LblSFuerza"
      Tab(1).Control(4)=   "LblSAgilidad"
      Tab(1).Control(5)=   "LblSCarisma"
      Tab(1).Control(6)=   "Label26"
      Tab(1).Control(7)=   "Label27"
      Tab(1).Control(8)=   "Label28"
      Tab(1).Control(9)=   "Label29"
      Tab(1).Control(10)=   "Label30"
      Tab(1).Control(11)=   "Label31"
      Tab(1).Control(12)=   "Label32"
      Tab(1).Control(13)=   "Label33"
      Tab(1).Control(14)=   "Label34"
      Tab(1).Control(15)=   "Label35"
      Tab(1).Control(16)=   "Label36"
      Tab(1).Control(17)=   "Label37"
      Tab(1).Control(18)=   "LblSHambre"
      Tab(1).Control(19)=   "LblSSed"
      Tab(1).Control(20)=   "Label40"
      Tab(1).Control(21)=   "Label41"
      Tab(1).Control(22)=   "Label42"
      Tab(1).Control(23)=   "Label43"
      Tab(1).Control(24)=   "CbSMana"
      Tab(1).Control(25)=   "CbSVida"
      Tab(1).Control(26)=   "CbSStamina"
      Tab(1).Control(27)=   "CbSFuerza"
      Tab(1).Control(28)=   "CbSAgilidad"
      Tab(1).Control(29)=   "CbSCarisma"
      Tab(1).Control(30)=   "TxtMpMin"
      Tab(1).Control(31)=   "TxtMpMax"
      Tab(1).Control(32)=   "TxtHpMin"
      Tab(1).Control(33)=   "TxtHpMax"
      Tab(1).Control(34)=   "TxtStMin"
      Tab(1).Control(35)=   "TxtStMax"
      Tab(1).Control(36)=   "TxtFzMin"
      Tab(1).Control(37)=   "TxtFzMax"
      Tab(1).Control(38)=   "TxtAgMin"
      Tab(1).Control(39)=   "TxtAgMax"
      Tab(1).Control(40)=   "TxtCaMin"
      Tab(1).Control(41)=   "TxtCaMax"
      Tab(1).Control(42)=   "CbSHambre"
      Tab(1).Control(43)=   "CbSSed"
      Tab(1).Control(44)=   "TxtHaMin"
      Tab(1).Control(45)=   "TxtHaMax"
      Tab(1).Control(46)=   "TxtSdMin"
      Tab(1).Control(47)=   "TxtSdMax"
      Tab(1).Control(48)=   "ChMP"
      Tab(1).Control(49)=   "ChHP"
      Tab(1).Control(50)=   "ChSP"
      Tab(1).Control(51)=   "ChStr"
      Tab(1).Control(52)=   "ChAgi"
      Tab(1).Control(53)=   "ChCa"
      Tab(1).Control(54)=   "ChHam"
      Tab(1).Control(55)=   "ChSed"
      Tab(1).ControlCount=   56
      TabCaption(2)   =   "Opciones"
      TabPicture(2)   =   "FrmHechizosCreator.frx":12D6E6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "CbMime"
      Tab(2).Control(3)=   "CbEstu"
      Tab(2).Control(4)=   "CbCeg"
      Tab(2).Control(5)=   "CbRev"
      Tab(2).Control(6)=   "CbCurEnv"
      Tab(2).Control(7)=   "CbEnv"
      Tab(2).Control(8)=   "Frame1"
      Tab(2).Control(9)=   "CbInmo"
      Tab(2).Control(10)=   "CbPara"
      Tab(2).Control(11)=   "CbInvi"
      Tab(2).Control(12)=   "Label55"
      Tab(2).Control(13)=   "Label54"
      Tab(2).Control(14)=   "Label53"
      Tab(2).Control(15)=   "Label52"
      Tab(2).Control(16)=   "Label51"
      Tab(2).Control(17)=   "Label50"
      Tab(2).Control(18)=   "Label46"
      Tab(2).Control(19)=   "Label45"
      Tab(2).Control(20)=   "Label44"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "Wav"
      TabPicture(3)   =   "FrmHechizosCreator.frx":12D702
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "WavList"
      Tab(3).Control(1)=   "TxtWav"
      Tab(3).Control(2)=   "LblWav"
      Tab(3).Control(3)=   "ImgPlay"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Fx"
      TabPicture(4)   =   "FrmHechizosCreator.frx":12D71E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "PbHFx"
      Tab(4).Control(1)=   "TFx"
      Tab(4).Control(2)=   "TxtSearchFx"
      Tab(4).Control(3)=   "LstHFx"
      Tab(4).Control(4)=   "TxtLoops"
      Tab(4).Control(5)=   "TxtFx"
      Tab(4).Control(6)=   "Label13"
      Tab(4).Control(7)=   "LblLoops"
      Tab(4).Control(8)=   "LblFx"
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "Agregados"
      TabPicture(5)   =   "FrmHechizosCreator.frx":12D73A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "CmdChange"
      Tab(5).Control(1)=   "CmdAgregar"
      Tab(5).Control(2)=   "CmdCargar"
      Tab(5).Control(3)=   "TxtAdd"
      Tab(5).Control(4)=   "LstAdd"
      Tab(5).Control(5)=   "Label2"
      Tab(5).Control(6)=   "LblAdd"
      Tab(5).ControlCount=   7
      Begin VB.CheckBox ChSed 
         Height          =   195
         Left            =   -70920
         TabIndex        =   141
         Top             =   3170
         Width           =   195
      End
      Begin VB.CheckBox ChHam 
         Height          =   195
         Left            =   -73080
         TabIndex        =   140
         Top             =   3170
         Width           =   195
      End
      Begin VB.CheckBox ChCa 
         Height          =   195
         Left            =   -68760
         TabIndex        =   139
         Top             =   1850
         Width           =   195
      End
      Begin VB.CheckBox ChAgi 
         Height          =   195
         Left            =   -70920
         TabIndex        =   138
         Top             =   1850
         Width           =   195
      End
      Begin VB.CheckBox ChStr 
         Height          =   195
         Left            =   -73200
         TabIndex        =   137
         Top             =   1850
         Width           =   195
      End
      Begin VB.CheckBox ChSP 
         Height          =   195
         Left            =   -68880
         TabIndex        =   136
         Top             =   530
         Width           =   195
      End
      Begin VB.CheckBox ChHP 
         Height          =   195
         Left            =   -71280
         TabIndex        =   135
         Top             =   530
         Width           =   195
      End
      Begin VB.CheckBox ChMP 
         Height          =   195
         Left            =   -73320
         TabIndex        =   134
         Top             =   530
         Width           =   195
      End
      Begin VB.CommandButton CmdChange 
         Caption         =   "Cambiar"
         Height          =   315
         Left            =   -74880
         TabIndex        =   128
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Parametros"
         Height          =   375
         Left            =   -74880
         TabIndex        =   127
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar Parametros"
         Height          =   375
         Left            =   -74880
         TabIndex        =   126
         Top             =   2160
         Width           =   1575
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
         Left            =   -74880
         TabIndex        =   125
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ListBox LstAdd 
         Height          =   2985
         ItemData        =   "FrmHechizosCreator.frx":12D756
         Left            =   -73200
         List            =   "FrmHechizosCreator.frx":12D758
         TabIndex        =   124
         Top             =   600
         Width           =   4575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Materializacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -71640
         TabIndex        =   116
         Top             =   1800
         Width           =   2415
         Begin VB.TextBox TxtNumObj 
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
            Left            =   1320
            TabIndex        =   119
            Top             =   840
            Width           =   735
         End
         Begin VB.ComboBox CbMate 
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
            Left            =   1320
            TabIndex        =   118
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Num Obj"
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
            TabIndex        =   120
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Materializa"
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
            TabIndex        =   117
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Invocacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74160
         TabIndex        =   109
         Top             =   1800
         Width           =   2295
         Begin VB.TextBox TxtCant 
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
            TabIndex        =   114
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtNumNpc 
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
            TabIndex        =   112
            Top             =   840
            Width           =   735
         End
         Begin VB.ComboBox CbInvo 
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
            Left            =   960
            TabIndex        =   111
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label58 
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
            Left            =   120
            TabIndex        =   115
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "Num Npc"
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
            TabIndex        =   113
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoca"
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
            TabIndex        =   110
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.ComboBox CbMime 
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
         Left            =   -69360
         TabIndex        =   108
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox CbEstu 
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
         Left            =   -71640
         TabIndex        =   106
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox CbCeg 
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
         Left            =   -73920
         TabIndex        =   104
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox CbRev 
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
         Left            =   -69360
         TabIndex        =   102
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox CbCurEnv 
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
         Left            =   -71280
         TabIndex        =   100
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox CbEnv 
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
         Left            =   -73680
         TabIndex        =   98
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Remover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         TabIndex        =   90
         Top             =   3600
         Width           =   6255
         Begin VB.ComboBox CbRemoEstu 
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
            Left            =   5400
            TabIndex        =   96
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox CbRemoPara 
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
            Left            =   3360
            TabIndex        =   94
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox CbRemoInvi 
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
            Left            =   1440
            TabIndex        =   92
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Estupidez"
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
            Left            =   4320
            TabIndex        =   95
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "Paralisis"
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
            Left            =   2400
            TabIndex        =   93
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Invisibilidad"
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
            TabIndex        =   91
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.ComboBox CbInmo 
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
         Left            =   -69360
         TabIndex        =   89
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox CbPara 
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
         Left            =   -71520
         TabIndex        =   87
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox CbInvi 
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
         Left            =   -73560
         TabIndex        =   85
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox CbSAffected 
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
         Left            =   5160
         TabIndex        =   83
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TxtSdMax 
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
         Left            =   -71880
         TabIndex        =   81
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TxtSdMin 
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
         Left            =   -71880
         TabIndex        =   79
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox TxtHaMax 
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
         Left            =   -74040
         TabIndex        =   77
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TxtHaMin 
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
         Left            =   -74040
         TabIndex        =   75
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox CbSSed 
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
         Left            =   -71880
         TabIndex        =   73
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox CbSHambre 
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
         Left            =   -74040
         TabIndex        =   71
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox TxtCaMax 
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
         Left            =   -69840
         TabIndex        =   69
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TxtCaMin 
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
         Left            =   -69840
         TabIndex        =   67
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TxtAgMax 
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
         Left            =   -71880
         TabIndex        =   65
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TxtAgMin 
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
         Left            =   -71880
         TabIndex        =   63
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TxtFzMax 
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
         Left            =   -74160
         TabIndex        =   61
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TxtFzMin 
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
         Left            =   -74160
         TabIndex        =   59
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TxtStMax 
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
         Left            =   -69840
         TabIndex        =   57
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtStMin 
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
         Left            =   -69840
         TabIndex        =   55
         Top             =   840
         Width           =   1215
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
         Left            =   -72240
         TabIndex        =   53
         Top             =   1200
         Width           =   1215
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
         Left            =   -72240
         TabIndex        =   51
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtMpMax 
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
         Left            =   -74280
         TabIndex        =   49
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtMpMin 
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
         Left            =   -74280
         TabIndex        =   47
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox CbSCarisma 
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
         Left            =   -69720
         TabIndex        =   45
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox CbSAgilidad 
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
         Left            =   -71880
         TabIndex        =   43
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox CbSFuerza 
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
         Left            =   -74160
         TabIndex        =   41
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox CbSStamina 
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
         Left            =   -69840
         TabIndex        =   39
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox CbSVida 
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
         Left            =   -72240
         TabIndex        =   37
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox CbSMana 
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
         Left            =   -74280
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox CbResis 
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
         Left            =   5640
         TabIndex        =   34
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox TxtNStaff 
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
         Left            =   1320
         TabIndex        =   30
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TxtStamina 
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
         Left            =   3240
         TabIndex        =   28
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtMana 
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
         Left            =   720
         TabIndex        =   26
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtSkills 
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
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox CbTarget 
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
         Left            =   3240
         TabIndex        =   22
         Top             =   480
         Width           =   1575
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
         Left            =   600
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.PictureBox PbHFx 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   3615
         Left            =   -74880
         ScaleHeight     =   237
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   317
         TabIndex        =   19
         Top             =   840
         Width           =   4815
      End
      Begin VB.Timer TFx 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   -71040
         Top             =   360
      End
      Begin VB.TextBox TxtSearchFx 
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
         Left            =   -69960
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.ListBox LstHFx 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         ItemData        =   "FrmHechizosCreator.frx":12D75A
         Left            =   -69960
         List            =   "FrmHechizosCreator.frx":12D75C
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TxtLoops 
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
         Left            =   -71760
         TabIndex        =   14
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxtFx 
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
         Left            =   -74520
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.FileListBox WavList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3600
         Left            =   -71760
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox TxtWav 
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
         Left            =   -72600
         TabIndex        =   9
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Left            =   -74880
         TabIndex        =   130
         Top             =   2640
         Width           =   1575
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
         Left            =   -74880
         TabIndex        =   129
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Mimetiza"
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
         Left            =   -70320
         TabIndex        =   107
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Estupidez"
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
         Left            =   -72720
         TabIndex        =   105
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Ceguera"
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
         Left            =   -74880
         TabIndex        =   103
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Revivir"
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
         Left            =   -70200
         TabIndex        =   101
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Cura Veneno"
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
         Left            =   -72720
         TabIndex        =   99
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Envenenar"
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
         Left            =   -74880
         TabIndex        =   97
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Inmoviliza"
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
         Left            =   -70440
         TabIndex        =   88
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Paraliza"
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
         Left            =   -72480
         TabIndex        =   86
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Invisibilidad"
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
         Left            =   -74880
         TabIndex        =   84
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   -72480
         TabIndex        =   82
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   -72480
         TabIndex        =   80
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         TabIndex        =   78
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         TabIndex        =   76
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label LblSSed 
         BackStyle       =   0  'Transparent
         Caption         =   "Sed"
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
         Left            =   -72480
         TabIndex        =   74
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label LblSHambre 
         BackStyle       =   0  'Transparent
         Caption         =   "Hambre"
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
         Left            =   -74880
         TabIndex        =   72
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   -70440
         TabIndex        =   70
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   -70440
         TabIndex        =   68
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   -72480
         TabIndex        =   66
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   -72480
         TabIndex        =   64
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         TabIndex        =   62
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         TabIndex        =   60
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   -70440
         TabIndex        =   58
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   -70440
         TabIndex        =   56
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   -72840
         TabIndex        =   54
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   -72840
         TabIndex        =   52
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   -74880
         TabIndex        =   50
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   -74880
         TabIndex        =   48
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LblSCarisma 
         BackStyle       =   0  'Transparent
         Caption         =   "Carisma"
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
         Left            =   -70560
         TabIndex        =   46
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label LblSAgilidad 
         BackStyle       =   0  'Transparent
         Caption         =   "Agilidad"
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
         Left            =   -72840
         TabIndex        =   44
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label LblSFuerza 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuerza"
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
         Left            =   -74880
         TabIndex        =   42
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label LblSStamina 
         BackStyle       =   0  'Transparent
         Caption         =   "Stamina"
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
         Left            =   -70680
         TabIndex        =   40
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblSVida 
         BackStyle       =   0  'Transparent
         Caption         =   "Vida"
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
         Left            =   -72840
         TabIndex        =   38
         Top             =   480
         Width           =   495
      End
      Begin VB.Label LblSMana 
         BackStyle       =   0  'Transparent
         Caption         =   "Mana"
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
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblResis 
         BackStyle       =   0  'Transparent
         Caption         =   "Resis"
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
         Left            =   4920
         TabIndex        =   33
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblSAffected 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Affected"
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
         Left            =   3600
         TabIndex        =   32
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label LblNStaff 
         BackStyle       =   0  'Transparent
         Caption         =   "Need Staff"
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
         TabIndex        =   31
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label LblStamina 
         BackStyle       =   0  'Transparent
         Caption         =   "Stamina"
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
         Left            =   2400
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LblMana 
         BackStyle       =   0  'Transparent
         Caption         =   "Mana"
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
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblSkills 
         BackStyle       =   0  'Transparent
         Caption         =   "Skills"
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
         Left            =   5040
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblTarget 
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
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
         Left            =   2400
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
      Begin VB.Label LblTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Fx"
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
         Left            =   -69960
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LblLoops 
         BackStyle       =   0  'Transparent
         Caption         =   "Loops"
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
         Left            =   -72480
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Label LblFx 
         BackStyle       =   0  'Transparent
         Caption         =   "FX"
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
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
      Begin VB.Label LblWav 
         BackStyle       =   0  'Transparent
         Caption         =   "WAV"
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
         Left            =   -73200
         TabIndex        =   10
         Top             =   540
         Width           =   615
      End
      Begin VB.Image ImgPlay 
         Height          =   360
         Left            =   -72360
         Picture         =   "FrmHechizosCreator.frx":12D75E
         Top             =   960
         Width           =   360
      End
   End
   Begin VB.TextBox TxtMPropio 
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
      Left            =   4920
      TabIndex        =   7
      Top             =   1860
      Width           =   4455
   End
   Begin VB.TextBox TxtMVictima 
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
      Left            =   4920
      TabIndex        =   6
      Top             =   1500
      Width           =   4455
   End
   Begin VB.TextBox TxtMUsuario 
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
      Left            =   4920
      TabIndex        =   5
      Top             =   1140
      Width           =   4455
   End
   Begin VB.TextBox TxtPMagicas 
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
      Left            =   5040
      TabIndex        =   4
      Top             =   780
      Width           =   4335
   End
   Begin VB.TextBox TxtDescripcion 
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
      TabIndex        =   3
      Top             =   420
      Width           =   5175
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
      Left            =   4200
      TabIndex        =   2
      Top             =   60
      Width           =   5175
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
      Left            =   960
      TabIndex        =   1
      Top             =   30
      Width           =   1935
   End
   Begin VB.Label CmdVolver 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2880
      TabIndex        =   133
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label CmdCrear 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   8640
      TabIndex        =   132
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label CmdModif 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4920
      TabIndex        =   131
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu InserParm 
         Caption         =   "Insertar Parametros"
         Begin VB.Menu HType 
            Caption         =   "Tipos de Hechizos"
         End
         Begin VB.Menu HTarget 
            Caption         =   "Targets(Objetivos)"
         End
      End
   End
   Begin VB.Menu Reloads 
      Caption         =   "Reloads"
      Begin VB.Menu HReload 
         Caption         =   "Reload Hechizos"
      End
      Begin VB.Menu HReloadFx 
         Caption         =   "Reload Fx"
      End
      Begin VB.Menu HWavReload 
         Caption         =   "Reload Wavs"
      End
      Begin VB.Menu HAllReload 
         Caption         =   "Reload All"
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
         Begin VB.Menu MObj 
            Caption         =   "Objetos"
         End
         Begin VB.Menu MNpcs 
            Caption         =   "Npcs"
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
Attribute VB_Name = "FrmHechizosCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LB_FINDSTRING = &H18F
Private Declare Function sendmessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Private Const SND_FILENAME = &H20000
Private Const SND_NODEFAULT = &H2
Private Const SND_RESOURCE = &H40004
Private Const SND_ASYNC = &H1

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long
Dim Sony$

Dim HechizoIndex As Integer
Dim CantHechizos As Integer
Dim GrhIndex As Integer
Dim CantFrames As Integer
Dim FrameActual As Integer
Dim HechizoPath As String
Dim HechizoPathSave As String
Dim Mensaje As String

Private Sub Check9_Click()

End Sub

Private Sub CmdCrear_Click()
If FileExist(Config.SaveDatPath & "\Hechizos.dat", vbNormal) Then
    CantHechizos = CantHechizos + 1
    HechizoIndex = CantHechizos
    Call CrearHechizo(True)
Else
    Call FileCopy(Config.DatPath & "\Hechizos.dat", Config.SaveDatPath & "\Hechizos.dat")
    CantHechizos = CantHechizos + 1
    HechizoIndex = CantHechizos
    Call CrearHechizo(True)
End If

If Mensaje <> "" Then
    MsgBox "El hechizo no se ha podido crear ya que se genero un error al momento de guardarlo"
    Mensaje = ""
Else
    MsgBox "El hechizo se ha creado con exito."
End If
End Sub

Private Sub CmdModif_Click()
If FileExist(Config.SaveDatPath & "\Hechizos.dat", vbNormal) Then
    Call CrearHechizo(False)
Else
    Call FileCopy(Config.DatPath & "\Hechizos.dat", Config.SaveDatPath & "\Hechizos.dat")
    Call CrearHechizo(False)
End If
If Mensaje <> "" Then
    MsgBox "El hechizo no se ha podido modificar ya que se genero un error al momento de guardarlo"
    Mensaje = ""
Else
    MsgBox "El hechizo se ha modificado con exito."
End If
End Sub

Private Sub CmdVolver_Click()
FrmDatMenu.Visible = True
Me.Visible = False
End Sub

Private Sub Form_Load()
HechizoPath = Config.DatPath & "\Hechizos.dat"
HechizoPathSave = Config.SaveDatPath & "\Hechizos.dat"
OpName.value = True
Call ReloadData
TxtWav.Enabled = False
TxtFx.Enabled = False
End Sub

Sub Reproducir_WAV(Archivo As String, Flags As Long)
    Dim ret As Long
    ret = PlaySound(Archivo, ByVal 0&, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmDatMenu.Visible = True
End Sub

Private Sub HAllReload_Click()
Call ReloadData
End Sub

Private Sub HReload_Click()
Dim Hechizo As Integer
Dim HechizoName As String
Dim CantHechizo As Integer
LstHechizos.Clear
LstHName.Clear
CantHechizo = GetVar(Config.DatPath & "\Hechizos.dat", "INIT", "NumeroHechizos")
For Hechizo = 1 To CantHechizo
    HechizoName = GetVar(Config.DatPath & "\Hechizos.dat", "HECHIZO" & Hechizo, "Nombre")
    If HechizoName <> "" Then
        LstHechizos.AddItem Hechizo & "-" & HechizoName
        LstHName.AddItem HechizoName
    End If
Next Hechizo
End Sub

Private Sub HReloadFx_Click()
Dim Cant As Integer
Dim Dato As String
Dim Cont As Integer
LstHFx.Clear
Cant = FxCountNew
For Cont = 1 To Cant
    If Fx(Cont).Animacion <> 0 Then
        LstHFx.AddItem Cont
    End If
Next Cont
End Sub

Private Sub HTarget_Click()
Dim NewTarget As Integer
Dim NewTargetName As String
NewTarget = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "HechizosTargets"))
NewTarget = NewTarget + 1
NewTargetName = InputBox("Ingrese el nuevo target(objetivo) para el hechizo", "Target(objetivo) para Hechizos")
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "HechizosTargets", NewTarget)
Call WriteVar(App.path & "\IndexerDats.dao", "HECHIZOTARGETS", "HechizoTarget" & NewTarget, NewTargetName)

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
CbTarget.Clear
Cant = GetVar(IndexDaterIni, "CANTS", "HechizosTargets")
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "HECHIZOTARGETS", "HechizoTarget" & Cont)
    CbTarget.AddItem Cont & "-" & Dato
Next Cont
CbTarget.AddItem "(Elegir)", Cant
CbTarget.ListIndex = Cant
End Sub

Private Sub HType_Click()
Dim NewType As Integer
Dim NewTypeName As String
NewType = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "HechizosType"))
NewType = NewType + 1
NewTypeName = InputBox("Ingrese el nuevo tipo de hechizo", "Tipo de Hechizos")
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "HechizosType", NewType)
Call WriteVar(App.path & "\IndexerDats.dao", "HECHIZOTYPES", "HechizoType" & NewType, NewTypeName)

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
CbTipo.Clear
Cant = GetVar(IndexDaterIni, "CANTS", "HechizosType")
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "HECHIZOTYPES", "HechizoType" & Cont)
    CbTipo.AddItem Cont & "-" & Dato
Next Cont
CbTipo.AddItem "(Elegir)", Cant
CbTipo.ListIndex = Cant
End Sub

Private Sub HWavReload_Click()
WavList.Refresh
End Sub

Private Sub ImgPlay_Click()
Sony = WavList.path & "\" & WavList.FileName
Call Reproducir_WAV(Sony, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT)
End Sub

Private Sub LstHechizos_DblClick()
On Error GoTo ErrHandler
Call SetComboValue

HechizoIndex = Mid(LstHechizos.Text, 1, InStr(1, LstHechizos.Text, "-") - 1)

TxtNombre.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Nombre")
TxtDescripcion.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Desc")
TxtPMagicas.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "PalabrasMagicas")
TxtMUsuario.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "HechizeroMsg")
TxtMVictima.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "TargetMsg")
TxtMPropio.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "PropioMsg")
CbTipo.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Tipo")) - 1
CbTarget.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Target")) - 1
TxtSkills.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinSkill")
TxtMana.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "ManaRequerido")
TxtStamina.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "StaRequerido")
CbResis.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Resis"))
TxtNStaff.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "NeedStaff")
CbSAffected.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "StaffAffected"))

Dim ValorSube As Byte
'Valores Sube o Baja entonces "(Elige)" si valor es 0, entonces ListIndex = 2
ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeMana"))
If ValorSube = 0 Then
    CbSMana.ListIndex = 0
    ChMP.value = 0
Else
    ChMP.value = 1
    CbSMana.ListIndex = ValorSube
    TxtMpMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinMana")
    TxtMpMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxMana")
End If

ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeHp"))
If ValorSube = 0 Then
    CbSVida.ListIndex = 0
    ChHP.value = 0
Else
    ChHP.value = 1
    CbSVida.ListIndex = ValorSube
    TxtHpMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinHP")
    TxtHpMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxHP")
End If

ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeSta"))
If ValorSube = 0 Then
    CbSStamina.ListIndex = 0
    ChSP.value = 0
Else
    ChSP.value = 1
    CbSStamina.ListIndex = ValorSube
    TxtStMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinSta")
    TxtStMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxSta")
End If

ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeFU"))
If ValorSube = 0 Then
    CbSFuerza.ListIndex = 0
    ChStr.value = 0
Else
    ChStr.value = 1
    CbSFuerza.ListIndex = ValorSube
    TxtFzMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinFU")
    TxtFzMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxFU")
End If

ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeAG"))
If ValorSube = 0 Then
    CbSAgilidad.ListIndex = 0
    ChAgi.value = 0
Else
    ChAgi.value = 1
    CbSAgilidad.ListIndex = ValorSube
    TxtAgMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinAG")
    TxtAgMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxAG")
End If

ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeCA"))
If ValorSube = 0 Then
    CbSCarisma.ListIndex = 0
    ChCa.value = 0
Else
    ChCa.value = 1
    CbSCarisma.ListIndex = ValorSube
    TxtCaMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinCA")
    TxtCaMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxCA")
End If

ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeHam"))
If ValorSube = 0 Then
    CbSHambre.ListIndex = 0
    ChHam.value = 0
Else
    ChHam.value = 1
    CbSHambre.ListIndex = ValorSube
    TxtHaMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinHam")
    TxtHaMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxHam")
End If

ValorSube = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "SubeSed"))
If ValorSube = 0 Then
    CbSSed.ListIndex = 0
    ChSed.value = 0
Else
    ChSed.value = 1
    CbSSed.ListIndex = ValorSube
    TxtSdMin.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MinSed")
    TxtSdMax.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "MaxSed")
End If

TxtWav.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Wav")
TxtFx.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "FXgrh")
TxtLoops.Text = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Loops")

CbInvi.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Invisibilidad"))
CbPara.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Paraliza"))
CbInmo.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Inmoviliza"))
CbEnv.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Envenena"))
CbCurEnv.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "CuraVeneno"))
CbRev.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Revivir"))
CbCeg.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Ceguera"))
CbEstu.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Estupidez"))
CbMime.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Mimetiza"))
CbInvo.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Invoca"))
TxtNumNpc.Text = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "NumNpc"))
TxtCant.Text = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "Cant"))
CbMate.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "materializa"))
TxtNumObj.Text = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "itemindex"))
CbRemoInvi.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "RemueveInvisibilidadParcial"))
CbRemoPara.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "RemoverParalisis"))
CbRemoEstu.ListIndex = Val(GetVar(HechizoPath, "HECHIZO" & HechizoIndex, "RemoverEstupidez"))

Dim T As Integer
Dim CantParm As Integer
CantParm = LstAdd.ListCount
LstAdd.Clear
For T = 1 To CantParm
    Dim ValorParm As String
    Dim Parametro As String
    Parametro = GetVar(IndexDaterIni, "PARAMETROSHECHIZOS", "Parametro" & T)
    ValorParm = GetVar(HechizoPath, "HECHIZO" & HechizoIndex, Parametro)
    If ValorParm = "" Then ValorParm = 0
    LstAdd.AddItem Parametro & "=" & ValorParm
Next T

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub

End Sub

Private Sub LstHFx_Click()
On Error GoTo ErrHandler
Dim FxIndex As Integer
FxIndex = Val(LstHFx.Text)

If FxIndex = 0 Then Exit Sub

TxtFx.Text = LstHFx.Text

GrhIndex = Fx(FxIndex).Animacion
If GrhIndex = 0 Then Exit Sub

CantFrames = GrhData(GrhIndex).NumFrames

If GrhData(GrhIndex).Speed <> 0 And GrhData(GrhIndex).NumFrames <> 0 Then
    If IndexMode = "12.1" Then
        TFx.Interval = Round(GrhData(GrhIndex).Speed / GrhData(GrhIndex).NumFrames)
    Else
        TFx.Interval = 100
    End If
    TFx.Enabled = True
Else
    TFx.Enabled = False
End If
FrameActual = 0
PbHFx.Cls

Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub MObjetos_Click()
FrmObjSelector.Visible = True
Me.Visible = False
End Sub

Private Sub MConversor_Click()
FrmConversor.Visible = True
End Sub

Private Sub MObj_Click()
FrmObjSelector.Visible = True
Me.Visible = False
End Sub

Private Sub MRutas_Click()
frmConfig.Visible = True
End Sub

Private Sub OpName_Click()
If OpName.value = True Then OpNum.value = False
End Sub

Private Sub OpNum_Click()
If OpNum.value = True Then OpName.value = False
End Sub

Private Sub TFx_Timer()
On Error GoTo ErrHandler
Dim AnimacionPosX As Integer
Dim AnimacionPosY As Integer
Dim AnimacionLongX As Integer
Dim AnimacionLongY As Integer
Dim GraficoPath As String
Dim GrhIndexAnim As Long

FrameActual = FrameActual + 1

GrhIndexAnim = GrhData(GrhIndex).Frames(FrameActual)

If GrhIndexAnim > 0 Then
    GraficoPath = Config.BmpPath & "\" & GrhData(GrhIndexAnim).FileNum & ".bmp"
    AnimacionPosX = GrhData(GrhIndexAnim).sX
    AnimacionPosY = GrhData(GrhIndexAnim).sY
    AnimacionLongX = GrhData(GrhIndexAnim).pixelWidth
    AnimacionLongY = GrhData(GrhIndexAnim).pixelHeight
    
    If FileExist(GraficoPath, vbNormal) = True Then
        PbHFx.PaintPicture LoadPicture(GraficoPath), 0, 0, AnimacionLongX, AnimacionLongY, AnimacionPosX, AnimacionPosY, AnimacionLongX, AnimacionLongY
    End If
End If

If FrameActual = CantFrames Then
    FrameActual = 0
End If

Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    TFx.Enabled = False
    Exit Sub
End Sub

Private Sub TxtSearch_Change()
    If OpNum.value = True Then
        LstHechizos.ListIndex = sendmessage(LstHechizos.hWnd, LB_FINDSTRING, -1, ByVal TxtSearch.Text)
    Else
        LstHName.ListIndex = sendmessage(LstHName.hWnd, LB_FINDSTRING, -1, ByVal TxtSearch.Text)
        LstHechizos.ListIndex = LstHName.ListIndex
    End If
End Sub

Private Sub TxtSearchFx_Change()
LstHFx.ListIndex = sendmessage(LstHFx.hWnd, LB_FINDSTRING, -1, ByVal TxtSearchFx.Text)
End Sub

Private Sub WavList_Click()
TxtWav.Text = Mid(WavList.FileName, 1, InStr(1, WavList.FileName, ".") - 1)
End Sub

Private Sub WavList_DblClick()
Sony = WavList.path & "\" & WavList.FileName
Call Reproducir_WAV(Sony, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT)
End Sub

Private Sub ReloadData()
Call ClearAll

Dim Hechizo As Integer
Dim HechizoName As String
CantHechizos = GetVar(Config.DatPath & "\Hechizos.dat", "INIT", "NumeroHechizos")
For Hechizo = 1 To CantHechizos
    HechizoName = GetVar(Config.DatPath & "\Hechizos.dat", "HECHIZO" & Hechizo, "Nombre")
    If HechizoName <> "" Then
        LstHechizos.AddItem Hechizo & "-" & HechizoName
        LstHName.AddItem HechizoName
    End If
Next Hechizo

WavList.path = Config.WavPath
WavList.Pattern = "*.WAV"

Dim Cant As Integer
Dim Dato As String
Dim Cont As Integer

Cant = FxCountNew
For Cont = 1 To Cant
    If Fx(Cont).Animacion <> 0 Then
        LstHFx.AddItem Cont
    End If
Next Cont

Cant = GetVar(IndexDaterIni, "CANTS", "HechizosType")
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "HECHIZOTYPES", "HechizoType" & Cont)
    CbTipo.AddItem Cont & "-" & Dato
Next Cont
CbTipo.AddItem "(Elegir)", Cant
CbTipo.ListIndex = Cant

Cant = GetVar(IndexDaterIni, "CANTS", "HechizosTargets")
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "HECHIZOTARGETS", "HechizoTarget" & Cont)
    CbTarget.AddItem Cont & "-" & Dato
Next Cont
CbTarget.AddItem "(Elegir)", Cant
CbTarget.ListIndex = Cant

Cant = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosHechizos"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROSHECHIZOS", "Parametro" & Cont)
    LstAdd.AddItem Dato & "=0"
Next Cont

'Valores "SI" o "NO" en Generales
CbResis.AddItem "No", 0
CbResis.AddItem "Si", 1
CbResis.AddItem "(Elegir)", 2

CbSAffected.AddItem "No", 0
CbSAffected.AddItem "Si", 1
CbSAffected.AddItem "(Elegir)", 2

'Valores "SI" o "NO" en Opciones
CbInvi.AddItem "No", 0
CbInvi.AddItem "Si", 1
CbInvi.AddItem "(Elegir)", 2

CbPara.AddItem "No", 0
CbPara.AddItem "Si", 1
CbPara.AddItem "(Elegir)", 2

CbInmo.AddItem "No", 0
CbInmo.AddItem "Si", 1
CbInmo.AddItem "(Elegir)", 2

CbEnv.AddItem "No", 0
CbEnv.AddItem "Si", 1
CbEnv.AddItem "(Elegir)", 2

CbCurEnv.AddItem "No", 0
CbCurEnv.AddItem "Si", 1
CbCurEnv.AddItem "(Elegir)", 2

CbRev.AddItem "No", 0
CbRev.AddItem "Si", 1
CbRev.AddItem "(Elegir)", 2

CbCeg.AddItem "No", 0
CbCeg.AddItem "Si", 1
CbCeg.AddItem "(Elegir)", 2

CbEstu.AddItem "No", 0
CbEstu.AddItem "Si", 1
CbEstu.AddItem "(Elegir)", 2

CbMime.AddItem "No", 0
CbMime.AddItem "Si", 1
CbMime.AddItem "(Elegir)", 2

CbInvo.AddItem "No", 0
CbInvo.AddItem "Si", 1
CbInvo.AddItem "(Elegir)", 2

CbMate.AddItem "No", 0
CbMate.AddItem "Si", 1
CbMate.AddItem "(Elegir)", 2

CbRemoInvi.AddItem "No", 0
CbRemoInvi.AddItem "Si", 1
CbRemoInvi.AddItem "(Elegir)", 2

CbRemoPara.AddItem "No", 0
CbRemoPara.AddItem "Si", 1
CbRemoPara.AddItem "(Elegir)", 2

CbRemoEstu.AddItem "No", 0
CbRemoEstu.AddItem "Si", 1
CbRemoEstu.AddItem "(Elegir)", 2

'Valores "Baja" y "Sube" para los daños
CbSMana.AddItem "(Elegir)", 0
CbSMana.AddItem "Sube", 1
CbSMana.AddItem "Baja", 2

CbSVida.AddItem "(Elegir)", 0
CbSVida.AddItem "Sube", 1
CbSVida.AddItem "Baja", 2

CbSStamina.AddItem "(Elegir)", 0
CbSStamina.AddItem "Sube", 1
CbSStamina.AddItem "Baja", 2

CbSFuerza.AddItem "(Elegir)", 0
CbSFuerza.AddItem "Sube", 1
CbSFuerza.AddItem "Baja", 2

CbSAgilidad.AddItem "(Elegir)", 0
CbSAgilidad.AddItem "Sube", 1
CbSAgilidad.AddItem "Baja", 2

CbSCarisma.AddItem "(Elegir)", 0
CbSCarisma.AddItem "Sube", 1
CbSCarisma.AddItem "Baja", 2

CbSHambre.AddItem "(Elegir)", 0
CbSHambre.AddItem "Sube", 1
CbSHambre.AddItem "Baja", 2

CbSSed.AddItem "(Elegir)", 0
CbSSed.AddItem "Sube", 1
CbSSed.AddItem "Baja", 2

Call SetComboValue
End Sub

Private Sub ClearAll()
LstHechizos.Clear
LstHName.Clear
CbTipo.Clear
CbTarget.Clear
CbResis.Clear
CbSAffected.Clear
CbInvi.Clear
CbPara.Clear
CbInmo.Clear
CbEnv.Clear
CbCurEnv.Clear
CbRev.Clear
CbCeg.Clear
CbEstu.Clear
CbMime.Clear
CbInvo.Clear
CbMate.Clear
CbRemoInvi.Clear
CbRemoPara.Clear
CbRemoEstu.Clear
CbSMana.Clear
CbSVida.Clear
CbSStamina.Clear
CbSFuerza.Clear
CbSAgilidad.Clear
CbSCarisma.Clear
CbSHambre.Clear
CbSSed.Clear
End Sub

Private Sub SetComboValue()
CbResis.ListIndex = 2
CbSAffected.ListIndex = 2
CbInvi.ListIndex = 2
CbPara.ListIndex = 2
CbInmo.ListIndex = 2
CbEnv.ListIndex = 2
CbCurEnv.ListIndex = 2
CbRev.ListIndex = 2
CbCeg.ListIndex = 2
CbEstu.ListIndex = 2
CbMime.ListIndex = 2
CbInvo.ListIndex = 2
CbMate.ListIndex = 2
CbRemoInvi.ListIndex = 2
CbRemoPara.ListIndex = 2
CbRemoEstu.ListIndex = 2
CbSMana.ListIndex = 0
CbSVida.ListIndex = 0
CbSStamina.ListIndex = 0
CbSFuerza.ListIndex = 0
CbSAgilidad.ListIndex = 0
CbSCarisma.ListIndex = 0
CbSHambre.ListIndex = 0
CbSSed.ListIndex = 0
TxtNombre.Text = ""
TxtDescripcion.Text = ""
TxtPMagicas.Text = ""
TxtMUsuario.Text = ""
TxtMVictima.Text = ""
TxtMPropio.Text = ""
TxtMana.Text = ""
TxtStamina.Text = ""
TxtSkills.Text = ""
TxtNStaff.Text = ""
TxtMpMin.Text = ""
TxtMpMax.Text = ""
TxtHpMin.Text = ""
TxtHpMax.Text = ""
TxtStMin.Text = ""
TxtStMax.Text = ""
TxtFzMin.Text = ""
TxtFzMax.Text = ""
TxtAgMin.Text = ""
TxtAgMax.Text = ""
TxtCaMin.Text = ""
TxtCaMax.Text = ""
TxtHaMin.Text = ""
TxtHaMax.Text = ""
TxtSdMin.Text = ""
TxtSdMax.Text = ""
TxtNumNpc.Text = ""
TxtCant.Text = ""
TxtNumObj.Text = ""
TxtWav.Text = ""
TxtFx.Text = ""
TxtLoops.Text = ""
End Sub

Private Sub CrearHechizo(Crear As Boolean)
If HzError = False Then
    MsgBox Mensaje
    Exit Sub
End If
If CbTipo.Text = "(Elegir)" Then
    MsgBox "Primero debes elegir un 'Tipo' de hechizo para poder continuar'"
    Exit Sub
End If

If Crear = True Then
    Call WriteVar(HechizoPathSave, "INIT", "NumeroHechizos", HechizoIndex)
End If

Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Nombre", TxtNombre.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Desc", TxtDescripcion.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "PalabrasMagicas", TxtPMagicas.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "HechizeroMsg", TxtMUsuario.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "TargetMsg", TxtMVictima.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "PropioMsg", TxtMPropio.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Tipo", Mid(CbTipo.Text, 1, InStr(1, CbTipo.Text, "-") - 1))
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "WAV", TxtWav.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "FXgrh", TxtFx.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Loops", TxtLoops.Text)

'Daños de los Hechizos
If CbSVida.ListIndex <> 0 And ChHP.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeHP", Str(CbSVida.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinHP", TxtHpMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxHP", TxtHpMax.Text)
End If
If CbSMana.ListIndex <> 0 And ChMP.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeMana", Str(CbSMana.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinMana", TxtMpMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxMana", TxtMpMax.Text)
End If
If CbSStamina.ListIndex <> 0 And ChSP.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeSta", Str(CbSStamina.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinSta", TxtStMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxSta", TxtStMax.Text)
End If
If CbSHambre.ListIndex <> 0 And ChHam.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeHam", Str(CbSHambre.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinHam", TxtHaMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxHam", TxtHaMax.Text)
End If
If CbSSed.ListIndex <> 0 And ChSed.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeSed", Str(CbSSed.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinSed", TxtSdMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxSed", TxtSdMax.Text)
End If
If CbSAgilidad.ListIndex <> 0 And ChAgi.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeAG", Str(CbSAgilidad.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinAG", TxtAgMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxAG", TxtAgMax.Text)
End If
If CbSFuerza.ListIndex <> 0 And ChStr.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeFU", Str(CbSFuerza.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinFU", TxtFzMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxFU", TxtFzMax.Text)
End If
If CbSCarisma.ListIndex <> 0 And ChCa.value = 1 Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "SubeCA", Str(CbSCarisma.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinCA", TxtCaMin.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MaxCA", TxtCaMax.Text)
End If

'Propiedades de los Hechizos
If CbInvi.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Invisibilidad", Str(CbInvi.ListIndex))
End If
If CbPara.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Paraliza", Str(CbPara.ListIndex))
End If
If CbCurEnv.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "CuraVeneno", Str(CbCurEnv.ListIndex))
End If
If CbEnv.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Envenena", Str(CbEnv.ListIndex))
End If
If CbInmo.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Inmoviliza", Str(CbInmo.ListIndex))
End If
If CbRev.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Revivir", Str(CbRev.ListIndex))
End If
If CbInvo.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Invoca", Str(CbInvo.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "NumNpc", TxtNumNpc.Text)
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Cant", TxtCant.Text)
End If
If CbMate.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "materializa", Str(CbMate.ListIndex))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "itemindex", TxtNumObj.Text)
End If
If CbEstu.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Estupidez", Str(CbEstu.ListIndex))
End If
If CbRemoEstu.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "RemoverEstupidez", Str(CbRemoEstu.ListIndex))
End If
If CbCeg.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Ceguera", Str(CbCeg.ListIndex))
End If
If CbMime.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Mimetiza", Str(CbMime.ListIndex))
End If
If CbRemoPara.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "RemoverParalisis", Str(CbRemoPara.ListIndex))
End If
If CbRemoInvi.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "RemueveInvisibilidadParcial", Str(CbRemoInvi.ListIndex))
End If
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "MinSkill", TxtSkills.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "ManaRequerido", TxtMana.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Target", Mid(CbTarget.Text, 1, InStr(1, CbTarget.Text, "-") - 1))
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "StaRequerido", TxtStamina.Text)
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "Resis", Str(CbResis.ListIndex))
Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "NeedStaff", TxtNStaff.Text)
If CbSAffected.Text <> "(Elegir)" Then
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, "StaffAffected", Str(CbSAffected.ListIndex))
End If

Dim Parms As Integer
For Parms = 0 To LstAdd.ListCount - 1
    Dim NombreParm As String
    Dim ValorParm As String
    NombreParm = Left(LstAdd.List(Parms), InStr(1, LstAdd.List(Parms), "=") - 1)
    ValorParm = Mid(LstAdd.List(Parms), InStr(1, LstAdd.List(Parms), "=") + 1, Trim(Len(LstAdd.List(Parms))))
    Call WriteVar(HechizoPathSave, "HECHIZO" & HechizoIndex, NombreParm, ValorParm)
Next Parms

End Sub

Private Sub MCreditos_Click()
FrmCreditos.Visible = True
Me.Visible = False
End Sub

Private Sub MDatosIndex_Click()
FrmAnimaciones.Visible = True
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

Private Sub MNpcs_Click()
FrmNpcCreator.Visible = True
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

CantParm = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosHechizos"))
TotalParm = CantParm + 1
Call WriteVar(IndexDaterIni, "CANTS", "ParametrosHechizos", TotalParm)
Call WriteVar(IndexDaterIni, "PARAMETROSHECHIZOS", "Parametro" & TotalParm, NewParm)

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
LstAdd.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosHechizos"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROSHECHIZOS", "Parametro" & Cont)
    LstAdd.AddItem Dato & "=0"
Next Cont

MsgBox "Parametro Ingresado"
End Sub

Private Sub CmdCargar_Click()
Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
LstAdd.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "ParametrosHechizos"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROSHECHIZOS", "Parametro" & Cont)
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

Public Function HzError() As Boolean
HzError = True
If CbTipo.Text = "(Elegir)" Then
    Mensaje = "Eliga un Tipo de Hechizo"
    HzError = False
End If
If CbTarget.Text = "(Elegir)" Then
    Mensaje = "Eliga un Objetivo(Target) para el Hechizo"
    HzError = False
End If
End Function
