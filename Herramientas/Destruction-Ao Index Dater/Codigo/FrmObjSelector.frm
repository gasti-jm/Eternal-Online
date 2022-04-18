VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmObjSelector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Objetos"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   10845
   Icon            =   "FrmObjSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmObjSelector.frx":08CA
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   723
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OpName 
      BackColor       =   &H00C00000&
      Height          =   195
      Left            =   1770
      TabIndex        =   157
      Top             =   480
      Width           =   180
   End
   Begin VB.OptionButton OpNum 
      BackColor       =   &H00C00000&
      Height          =   195
      Left            =   690
      TabIndex        =   156
      Top             =   480
      Width           =   180
   End
   Begin VB.ListBox ObjList 
      Height          =   6495
      ItemData        =   "FrmObjSelector.frx":157B56
      Left            =   0
      List            =   "FrmObjSelector.frx":157B58
      TabIndex        =   63
      Top             =   720
      Width           =   2775
   End
   Begin VB.ListBox ObjListCopy 
      Height          =   6885
      ItemData        =   "FrmObjSelector.frx":157B5A
      Left            =   0
      List            =   "FrmObjSelector.frx":157B5C
      TabIndex        =   155
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox TxtSearch 
      Height          =   285
      Left            =   1080
      TabIndex        =   81
      Top             =   30
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6855
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmObjSelector.frx":157B5E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblNombre"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblPociones"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblsubTipo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblTipo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblGrhIndex"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblAgarrable"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblAlineacion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LblCrucial"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblValor"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblRaza"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblHIndex"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LblDuracion"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtNombre"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ChNombre"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ChPociones"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CbPociones"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ChSubTipo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CbTipo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ChTipo"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtGrhIndex"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ChGrhIndex"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CbAgarrable"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ChAgarrable"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CbAlineacion"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "ChAlineacion"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ChCrucial"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtValor"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ChValor"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "CbCrucial"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ChRaza"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CbRaza"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtHIndex"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "ChHindex"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtDuracion"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "ChDuracion"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "LstGrhIndex"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "PbGrhIndex"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "CbSubTipo"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtSGrhIndex"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtSHIndex"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "THechizo"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "LstHIndex"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Ropajes"
      TabPicture(1)   =   "FrmObjSelector.frx":157B7A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TAllDirections"
      Tab(1).Control(1)=   "TxtSRopaje"
      Tab(1).Control(2)=   "PbRopaje(4)"
      Tab(1).Control(3)=   "PbRopaje(3)"
      Tab(1).Control(4)=   "PbRopaje(2)"
      Tab(1).Control(5)=   "PbRopaje(1)"
      Tab(1).Control(6)=   "LstRopaje"
      Tab(1).Control(7)=   "ChRopaje"
      Tab(1).Control(8)=   "TxtRopaje"
      Tab(1).Control(9)=   "ChDefMin"
      Tab(1).Control(10)=   "TxtDefMin"
      Tab(1).Control(11)=   "ChDefMax"
      Tab(1).Control(12)=   "TxtDefMax"
      Tab(1).Control(13)=   "Label6"
      Tab(1).Control(14)=   "Label3"
      Tab(1).Control(15)=   "LblRopaje"
      Tab(1).Control(16)=   "LblDefMin"
      Tab(1).Control(17)=   "LblDefMax"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Armas"
      TabPicture(2)   =   "FrmObjSelector.frx":157B96
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtSAnimacion"
      Tab(2).Control(1)=   "PbAnimacion(4)"
      Tab(2).Control(2)=   "PbAnimacion(3)"
      Tab(2).Control(3)=   "PbAnimacion(2)"
      Tab(2).Control(4)=   "PbAnimacion(1)"
      Tab(2).Control(5)=   "LstAnimaciones"
      Tab(2).Control(6)=   "ChAnimacion"
      Tab(2).Control(7)=   "TxtAnimacion"
      Tab(2).Control(8)=   "ChDanMin"
      Tab(2).Control(9)=   "TxtDanMin"
      Tab(2).Control(10)=   "ChDanMax"
      Tab(2).Control(11)=   "TxtDanMax"
      Tab(2).Control(12)=   "ChApu"
      Tab(2).Control(13)=   "CbApu"
      Tab(2).Control(14)=   "ChMunicion"
      Tab(2).Control(15)=   "CbMunicion"
      Tab(2).Control(16)=   "CbProyectil"
      Tab(2).Control(17)=   "ChProyectil"
      Tab(2).Control(18)=   "Label7"
      Tab(2).Control(19)=   "Label4"
      Tab(2).Control(20)=   "LblAnimacion"
      Tab(2).Control(21)=   "LblDanMin"
      Tab(2).Control(22)=   "LblDanMax"
      Tab(2).Control(23)=   "LblApun"
      Tab(2).Control(24)=   "LblMunicion"
      Tab(2).Control(25)=   "LblProyectil"
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "Escudos"
      TabPicture(3)   =   "FrmObjSelector.frx":157BB2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TxtEscMaxDef"
      Tab(3).Control(1)=   "ChEscMaxDef"
      Tab(3).Control(2)=   "TxtEscMinDef"
      Tab(3).Control(3)=   "ChEscMinDef"
      Tab(3).Control(4)=   "TxtEscAnim"
      Tab(3).Control(5)=   "ChEscAnim"
      Tab(3).Control(6)=   "LstEscudos"
      Tab(3).Control(7)=   "PbEscudo(1)"
      Tab(3).Control(8)=   "PbEscudo(2)"
      Tab(3).Control(9)=   "PbEscudo(3)"
      Tab(3).Control(10)=   "PbEscudo(4)"
      Tab(3).Control(11)=   "TxtEscSearch"
      Tab(3).Control(12)=   "LblEscMaxdef"
      Tab(3).Control(13)=   "LblEscMinDef"
      Tab(3).Control(14)=   "Label10"
      Tab(3).Control(15)=   "Label9"
      Tab(3).Control(16)=   "Label8"
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Cabezas"
      TabPicture(4)   =   "FrmObjSelector.frx":157BCE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TxtCascoSearch"
      Tab(4).Control(1)=   "LstCascos"
      Tab(4).Control(2)=   "TxtHeadSearch"
      Tab(4).Control(3)=   "PbHeads(4)"
      Tab(4).Control(4)=   "PbHeads(3)"
      Tab(4).Control(5)=   "PbHeads(2)"
      Tab(4).Control(6)=   "PbHeads(1)"
      Tab(4).Control(7)=   "LstHeads"
      Tab(4).Control(8)=   "ChHeadAnim"
      Tab(4).Control(9)=   "TxtHeadAnim"
      Tab(4).Control(10)=   "ChHeadDefMin"
      Tab(4).Control(11)=   "TxtHeadDefMin"
      Tab(4).Control(12)=   "ChHeadDefMax"
      Tab(4).Control(13)=   "TxtHeadDefMax"
      Tab(4).Control(14)=   "Label12"
      Tab(4).Control(15)=   "Label11"
      Tab(4).Control(16)=   "Label15"
      Tab(4).Control(17)=   "Label14"
      Tab(4).Control(18)=   "LblHeadAnim"
      Tab(4).Control(19)=   "LblHeadDefMin"
      Tab(4).Control(20)=   "LblHeadDefMax"
      Tab(4).ControlCount=   21
      TabCaption(5)   =   "Skill/Prima"
      TabPicture(5)   =   "FrmObjSelector.frx":157BEA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ChNavegacion"
      Tab(5).Control(1)=   "TxtNavegacion"
      Tab(5).Control(2)=   "ChCarpinteria"
      Tab(5).Control(3)=   "TxtCarpinteria"
      Tab(5).Control(4)=   "ChHerreria"
      Tab(5).Control(5)=   "TxtHerreria"
      Tab(5).Control(6)=   "ChLingOro"
      Tab(5).Control(7)=   "TxtLingOro"
      Tab(5).Control(8)=   "ChLingPlata"
      Tab(5).Control(9)=   "TxtLingPlata"
      Tab(5).Control(10)=   "ChLingHierro"
      Tab(5).Control(11)=   "TxtLingHierro"
      Tab(5).Control(12)=   "ChMadera"
      Tab(5).Control(13)=   "TxtMadera"
      Tab(5).Control(14)=   "LblNavegacion(1)"
      Tab(5).Control(15)=   "LblCarpinteria"
      Tab(5).Control(16)=   "LblHerreria"
      Tab(5).Control(17)=   "LblLingOro"
      Tab(5).Control(18)=   "LblLingPlata"
      Tab(5).Control(19)=   "LblLingHierro"
      Tab(5).Control(20)=   "LblMadera"
      Tab(5).ControlCount=   21
      TabCaption(6)   =   "Clases"
      TabPicture(6)   =   "FrmObjSelector.frx":157C06
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "LstClases"
      Tab(6).Control(1)=   "ChClases"
      Tab(6).Control(2)=   "LblClases"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Agregados"
      TabPicture(7)   =   "FrmObjSelector.frx":157C22
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "LstAdd"
      Tab(7).Control(1)=   "TxtAdd"
      Tab(7).Control(2)=   "CmdChange"
      Tab(7).Control(3)=   "CmdAgregar"
      Tab(7).Control(4)=   "ChAdd"
      Tab(7).Control(5)=   "CmdCargar"
      Tab(7).Control(6)=   "Label16"
      Tab(7).Control(7)=   "LblAdd"
      Tab(7).ControlCount=   8
      Begin VB.ListBox LstHIndex 
         Height          =   2790
         ItemData        =   "FrmObjSelector.frx":157C3E
         Left            =   6240
         List            =   "FrmObjSelector.frx":157C40
         TabIndex        =   66
         Top             =   3760
         Width           =   1575
      End
      Begin VB.ListBox LstAdd 
         Height          =   2985
         ItemData        =   "FrmObjSelector.frx":157C42
         Left            =   -72840
         List            =   "FrmObjSelector.frx":157C44
         TabIndex        =   161
         Top             =   480
         Width           =   5655
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
         Left            =   -74520
         TabIndex        =   160
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton CmdChange 
         Caption         =   "Cambiar"
         Height          =   315
         Left            =   -74520
         TabIndex        =   159
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox TxtCascoSearch 
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
         Left            =   -67920
         TabIndex        =   152
         Top             =   1755
         Width           =   735
      End
      Begin VB.ListBox LstCascos 
         Height          =   4350
         ItemData        =   "FrmObjSelector.frx":157C46
         Left            =   -68760
         List            =   "FrmObjSelector.frx":157C48
         TabIndex        =   151
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TxtHeadSearch 
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
         TabIndex        =   145
         Top             =   1740
         Width           =   735
      End
      Begin VB.PictureBox PbHeads 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1935
         Index           =   4
         Left            =   -70920
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   144
         Top             =   4200
         Width           =   2055
      End
      Begin VB.PictureBox PbHeads 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1935
         Index           =   3
         Left            =   -73200
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   143
         Top             =   4200
         Width           =   2055
      End
      Begin VB.PictureBox PbHeads 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1935
         Index           =   2
         Left            =   -70920
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   142
         Top             =   2040
         Width           =   2055
      End
      Begin VB.PictureBox PbHeads 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1935
         Index           =   1
         Left            =   -73200
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   141
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ListBox LstHeads 
         Height          =   4350
         ItemData        =   "FrmObjSelector.frx":157C4A
         Left            =   -74880
         List            =   "FrmObjSelector.frx":157C4C
         TabIndex        =   140
         Top             =   2025
         Width           =   1575
      End
      Begin VB.CheckBox ChHeadAnim 
         Height          =   255
         Left            =   -74640
         TabIndex        =   139
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxtHeadAnim 
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
         Left            =   -73560
         TabIndex        =   138
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox ChHeadDefMin 
         Height          =   255
         Left            =   -72120
         TabIndex        =   137
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxtHeadDefMin 
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
         Left            =   -70920
         TabIndex        =   136
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox ChHeadDefMax 
         Height          =   255
         Left            =   -69720
         TabIndex        =   135
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxtHeadDefMax 
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
         Left            =   -68400
         TabIndex        =   134
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox ChNavegacion 
         Height          =   255
         Left            =   -74760
         TabIndex        =   126
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxtNavegacion 
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
         Left            =   -73200
         TabIndex        =   125
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox ChCarpinteria 
         Height          =   255
         Left            =   -72240
         TabIndex        =   124
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxtCarpinteria 
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
         Left            =   -70800
         TabIndex        =   123
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox ChHerreria 
         Height          =   255
         Left            =   -69600
         TabIndex        =   122
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxtHerreria 
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
         Left            =   -68160
         TabIndex        =   121
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox ChLingOro 
         Height          =   255
         Left            =   -74760
         TabIndex        =   120
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox TxtLingOro 
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
         Left            =   -73200
         TabIndex        =   119
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox ChLingPlata 
         Height          =   255
         Left            =   -72240
         TabIndex        =   118
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox TxtLingPlata 
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
         Left            =   -70800
         TabIndex        =   117
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox ChLingHierro 
         Height          =   255
         Left            =   -69600
         TabIndex        =   116
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox TxtLingHierro 
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
         Left            =   -68160
         TabIndex        =   115
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox ChMadera 
         Height          =   255
         Left            =   -74760
         TabIndex        =   114
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox TxtMadera 
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
         Left            =   -73200
         TabIndex        =   113
         Top             =   1440
         Width           =   735
      End
      Begin VB.ListBox LstClases 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         ItemData        =   "FrmObjSelector.frx":157C4E
         Left            =   -70800
         List            =   "FrmObjSelector.frx":157C50
         MultiSelect     =   2  'Extended
         TabIndex        =   111
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox ChClases 
         Height          =   255
         Left            =   -73080
         TabIndex        =   110
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Parametros"
         Height          =   375
         Left            =   -74520
         TabIndex        =   108
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox ChAdd 
         Height          =   255
         Left            =   -74760
         TabIndex        =   107
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar Parametros"
         Height          =   375
         Left            =   -74520
         TabIndex        =   106
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TxtEscMaxDef 
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
         Left            =   -68400
         TabIndex        =   100
         Top             =   500
         Width           =   1215
      End
      Begin VB.CheckBox ChEscMaxDef 
         Height          =   255
         Left            =   -69720
         TabIndex        =   99
         Top             =   500
         Width           =   255
      End
      Begin VB.TextBox TxtEscMinDef 
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
         Left            =   -70920
         TabIndex        =   98
         Top             =   500
         Width           =   1095
      End
      Begin VB.CheckBox ChEscMinDef 
         Height          =   255
         Left            =   -72120
         TabIndex        =   97
         Top             =   500
         Width           =   255
      End
      Begin VB.TextBox TxtEscAnim 
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
         Left            =   -73560
         TabIndex        =   96
         Top             =   500
         Width           =   1215
      End
      Begin VB.CheckBox ChEscAnim 
         Height          =   255
         Left            =   -74880
         TabIndex        =   95
         Top             =   500
         Width           =   255
      End
      Begin VB.ListBox LstEscudos 
         Height          =   4350
         ItemData        =   "FrmObjSelector.frx":157C52
         Left            =   -74880
         List            =   "FrmObjSelector.frx":157C54
         TabIndex        =   94
         Top             =   1925
         Width           =   2055
      End
      Begin VB.PictureBox PbEscudo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   1
         Left            =   -72720
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   93
         Top             =   1565
         Width           =   2415
      End
      Begin VB.PictureBox PbEscudo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   2
         Left            =   -69960
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   92
         Top             =   1565
         Width           =   2415
      End
      Begin VB.PictureBox PbEscudo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   3
         Left            =   -72720
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   91
         Top             =   4085
         Width           =   2415
      End
      Begin VB.PictureBox PbEscudo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   4
         Left            =   -69960
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   90
         Top             =   4085
         Width           =   2415
      End
      Begin VB.TextBox TxtEscSearch 
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
         TabIndex        =   89
         Top             =   1640
         Width           =   1215
      End
      Begin VB.Timer THechizo 
         Left            =   7200
         Top             =   2300
      End
      Begin VB.Timer TAllDirections 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   -67680
         Top             =   880
      End
      Begin VB.TextBox TxtSHIndex 
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
         TabIndex        =   88
         Top             =   3400
         Width           =   1575
      End
      Begin VB.TextBox TxtSAnimacion 
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
         TabIndex        =   86
         Top             =   2080
         Width           =   1215
      End
      Begin VB.TextBox TxtSGrhIndex 
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
         Left            =   120
         TabIndex        =   85
         Top             =   3400
         Width           =   1575
      End
      Begin VB.TextBox TxtSRopaje 
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
         TabIndex        =   84
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox CbSubTipo 
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
         ItemData        =   "FrmObjSelector.frx":157C56
         Left            =   3960
         List            =   "FrmObjSelector.frx":157C58
         TabIndex        =   82
         Top             =   1000
         Width           =   1095
      End
      Begin VB.PictureBox PbAnimacion 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   4
         Left            =   -69960
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   80
         Top             =   4480
         Width           =   2415
      End
      Begin VB.PictureBox PbAnimacion 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   3
         Left            =   -72720
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   79
         Top             =   4480
         Width           =   2415
      End
      Begin VB.PictureBox PbAnimacion 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   2
         Left            =   -69960
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   78
         Top             =   1960
         Width           =   2415
      End
      Begin VB.PictureBox PbAnimacion 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   1
         Left            =   -72720
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   77
         Top             =   1960
         Width           =   2415
      End
      Begin VB.ListBox LstAnimaciones 
         Height          =   4155
         ItemData        =   "FrmObjSelector.frx":157C5A
         Left            =   -74880
         List            =   "FrmObjSelector.frx":157C5C
         TabIndex        =   75
         Top             =   2440
         Width           =   2055
      End
      Begin VB.PictureBox PbRopaje 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   4
         Left            =   -69960
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   74
         Top             =   4120
         Width           =   2415
      End
      Begin VB.PictureBox PbRopaje 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   3
         Left            =   -72720
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   73
         Top             =   4120
         Width           =   2415
      End
      Begin VB.PictureBox PbRopaje 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   2
         Left            =   -69960
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   72
         Top             =   1600
         Width           =   2415
      End
      Begin VB.PictureBox PbRopaje 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Index           =   1
         Left            =   -72720
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   71
         Top             =   1600
         Width           =   2415
      End
      Begin VB.ListBox LstRopaje 
         Height          =   4350
         ItemData        =   "FrmObjSelector.frx":157C5E
         Left            =   -74880
         List            =   "FrmObjSelector.frx":157C60
         TabIndex        =   69
         Top             =   1960
         Width           =   2055
      End
      Begin VB.PictureBox PbGrhIndex 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   3375
         Left            =   1800
         ScaleHeight     =   221
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   285
         TabIndex        =   68
         Top             =   3400
         Width           =   4335
      End
      Begin VB.ListBox LstGrhIndex 
         Height          =   2790
         ItemData        =   "FrmObjSelector.frx":157C62
         Left            =   120
         List            =   "FrmObjSelector.frx":157C64
         TabIndex        =   64
         Top             =   3760
         Width           =   1575
      End
      Begin VB.CheckBox ChDuracion 
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2460
         Width           =   255
      End
      Begin VB.TextBox TxtDuracion 
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
         Left            =   1440
         TabIndex        =   59
         Top             =   2460
         Width           =   975
      End
      Begin VB.CheckBox ChHindex 
         Height          =   255
         Left            =   2640
         TabIndex        =   58
         Top             =   2460
         Width           =   255
      End
      Begin VB.TextBox TxtHIndex 
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
         Left            =   3840
         TabIndex        =   57
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CheckBox ChAnimacion 
         Height          =   255
         Left            =   -74880
         TabIndex        =   50
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox TxtAnimacion 
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
         Left            =   -73440
         TabIndex        =   49
         Top             =   540
         Width           =   1215
      End
      Begin VB.CheckBox ChDanMin 
         Height          =   255
         Left            =   -72120
         TabIndex        =   48
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox TxtDanMin 
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
         Left            =   -70800
         TabIndex        =   47
         Top             =   540
         Width           =   975
      End
      Begin VB.CheckBox ChDanMax 
         Height          =   255
         Left            =   -69720
         TabIndex        =   46
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox TxtDanMax 
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
         Left            =   -68400
         TabIndex        =   45
         Top             =   540
         Width           =   1095
      End
      Begin VB.CheckBox ChApu 
         Height          =   255
         Left            =   -69720
         TabIndex        =   44
         Top             =   1050
         Width           =   255
      End
      Begin VB.ComboBox CbApu 
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
         ItemData        =   "FrmObjSelector.frx":157C66
         Left            =   -68400
         List            =   "FrmObjSelector.frx":157C68
         TabIndex        =   43
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CheckBox ChMunicion 
         Height          =   255
         Left            =   -72360
         TabIndex        =   42
         Top             =   1020
         Width           =   255
      End
      Begin VB.ComboBox CbMunicion 
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
         ItemData        =   "FrmObjSelector.frx":157C6A
         Left            =   -70800
         List            =   "FrmObjSelector.frx":157C6C
         TabIndex        =   41
         Top             =   1020
         Width           =   975
      End
      Begin VB.ComboBox CbProyectil 
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
         ItemData        =   "FrmObjSelector.frx":157C6E
         Left            =   -73560
         List            =   "FrmObjSelector.frx":157C70
         TabIndex        =   40
         Top             =   1020
         Width           =   975
      End
      Begin VB.CheckBox ChProyectil 
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   1050
         Width           =   255
      End
      Begin VB.CheckBox ChRopaje 
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox TxtRopaje 
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
         Left            =   -73680
         TabIndex        =   34
         Top             =   540
         Width           =   1335
      End
      Begin VB.CheckBox ChDefMin 
         Height          =   255
         Left            =   -72120
         TabIndex        =   33
         Top             =   535
         Width           =   255
      End
      Begin VB.TextBox TxtDefMin 
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
         Left            =   -70920
         TabIndex        =   32
         Top             =   535
         Width           =   1095
      End
      Begin VB.CheckBox ChDefMax 
         Height          =   255
         Left            =   -69720
         TabIndex        =   31
         Top             =   535
         Width           =   255
      End
      Begin VB.TextBox TxtDefMax 
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
         Left            =   -68400
         TabIndex        =   30
         Top             =   500
         Width           =   1215
      End
      Begin VB.ComboBox CbRaza 
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
         ItemData        =   "FrmObjSelector.frx":157C72
         Left            =   6240
         List            =   "FrmObjSelector.frx":157C74
         TabIndex        =   28
         Top             =   1980
         Width           =   1455
      End
      Begin VB.CheckBox ChRaza 
         Height          =   255
         Left            =   5280
         TabIndex        =   27
         Top             =   2010
         Width           =   255
      End
      Begin VB.ComboBox CbCrucial 
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
         ItemData        =   "FrmObjSelector.frx":157C76
         Left            =   6360
         List            =   "FrmObjSelector.frx":157C78
         TabIndex        =   26
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CheckBox ChValor 
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   1500
         Width           =   255
      End
      Begin VB.TextBox TxtValor 
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
         TabIndex        =   20
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CheckBox ChCrucial 
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   1530
         Width           =   255
      End
      Begin VB.CheckBox ChAlineacion 
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   2010
         Width           =   255
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
         ItemData        =   "FrmObjSelector.frx":157C7A
         Left            =   4080
         List            =   "FrmObjSelector.frx":157C7C
         TabIndex        =   17
         Top             =   1980
         Width           =   975
      End
      Begin VB.CheckBox ChAgarrable 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2010
         Width           =   255
      End
      Begin VB.ComboBox CbAgarrable 
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
         ItemData        =   "FrmObjSelector.frx":157C7E
         Left            =   1440
         List            =   "FrmObjSelector.frx":157C80
         TabIndex        =   15
         Top             =   1980
         Width           =   975
      End
      Begin VB.CheckBox ChGrhIndex 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1500
         Width           =   255
      End
      Begin VB.TextBox TxtGrhIndex 
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
         Left            =   1440
         TabIndex        =   12
         Top             =   1500
         Width           =   975
      End
      Begin VB.CheckBox ChTipo 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1050
         Width           =   255
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
         ItemData        =   "FrmObjSelector.frx":157C82
         Left            =   960
         List            =   "FrmObjSelector.frx":157C84
         TabIndex        =   7
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CheckBox ChSubTipo 
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   1050
         Width           =   255
      End
      Begin VB.ComboBox CbPociones 
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
         ItemData        =   "FrmObjSelector.frx":157C86
         Left            =   6600
         List            =   "FrmObjSelector.frx":157C88
         TabIndex        =   5
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CheckBox ChPociones 
         Height          =   255
         Left            =   5280
         TabIndex        =   4
         Top             =   1050
         Width           =   255
      End
      Begin VB.CheckBox ChNombre 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   495
         Width           =   255
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
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   420
         Width           =   6375
      End
      Begin VB.Label Label16 
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
         Left            =   -74520
         TabIndex        =   158
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label12 
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
         Left            =   -68760
         TabIndex        =   154
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Seleccion de Casco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68760
         TabIndex        =   153
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label15 
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
         Left            =   -74880
         TabIndex        =   150
         Top             =   1785
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Seleccion de Cabeza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   149
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label LblHeadAnim 
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
         Left            =   -74280
         TabIndex        =   148
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblHeadDefMin 
         Caption         =   "Def. Min."
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
         Left            =   -71880
         TabIndex        =   147
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LblHeadDefMax 
         Caption         =   "Def. Max."
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
         Left            =   -69480
         TabIndex        =   146
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LblNavegacion 
         Caption         =   "Navegacion"
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
         Index           =   1
         Left            =   -74520
         TabIndex        =   133
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label LblCarpinteria 
         Caption         =   "Carpinteria"
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
         TabIndex        =   132
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LblHerreria 
         Caption         =   "Herreria"
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
         Left            =   -69360
         TabIndex        =   131
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LblLingOro 
         Caption         =   "Ling. Oro"
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
         TabIndex        =   130
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LblLingPlata 
         Caption         =   "Ling. Plata"
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
         TabIndex        =   129
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LblLingHierro 
         Caption         =   "Ling. Hierro"
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
         Left            =   -69360
         TabIndex        =   128
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LblMadera 
         Caption         =   "Madera"
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
         TabIndex        =   127
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LblClases 
         Caption         =   "Clases Prohibidas"
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
         TabIndex        =   112
         Top             =   600
         Width           =   1935
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
         Left            =   -74520
         TabIndex        =   109
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label LblEscMaxdef 
         Caption         =   "Def. Max."
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
         Left            =   -69480
         TabIndex        =   105
         Top             =   500
         Width           =   975
      End
      Begin VB.Label LblEscMinDef 
         Caption         =   "Def. Min."
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
         Left            =   -71880
         TabIndex        =   104
         Top             =   500
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Esc. Anim"
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
         TabIndex        =   103
         Top             =   500
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Seleccion de Animacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74640
         TabIndex        =   102
         Top             =   1085
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Left            =   -74880
         TabIndex        =   101
         Top             =   1685
         Width           =   735
      End
      Begin VB.Label Label7 
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
         Left            =   -74880
         TabIndex        =   87
         Top             =   2125
         Width           =   735
      End
      Begin VB.Label Label6 
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
         Left            =   -74880
         TabIndex        =   83
         Top             =   1720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Seleccion Animacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   76
         Top             =   1480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Seleccion de Ropaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74640
         TabIndex        =   70
         Top             =   1120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccion de Hechizos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   67
         Top             =   2920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccion de GrhIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   65
         Top             =   2920
         Width           =   1335
      End
      Begin VB.Label LblDuracion 
         Caption         =   "Dur. Efec."
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
         Left            =   360
         TabIndex        =   62
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label LblHIndex 
         Caption         =   "H. Index"
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
         TabIndex        =   61
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label LblAnimacion 
         Caption         =   "Animacion"
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
         TabIndex        =   56
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label LblDanMin 
         Caption         =   "Daño Min"
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
         Left            =   -71880
         TabIndex        =   55
         Top             =   540
         Width           =   975
      End
      Begin VB.Label LblDanMax 
         Caption         =   "Daño Max"
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
         Left            =   -69480
         TabIndex        =   54
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label LblApun 
         Caption         =   "Apuñala"
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
         Left            =   -69480
         TabIndex        =   53
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label LblMunicion 
         Caption         =   "Municiones"
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
         TabIndex        =   52
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label LblProyectil 
         Caption         =   "Proyectil"
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
         TabIndex        =   51
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label LblRopaje 
         Caption         =   "Ropaje"
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
         TabIndex        =   38
         Top             =   540
         Width           =   855
      End
      Begin VB.Label LblDefMin 
         Caption         =   "Def. Min."
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
         Left            =   -71880
         TabIndex        =   37
         Top             =   535
         Width           =   975
      End
      Begin VB.Label LblDefMax 
         Caption         =   "Def. Max."
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
         Left            =   -69480
         TabIndex        =   36
         Top             =   535
         Width           =   975
      End
      Begin VB.Label LblRaza 
         Caption         =   "Raza"
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
         Left            =   5520
         TabIndex        =   29
         Top             =   2010
         Width           =   735
      End
      Begin VB.Label LblValor 
         Caption         =   "Valor"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label LblCrucial 
         Caption         =   "Crucial"
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
         Left            =   5520
         TabIndex        =   24
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label LblAlineacion 
         Caption         =   "Alineacion"
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
         TabIndex        =   23
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label LblAgarrable 
         Caption         =   "Agarrable"
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
         Left            =   360
         TabIndex        =   22
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label LblGrhIndex 
         Caption         =   "GrhIndex"
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
         Left            =   360
         TabIndex        =   14
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label LblTipo 
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
         Left            =   360
         TabIndex        =   11
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label LblsubTipo 
         Caption         =   "SubTipo"
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
         TabIndex        =   10
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label LblPociones 
         Caption         =   "Pociones"
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
         Left            =   5520
         TabIndex        =   9
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label LblNombre 
         Caption         =   "Nombre"
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
         Left            =   360
         TabIndex        =   3
         Top             =   495
         Width           =   855
      End
   End
   Begin VB.Label CmdVolver 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3000
      TabIndex        =   164
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label CmdCrear 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   9960
      TabIndex        =   163
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label CmdGuardar 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   5760
      TabIndex        =   162
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Checks 
         Caption         =   "Setear Checks"
         Begin VB.Menu Weapons 
            Caption         =   "Armas"
         End
         Begin VB.Menu Armaduras 
            Caption         =   "Armaduras/Tunicas"
         End
         Begin VB.Menu CascSomb 
            Caption         =   "Cascos/Sombreros"
         End
         Begin VB.Menu Shiled 
            Caption         =   "Escudos"
         End
      End
      Begin VB.Menu Parametros 
         Caption         =   "Ingresar Mas Parametros"
         Begin VB.Menu Tipo 
            Caption         =   "Agregar un Tipo"
         End
         Begin VB.Menu SubTipo 
            Caption         =   "Agregar un Sub Tipo"
         End
         Begin VB.Menu Pocion 
            Caption         =   "Agregar un tipo de Pocion"
         End
         Begin VB.Menu Alineacion 
            Caption         =   "Agregar una Alineacion"
         End
      End
   End
   Begin VB.Menu Reloads 
      Caption         =   "Reloads"
      Begin VB.Menu ReloadIndex 
         Caption         =   "Recargar Indices"
         Begin VB.Menu ReloadHechizos 
            Caption         =   "Recargar Hechizos"
         End
         Begin VB.Menu ReloadEscudos 
            Caption         =   "Recargar Escudos"
         End
         Begin VB.Menu ReloadArmas 
            Caption         =   "Recargar Armas"
         End
         Begin VB.Menu ReloadBodys 
            Caption         =   "Recargar Ropajes"
         End
      End
      Begin VB.Menu ReloadObjSave 
         Caption         =   "Recargar Objetos Guardados"
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
         Begin VB.Menu MNpcs 
            Caption         =   "Npcs"
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
Attribute VB_Name = "FrmObjSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LB_FINDSTRING = &H18F
Private Declare Function sendmessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Dim SelectedItem As Integer
Dim CantItems As Integer
Dim CantHechizos As Integer
Dim AnimType As Byte
Dim FrameActual As Long
Dim FrameActualAnim(1 To 4) As Long
Dim Index As Integer
Dim CantFramesAnim(1 To 4) As Byte
Dim GrhIndex As Long
Dim CantFrames As Byte
Dim ObjNumber As Integer
Dim ReloadSave As Boolean

Private Sub Alineacion_Click()
Dim Insert As String
Dim Nuevo As Integer
Insert = InputBox("Ingrese la nueva 'Alineacion'", "Agregando 'Alineacion'")
Nuevo = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "AlineacionType")) + 1
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "AlineacionType", Nuevo)
Call WriteVar(App.path & "\IndexerDats.dao", "ALINEACIONTYPE", "Alineacion" & Nuevo, Insert)

Call ReloadCombo
End Sub

Private Sub Clase_Click()
Dim Insert As String
Dim Nuevo As Integer
Insert = InputBox("Ingrese la nueva 'Clase'", "Agregando 'Clase'")
Nuevo = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "Clases")) + 1
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "Clases", Nuevo)
Call WriteVar(App.path & "\IndexerDats.dao", "CLASES", "Clase" & Nuevo, Insert)

Call ReloadCombo
End Sub

Private Sub Armaduras_Click()
Call NormalCH

ChDefMin.value = 1
ChDefMax.value = 1
ChRopaje.value = 1
ChRaza.value = 1
End Sub

Private Sub CascSomb_Click()
Call NormalCH

ChHeadAnim.value = 1
ChHeadDefMin.value = 1
ChHeadDefMax.value = 1
ChSubTipo.value = 1
End Sub

Private Sub CmdAgregar_Click()
Dim NewParm As String
Dim CantParm As Integer
Dim TotalParm As Integer
NewParm = InputBox("Agregue el nuevo parametro, SIN valor inicia y SIN el '='.", "Agregar Parametros")

If NewParm = "" Then Exit Sub

CantParm = Val(GetVar(IndexDaterIni, "CANTS", "Parametros"))
TotalParm = CantParm + 1
Call WriteVar(IndexDaterIni, "CANTS", "Parametros", TotalParm)
Call WriteVar(IndexDaterIni, "PARAMETROS", "Parametro" & TotalParm, NewParm)

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
LstAdd.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "Parametros"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROS", "Parametro" & Cont)
    LstAdd.AddItem Dato & "=0"
Next Cont

MsgBox "Parametro Ingresado"
End Sub

Private Sub CmdCargar_Click()
Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String
LstAdd.Clear
Cant = Val(GetVar(IndexDaterIni, "CANTS", "Parametros"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROS", "Parametro" & Cont)
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

Private Sub CmdCrear_Click()
ObjNumber = CantItems + 1
If FileExist(Config.SaveDatPath & "\Obj.dat", vbNormal) Then
    Call SaveObj
Else
    Call FileCopy(Config.DatPath & "\Obj.dat", Config.SaveDatPath & "\Obj.dat")
    Call SaveObj
End If

MsgBox "El objeto se ha creado con exito."
End Sub

Private Sub CmdGuardar_Click()
If FileExist(Config.SaveDatPath & "\Obj.dat", vbNormal) Then
    Call SaveObj
Else
    Call FileCopy(Config.DatPath & "\Obj.dat", Config.SaveDatPath & "\Obj.dat")
    Call SaveObj
End If

MsgBox "El Objeto se ha modificado con exito"
End Sub

Private Sub CmdVolver_Click()
FrmDatMenu.Visible = True
Me.Visible = False
End Sub

Private Sub Form_Load()
OpName.value = True
ReloadSave = False
Call ReloadAllDats

End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmDatMenu.Visible = True
End Sub

Private Sub LstAnimaciones_Click()
Dim ArmaIndex As Integer
Dim T As Byte
If LstAnimaciones.Text <> "" Then
    ArmaIndex = Val(LstAnimaciones.Text)
    For T = 1 To 4
        CantFramesAnim(T) = GrhData(Armas(ArmaIndex).Arma(T)).NumFrames
        FrameActualAnim(T) = 0
        PbAnimacion(T).Cls
    Next T

    THechizo.Enabled = False
    TAllDirections.Enabled = True
    Index = ArmaIndex
    AnimType = 1
    TxtAnimacion.Text = LstAnimaciones.Text
End If
End Sub

Private Sub LstCascos_Click()
Dim GraficosPath(1 To 4) As String
Dim CascoIndex As Integer
Dim IndexCasco As Long
Dim LongXCasco As Integer
Dim LongYCasco As Integer
Dim PosXCasco As Integer
Dim PosYCasco As Integer
CascoIndex = Val(LstCascos.Text)

If CascoIndex = 0 Then Exit Sub

TxtHeadAnim.Text = Str(CascoIndex)

Dim T As Byte
For T = 1 To 4
    IndexCasco = Cascos(CascoIndex).Casco(T)
    LongXCasco = GrhData(IndexCasco).pixelWidth
    LongYCasco = GrhData(IndexCasco).pixelHeight
    PosXCasco = GrhData(IndexCasco).sX
    PosYCasco = GrhData(IndexCasco).sY
    
    GraficosPath(T) = Config.BmpPath & "\" & GrhData(IndexCasco).FileNum & ".bmp"
    If FileExist(GraficosPath(T), vbNormal) = True Then
        PbHeads(T).Visible = True
        PbHeads(T).Cls
        PbHeads(T).PaintPicture LoadPicture(GraficosPath(T)), 0, 0, LongXCasco, LongYCasco, PosXCasco, PosYCasco, LongXCasco, LongYCasco
    End If
Next T

End Sub

Private Sub LstEscudos_Click()
Dim EscudoIndex As Integer
Dim T As Byte
If LstEscudos.Text <> "" Then
    EscudoIndex = Val(LstEscudos.Text)
    For T = 1 To 4
        CantFramesAnim(T) = GrhData(Escudos(EscudoIndex).Escudo(T)).NumFrames
        FrameActualAnim(T) = 0
        PbAnimacion(T).Cls
    Next T

    THechizo.Enabled = False
    TAllDirections.Enabled = True
    Index = EscudoIndex
    AnimType = 2
    TxtEscAnim.Text = LstEscudos.Text
End If

End Sub

Private Sub LstGrhIndex_Click()
Dim GraficoPath As String
Dim PosicionX As Integer
Dim PosicionY As Integer
Dim LongitudX As Integer
Dim LongitudY As Integer
Dim GrhIndex As Long
Dim Frames As Byte

TxtGrhIndex.Text = LstGrhIndex.Text
GrhIndex = Val(LstGrhIndex.Text)

If GrhIndex > 0 Then
    THechizo.Enabled = False
    If GrhData(GrhIndex).FileNum > 0 Then
        GraficoPath = Config.BmpPath & "\" & GrhData(GrhIndex).FileNum & ".bmp"
        PosicionX = GrhData(GrhIndex).sX
        PosicionY = GrhData(GrhIndex).sY
        LongitudX = GrhData(GrhIndex).pixelWidth
        LongitudY = GrhData(GrhIndex).pixelHeight
        If FileExist(GraficoPath, vbNormal) = True Then
            PbGrhIndex.Cls
            PbGrhIndex.PaintPicture LoadPicture(GraficoPath), 0, 0, LongitudX, LongitudY, PosicionX, PosicionY, LongitudX, LongitudY
        End If
    End If
End If
End Sub

Private Sub LstHeads_Click()
Dim GraficosPath(1 To 4) As String
Dim HeadIndex As Integer
Dim IndexHead As Long
Dim LongXHead As Integer
Dim LongYHead As Integer
Dim PosXHead As Integer
Dim PosYHead As Integer
HeadIndex = Val(LstHeads.Text)

If HeadIndex = 0 Then Exit Sub

TxtHeadAnim.Text = HeadIndex

Dim T As Byte
For T = 1 To 4
    IndexHead = Heads(HeadIndex).Head(T)
    LongXHead = GrhData(IndexHead).pixelWidth
    LongYHead = GrhData(IndexHead).pixelHeight
    PosXHead = GrhData(IndexHead).sX
    PosYHead = GrhData(IndexHead).sY

    GraficosPath(T) = Config.BmpPath & "\" & GrhData(IndexHead).FileNum & ".bmp"
    If FileExist(GraficosPath(T), vbNormal) = True Then
        PbHeads(T).Visible = True
        PbHeads(T).Cls
        PbHeads(T).PaintPicture LoadPicture(GraficosPath(T)), 0, 0, LongXHead, LongYHead, PosXHead, PosYHead, LongXHead, LongYHead
    End If
Next T

End Sub

Private Sub LstHIndex_Click()
Dim HechizoSelect As Integer
Dim HechizoFx As Integer
If Val(LstHIndex.Text) > 0 Then
    TxtHIndex.Text = Val(Left(LstHIndex.Text, InStr(1, LstHIndex.Text, "-") - 1))
    HechizoSelect = Val(TxtHIndex.Text)
    
    HechizoFx = Val(GetVar(Config.DatPath & "\Hechizos.dat", "HECHIZO" & HechizoSelect, "FXgrh"))
    
    If HechizoFx = 0 Then
        PbGrhIndex.Cls
        Exit Sub
    End If
    
    GrhIndex = Fx(HechizoFx).Animacion
    CantFrames = GrhData(GrhIndex).NumFrames
    
    If GrhData(GrhIndex).Speed <> 0 And GrhData(GrhIndex).NumFrames <> 0 Then
    If IndexMode = "12.1" Then
        THechizo.Interval = Round(GrhData(GrhIndex).Speed / GrhData(GrhIndex).NumFrames)
    Else
        THechizo.Interval = 100
    End If
    THechizo.Enabled = True
    Else
    THechizo.Enabled = False
    End If
    FrameActual = 0
    TAllDirections.Enabled = False
    PbGrhIndex.Cls
    
End If
End Sub

Private Sub LstRopaje_Click()
Dim BodyIndex As Integer
Dim T As Byte
TxtRopaje.Text = LstRopaje.Text
If LstRopaje.Text <> "" Then
    BodyIndex = Val(LstRopaje.Text)
    For T = 1 To 4
        CantFramesAnim(T) = GrhData(Bodys(BodyIndex).Body(T)).NumFrames
        FrameActualAnim(T) = 0
        PbRopaje(T).Cls
    Next T

    THechizo.Enabled = False
    TAllDirections.Enabled = True
    Index = BodyIndex
    AnimType = 0
End If
End Sub

Private Sub MConversor_Click()
FrmConversor.Visible = True
End Sub

Private Sub MRutas_Click()
frmConfig.Visible = True
End Sub

Private Sub ObjList_DblClick()
On Error GoTo ErrHandler
Dim T As Integer
Dim Imag As Integer
Dim GraficoPath As String
Dim PosicionX As Integer
Dim PosicionY As Integer
Dim LongitudX As Integer
Dim LongitudY As Integer
Dim GrhIndex As Long
Dim Frames As Byte
Dim ObjRut As String

Call SetAllBlank
Call NormalCH

If ReloadSave = True Then
    ObjRut = Config.SaveDatPath
Else
    ObjRut = Config.DatPath
End If

ObjNumber = Left(ObjList.Text, InStr(1, ObjList.Text, "-") - 1)

TxtNombre.Text = GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Name")
CbTipo.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Objtype")) - 1
CbSubTipo.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Subtipo")) - 1
CbPociones.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "TipoPocion")) - 1
If CbPociones.Text <> "" Then ChPociones.value = 1
TxtGrhIndex.Text = GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "GrhIndex")
Imag = Val(TxtGrhIndex.Text)
If GrhData(Imag).FileNum > 0 Then
    GraficoPath = Config.BmpPath & "\" & GrhData(Imag).FileNum & ".bmp"
    PosicionX = GrhData(Imag).sX
    PosicionY = GrhData(Imag).sY
    LongitudX = GrhData(Imag).pixelWidth
    LongitudY = GrhData(Imag).pixelHeight
    If FileExist(GraficoPath, vbNormal) = True Then
        PbGrhIndex.Cls
        PbGrhIndex.PaintPicture LoadPicture(GraficoPath), 0, 0, LongitudX, LongitudY, PosicionX, PosicionY, LongitudX, LongitudY
    End If
End If


TxtValor.Text = GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Valor")
CbCrucial.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Crucial"))
CbAgarrable.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Agarrable"))
If CbAgarrable.ListIndex = 1 Then ChAgarrable.value = 1
Dim Alin As String
Alin = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Real"))
If Alin = 1 Then
    CbAlineacion.ListIndex = 0
    ChAlineacion.value = 1
Else
    Alin = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Caos"))
    If Alin = 1 Then
        CbAlineacion.ListIndex = 1
        ChAlineacion.value = 1
    Else
        CbAlineacion.ListIndex = CbAlineacion.ListCount - 1
    End If
End If

Dim Raza As Boolean

Raza = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "RazaEnana"))
If Raza = True Then
    CbRaza.ListIndex = 1
    ChRaza.value = 1
Else
    CbRaza.ListIndex = 2
End If

TxtDuracion.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "DuracionEfecto"))
If Val(TxtDuracion.Text) > 0 Then ChDuracion.value = 1
TxtHIndex.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "HechizoIndex"))
If Val(TxtHIndex.Text) > 0 Then ChHindex.value = 1
TxtRopaje.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "NumRopaje"))
If Val(TxtRopaje.Text) > 0 Then ChRopaje.value = 1

If CbTipo.ListIndex = 1 Then
    TxtAnimacion.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Anim"))
    TxtDanMin.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MinHIT"))
    TxtDanMax.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MaxHIT"))
    ChAnimacion.value = 1
    ChDanMin.value = 1
    ChDanMax.value = 1
End If
If CbSubTipo.ListIndex = 1 Then
    TxtEscMinDef.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MINDEF"))
    TxtEscMaxDef.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MAXDEF"))
    TxtEscAnim.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Anim"))
    ChEscAnim.value = 1
    ChEscMaxDef.value = 1
    ChEscMinDef.value = 1
End If
If CbSubTipo.ListIndex = 0 Then
    TxtHeadDefMin.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MINDEF"))
    TxtHeadDefMax.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MAXDEF"))
    TxtHeadAnim.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Anim"))
    ChHeadAnim.value = 1
    ChHeadDefMax.value = 1
    ChHeadDefMin.value = 1
End If

If CbSubTipo.Text = "" Then
    TxtDefMin.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MINDEF"))
    TxtDefMax.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MAXDEF"))
    If Val(TxtDefMin.Text) > 0 Then ChDefMin.value = 1
    If Val(TxtDefMax.Text) > 0 Then ChDefMax.value = 1
End If
CbProyectil.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Proyectil"))
If CbProyectil.ListIndex = 1 Then ChProyectil.value = 1
CbMunicion.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Municiones"))
If CbMunicion.ListIndex = 1 Then ChMunicion.value = 1
CbApu.ListIndex = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Apuñala"))
If CbApu.ListIndex = 1 Then ChApu.value = 1
TxtNavegacion.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "MinSkill"))
If Val(TxtNavegacion.Text) > 0 Then ChNavegacion.value = 1
TxtHerreria.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "SkHerreria"))
If Val(TxtHerreria.Text) > 0 Then ChHerreria.value = 1
TxtCarpinteria.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "SkCarpinteria"))
If Val(TxtCarpinteria.Text) > 0 Then ChCarpinteria.value = 1
TxtLingOro.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "LingO"))
If Val(TxtLingOro.Text) > 0 Then ChLingOro.value = 1
TxtLingPlata.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "LingP"))
If Val(TxtLingPlata.Text) > 0 Then ChLingPlata.value = 1
TxtLingHierro.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "LingH"))
If Val(TxtLingHierro.Text) > 0 Then ChLingHierro.value = 1
TxtMadera.Text = Val(GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "Madera"))
If Val(TxtMadera.Text) > 0 Then ChMadera.value = 1

LstAdd.Clear
For T = 0 To LstAdd.ListCount - 1
    Dim ValorParm As String
    Dim Parametro As String
    Parametro = GetVar(IndexDaterIni, "PARAMETROS", "Parametro" & T)
    ValorParm = GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, Parametro)
    If ValorParm = "" Then ValorParm = 0
    LstAdd.AddItem Parametro & "=" & ValorParm
Next T
If LstAdd.ListCount <> 0 Then ChAdd.value = 1

Dim Clases As Byte
Dim Clase As String
For Clases = 1 To 16
    Clase = GetVar(ObjRut & "\Obj.dat", "OBJ" & ObjNumber, "CP" & Clases)
    If Clase <> "" Then
        For T = 0 To LstClases.ListCount
            If UCase$(LstClases.List(T)) = UCase$(Clase) Then LstClases.Selected(T) = True
        Next T
    End If
Next Clases
Exit Sub
ErrHandler:
    MsgBox "Se Produjo el siguiente error: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Contactarse via MAIL a: soporte@aodestruction.com.ar" & vbCrLf & "Disculpen las molestias."
    Exit Sub
End Sub

Private Sub OpName_Click()
OpNum.value = False
End Sub

Private Sub OpNum_Click()
OpName.value = False
End Sub

Private Sub Pocion_Click()
Dim Insert As String
Dim Nuevo As Integer
Insert = InputBox("Ingrese el nuevo tipo de 'Pocion' para los objetos", "Agregando tipo de 'Pocion' para los Objetos")
Nuevo = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "PotionType")) + 1
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "PotionType", Nuevo)
Call WriteVar(App.path & "\IndexerDats.dao", "POTIONTYPE", "PotionType" & Nuevo, Insert)

Call ReloadCombo
End Sub

Private Sub Raza_Click()
Dim Insert As String
Dim Nuevo As Integer
Insert = InputBox("Ingrese la nueva 'Raza'", "Agregando 'Raza'")
Nuevo = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "RazaType")) + 1
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "RazaType", Nuevo)
Call WriteVar(App.path & "\IndexerDats.dao", "RAZATYPE", "Raza" & Nuevo, Insert)

Call ReloadCombo
End Sub

Private Sub ReloadAll_Click()

Call ReloadAllDats

End Sub

Private Sub ReloadArmas_Click()
LstAnimaciones.Clear

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

Cant = ArmasCountNew
For Cont = 1 To Cant
    If Armas(Cont).Arma(1) > 0 Then
        LstAnimaciones.AddItem Cont
    End If
Next Cont

End Sub

Private Sub ReloadBodys_Click()
LstRopaje.Clear

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

Cant = BodysCountNew
For Cont = 1 To Cant
    If Bodys(Cont).Body(1) > 0 Then
        LstRopaje.AddItem Cont
    End If
Next Cont

End Sub

Private Sub ReloadEscudos_Click()
LstEscudos.Clear

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

Cant = EscudosCountNew
For Cont = 1 To EscudosCountNew
    If Escudos(Cont).Escudo(1) > 0 Then
        LstEscudos.AddItem Cont
    End If
Next Cont

End Sub

Private Sub ReloadHechizos_Click()
LstHIndex.Clear

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

Dim Hechizo As Integer
Dim HechizoName As String
CantHechizos = GetVar(Config.DatPath & "\Hechizos.dat", "INIT", "NumeroHechizos")
For Hechizo = 1 To CantHechizos
    HechizoName = GetVar(Config.DatPath & "\Hechizos.dat", "HECHIZO" & Hechizo, "Nombre")
    If HechizoName <> "" Then
        LstHIndex.AddItem Hechizo & "-" & HechizoName
    End If
Next Hechizo

End Sub

Private Sub ReloadObj_Click()
ObjList.Clear

Dim Obj As Integer
Dim ObjName As String
CantItems = GetVar(Config.DatPath & "\Obj.dat", "INIT", "NumOBJs")
For Obj = 1 To CantItems
    ObjName = GetVar(Config.DatPath & "\Obj.dat", "OBJ" & Obj, "Name")
    If ObjName <> "" Then
        ObjList.AddItem Obj & "-" & ObjName
    End If
Next Obj

End Sub

Private Sub ReloadObjSave_Click()
ObjList.Clear

Dim Obj As Integer
Dim ObjName As String
CantItems = GetVar(Config.SaveDatPath & "\Obj.dat", "INIT", "NumOBJs")
For Obj = 1 To CantItems
    ObjName = GetVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & Obj, "Name")
    If ObjName <> "" Then
        ObjList.AddItem Obj & "-" & ObjName
    End If
Next Obj
ReloadSave = True

End Sub

Private Sub Shiled_Click()
Call NormalCH

ChSubTipo.value = 1
ChEscAnim.value = 1
ChEscMinDef.value = 1
ChEscMaxDef.value = 1
End Sub

Private Sub SubTipo_Click()
Dim Insert As String
Dim Nuevo As Integer
Insert = InputBox("Ingrese el nuevo 'Sub Tipo' de objeto", "Agregando 'Sub Tipo' de Objeto")
Nuevo = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "SubType")) + 1
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "SubType", Nuevo)
Call WriteVar(App.path & "\IndexerDats.dao", "SUBTYPE", "SubType" & Nuevo, Insert)

Call ReloadCombo
End Sub

Private Sub TAllDirections_Timer()
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
    
    If AnimType = 0 Then 'Chequeo Tipo de Animacion
        Anim(T) = Bodys(Index).Body(T)
    ElseIf AnimType = 1 Then
        Anim(T) = Armas(Index).Arma(T)
    Else
        Anim(T) = Escudos(Index).Escudo(T)
    End If
    
    GrhIndexAnim(T) = GrhData(Anim(T)).Frames(FrameActualAnim(T))
    
    If GrhIndexAnim(T) = 0 Then 'Por si hay error y no existe.
        TAllDirections.Enabled = False
        Exit Sub
    End If
    
    GraficoPath(T) = Config.BmpPath & "\" & GrhData(GrhIndexAnim(T)).FileNum & ".bmp" 'Busco la Imagen
    
    If Not FileExist(GraficoPath(T), vbNormal) = True Then 'Me fijo si Existe
        TAllDirections.Enabled = False
        Exit Sub
    End If
    
    AnimacionPosX(T) = GrhData(GrhIndexAnim(T)).sX 'Coordenada X
    AnimacionPosY(T) = GrhData(GrhIndexAnim(T)).sY 'Coordenada Y
    AnimacionLongX(T) = GrhData(GrhIndexAnim(T)).pixelWidth 'Longitud sobre X
    AnimacionLongY(T) = GrhData(GrhIndexAnim(T)).pixelHeight 'Longitud sobre Y
    
    If AnimType = 0 Then
        PbRopaje(T).PaintPicture LoadPicture(GraficoPath(T)), 0, 0, AnimacionLongX(T), AnimacionLongY(T), AnimacionPosX(T), AnimacionPosY(T), AnimacionLongX(T), AnimacionLongY(T)
    ElseIf AnimType = 1 Then
        PbAnimacion(T).PaintPicture LoadPicture(GraficoPath(T)), 0, 0, AnimacionLongX(T), AnimacionLongY(T), AnimacionPosX(T), AnimacionPosY(T), AnimacionLongX(T), AnimacionLongY(T)
    Else
        PbEscudo(T).PaintPicture LoadPicture(GraficoPath(T)), 0, 0, AnimacionLongX(T), AnimacionLongY(T), AnimacionPosX(T), AnimacionPosY(T), AnimacionLongX(T), AnimacionLongY(T)
    End If
    
    
    If FrameActualAnim(T) = CantFramesAnim(T) Then
        FrameActualAnim(T) = 0
    End If
Next T

Exit Sub

ErrorAnim:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    TAllDirections.Enabled = False
    Exit Sub
'For T = 1 To 4
'    FrameActualAnim(T) = 0
'Next T

End Sub

Private Sub THechizo_Timer()
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
        PbGrhIndex.PaintPicture LoadPicture(GraficoPath), 0, 0, AnimacionLongX, AnimacionLongY, AnimacionPosX, AnimacionPosY, AnimacionLongX, AnimacionLongY
    End If
End If

If FrameActual = CantFrames Then
    FrameActual = 0
End If
Exit Sub
ErrHandler:
    MsgBox "Se Produjo un error, se detendra la reproduccion de la animacion." & vbCrLf & Err.Description
    THechizo.Enabled = False
    Exit Sub
End Sub

Private Sub Tipo_Click()
Dim Insert As String
Dim Nuevo As Integer
Insert = InputBox("Ingrese el nuevo 'Tipo' de objeto", "Agregando 'Tipo' de Objeto")
Nuevo = Val(GetVar(App.path & "\IndexerDats.dao", "CANTS", "ObjType")) + 1
Call WriteVar(App.path & "\IndexerDats.dao", "CANTS", "ObjType", Nuevo)
Call WriteVar(App.path & "\IndexerDats.dao", "OBJTYPES", "Type" & Nuevo, Insert)

Call ReloadCombo
End Sub

Private Sub TxtCascoSearch_Change()
LstCascos.ListIndex = sendmessage(LstCascos.hWnd, LB_FINDSTRING, -1, ByVal TxtCascoSearch.Text)
End Sub

Private Sub TxtEscSearch_Change()
LstEscudos.ListIndex = sendmessage(LstEscudos.hWnd, LB_FINDSTRING, -1, ByVal TxtEscSearch.Text)
End Sub

Private Sub TxtHeadSearch_Change()
LstHeads.ListIndex = sendmessage(LstHeads.hWnd, LB_FINDSTRING, -1, ByVal TxtHeadSearch.Text)
End Sub


Private Sub TxtSAnimacion_Change()
LstAnimaciones.ListIndex = sendmessage(LstAnimaciones.hWnd, LB_FINDSTRING, -1, ByVal TxtSAnimacion.Text)
End Sub

Private Sub TxtSGrhIndex_Change()
LstGrhIndex.ListIndex = sendmessage(LstGrhIndex.hWnd, LB_FINDSTRING, -1, ByVal TxtSGrhIndex.Text)
End Sub

Private Sub TxtSHIndex_Change()
LstHIndex.ListIndex = sendmessage(LstHIndex.hWnd, LB_FINDSTRING, -1, ByVal TxtSHIndex.Text)
End Sub

Private Sub TxtSRopaje_Change()
    LstRopaje.ListIndex = sendmessage(LstRopaje.hWnd, LB_FINDSTRING, -1, ByVal TxtSRopaje.Text)
End Sub

Private Sub TxtSearch_Change()
If OpNum.value = True Then
    ObjList.ListIndex = sendmessage(ObjList.hWnd, LB_FINDSTRING, -1, ByVal TxtSearch.Text)
Else
    ObjListCopy.ListIndex = sendmessage(ObjListCopy.hWnd, LB_FINDSTRING, -1, ByVal TxtSearch.Text)
    ObjList.ListIndex = ObjListCopy.ListIndex
End If
End Sub

Private Sub SetAllBlank()
    TxtNombre.Text = ""
    TxtGrhIndex.Text = ""
    TxtValor.Text = ""
    TxtDuracion.Text = ""
    TxtHIndex.Text = ""
    TxtRopaje.Text = ""
    TxtDefMin.Text = ""
    TxtDefMax.Text = ""
    TxtDanMin.Text = ""
    TxtDanMax.Text = ""
    TxtAnimacion.Text = ""
    TxtEscAnim.Text = ""
    TxtEscMinDef.Text = ""
    TxtEscMaxDef.Text = ""
    TxtNavegacion.Text = ""
    TxtCarpinteria.Text = ""
    TxtHerreria.Text = ""
    TxtLingOro.Text = ""
    TxtLingPlata.Text = ""
    TxtLingHierro.Text = ""
    TxtMadera.Text = ""
    'LstAdd.Clear
    TxtHeadDefMin.Text = ""
    TxtHeadDefMax.Text = ""
    TxtHeadAnim.Text = ""
    
    CbTipo.ListIndex = CbTipo.ListCount - 1
    CbSubTipo.ListIndex = CbSubTipo.ListCount - 1
    CbPociones.ListIndex = CbPociones.ListCount - 1
    CbCrucial.ListIndex = CbCrucial.ListCount - 1
    CbAgarrable.ListIndex = CbAgarrable.ListCount - 1
    CbRaza.ListIndex = CbRaza.ListCount - 1
    CbAlineacion.ListIndex = CbAlineacion.ListCount - 1
    CbTipo.ListIndex = CbTipo.ListCount - 1
    CbApu.ListIndex = CbApu.ListCount - 1
    CbProyectil.ListIndex = CbProyectil.ListCount - 1
    CbMunicion.ListIndex = CbMunicion.ListCount - 1
    
    Dim T As Byte
    For T = 0 To LstClases.ListCount - 1
        LstClases.Selected(T) = False
    Next T
End Sub

Private Sub ReloadCombo()
Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

CbTipo.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "ObjType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "OBJTYPES", "Type" & Cont)
    CbTipo.AddItem Dato
Next Cont
CbTipo.AddItem "(Elegir)", Cant

CbSubTipo.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "SubType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "SUBTYPE", "SubType" & Cont)
    CbSubTipo.AddItem Dato
Next Cont
CbSubTipo.AddItem "(Elegir)", Cant

CbPociones.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "PotionType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "POTIONTYPE", "PotionType" & Cont)
    CbPociones.AddItem Dato
Next Cont
CbPociones.AddItem "(Elegir)", Cant

CbAlineacion.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "AlineacionType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "ALINEACIONTYPE", "Alineacion" & Cont)
    CbAlineacion.AddItem Dato
Next Cont
CbAlineacion.AddItem "(Elegir)", Cant

CbRaza.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "RazaType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "RAZATYPE", "Raza" & Cont)
    CbRaza.AddItem Dato
Next Cont
CbRaza.AddItem "(Elegir)", Cant

LstClases.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "Clases"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "CLASES", "Clase" & Cont)
    LstClases.AddItem Dato
Next Cont

End Sub

Private Sub ReloadAllDats()

ObjList.Clear
ObjListCopy.Clear

Dim Obj As Integer
Dim ObjName As String
CantItems = GetVar(Config.DatPath & "\Obj.dat", "INIT", "NumOBJs")
For Obj = 1 To CantItems
    ObjName = GetVar(Config.DatPath & "\Obj.dat", "OBJ" & Obj, "Name")
    If ObjName <> "" Then
        ObjList.AddItem Obj & "-" & ObjName
        ObjListCopy.AddItem ObjName
    End If
Next Obj

LstHIndex.Clear

Dim Hechizo As Integer
Dim HechizoName As String
CantHechizos = GetVar(Config.DatPath & "\Hechizos.dat", "INIT", "NumeroHechizos")
For Hechizo = 1 To CantHechizos
    HechizoName = GetVar(Config.DatPath & "\Hechizos.dat", "HECHIZO" & Hechizo, "Nombre")
    If HechizoName <> "" Then
        LstHIndex.AddItem Hechizo & "-" & HechizoName
    End If
Next Hechizo

Dim Cant As Integer
Dim Cont As Integer
Dim Dato As String

CbTipo.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "ObjType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "OBJTYPES", "Type" & Cont)
    CbTipo.AddItem Dato
Next Cont
CbTipo.AddItem "(Elegir)", Cant

LstRopaje.Clear

Cant = BodysCountNew
For Cont = 1 To Cant
    If Bodys(Cont).Body(1) > 0 Then
        LstRopaje.AddItem Cont
    End If
Next Cont

CbSubTipo.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "SubType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "SUBTYPE", "SubType" & Cont)
    CbSubTipo.AddItem Dato
Next Cont
CbSubTipo.AddItem "(Elegir)", Cant

CbPociones.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "PotionType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "POTIONTYPE", "PotionType" & Cont)
    CbPociones.AddItem Dato
Next Cont
CbPociones.AddItem "(Elegir)", Cant

CbAlineacion.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "AlineacionType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "ALINEACIONTYPE", "Alineacion" & Cont)
    CbAlineacion.AddItem Dato
Next Cont
CbAlineacion.AddItem "(Elegir)", Cant

CbRaza.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "RazaType"))
For Cont = 1 To Cant
    Dato = Cont & "-" & GetVar(IndexDaterIni, "RAZATYPE", "Raza" & Cont)
    CbRaza.AddItem Dato
Next Cont
CbRaza.AddItem "(Elegir)", Cant

LstGrhIndex.Clear

Cant = AllGrhData
For Cont = 1 To AllGrhData
    If GrhData(Cont).NumFrames > 0 Then
        If GrhData(Cont).NumFrames < 2 Then
            LstGrhIndex.AddItem Cont
        End If
    End If
Next Cont

LstAnimaciones.Clear

Cant = ArmasCountNew
For Cont = 1 To Cant
    If Armas(Cont).Arma(1) > 0 Then
        LstAnimaciones.AddItem Cont
    End If
Next Cont

LstEscudos.Clear

Cant = EscudosCountNew
For Cont = 1 To EscudosCountNew
    If Escudos(Cont).Escudo(1) > 0 Then
        LstEscudos.AddItem Cont
    End If
Next Cont

LstClases.Clear

Cant = Val(GetVar(IndexDaterIni, "CANTS", "Clases"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "CLASES", "Clase" & Cont)
    LstClases.AddItem Dato
Next Cont

Cant = Val(GetVar(IndexDaterIni, "CANTS", "Parametros"))
For Cont = 1 To Cant
    Dato = GetVar(IndexDaterIni, "PARAMETROS", "Parametro" & Cont)
    LstAdd.AddItem Dato & "=0"
Next Cont

LstHeads.Clear

Cant = HeadsCountNew
For Cont = 1 To Cant
    If Heads(Cont).Head(1) > 0 Then
        LstHeads.AddItem Cont
    End If
Next Cont

LstCascos.Clear

Cant = CascosCountNew
For Cont = 1 To Cant
    If Cascos(Cont).Casco(1) > 0 Then
        LstCascos.AddItem Cont
    End If
Next Cont

CbCrucial.AddItem "No", 0
CbCrucial.AddItem "Si", 1
CbCrucial.AddItem "(Elegir)", 2

CbAgarrable.AddItem "Si", 0
CbAgarrable.AddItem "No", 1
CbAgarrable.AddItem "(Elegir)", 2

CbProyectil.AddItem "No", 0
CbProyectil.AddItem "Si", 1
CbProyectil.AddItem "(Elegir)", 2

CbMunicion.AddItem "No", 0
CbMunicion.AddItem "Si", 1
CbMunicion.AddItem "(Elegir)", 2

CbApu.AddItem "No", 0
CbApu.AddItem "Si", 1
CbApu.AddItem "(Elegir)", 2

TxtGrhIndex.Enabled = False
TxtHIndex.Enabled = False
TxtRopaje.Enabled = False
TxtAnimacion.Enabled = False
TxtEscAnim.Enabled = False
TxtHeadAnim.Enabled = False

Call SetAllBlank

End Sub

Private Sub Weapons_Click()
Call NormalCH

ChSubTipo.value = 1
ChAnimacion.value = 1
ChDanMin.value = 1
ChDanMax.value = 1
End Sub

Private Sub NormalCH()
ChNombre.value = 0
ChTipo.value = 0
ChSubTipo.value = 0
ChGrhIndex.value = 0
ChPociones.value = 0
ChValor.value = 0
ChHindex.value = 0
ChDuracion.value = 0
ChCrucial.value = 0
ChRaza.value = 0
ChAlineacion.value = 0
ChAgarrable.value = 0
ChRopaje.value = 0
ChDefMin.value = 0
ChDefMax.value = 0
ChAnimacion.value = 0
ChDanMax.value = 0
ChDanMin.value = 0
ChApu.value = 0
ChProyectil.value = 0
ChMunicion.value = 0
ChEscAnim.value = 0
ChEscMinDef.value = 0
ChEscMaxDef.value = 0
ChHeadAnim.value = 0
ChHeadDefMin.value = 0
ChHeadDefMax.value = 0
ChNavegacion.value = 0
ChHerreria.value = 0
ChCarpinteria.value = 0
ChLingOro.value = 0
ChLingPlata.value = 0
ChLingHierro.value = 0
ChMadera.value = 0
ChClases.value = 0
ChAdd.value = 0
ChAgarrable.value = 0

ChNombre.value = 1
ChGrhIndex.value = 1
ChTipo.value = 1
ChValor.value = 1
ChCrucial.value = 1
ChClases.value = 1
End Sub

Private Sub SaveObj()
If ObjNumber > CantItems Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "INIT", "NumOBJs", ObjNumber)
    CantItems = CantItems + 1
End If

If ChNombre.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Name", TxtNombre.Text)
End If
If ChTipo.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Objtype", CbTipo.ListIndex + 1)
End If
If ChSubTipo.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Subtipo", CbSubTipo.ListIndex + 1)
End If
If ChPociones.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "TipoPocion", CbPociones.ListIndex + 1)
End If
If ChGrhIndex.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "GrhIndex", TxtGrhIndex.Text)
End If
If ChValor.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Valor", TxtValor.Text)
End If
If ChCrucial.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Crucial", CbCrucial.ListIndex)
End If
If ChAgarrable.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Agarrable", CbAgarrable.ListIndex)
End If
If ChAlineacion.value = 1 Then
    If CbAlineacion.ListIndex = 1 Then
        Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Caos", 1)
    ElseIf CbAlineacion.ListIndex = 0 Then
        Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Real", 1)
    End If
End If
If ChRaza.value = 1 Then
    If CbAlineacion.ListIndex = 1 Then
        Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "RazaEnana", 1)
    End If
End If
If ChDuracion.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "DuracionEfecto", TxtDuracion.Text)
End If
If ChHindex.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "HechizoIndex", TxtHIndex.Text)
End If
If ChRopaje.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "NumRopaje", TxtRopaje.Text)
End If
If ChDefMin.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MINDEF", TxtDefMin.Text)
End If
If ChDefMax.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MAXDEF", TxtDefMax.Text)
End If
If ChAnimacion.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Anim", TxtAnimacion.Text)
End If
If ChDanMin.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MinHIT", TxtDanMin.Text)
End If
If ChDanMax.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MaxHIT", TxtDanMax.Text)
End If
If ChProyectil.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Proyectil", CbProyectil.ListIndex)
End If
If ChMunicion.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Municiones", CbMunicion.ListIndex)
End If
If ChApu.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Apuñala", CbApu.ListIndex)
End If
If ChEscAnim.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Anim", TxtEscAnim.Text)
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "NumRopaje", "2")
End If
If ChEscMinDef.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MINDEF", TxtEscMinDef.Text)
End If
If ChEscMaxDef.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MAXDEF", TxtEscMaxDef.Text)
End If
If ChNavegacion.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MinSkill", TxtNavegacion.Text)
End If
If ChHerreria.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "SkHerreria", TxtHerreria.Text)
End If
If ChCarpinteria.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "SkCarpinteria", TxtCarpinteria.Text)
End If
If ChLingOro.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "LingO", TxtLingOro.Text)
End If
If ChLingPlata.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "LingP", TxtLingPlata.Text)
End If
If ChLingHierro.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "LingH", TxtLingHierro.Text)
End If
If ChMadera.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Madera", TxtMadera.Text)
End If
If ChHeadAnim.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "Anim", TxtHeadAnim.Text)
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "NumRopaje", "2")
End If
If ChHeadDefMin.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MINDEF", TxtHeadDefMin.Text)
End If
If ChHeadDefMax.value = 1 Then
    Call WriteVar(Config.SaveDatPath & "\Obj.dat", "OBJ" & ObjNumber, "MAXDEF", TxtHeadDefMax.Text)
End If
If ChClases.value = 1 Then
    Dim Class As Byte
    Dim CPCont As Byte
    For Class = 0 To LstClases.ListCount - 1
        If LstClases.Selected(Class) Then
            CPCont = CPCont + 1
            Call WriteVar(Config.SaveDatPath & "\obj.dat", "OBJ" & ObjNumber, "CP" & CPCont, LstClases.List(Class))
        End If
    Next Class
End If
If ChAdd.value = 1 Then
    Dim Parms As Integer
    For Parms = 0 To LstAdd.ListCount - 1
        Dim NombreParm As String
        Dim ValorParm As String
        NombreParm = Left(LstAdd.List(Parms), InStr(1, LstAdd.List(Parms), "=") - 1)
        ValorParm = Mid(LstAdd.List(Parms), InStr(1, LstAdd.List(Parms), "=") + 1, Trim(Len(LstAdd.List(Parms))))
        Call WriteVar(Config.SaveDatPath & "\obj.dat", "OBJ" & ObjNumber, NombreParm, ValorParm)
    Next Parms
End If

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

Private Sub MNpcs_Click()
FrmNpcCreator.Visible = True
Me.Visible = False
End Sub

Private Sub MSalir_Click()
End
End Sub
