VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_opsdesp 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Utilitarios DESPACHO"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14175
   Icon            =   "frm_opsdesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frm_opsdesp.frx":058A
   ScaleHeight     =   8130
   ScaleWidth      =   14175
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport crmsp 
      Left            =   6360
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   12000
      Left            =   6480
      Top             =   3360
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   12090887
      MouseIcon       =   "frm_opsdesp.frx":0B14
      TabCaption(0)   =   "Móviles"
      TabPicture(0)   =   "frm_opsdesp.frx":0B30
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6(0)"
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(2)=   "Label3(0)"
      Tab(0).Control(3)=   "labchof"
      Tab(0).Control(4)=   "Label2(0)"
      Tab(0).Control(5)=   "Label5(0)"
      Tab(0).Control(6)=   "labenf"
      Tab(0).Control(7)=   "Label4(0)"
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(9)=   "Line1"
      Tab(0).Control(10)=   "Line2"
      Tab(0).Control(11)=   "Line3"
      Tab(0).Control(12)=   "Line4"
      Tab(0).Control(13)=   "Line5"
      Tab(0).Control(14)=   "dbmedic"
      Tab(0).Control(15)=   "txt_nro"
      Tab(0).Control(16)=   "txt_codmed"
      Tab(0).Control(17)=   "Command1(0)"
      Tab(0).Control(18)=   "t_chof(0)"
      Tab(0).Control(19)=   "mfec"
      Tab(0).Control(20)=   "t_base"
      Tab(0).Control(21)=   "t_enf(0)"
      Tab(0).Control(22)=   "Command2(0)"
      Tab(0).Control(23)=   "cr1"
      Tab(0).Control(24)=   "b_nuevo"
      Tab(0).Control(25)=   "b_modif(0)"
      Tab(0).Control(26)=   "b_graba(0)"
      Tab(0).Control(27)=   "b_cance(0)"
      Tab(0).Control(28)=   "b_elimina"
      Tab(0).Control(29)=   "b_busca"
      Tab(0).Control(30)=   "b_imprime"
      Tab(0).Control(31)=   "data_chof(0)"
      Tab(0).Control(32)=   "data_med(0)"
      Tab(0).Control(33)=   "data_mov"
      Tab(0).Control(34)=   "data_enf(0)"
      Tab(0).Control(35)=   "data_inf(0)"
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "Acondicionamientos"
      TabPicture(1)   =   "frm_opsdesp.frx":1020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "data_chof(1)"
      Tab(1).Control(1)=   "data_med(1)"
      Tab(1).Control(2)=   "data_enf(1)"
      Tab(1).Control(3)=   "data_graba"
      Tab(1).Control(4)=   "data_busca"
      Tab(1).Control(5)=   "Check1"
      Tab(1).Control(6)=   "Command222(1)"
      Tab(1).Control(7)=   "t_movb"
      Tab(1).Control(8)=   "b_imp"
      Tab(1).Control(9)=   "b_eli"
      Tab(1).Control(10)=   "b_cancee(1)"
      Tab(1).Control(11)=   "b_grabaa(1)"
      Tab(1).Control(12)=   "b_modiff(1)"
      Tab(1).Control(13)=   "b_alta"
      Tab(1).Control(14)=   "Frame1"
      Tab(1).Control(15)=   "DBGrid13(0)"
      Tab(1).Control(16)=   "Label8"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Médicos"
      TabPicture(2)   =   "frm_opsdesp.frx":1510
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label18"
      Tab(2).Control(1)=   "Label39"
      Tab(2).Control(2)=   "Label45"
      Tab(2).Control(3)=   "CrystalReport1"
      Tab(2).Control(4)=   "DBGrid112(1)"
      Tab(2).Control(5)=   "Frame2"
      Tab(2).Control(6)=   "txt_bcob"
      Tab(2).Control(7)=   "bnuevo"
      Tab(2).Control(8)=   "bgraba"
      Tab(2).Control(9)=   "bmodif"
      Tab(2).Control(10)=   "bcance"
      Tab(2).Control(11)=   "bbusca"
      Tab(2).Control(12)=   "bimp"
      Tab(2).Control(13)=   "data_cob"
      Tab(2).Control(14)=   "Data1"
      Tab(2).Control(15)=   "data_medhc"
      Tab(2).Control(16)=   "List4"
      Tab(2).Control(17)=   "Command10"
      Tab(2).Control(18)=   "data_listamedicos"
      Tab(2).Control(19)=   "mdmed"
      Tab(2).Control(20)=   "b_vermedcmt"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "Horarios"
      TabPicture(3)   =   "frm_opsdesp.frx":1AAA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "crhs"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Informe por turno"
      TabPicture(4)   =   "frm_opsdesp.frx":2044
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label29"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label31"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "dbturno"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "mfbusca"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "b_nuetur"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "b_gratur"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "b_cantur"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "b_inftur"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "data_grabatur"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "data_vertur"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "data_inftur"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "data_llamtur"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "b_busturn"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "Check3"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "crturno"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "b_envcor"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "Winsock1"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "Timer1"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "Command7"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).ControlCount=   20
      TabCaption(5)   =   "Resumen diario"
      TabPicture(5)   =   "frm_opsdesp.frx":25DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label34"
      Tab(5).Control(1)=   "labcmtdesp"
      Tab(5).Control(2)=   "Label36"
      Tab(5).Control(3)=   "labcmtpol"
      Tab(5).Control(4)=   "Label38"
      Tab(5).Control(5)=   "labpenddesp"
      Tab(5).Control(6)=   "Label40"
      Tab(5).Control(7)=   "labpendpol"
      Tab(5).Control(8)=   "Label42"
      Tab(5).Control(9)=   "Label35"
      Tab(5).Control(10)=   "labseg"
      Tab(5).Control(11)=   "Label37"
      Tab(5).Control(12)=   "labpendposi"
      Tab(5).Control(13)=   "Label41"
      Tab(5).Control(14)=   "labtotmg"
      Tab(5).Control(15)=   "Label44"
      Tab(5).Control(16)=   "labpendmg"
      Tab(5).Control(17)=   "Label43"
      Tab(5).Control(18)=   "labtotres"
      Tab(5).Control(19)=   "Label46"
      Tab(5).Control(20)=   "labpendres"
      Tab(5).Control(21)=   "List1"
      Tab(5).Control(22)=   "List2"
      Tab(5).Control(23)=   "mdcmt"
      Tab(5).Control(24)=   "mhcmt"
      Tab(5).Control(25)=   "Command8"
      Tab(5).Control(26)=   "List3"
      Tab(5).Control(27)=   "List5"
      Tab(5).Control(28)=   "List6"
      Tab(5).ControlCount=   29
      Begin VB.CommandButton b_vermedcmt 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   -66000
         Picture         =   "frm_opsdesp.frx":2B78
         Style           =   1  'Graphical
         TabIndex        =   175
         ToolTipText     =   "Muestra los médicos ingresados para la fecha seleccionada"
         Top             =   5760
         Width           =   495
      End
      Begin MSMask.MaskEdBox mdmed 
         Height          =   375
         Left            =   -64080
         TabIndex        =   174
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ListBox List6 
         Height          =   2400
         Left            =   -64560
         TabIndex        =   170
         Top             =   5040
         Width           =   3255
      End
      Begin VB.ListBox List5 
         Height          =   2790
         Left            =   -68040
         TabIndex        =   163
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Data data_listamedicos 
         Caption         =   "data_listamedicos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   -65280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpiar lista"
         Height          =   375
         Left            =   -65040
         Picture         =   "frm_opsdesp.frx":3102
         Style           =   1  'Graphical
         TabIndex        =   162
         ToolTipText     =   "Borra todos los médicos de la lista"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.ListBox List4 
         Height          =   3570
         Left            =   -66000
         TabIndex        =   160
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ListBox List3 
         Height          =   2790
         Left            =   -64560
         TabIndex        =   155
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualizar datos"
         Height          =   975
         Left            =   -69600
         Picture         =   "frm_opsdesp.frx":368C
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   5040
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mhcmt 
         Height          =   375
         Left            =   -72120
         TabIndex        =   152
         Top             =   5640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mdcmt 
         Height          =   375
         Left            =   -72120
         TabIndex        =   151
         Top             =   5160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ListBox List2 
         Height          =   2790
         Left            =   -71400
         TabIndex        =   145
         Top             =   1200
         Width           =   3135
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   -74760
         TabIndex        =   144
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Positivos CCOU"
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   6540
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data data_medhc 
         Caption         =   "data_medhc"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   -72120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4320
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   6840
         Top             =   6180
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   9000
         Top             =   6060
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton b_envcor 
         Caption         =   "Command3"
         Height          =   495
         Left            =   7680
         TabIndex        =   126
         Top             =   6060
         Visible         =   0   'False
         Width           =   615
      End
      Begin Crystal.CrystalReport crturno 
         Left            =   6000
         Top             =   6060
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Enviar informe por mail"
         Height          =   255
         Left            =   2520
         TabIndex        =   125
         Top             =   6060
         Width           =   2295
      End
      Begin VB.CommandButton b_busturn 
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   13200
         Picture         =   "frm_opsdesp.frx":3C16
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   1380
         Width           =   375
      End
      Begin VB.Data data_llamtur 
         Caption         =   "data_llamtur"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6180
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Data data_inftur 
         Caption         =   "data_inftur"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   9960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6420
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_vertur 
         Caption         =   "data_vertur"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "mant_sol_hc"
         Top             =   6300
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_grabatur 
         Caption         =   "data_grabatur"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6540
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton b_inftur 
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   2040
         Picture         =   "frm_opsdesp.frx":41A0
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Informe del turno seleccionado"
         Top             =   5940
         Width           =   375
      End
      Begin VB.CommandButton b_cantur 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         Picture         =   "frm_opsdesp.frx":472A
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   5940
         Width           =   375
      End
      Begin VB.CommandButton b_gratur 
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   840
         Picture         =   "frm_opsdesp.frx":4CB4
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   5940
         Width           =   375
      End
      Begin VB.CommandButton b_nuetur 
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         Picture         =   "frm_opsdesp.frx":523E
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   5940
         Width           =   375
      End
      Begin MSMask.MaskEdBox mfbusca 
         Height          =   375
         Left            =   11640
         TabIndex        =   111
         Top             =   1380
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12640511
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDBGrid.DBGrid dbturno 
         Bindings        =   "frm_opsdesp.frx":57C8
         Height          =   4215
         Left            =   9840
         OleObjectBlob   =   "frm_opsdesp.frx":57E2
         TabIndex        =   109
         Top             =   1740
         Width           =   3975
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Informe por turno"
         Height          =   5055
         Left            =   240
         TabIndex        =   100
         Top             =   900
         Width           =   9615
         Begin VB.CommandButton b_samc 
            Caption         =   "SAMC"
            Height          =   495
            Left            =   240
            TabIndex        =   176
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "SMI Covid"
            Height          =   375
            Left            =   240
            TabIndex        =   138
            Top             =   4440
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Data data_paraci 
            Caption         =   "data_paraci"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1920
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton Command5 
            Caption         =   "EVANG Covid"
            Height          =   495
            Left            =   240
            TabIndex        =   137
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Data data_covidT 
            Caption         =   "data_covidT"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2160
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Data data_covid 
            Caption         =   "data_covid"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   4200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Data data_llam 
            Caption         =   "data_llam"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data data_ctrf 
            Caption         =   "data_ctrf"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2520
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "CCOU Covid"
            Height          =   615
            Left            =   240
            TabIndex        =   136
            Top             =   3000
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   2400
            TabIndex        =   135
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox t_envmsp 
            Height          =   375
            Left            =   5880
            TabIndex        =   134
            Top             =   3240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data data_llam2 
            Caption         =   "data_llam2"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   5160
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Data data_inftr 
            Caption         =   "data_inftr"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   5520
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1080
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Data data_fecmsp 
            Caption         =   "data_fecmsp"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   2040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3120
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.CommandButton Command3 
            Caption         =   "msp"
            Height          =   375
            Left            =   2160
            TabIndex        =   133
            Top             =   2280
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H0080FF80&
            Caption         =   "CERRAR TURNO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   124
            Top             =   4560
            Width           =   2895
         End
         Begin MSMask.MaskEdBox mhfin 
            Height          =   375
            Left            =   4320
            TabIndex        =   119
            Top             =   960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mffin 
            Height          =   375
            Left            =   1680
            TabIndex        =   117
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox t_obsturno 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   108
            Top             =   1560
            Width           =   7695
         End
         Begin MSMask.MaskEdBox mhor 
            Height          =   375
            Left            =   4320
            TabIndex        =   104
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mf 
            Height          =   375
            Left            =   1680
            TabIndex        =   102
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label labnomu 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   5640
            TabIndex        =   123
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label labidtur 
            Height          =   375
            Left            =   5640
            TabIndex        =   122
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Hora:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   118
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0E0FF&
            Caption         =   "FINALIZA:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   116
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Observaciones:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   107
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label labu 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6960
            TabIndex        =   106
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Usuario:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   105
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Hora:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   103
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0E0FF&
            Caption         =   "INICIO:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   101
            Top             =   480
            Width           =   1455
         End
      End
      Begin Crystal.CrystalReport crhs 
         Left            =   -65280
         Top             =   5340
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080FF80&
         Caption         =   "Entradas y Salidas Médicos No Dep."
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   -73920
         TabIndex        =   81
         Top             =   1140
         Width           =   8415
         Begin MSAdodcLib.Adodc data_medhcc 
            Height          =   375
            Left            =   120
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "data_medhcc"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc data_hsmed2 
            Height          =   375
            Left            =   3480
            Top             =   840
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "data_hsmed2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc data_verhsme 
            Height          =   375
            Left            =   360
            Top             =   1560
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "data_verhsme"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Data data_verhsmed 
            Caption         =   "data_verhsmed"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   495
            Left            =   4320
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   120
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Data data_par 
            Caption         =   "data_par"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   840
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2280
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox t_nromov 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   128
            Top             =   1200
            Width           =   975
         End
         Begin VB.Data data_infhs 
            Caption         =   "data_infhs"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   4800
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1680
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Data data_m 
            Caption         =   "data_m"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   4920
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3120
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Data data_horasmed 
            Caption         =   "data_horasmed"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   4680
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2640
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.TextBox t_codmedb 
            Height          =   375
            Left            =   2160
            TabIndex        =   98
            Top             =   3120
            Width           =   735
         End
         Begin MSDBGrid.DBGrid dbensal 
            Bindings        =   "frm_opsdesp.frx":6365
            Height          =   1815
            Left            =   240
            OleObjectBlob   =   "frm_opsdesp.frx":6381
            TabIndex        =   96
            ToolTipText     =   "doble click para editar los datos"
            Top             =   3480
            Width           =   6615
         End
         Begin VB.CommandButton b_in 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   3120
            Picture         =   "frm_opsdesp.frx":70B8
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Ver en pantalla o impresora los registros del médico seleccionado o todos por fecha"
            Top             =   2640
            Width           =   495
         End
         Begin VB.CommandButton b_ca 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            Picture         =   "frm_opsdesp.frx":7642
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   2640
            Width           =   495
         End
         Begin VB.CommandButton b_gr 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            Picture         =   "frm_opsdesp.frx":7BCC
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   2640
            Width           =   495
         End
         Begin VB.CommandButton b_ed 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   960
            Picture         =   "frm_opsdesp.frx":8156
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   2640
            Width           =   495
         End
         Begin VB.CommandButton b_alt 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   240
            Picture         =   "frm_opsdesp.frx":86E0
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   2640
            Width           =   495
         End
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3120
            TabIndex        =   90
            Top             =   1920
            Width           =   5055
         End
         Begin VB.TextBox t_codmedm 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   89
            Top             =   1920
            Width           =   855
         End
         Begin MSMask.MaskEdBox mhmed 
            Height          =   495
            Left            =   5520
            TabIndex        =   87
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   393216
            ForeColor       =   255
            Enabled         =   0   'False
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboes 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "frm_opsdesp.frx":8C6A
            Left            =   2160
            List            =   "frm_opsdesp.frx":8C74
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   1200
            Width           =   2415
         End
         Begin MSMask.MaskEdBox mfmed 
            Height          =   495
            Left            =   2160
            TabIndex        =   83
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            _Version        =   393216
            ForeColor       =   255
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0FFFF&
            Caption         =   "BASE/MOVIL:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   127
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label labid 
            Height          =   375
            Left            =   6960
            TabIndex        =   99
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Buscar por médico:"
            Height          =   375
            Left            =   240
            TabIndex        =   97
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Line Line8 
            BorderColor     =   &H0000C000&
            BorderWidth     =   2
            X1              =   0
            X2              =   8400
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0FFFF&
            Caption         =   "COD.MEDICO:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   88
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0FFFF&
            Caption         =   "HORA:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4440
            TabIndex        =   86
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ACCION:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0FFFF&
            Caption         =   "FECHA:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   82
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "medicos"
         Top             =   4920
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data data_cob 
         Caption         =   "data_cob"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -71400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "medicos"
         Top             =   5280
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton bimp 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -69120
         Picture         =   "frm_opsdesp.frx":8C89
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Informes"
         Top             =   3900
         Width           =   495
      End
      Begin VB.CommandButton bbusca 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -69960
         Picture         =   "frm_opsdesp.frx":9213
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Buscar"
         Top             =   3900
         Width           =   495
      End
      Begin VB.CommandButton bcance 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   -70800
         Picture         =   "frm_opsdesp.frx":979D
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Cancelar acción"
         Top             =   3900
         Width           =   495
      End
      Begin VB.CommandButton bmodif 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -71640
         Picture         =   "frm_opsdesp.frx":9D27
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Modificar datos"
         Top             =   3900
         Width           =   495
      End
      Begin VB.CommandButton bgraba 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   -72480
         Picture         =   "frm_opsdesp.frx":A2B1
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Grabar datos"
         Top             =   3900
         Width           =   495
      End
      Begin VB.CommandButton bnuevo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73320
         Picture         =   "frm_opsdesp.frx":A83B
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Nuevo registro"
         Top             =   3900
         Width           =   495
      End
      Begin VB.TextBox txt_bcob 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70800
         TabIndex        =   73
         Top             =   4740
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Médicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -73440
         TabIndex        =   63
         Top             =   780
         Width           =   6855
         Begin VB.CommandButton Command9 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5520
            Picture         =   "frm_opsdesp.frx":ADC5
            Style           =   1  'Graphical
            TabIndex        =   159
            ToolTipText     =   "Agregar el médico seleccionado a la lista"
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H0080FFFF&
            Caption         =   "Médico de Cooperativa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   132
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox t_codced 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   131
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox t_ced 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   130
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txt_tel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   71
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txt_espec 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   70
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txt_nomcob 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   69
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox txt_nrocob 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   68
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label33 
            BackColor       =   &H0000C000&
            Caption         =   "Cédula:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   129
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Line Line7 
            X1              =   0
            X2              =   6840
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   6840
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label Label17 
            BackColor       =   &H0000C000&
            Caption         =   "Teléfonos:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label16 
            BackColor       =   &H0000C000&
            Caption         =   "Especialidad:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label15 
            BackColor       =   &H0000C000&
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   65
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackColor       =   &H0000C000&
            Caption         =   "Cod.Médico:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Data data_chof 
         Caption         =   "data_chof"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   -65160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2880
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   -64800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data data_enf 
         Caption         =   "data_enf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   -66840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data data_graba 
         Caption         =   "data_graba"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -67080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Data data_busca 
         Caption         =   "data_busca"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -69360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   7080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Ver TODO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67080
         TabIndex        =   62
         Top             =   6300
         Width           =   2775
      End
      Begin VB.CommandButton Command222 
         Caption         =   "Tabla Chóferes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -63000
         TabIndex        =   61
         Top             =   3720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox t_movb 
         Height          =   375
         Left            =   -65040
         TabIndex        =   59
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton b_imp 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -70680
         Picture         =   "frm_opsdesp.frx":B34F
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Informes"
         Top             =   6420
         Width           =   615
      End
      Begin VB.CommandButton b_eli 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -71520
         Picture         =   "frm_opsdesp.frx":B8D9
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Eliminar registro"
         Top             =   6420
         Width           =   615
      End
      Begin VB.CommandButton b_cancee 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   -72360
         Picture         =   "frm_opsdesp.frx":BE63
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Cancelar acción"
         Top             =   6420
         Width           =   615
      End
      Begin VB.CommandButton b_grabaa 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   -73200
         Picture         =   "frm_opsdesp.frx":C3ED
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Grabar datos"
         Top             =   6420
         Width           =   615
      End
      Begin VB.CommandButton b_modiff 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   -74040
         Picture         =   "frm_opsdesp.frx":C977
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Editar registro"
         Top             =   6420
         Width           =   615
      End
      Begin VB.CommandButton b_alta 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -74880
         Picture         =   "frm_opsdesp.frx":CF01
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Nuevo registro"
         Top             =   6420
         Width           =   615
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Datos de registro del acondicionamiento"
         Enabled         =   0   'False
         Height          =   5295
         Left            =   -74880
         TabIndex        =   26
         Top             =   1020
         Width           =   7815
         Begin VB.Data data_inf2 
            Caption         =   "data_inf2"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   5160
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   4920
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox t_mov 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox t_choff 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   36
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox t_obs 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2160
            MaxLength       =   35
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   3960
            Width           =   5175
         End
         Begin VB.CommandButton Command111 
            BackColor       =   &H00FF8080&
            Caption         =   "Buscar..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   6000
            Picture         =   "frm_opsdesp.frx":D48B
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Ver tabla de choferes..."
            Top             =   840
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Data data_inf 
            Caption         =   "data_inf"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   1
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   4560
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frm_opsdesp.frx":DA15
            Left            =   2160
            List            =   "frm_opsdesp.frx":DA28
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   3480
            Width           =   3015
         End
         Begin VB.TextBox t_med 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2160
            TabIndex        =   28
            Top             =   1800
            Width           =   3495
         End
         Begin VB.TextBox t_enff 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   2160
            TabIndex        =   27
            Top             =   1320
            Width           =   3495
         End
         Begin MSMask.MaskEdBox mhorh 
            Height          =   375
            Left            =   3960
            TabIndex        =   32
            Top             =   2880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "HH:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mh 
            Height          =   375
            Left            =   2160
            TabIndex        =   33
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mhord 
            Height          =   375
            Left            =   3960
            TabIndex        =   34
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "HH:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox md 
            Height          =   375
            Left            =   2160
            TabIndex        =   35
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Número de móvil:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Chofer:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   50
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Fecha/Hora de comienzo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Fecha/Hora de finalización:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   48
            Top             =   2880
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Observaciones:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   3960
            Width           =   1935
         End
         Begin VB.Label Label6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   5880
            TabIndex        =   46
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Demora total:"
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
            Left            =   5880
            TabIndex        =   45
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label9 
            Height          =   375
            Left            =   2880
            TabIndex        =   44
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label labcod 
            Height          =   255
            Left            =   3720
            TabIndex        =   43
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label laben 
            Height          =   375
            Left            =   4800
            TabIndex        =   42
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label labmed 
            Height          =   375
            Left            =   6000
            TabIndex        =   41
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Servicio:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   3480
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Médico:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Enfermería:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   0
         Left            =   -74040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5280
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data data_enf 
         Caption         =   "data_enf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -74760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_mov 
         Caption         =   "data_mov"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -69840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4440
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Data data_chof 
         Caption         =   "data_chof"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   0
         Left            =   -70200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5400
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton b_imprime 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -69600
         Picture         =   "frm_opsdesp.frx":DA68
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Informes"
         Top             =   5940
         Width           =   495
      End
      Begin VB.CommandButton b_busca 
         BackColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   -70440
         Picture         =   "frm_opsdesp.frx":DFF2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Buscar datos"
         Top             =   6060
         Width           =   495
      End
      Begin VB.CommandButton b_elimina 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -71280
         Picture         =   "frm_opsdesp.frx":E57C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Elimina el registro del móvil seleccionado"
         Top             =   6000
         Width           =   495
      End
      Begin VB.CommandButton b_cance 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   -72120
         Picture         =   "frm_opsdesp.frx":EB06
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancelar la acción a realizar"
         Top             =   6000
         Width           =   495
      End
      Begin VB.CommandButton b_graba 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   -72960
         Picture         =   "frm_opsdesp.frx":F090
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Graba los datos"
         Top             =   6000
         Width           =   495
      End
      Begin VB.CommandButton b_modif 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   -73800
         Picture         =   "frm_opsdesp.frx":F61A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6000
         Width           =   495
      End
      Begin VB.CommandButton b_nuevo 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -74640
         Picture         =   "frm_opsdesp.frx":FBA4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Agregar registro NUEVO"
         Top             =   6000
         Width           =   495
      End
      Begin Crystal.CrystalReport cr1 
         Left            =   -67080
         Top             =   1980
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Buscar..."
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   -68040
         Picture         =   "frm_opsdesp.frx":1012E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox t_enf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -72600
         TabIndex        =   17
         Top             =   3840
         Width           =   4455
      End
      Begin VB.TextBox t_base 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   16
         Top             =   4800
         Width           =   615
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   -72600
         TabIndex        =   14
         Top             =   4800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_chof 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -72600
         TabIndex        =   8
         Top             =   2880
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Buscar..."
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   -68040
         Picture         =   "frm_opsdesp.frx":106B8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txt_codmed 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   4
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   -72600
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1440
         Width           =   1695
      End
      Begin MSDBCtls.DBCombo dbmedic 
         Bindings        =   "frm_opsdesp.frx":10C42
         Height          =   360
         Left            =   -72600
         TabIndex        =   5
         Top             =   2040
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MED_NOMBRE"
         BoundColumn     =   "MED_NOMBRE"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDBGrid.DBGrid DBGrid13 
         Bindings        =   "frm_opsdesp.frx":10C5C
         Height          =   2175
         Index           =   0
         Left            =   -67080
         OleObjectBlob   =   "frm_opsdesp.frx":10C75
         TabIndex        =   60
         Top             =   4140
         Width           =   5895
      End
      Begin MSDBGrid.DBGrid DBGrid112 
         Bindings        =   "frm_opsdesp.frx":117F0
         Height          =   1935
         Index           =   1
         Left            =   -73440
         OleObjectBlob   =   "frm_opsdesp.frx":11807
         TabIndex        =   72
         Top             =   5100
         Width           =   6855
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   -74760
         Top             =   1140
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label45 
         BackColor       =   &H00404040&
         Caption         =   "Fecha (opcional)"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -66000
         TabIndex        =   173
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label labpendres 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -62160
         TabIndex        =   172
         Top             =   7440
         Width           =   855
      End
      Begin VB.Label Label46 
         BackColor       =   &H000040C0&
         Caption         =   "Pendientes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -64560
         TabIndex        =   171
         Top             =   7440
         Width           =   2055
      End
      Begin VB.Label labtotres 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -62160
         TabIndex        =   169
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label43 
         BackColor       =   &H00404040&
         Caption         =   "CMT Resultados Covid-19"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -64560
         TabIndex        =   168
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label labpendmg 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -65880
         TabIndex        =   167
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label44 
         BackColor       =   &H000040C0&
         Caption         =   "Pendientes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -68040
         TabIndex        =   166
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label labtotmg 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -65760
         TabIndex        =   165
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label41 
         BackColor       =   &H00404040&
         Caption         =   "Policlínica presenciales MG:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -68040
         TabIndex        =   164
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label39 
         BackColor       =   &H000080FF&
         Caption         =   "Médicos de guardia para CMT y Polic. de Medicina General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -66000
         TabIndex        =   161
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label labpendposi 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -62160
         TabIndex        =   158
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label37 
         BackColor       =   &H000040C0&
         Caption         =   "Pendientes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -64560
         TabIndex        =   157
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label labseg 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -62040
         TabIndex        =   156
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label35 
         BackColor       =   &H00404040&
         Caption         =   "Covid positivos"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -64560
         TabIndex        =   154
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label42 
         BackColor       =   &H00404040&
         Caption         =   "Rango de fechas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -74760
         TabIndex        =   150
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label labpendpol 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -69240
         TabIndex        =   149
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label40 
         BackColor       =   &H000040C0&
         Caption         =   "Pendientes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -71400
         TabIndex        =   148
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label labpenddesp 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -72600
         TabIndex        =   147
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H000040C0&
         Caption         =   "Pendientes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74760
         TabIndex        =   146
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label labcmtpol 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -69120
         TabIndex        =   143
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label36 
         BackColor       =   &H00404040&
         Caption         =   "CMT Policlínicas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -71400
         TabIndex        =   142
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label labcmtdesp 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -72480
         TabIndex        =   141
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label34 
         BackColor       =   &H00404040&
         Caption         =   "CMT Despacho"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74760
         TabIndex        =   140
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H0080FFFF&
         Caption         =   "Doble click para editar un turno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   120
         Top             =   5940
         Width           =   3975
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Buscar fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   110
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackColor       =   &H0000FF00&
         Caption         =   "Nombre a buscar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   74
         Top             =   4740
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackColor       =   &H000080FF&
         Caption         =   "Buscar por móvil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66960
         TabIndex        =   58
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   -74880
         X2              =   -66360
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   -66360
         X2              =   -66360
         Y1              =   6780
         Y2              =   1020
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   -66360
         X2              =   -74880
         Y1              =   6780
         Y2              =   6780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   -74880
         X2              =   -74880
         Y1              =   6780
         Y2              =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   -74880
         X2              =   -66360
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Base Fact:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -70560
         TabIndex        =   15
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha actualización:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   13
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label labenf 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   -74640
         TabIndex        =   12
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enfermería"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   11
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Chófer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   10
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label labchof 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   -74280
         TabIndex        =   9
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Médico:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nro. MOVIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   0
         Left            =   -70800
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   12360
      Picture         =   "frm_opsdesp.frx":1239B
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1815
   End
End
Attribute VB_Name = "frm_opsdesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_alt_Click()

mfmed.Enabled = True
mhmed.Enabled = True
cboes.Enabled = True
t_codmedm.Enabled = True
Combo2.Enabled = True
dbensal.Enabled = False
t_codmedb.Enabled = False

mfmed.Text = Format(Date, "dd/mm/yyyy")
mhmed.Text = Format(Time, "HH:mm")
data_horasmed.RecordSource = "Select * from hc_archotro order by id DESC"
data_horasmed.Refresh
If data_horasmed.Recordset.RecordCount > 0 Then
   data_horasmed.Recordset.MoveFirst
   labid.Caption = data_horasmed.Recordset("id") + 1
Else
   labid.Caption = 1
End If
cboes.ListIndex = 0
cboes.SetFocus

b_alt.Enabled = False
b_ed.Enabled = False
b_gr.Enabled = True
b_ca.Enabled = True
b_in.Enabled = False

XAlta = 1


End Sub

Private Sub b_alta_Click()
XAlta = 1
deshab
limpia
If data_graba.Recordset.RecordCount > 0 Then
   data_graba.Recordset.MoveLast
   labcod.Caption = data_graba.Recordset("nrolla") + 1
Else
   labcod.Caption = 1
End If
t_mov.SetFocus

End Sub

Private Sub b_busturn_Click()
If mfbusca.Text <> "__/__/____" Then
   If XWeltipoU = "ADMINISTRADOR" Then
      data_vertur.RecordSource = "Select * from mant_sol_hc where cl_fnac >=#" & Format(mfbusca.Text, "yyyy/mm/dd") & "# order by cl_fnac DESC"
      data_vertur.Refresh
   Else
      data_vertur.RecordSource = "Select * from mant_sol_hc where cl_fnac >=#" & Format(mfbusca.Text, "yyyy/mm/dd") & "# and cl_descpag ='" & WElusuario & "' order by cl_fnac DESC"
      data_vertur.Refresh
   End If
   dbturno.SetFocus
Else
   MsgBox "No ingresó fecha a buscar"
End If

End Sub

Private Sub b_ca_Click()
XAlta = 0
mfmed.Enabled = False
mhmed.Enabled = False
cboes.Enabled = False
t_codmedm.Enabled = False
Combo2.Enabled = False
dbensal.Enabled = True
t_codmedb.Enabled = True

b_alt.Enabled = True
b_ed.Enabled = True
b_gr.Enabled = False
b_ca.Enabled = False
b_in.Enabled = True

End Sub

Private Sub b_cance_Click(index As Integer)

If XAlta = 1 Then
   data_mov.Recordset.CancelUpdate
End If
XAlta = 0
Command1(0).Enabled = False
Command2(0).Enabled = False

data_mov.Recordset.MoveFirst
habilitamov
igualacuadros

'Frame1.Enabled = False

End Sub

Private Sub b_cancee_Click(index As Integer)
limpia
habilita

End Sub

Private Sub b_cantur_Click()
If XAlta = 1 Then
   data_grabatur.Recordset.CancelUpdate
   
   mf.Text = "__/__/____"
   mhor.Text = "__:__"
   mffin.Text = "__/__/____"
   mhfin.Text = "__:__"
   labu.Caption = ""
   t_obsturno.Text = ""
   labnomu.Caption = ""
   labidtur.Caption = ""
   dbturno.Enabled = True
   b_nuetur.Enabled = True
   b_cantur.Enabled = True
   b_inftur.Enabled = True
   b_busturn.Enabled = True
Else
   mf.Text = "__/__/____"
   mhor.Text = "__:__"
   mffin.Text = "__/__/____"
   mhfin.Text = "__:__"
   labu.Caption = ""
   t_obsturno.Text = ""
   labnomu.Caption = ""
   labidtur.Caption = ""
End If


End Sub

Private Sub b_ed_Click()
mfmed.Enabled = True
mhmed.Enabled = True
cboes.Enabled = True
t_codmedm.Enabled = True
Combo2.Enabled = True
dbensal.Enabled = False
t_codmedb.Enabled = False

If labid.Caption <> "" Then
   igualahs
End If

b_alt.Enabled = False
b_ed.Enabled = False
b_gr.Enabled = True
b_ca.Enabled = True
b_in.Enabled = False

XAlta = 0
End Sub

Private Sub b_eli_Click()
Dim Xelmen As String
Xelmen = MsgBox("Desea eliminar el registro seleccionado?", vbInformation + vbYesNo, "SAPP")
If Xelmen = vbYes Then
   data_graba.Recordset.FindFirst "nrolla =" & data_busca.Recordset("nrolla")
   If Not data_graba.Recordset.NoMatch Then
      data_graba.Recordset.Delete
      data_graba.Refresh
      data_busca.Refresh
      limpia
   End If
End If

End Sub

Private Sub b_envcor_Click()
Dim MenCorreo As String
Dim oMail As Class1
     
Dim lineatexto, textocorreo As String

Open App.path & "\correosll.txt" For Input As #1
Do While Not EOF(1)
   Line Input #1, lineatexto
Loop
textocorreo = lineatexto
Close #1
     
     Set oMail = New Class1
     With oMail
         .servidor = "smtp.office365.com"
         .puerto = 25
         .UseAuntentificacion = True
         .ssl = True
         .Usuario = "despacho@sapp.com.uy"
         .PassWord = "Salinas1987"
         .Asunto = "Informe por turno " & mf.Text & " " & labnomu.Caption
         .de = "despacho@sapp.com.uy"
         .para = textocorreo
'         .para = "sappjorge@hotmail.com; despachosapp@hotmail.com; sappsusanadominguez@hotmail.com; sappdirecciontecnica@hotmail.com; sappenrique@hotmail.com"
         .Adjunto = "c:\planillas\inftr.pdf"
         .Mensaje = "Informe por turno"
         .Enviar_Backup ' manda el mail
     End With
     Set oMail = Nothing
'     MsgBox "Correo enviado..."
     Command3_Click
''     Timer3.Enabled = True


End Sub

Private Sub b_gr_Click()
On Error GoTo Queeres

If mfmed.Text <> "__/__/____" And mhmed.Text <> "__:__" And t_codmedm.Text <> "" Then
   If XAlta = 1 Then
      data_par.Recordset.Edit
      data_par.Recordset("nro_reg") = data_par.Recordset("nro_reg") + 1
      data_par.Recordset.Update
      data_par.Refresh
      
      data_horasmed.Recordset.AddNew
      data_horasmed.Recordset("id") = data_par.Recordset("nro_reg")
      data_horasmed.Recordset("hc_mat") = t_codmedm.Text
      data_horasmed.Recordset("hc_nro") = cboes.ListIndex
      data_horasmed.Recordset("hc_fecha") = mfmed.Text
      data_horasmed.Recordset("hc_hora") = mhmed.Text
      data_horasmed.Recordset("hc_descrip") = WElusuario
      If Combo2.Text <> "" Then
         data_horasmed.Recordset("hc_lugar") = Combo2.Text
      End If
      data_horasmed.Recordset.Update
      
      data_hsmed2.RecordSource = "Select * from hc_viaae where id =" & data_par.Recordset("nro_reg")
      data_hsmed2.Refresh
   
      data_hsmed2.Recordset.AddNew
      data_hsmed2.Recordset("id") = data_par.Recordset("nro_reg")
      data_hsmed2.Recordset("hc_cod") = t_nromov.Text
      data_hsmed2.Recordset.Update
      data_hsmed2.Refresh
      data_verhsmed.Refresh
      mfmed.Text = "__/__/____"
      mhmed.Text = "__:__"
      cboes.ListIndex = -1
      t_codmedm.Text = ""
      t_nromov.Text = ""
      Combo2.ListIndex = -1
      mfmed.Enabled = False
      mhmed.Enabled = False
      cboes.Enabled = False
      t_codmedm.Enabled = False
      Combo2.Enabled = False
      dbensal.Enabled = True
      t_codmedb.Enabled = True
      b_alt.Enabled = True
      b_ed.Enabled = True
      b_gr.Enabled = False
      b_ca.Enabled = False
      b_in.Enabled = True
   Else
      data_horasmed.Recordset.Edit
      data_horasmed.Recordset("hc_nro") = cboes.ListIndex
      data_horasmed.Recordset("hc_fecha") = mfmed.Text
      data_horasmed.Recordset("hc_hora") = mhmed.Text
      data_horasmed.Recordset("hc_descrip") = WElusuario
      If Combo2.Text <> "" Then
         data_horasmed.Recordset("hc_lugar") = Combo2.Text
      End If
      data_horasmed.Recordset.Update
      data_hsmed2.RecordSource = "Select * from hc_viaae where id =" & data_par.Recordset("id")
      data_hsmed2.Refresh
      If data_hsmed2.Recordset.RecordCount > 0 Then
'         data_hsmed2.Recordset.Edit
         data_hsmed2.Recordset("hc_cod") = t_nromov.Text
         data_hsmed2.Recordset.Update
         data_hsmed2.Refresh
      End If
      data_verhsmed.Refresh
      mfmed.Text = "__/__/____"
      mhmed.Text = "__:__"
      cboes.ListIndex = -1
      t_codmedm.Text = ""
      t_nromov.Text = ""
      Combo2.ListIndex = -1
      mfmed.Enabled = False
      mhmed.Enabled = False
      cboes.Enabled = False
      t_codmedm.Enabled = False
      Combo2.Enabled = False
      dbensal.Enabled = True
      t_codmedb.Enabled = True
      b_alt.Enabled = True
      b_ed.Enabled = True
      b_gr.Enabled = False
      b_ca.Enabled = False
      b_in.Enabled = True
   End If
End If
XAlta = 0

Exit Sub

Queeres:
        If Err.Number = 3155 Then
           MsgBox "No hay modificaciones a grabar"
           b_ca_Click
        Else
           MsgBox "Error al grabar, verifique datos"
        End If

End Sub

Private Sub b_graba_Click(index As Integer)

If txt_nro.Text <> "" Then
   If txt_nro.Text > 0 Then
      If XAlta = 1 Then
         data_mov.Recordset("nroreg") = Label6(0).Caption
         data_mov.Recordset("movil") = txt_nro.Text
         If txt_codmed.Text = "" Then
            txt_codmed.Text = 0
         End If
         data_mov.Recordset("codmed") = txt_codmed.Text
         If dbmedic.Text = "" Then
         Else
            data_mov.Recordset("nommed") = dbmedic.Text
         End If
         If mfec.Text = "__/__/____" Then
         Else
            data_mov.Recordset("fecha_act") = mfec.Text
         End If
         data_mov.Recordset("hora_act") = Format(Time, "HH:mm")
         If t_base.Text = "" Then
            t_base.Text = 0
         End If
         data_mov.Recordset("ano") = t_base.Text
         If labchof.Caption = "" Then
            labchof.Caption = 0
         End If
         data_mov.Recordset("codchof") = labchof.Caption
         If t_chof(0).Text = "" Then
         Else
            data_mov.Recordset("nomchof") = t_chof(0).Text
         End If
         If labenf.Caption = "" Then
            labenf.Caption = 0
         End If
         data_mov.Recordset("codenf") = labenf.Caption
         If t_enf(0).Text = "" Then
         Else
            data_mov.Recordset("nomenf") = t_enf(0).Text
         End If
         data_mov.Recordset.Update
         data_mov.Refresh
         borra_mov
         XAlta = 0
      End If
      If XAlta = 2 Then
         data_mov.Recordset.Edit
         data_mov.Recordset("movil") = txt_nro.Text
         If txt_codmed.Text = "" Then
            txt_codmed.Text = 0
         End If
         data_mov.Recordset("codmed") = txt_codmed.Text
         If dbmedic.Text = "" Then
         Else
            data_mov.Recordset("nommed") = dbmedic.Text
         End If
         If mfec.Text = "__/__/____" Then
         Else
            data_mov.Recordset("fecha_act") = mfec.Text
         End If
         data_mov.Recordset("hora_act") = Format(Time, "HH:mm")
         If t_base.Text = "" Then
            t_base.Text = 0
         End If
         data_mov.Recordset("ano") = t_base.Text
         If labchof.Caption = "" Then
            labchof.Caption = 0
         End If
         data_mov.Recordset("codchof") = labchof.Caption
         If t_chof(0).Text = "" Then
         Else
            data_mov.Recordset("nomchof") = t_chof(0).Text
         End If
         If labenf.Caption = "" Then
            labenf.Caption = 0
         End If
         data_mov.Recordset("codenf") = labenf.Caption
         If t_enf(0).Text = "" Then
         Else
            data_mov.Recordset("nomenf") = t_enf(0).Text
         End If
         data_mov.Recordset.Update
         data_mov.Refresh
         XAlta = 0
      End If
      Command1(0).Enabled = False
      Command2(0).Enabled = False
      
      borra_mov
      habilitamov
      Frame1.Enabled = False
   Else
      MsgBox "No puede ser CERO el número de móvil", vbInformation, "Móviles"
   End If
Else
   MsgBox "No ingresó número de móvil", vbInformation, "Mensaje"

End If


End Sub

Private Sub b_grabaa_Click(index As Integer)
Dim xhh, xmm, Xss As Long
Dim Xhhh, Xmmh, Xssh As Long
Dim Xtoth, Xtotm, Xtots As Long

On Error GoTo Quepasaalg

If t_mov.Text <> "" Then
   If t_choff(1).Text <> "" Then
      If XAlta = 1 Then
         data_graba.Recordset.AddNew
         data_graba.Recordset("nromov") = 999
         data_graba.Recordset("nrolla") = labcod.Caption
         data_graba.Recordset("codmed") = t_mov.Text
         data_graba.Recordset("medico") = t_choff(1).Text
         data_graba.Recordset("zona") = Label9.Caption
         If md.Text = "__/__/____" Then
         Else
            data_graba.Recordset("fecha") = md.Text
         End If
         If mhord.Text = "__:__:__" Then
         Else
            data_graba.Recordset("usuario") = mhord.Text
         End If
         If mh.Text = "__/__/____" Then
         Else
            data_graba.Recordset("fecmod") = mh.Text
         End If
         If mhorh.Text = "__:__:__" Then
         Else
            data_graba.Recordset("matricm") = mhorh.Text
         End If
         If laben.Caption = "" Then
            data_graba.Recordset("ult_kms") = 0
         Else
            data_graba.Recordset("ult_kms") = laben.Caption
         End If
         If labmed.Caption = "" Then
            data_graba.Recordset("pro_kms") = 0
         Else
            data_graba.Recordset("pro_kms") = labmed.Caption
         End If
         If t_obs.Text <> "" Then
            data_graba.Recordset("motivo") = t_obs.Text
         End If
         If mhord.Text = "__:__:__" Then
            Label6(1).Caption = "00:00:00"
         Else
            If mhorh.Text = "__:__:__" Then
               Label6(1).Caption = "00:00:00"
            Else
               xhh = Mid(mhord.Text, 1, 2)
               xmm = Mid(mhord.Text, 4, 2)
               Xss = Mid(mhord.Text, 7, 2)
               Xhhh = Mid(mhorh.Text, 1, 2)
               Xmmh = Mid(mhorh.Text, 4, 2)
               Xssh = Mid(mhorh.Text, 7, 2)
               Xtoth = Xhhh - xhh
               Xtotm = Xmmh - xmm
               Xtots = Xssh - Xss
               If Xtots < 0 Then
                  Xtots = Xtots + 60
                  Xtotm = Xtotm - 1
               End If
               If Xtotm < 0 Then
                  Xtotm = Xtotm + 60
                  Xtoth = Xtoth - 1
               End If
               If Xtoth > 9 Then
                  If Xtotm > 9 Then
                     If Xtots > 9 Then
                        Label6(1).Caption = Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  Else
                     If Xtots > 9 Then
                        Label6(1).Caption = Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  End If
               Else
                  If Xtotm > 9 Then
                     If Xtots > 9 Then
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  Else
                     If Xtots > 9 Then
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  End If
               End If
            End If
         End If
         data_graba.Recordset("nroseg") = Label6(1).Caption
         If Combo1.ListIndex >= 0 Then
            data_graba.Recordset("kmactu") = Combo1.ListIndex
            data_graba.Recordset("desbase") = Combo1.Text
         Else
            data_graba.Recordset("kmactu") = Combo1.ListIndex
         End If
         data_graba.Recordset.Update
         data_graba.Refresh
         data_busca.Refresh
         limpia
         habilita
      Else
         data_graba.Recordset.Edit
         data_graba.Recordset("codmed") = t_mov.Text
         data_graba.Recordset("medico") = t_choff(1).Text
         data_graba.Recordset("zona") = Label9.Caption
         If md.Text = "__/__/____" Then
         Else
            data_graba.Recordset("fecha") = md.Text
         End If
         If mhord.Text = "__:__:__" Then
         Else
            data_graba.Recordset("usuario") = mhord.Text
         End If
         If mh.Text = "__/__/____" Then
         Else
            data_graba.Recordset("fecmod") = mh.Text
         End If
         If mhorh.Text = "__:__:__" Then
         Else
            data_graba.Recordset("matricm") = mhorh.Text
         End If
         If laben.Caption = "" Then
            data_graba.Recordset("ult_kms") = 0
         Else
            data_graba.Recordset("ult_kms") = laben.Caption
         End If
         If labmed.Caption = "" Then
            data_graba.Recordset("pro_kms") = 0
         Else
            data_graba.Recordset("pro_kms") = labmed.Caption
         End If
         If t_obs.Text <> "" Then
            data_graba.Recordset("motivo") = t_obs.Text
         End If
         If Combo1.ListIndex >= 0 Then
            data_graba.Recordset("kmactu") = Combo1.ListIndex
            data_graba.Recordset("desbase") = Combo1.Text
         Else
            data_graba.Recordset("kmactu") = Combo1.ListIndex
         End If

         If mhord.Text = "__:__:__" Then
            Label6(1).Caption = "00:00:00"
         Else
            If mhorh.Text = "__:__:__" Then
               Label6(1).Caption = "00:00:00"
            Else
               xhh = Mid(mhord.Text, 1, 2)
               xmm = Mid(mhord.Text, 4, 2)
               Xss = Mid(mhord.Text, 7, 2)
               Xhhh = Mid(mhorh.Text, 1, 2)
               Xmmh = Mid(mhorh.Text, 4, 2)
               Xssh = Mid(mhorh.Text, 7, 2)
               Xtoth = Xhhh - xhh
               Xtotm = Xmmh - xmm
               Xtots = Xssh - Xss
               If Xtots < 0 Then
                  Xtots = Xtots + 60
               End If
               If Xtotm < 0 Then
                  Xtotm = Xtotm + 60
                  Xtoth = Xtoth - 1
               End If
               If Xtoth > 9 Then
                  If Xtotm > 9 Then
                     If Xtots > 9 Then
                        Label6(1).Caption = Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  Else
                     If Xtots > 9 Then
                        Label6(1).Caption = Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  End If
               Else
                  If Xtotm > 9 Then
                     If Xtots > 9 Then
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  Else
                     If Xtots > 9 Then
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":" + Trim(str(Xtots))
                     Else
                        Label6(1).Caption = "0" + Trim(str(Xtoth)) + ":0" + Trim(str(Xtotm)) + ":0" + Trim(str(Xtots))
                     End If
                  End If
               End If
            End If
         End If
         data_graba.Recordset("nroseg") = Label6(1).Caption
         data_graba.Recordset.Update
         data_graba.Refresh
         data_busca.Refresh
         limpia
         habilita
      End If
   Else
      MsgBox "Ingrese dato de chofer"
   End If
Else
   MsgBox "Ingrese dato de móvil"
End If

Exit Sub

Quepasaalg:
           If Err.Number = 3155 Then
              MsgBox "Error al grabar, verifique datos!"
           Else
              MsgBox "Error al grabar, verifique médico, chofer, enfermero."
           End If
           
End Sub

Private Sub b_gratur_Click()
On Error GoTo Queesturno

If mf.Text <> "__/__/____" And mffin.Text <> "__/__/____" And labidtur.Caption <> "" And labu.Caption <> "" Then
   If XAlta = 1 Then
      data_grabatur.Recordset("cl_codigo") = labidtur.Caption
      data_grabatur.Recordset("cl_fnac") = mf.Text
      data_grabatur.Recordset("cl_ruc") = mhor.Text
      data_grabatur.Recordset("cl_fultmov") = mffin.Text
      data_grabatur.Recordset("cl_fax") = mhfin.Text
      data_grabatur.Recordset("cl_descpag") = labu.Caption
      data_grabatur.Recordset("cl_nom_sup") = Mid(labnomu.Caption, 1, 30)
      data_grabatur.Recordset("cl_atrasoa") = Check2.Value
      If t_obsturno.Text <> "" Then
         data_grabatur.Recordset("info_debit") = t_obsturno.Text
      End If
      If Check2.Value = 1 Then
         data_grabatur.Recordset("cl_fultpag") = Date
         data_grabatur.Recordset("cl_codconv") = Format(Time, "HH:mm")
      End If
      data_grabatur.Recordset.Update
      data_grabatur.Refresh
      XAlta = 0
      b_nuetur.Enabled = True
      b_cantur.Enabled = False
      b_inftur.Enabled = True
      b_busturn.Enabled = True
      dbturno.Enabled = True
      data_vertur.Refresh
   Else
      data_grabatur.RecordSource = "Select * from mant_sol_hc where cl_codigo =" & labidtur.Caption
      data_grabatur.Refresh
      If data_grabatur.Recordset.RecordCount > 0 Then
         If IsNull(data_grabatur.Recordset("cl_atrasoa")) = False Then
            If data_grabatur.Recordset("cl_atrasoa") <> Check2.Value Then
               data_grabatur.Recordset.Edit
               data_grabatur.Recordset("cl_atrasoa") = Check2.Value
               data_grabatur.Recordset.Update
            End If
         Else
            data_grabatur.Recordset.Edit
            data_grabatur.Recordset("cl_atrasoa") = Check2.Value
            data_grabatur.Recordset.Update
         End If
         If t_obsturno.Text <> "" Then
            If IsNull(data_grabatur.Recordset("info_debit")) = False Then
               If t_obsturno.Text <> data_grabatur.Recordset("info_debit") Then
                  data_grabatur.Recordset.Edit
                  data_grabatur.Recordset("info_debit") = t_obsturno.Text
                  data_grabatur.Recordset.Update
               End If
            Else
               data_grabatur.Recordset.Edit
               data_grabatur.Recordset("info_debit") = t_obsturno.Text
               data_grabatur.Recordset.Update
            End If
         Else
            If IsNull(data_grabatur.Recordset("info_debit")) = False Then
               data_grabatur.Recordset.Edit
               data_grabatur.Recordset("info_debit") = Null
               data_grabatur.Recordset.Update
            End If
         End If
         If IsNull(data_grabatur.Recordset("cl_fultpag")) = False Then
            If Check2.Value = 1 Then
               If Format(data_grabatur.Recordset("cl_fultpag"), "dd/mm/yyyy") <> Format(Date, "dd/mm/yyyy") Then
                  data_grabatur.Recordset.Edit
                  data_grabatur.Recordset("cl_fultpag") = Date
                  data_grabatur.Recordset.Update
               End If
               If IsNull(data_grabatur.Recordset("cl_codconv")) = False Then
                  If Format(data_grabatur.Recordset("cl_codconv"), "HH:mm") <> Format(Time, "HH:mm") Then
                     data_grabatur.Recordset.Edit
                     data_grabatur.Recordset("cl_codconv") = Format(Time, "HH:mm")
                     data_grabatur.Recordset.Update
                  End If
               Else
                  data_grabatur.Recordset.Edit
                  data_grabatur.Recordset("cl_codconv") = Format(Time, "HH:mm")
                  data_grabatur.Recordset.Update
               End If
            Else
               data_grabatur.Recordset.Edit
               data_grabatur.Recordset("cl_fultpag") = Null
               data_grabatur.Recordset("cl_codconv") = Null
               data_grabatur.Recordset.Update
            End If
         Else
            If Check2.Value = 1 Then
               data_grabatur.Recordset.Edit
               data_grabatur.Recordset("cl_fultpag") = Date
               data_grabatur.Recordset("cl_codconv") = Format(Time, "HH:mm")
               data_grabatur.Recordset.Update
            End If
         End If
         data_vertur.Refresh
      Else
         MsgBox "No se encuentra el registro, verifique en la lista"
      End If
   End If
Else
   MsgBox "No hay datos seleccionados o NUEVO REGISTRO, Verifique!!", vbExclamation
End If

Exit Sub

Queesturno:
           If Err.Number = 3155 Then
              MsgBox "No hay modificaciones para grabar"
           Else
              MsgBox "Error al grabar Nro:" & Err.Number & " " & Err.Description
           End If


End Sub

Private Sub b_imp_Click()
Dim Fdes, Fhas As String
Dim Xmas15 As String
Fdes = InputBox("Ingrese DESDE que FECHA en formato dd/mm/aaaa", "Informes")
Fhas = InputBox("Ingrese HASTA que FECHA en formato dd/mm/aaaa", "Informes")
data_inf2.DatabaseName = App.path & "\informes.mdb"
data_inf2.RecordSource = "Select * from infvtas"
data_inf2.Refresh
If data_inf2.Recordset.RecordCount > 0 Then
   data_inf2.Recordset.MoveFirst
   Do While Not data_inf2.Recordset.EOF
      data_inf2.Recordset.Delete
      data_inf2.Recordset.MoveNext
   Loop
End If
Dim Xquehace As String
Xquehace = MsgBox("Desea imprimir SOLO TIEMPOS MAYORES A 30 MINUTOS?", vbInformation + vbYesNo, "SAPP")

Xmas15 = MsgBox("Desea imprimir SOLO TIEMPOS DE ALMUERZO ASIGNADOS >A 15HORAS?", vbInformation + vbYesNo, "SAPP")

Frame1.Enabled = False
If Xquehace = vbYes Then
    If Fdes <> "" Then
       If Fhas <> "" Then
          data_inf(1).Connect = "odbc;dsn=" & Xconexrmt & ";"
          data_inf(1).RecordSource = "Select * from movil where nromov =" & 999 & " and fecha >=#" & Format(Fdes, "yyyy/mm/dd") & "# and fecha <=#" & Format(Fhas, "yyyy/mm/dd") & "# and nroseg >='" & "00:30:00" & "' order by nrolla"
          data_inf(1).Refresh
          If data_inf(1).Recordset.RecordCount > 0 Then
             data_inf(1).Recordset.MoveFirst
             Do While Not data_inf(1).Recordset.EOF
                data_inf2.Recordset.AddNew
                data_inf2.Recordset("fecha") = data_inf(1).Recordset("fecha")
                data_inf2.Recordset("tipo") = data_inf(1).Recordset("usuario")
                data_inf2.Recordset("realizada") = data_inf(1).Recordset("fecmod")
                data_inf2.Recordset("operador") = data_inf(1).Recordset("matricm")
                data_inf2.Recordset("nom_cli") = data_inf(1).Recordset("nroseg")
                data_inf2.Recordset("nom_prod") = data_inf(1).Recordset("medico")
                data_inf2.Recordset("cod_cli") = data_inf(1).Recordset("zona")
                data_inf2.Recordset("factura") = data_inf(1).Recordset("codmed")
                data_inf2.Recordset("nom_superv") = data_inf(1).Recordset("desbase")
                data_inf2.Recordset("base") = data_inf(1).Recordset("kmactu")
                If IsNull(data_inf(1).Recordset("pro_kms")) = False Then
                   data_med(1).RecordSource = "Select * from medicos where med_cod =" & data_inf(1).Recordset("pro_kms")
                   data_med(1).Refresh
                   If data_med(1).Recordset.RecordCount > 0 Then
                      data_inf2.Recordset("nom_medic") = data_med(1).Recordset("med_nombre")
                   End If
                End If
                data_inf2.Recordset.Update
                data_inf(1).Recordset.MoveNext
             Loop
             If Xmas15 = vbYes Then
                data_inf2.RecordSource = "Select * from infvtas where tipo <='" & "15:00" & "'"
                data_inf2.Refresh
                If data_inf2.Recordset.RecordCount > 0 Then
                   Do While Not data_inf2.Recordset.EOF
                      data_inf2.Recordset.Delete
                      data_inf2.Recordset.MoveNext
                   Loop
                End If
                data_inf2.RecordSource = "Select * from infvtas where nom_superv <>'" & "ALMUERZO" & "'"
                data_inf2.Refresh
                If data_inf2.Recordset.RecordCount > 0 Then
                   Do While Not data_inf2.Recordset.EOF
                      data_inf2.Recordset.Delete
                      data_inf2.Recordset.MoveNext
                   Loop
                End If
                data_inf2.RecordSource = "Select * from infvtas"
                data_inf2.Refresh
             Else
                data_inf2.RecordSource = "Select * from infvtas"
                data_inf2.Refresh
             End If
             cr1.ReportFileName = App.path & "\infaconmov.rpt"
             cr1.ReportTitle = "INFORME DEMORAS EN ACONDICIONAMIENTO MOVILES --FECHA: " & Fdes & " " & Fhas
             cr1.Action = 1
             
          End If
       End If
    End If
Else
    If Fdes <> "" Then
       If Fhas <> "" Then
          data_inf(1).Connect = "odbc;dsn=" & Xconexrmt & ";"
          data_inf(1).RecordSource = "Select * from movil where nromov =" & 999 & " and fecha >=#" & Format(Fdes, "yyyy/mm/dd") & "# and fecha <=#" & Format(Fhas, "yyyy/mm/dd") & "# order by nrolla"
          data_inf(1).Refresh
          If data_inf(1).Recordset.RecordCount > 0 Then
             data_inf(1).Recordset.MoveFirst
             Do While Not data_inf(1).Recordset.EOF
                data_inf2.Recordset.AddNew
                data_inf2.Recordset("fecha") = data_inf(1).Recordset("fecha")
                data_inf2.Recordset("tipo") = data_inf(1).Recordset("usuario")
                data_inf2.Recordset("realizada") = data_inf(1).Recordset("fecmod")
                data_inf2.Recordset("operador") = data_inf(1).Recordset("matricm")
                data_inf2.Recordset("nom_cli") = data_inf(1).Recordset("nroseg")
                data_inf2.Recordset("nom_prod") = data_inf(1).Recordset("medico")
                data_inf2.Recordset("cod_cli") = data_inf(1).Recordset("zona")
                data_inf2.Recordset("factura") = data_inf(1).Recordset("codmed")
                data_inf2.Recordset("nom_superv") = data_inf(1).Recordset("desbase")
                data_med(1).RecordSource = "Select * from medicos where med_cod =" & data_inf(1).Recordset("pro_kms")
                data_med(1).Refresh
                If data_med(1).Recordset.RecordCount > 0 Then
                   data_inf2.Recordset("nom_medic") = data_med(1).Recordset("med_nombre")
                End If
                data_inf2.Recordset.Update
                data_inf(1).Recordset.MoveNext
             Loop
             If Xmas15 = vbYes Then
                data_inf2.RecordSource = "Select * from infvtas where tipo <='" & "15:00" & "'"
                data_inf2.Refresh
                If data_inf2.Recordset.RecordCount > 0 Then
                   Do While Not data_inf2.Recordset.EOF
                      data_inf2.Recordset.Delete
                      data_inf2.Recordset.MoveNext
                   Loop
                End If
                data_inf2.RecordSource = "Select * from infvtas where nom_superv <>'" & "ALMUERZO" & "'"
                data_inf2.Refresh
                If data_inf2.Recordset.RecordCount > 0 Then
                   Do While Not data_inf2.Recordset.EOF
                      data_inf2.Recordset.Delete
                      data_inf2.Recordset.MoveNext
                   Loop
                End If
                data_inf2.RecordSource = "Select * from infvtas"
                data_inf2.Refresh
             Else
                data_inf2.RecordSource = "Select * from infvtas"
                data_inf2.Refresh
             End If
             cr1.ReportFileName = App.path & "\infaconmov.rpt"
             cr1.ReportTitle = "INFORME DEMORAS EN ACONDICIONAMIENTO MOVILES --FECHA: " & Fdes & " " & Fhas
             cr1.Action = 1
             
          End If
       End If
    End If
End If

Frame1.Enabled = True

MsgBox "Terminado"

End Sub

Private Sub b_in_Click()
Dim Xfd, Xfh, Xquehace, Xnom As String
Dim Xhord, Xhorh, Xmind, Xminh, Xtoth, Xtotn As Integer
Dim Xent As Integer
Dim Xfechs As Date

Xent = 0
On Error GoTo Vererrinfohs

Xquehace = MsgBox("Desea imprimir el Médico Seleccionado?", vbInformation + vbYesNo)
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_infhs.RecordSource = "infvtas"
data_infhs.Refresh

If Xquehace = vbYes Then
   If t_codmedm.Text <> "" Then
      Xfd = InputBox("Ingrese desde que fecha:")
      Xfh = InputBox("Ingrese hasta que fecha:")
      If Xfd <> "" And Xfh <> "" Then
         data_verhsme.RecordSource = "Select * from hc_archotro where hc_fecha >='" & Format(Xfd, "yyyy-mm-dd") & "' and hc_fecha <='" & Format(Xfh, "yyyy-mm-dd") & "' and hc_mat =" & t_codmedm.Text
         data_verhsme.Refresh
         If data_verhsme.Recordset.RecordCount > 0 Then
            frm_opsdesp.MousePointer = 11
            data_verhsme.Recordset.MoveFirst
            Do While Not data_verhsme.Recordset.EOF
               data_infhs.Recordset.AddNew
               data_infhs.Recordset("fecha") = data_verhsme.Recordset("hc_fecha")
               data_infhs.Recordset("hora") = data_verhsme.Recordset("hc_hora")
               data_infhs.Recordset("cod_cli") = data_verhsme.Recordset("hc_mat")
               data_infhs.Recordset("nom_prod") = Mid(data_verhsme.Recordset("hc_lugar"), 1, 50)
               If data_verhsme.Recordset("hc_nro") = 0 Then
                  data_infhs.Recordset("tipo") = "ENTRADA"
               Else
                  data_infhs.Recordset("tipo") = "SALIDA"
               End If
               data_hsmed2.RecordSource = "Select * from hc_viaae where id =" & data_verhsme.Recordset("id")
               data_hsmed2.Refresh
               If data_hsmed2.Recordset.RecordCount > 0 Then
                  data_infhs.Recordset("cod_prod") = data_hsmed2.Recordset("hc_cod")
               End If
               data_medhcc.RecordSource = "Select * from meta_tres where m_mat =" & data_verhsmed.Recordset("hc_mat")
               data_medhcc.Refresh
               If data_medhcc.Recordset.RecordCount > 0 Then
                  If IsNull(data_medhcc.Recordset("m_codmed")) = False Then
                     If data_medhcc.Recordset("m_codmed") = 1 Then
                        data_infhs.Recordset("convenio") = "COOP"
                     Else
                        data_infhs.Recordset("convenio") = "SAPP"
                     End If
                  Else
                     data_infhs.Recordset("convenio") = "SAPP"
                  End If
               Else
                  data_medhcc.RecordSource = "Select * from meta_tres where m_nrofrm ='" & Trim(str(data_verhsme.Recordset("hc_mat"))) & "'"
                  data_medhcc.Refresh
                  If data_medhcc.Recordset.RecordCount > 0 Then
                     If IsNull(data_medhcc.Recordset("m_codmed")) = False Then
                        If data_medhcc.Recordset("m_codmed") = 1 Then
                           data_infhs.Recordset("convenio") = "COOP"
                        Else
                           data_infhs.Recordset("convenio") = "SAPP"
                        End If
                     Else
                        data_infhs.Recordset("convenio") = "SAPP"
                     End If
                  End If
               End If
               data_infhs.Recordset.Update
               data_verhsme.Recordset.MoveNext
            Loop
            frm_opsdesp.MousePointer = 0
            
            MsgBox "Proceso terminado"
            crhs.ReportFileName = App.path & "\infhsmed.rpt"
            crhs.ReportTitle = "Informe de registros de ENTRADA/SALIDA Médicos no Dep. desde:" & Xfd & " hasta:" & Xfh
            crhs.Action = 1
         Else
            MsgBox "No hay datos"
         End If
      End If
   Else
      MsgBox "Ingrese médico a buscar"
   End If
Else
    Xfd = InputBox("Ingrese desde que fecha:")
    Xfh = InputBox("Ingrese hasta que fecha:")
    If Xfd <> "" And Xfh <> "" Then
       data_verhsme.RecordSource = "Select * from hc_archotro where hc_fecha >='" & Format(Xfd, "yyyy-mm-dd") & "' and hc_fecha <='" & Format(Xfh, "yyyy-mm-dd") & "'"
       data_verhsme.Refresh
       If data_verhsme.Recordset.RecordCount > 0 Then
          data_verhsme.Recordset.MoveFirst
            frm_opsdesp.MousePointer = 11
          
          Do While Not data_verhsme.Recordset.EOF
             data_infhs.Recordset.AddNew
             data_infhs.Recordset("fecha") = data_verhsme.Recordset("hc_fecha")
             data_infhs.Recordset("hora") = data_verhsme.Recordset("hc_hora")
             data_infhs.Recordset("cod_cli") = data_verhsme.Recordset("hc_mat")
             data_infhs.Recordset("nom_prod") = Mid(data_verhsme.Recordset("hc_lugar"), 1, 50)
             If data_verhsme.Recordset("hc_nro") = 0 Then
                data_infhs.Recordset("tipo") = "ENTRADA"
             Else
                data_infhs.Recordset("tipo") = "SALIDA"
             End If
             data_hsmed2.RecordSource = "Select * from hc_viaae where id =" & data_verhsme.Recordset("id")
             data_hsmed2.Refresh
             If data_hsmed2.Recordset.RecordCount > 0 Then
                data_infhs.Recordset("cod_prod") = data_hsmed2.Recordset("hc_cod")
             End If
             data_medhcc.RecordSource = "Select * from meta_tres where m_mat =" & data_verhsme.Recordset("hc_mat")
             data_medhcc.Refresh
             If data_medhcc.Recordset.RecordCount > 0 Then
                If IsNull(data_medhcc.Recordset("m_codmed")) = False Then
                   If data_medhcc.Recordset("m_codmed") = 1 Then
                      data_infhs.Recordset("convenio") = "COOP"
                   Else
                      data_infhs.Recordset("convenio") = "SAPP"
                   End If
                Else
                   data_infhs.Recordset("convenio") = "SAPP"
                End If
             Else
                data_medhcc.RecordSource = "Select * from meta_tres where m_nrofrm ='" & Trim(str(data_verhsme.Recordset("hc_mat"))) & "'"
                data_medhcc.Refresh
                If data_medhcc.Recordset.RecordCount > 0 Then
                   If IsNull(data_medhcc.Recordset("m_codmed")) = False Then
                      If data_medhcc.Recordset("m_codmed") = 1 Then
                         data_infhs.Recordset("convenio") = "COOP"
                      Else
                         data_infhs.Recordset("convenio") = "SAPP"
                      End If
                   Else
                      data_infhs.Recordset("convenio") = "SAPP"
                   End If
                End If
             End If
             data_infhs.Recordset.Update
             data_verhsme.Recordset.MoveNext
          Loop
          data_infhs.RecordSource = "Select * from infvtas order by nom_prod,fecha,tipo"
          data_infhs.Refresh
'          If data_infhs.Recordset.RecordCount > 0 Then
'             data_infhs.Recordset.MoveFirst
'             Xnom = data_infhs.Recordset("nom_prod")
'             Do While Not data_infhs.Recordset.EOF
'                If data_infhs.Recordset.EOF = False Then
'                    Do While UCase(Xnom) = UCase(data_infhs.Recordset("nom_prod"))
'                       If data_infhs.Recordset("tipo") = "ENTRADA" Then
'                          Xhord = Val(Mid(data_infhs.Recordset("hora"), 1, 2))
'                          Xmind = Val(Mid(data_infhs.Recordset("hora"), 4, 2))
'                          Xent = 1
'                          Xfechs = Format(data_infhs.Recordset("fecha"), "dd/mm/yyyy")
'                       Else
'                          If data_infhs.Recordset("tipo") = "SALIDA" Then
'                             If Xent = 1 Then
'                                Xhorh = Val(Mid(data_infhs.Recordset("hora"), 1, 2))
'                                Xminh = Val(Mid(data_infhs.Recordset("hora"), 4, 2))
'                                If Xmind > 50 Then
'                                   Xhord = Xhord + 1
'                                End If
'                                If Xminh > 50 Then
'                                   Xhorh = Xhorh + 1
'                                End If
'                                Xtoth = Xhord - Xhorh
'                                Xtotn = 0
'                                If Xtoth <= 0 Then
'                                   If Xtoth = 0 Then
'                                      If Format(Xfechs, "dd/mm/yyyy") = Format(data_infhs.Recordset("fecha"), "dd/mm/yyyy") Then
'                                         Xtoth = 0
'                                      Else
'                                         Xtoth = Xtoth + 24
'                                      End If
'                                   Else
'                                      Xtoth = Xhorh - Xhord
'                                   End If
'                                End If
'                                If Xhorh > 22 Then
'                                   Xtotn = 24 - Xhorh
'                                Else
'                                   If Xhorh >= 0 And Xhorh <= 6 Then
'                                      Xtotn = 6 - Xhorh
'                                      If Xtotn = 0 Then
'                                         Xtotn = 6
'                                      End If
'                                   End If
'                                End If
'
'                                If Xhord >= 23 Then
'                                   Xtotd = 24 - Xhorh
'                                Else
 '                                  If Xhord >= 0 And Xhord <= 6 Then
 '                                     Xtotn = 6 - Xhord
 '                                     If Xtotn = 0 Then
'                                         Xtotn = 6
'                                      End If
'                                   End If
'                                End If
'                                If Xtotn > 0 Then
'                                   Xtoth = Xtoth - Xtotn
'                                End If
'                                Xent = 0
'                             Else
'                                Xtoth = 0
'                                Xtotn = 0
 '                            End If
'                             Xent = 0
'                          Else
'                             Xtoth = 0
'                             Xtotn = 0
'                          End If
'                          data_infhs.Recordset.Edit
'                          data_infhs.Recordset("nro_flia") = Xtoth
'                          data_infhs.Recordset("nro_superv") = Xtotn
'                          data_infhs.Recordset.Update
'                       End If
'                       Xnom = data_infhs.Recordset("nom_prod")
'                       data_infhs.Recordset.MoveNext
'                       Xtoth = 0
'                       Xtotn = 0
'                    Loop
'                    Xnom = data_infhs.Recordset("nom_prod")
'                    Xent = 0
'                End If
'             Loop
'          End If
          frm_opsdesp.MousePointer = 0
          MsgBox "Proceso terminado"
'          data_infhs.Refresh
          crhs.ReportFileName = App.path & "\infhsmed2.rpt"
          crhs.ReportTitle = "Informe de registros de ENTRADA/SALIDA Médicos no Dep. desde:" & Xfd & " hasta:" & Xfh
          crhs.Action = 1
       Else
          frm_opsdesp.MousePointer = 0
          MsgBox "No hay datos"
       End If
    End If
End If

Exit Sub

Vererrinfohs:
             If Err.Number = 3155 Then
                MsgBox "Error al grabar"
             Else
                frm_opsdesp.MousePointer = 0
                MsgBox "Proceso terminado"
                data_infhs.Refresh
                crhs.ReportFileName = App.path & "\infhsmed2.rpt"
                crhs.ReportTitle = "Informe de registros de ENTRADA/SALIDA Médicos no Dep. desde:" & Xfd & " hasta:" & Xfh
                crhs.Action = 1
             End If

End Sub

Private Sub b_inftur_Click()
Dim xhh, Xhhh, xmm, Xmmh As Integer
Dim Xtoth, Xtotm, Xtotgral As Integer
Dim Xnumera As Integer
Dim Xelarch As String
Xelarch = "c:\planillas\inftr.pdf"
Dim Xresulta As String
Xresulta = Dir$(Xelarch)

''On Error GoTo Quepasainf

frm_opsdesp.MousePointer = 11
b_inftur.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

data_inftur.RecordSource = "infcli"
data_inftur.Refresh

MiBaseact.Execute "Delete * from inflla"
data_inftr.RecordSource = "inflla"
data_inftr.Refresh

Dim Xsionocerrado As String

If Xresulta <> "" Then
   Kill ("c:\planillas\inftr.pdf")
End If

'''''Command4_Click 'ccou covid

'''''Command6_Click 'smi

''''Command5_Click 'evang
If Check3.Value = 1 Then
   b_samc_Click
End If

data_llamtur.RecordSource = "Select * from mant_sol_hc where cl_codigo =" & labidtur.Caption
data_llamtur.Refresh
If data_llamtur.Recordset.RecordCount > 0 Then
   If IsNull(data_llamtur.Recordset("cl_atrasoa")) = False Then
      If data_llamtur.Recordset("cl_atrasoa") = 1 Then

        If mf.Text <> "__/__/____" And mffin.Text <> "__/__/____" And labidtur.Caption <> "" Then
           data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "'" & _
           " and categ not in ('UDEMM','CERSEM','CERADU','CERDGI','CERSAP','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS','CERANT','CERESS') and movilpas <>" & 99
           data_llamtur.Refresh
           If data_llamtur.Recordset.RecordCount > 0 Then
              data_llamtur.Recordset.MoveFirst
              Xnumera = 1
              Do While Not data_llamtur.Recordset.EOF
                 If IsNull(data_llamtur.Recordset("hora")) = False Then
                    xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                    xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                 Else
                    xhh = 0
                    xmm = 0
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                    If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                       Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                       Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                    Else
                       Xhhh = 0
                       Xmmh = 0
                    End If
                 Else
                    Xhhh = 0
                    Xmmh = 0
                 End If
                 If xhh = Xhhh Then
                    Xtoth = 0
                 Else
                    Xtoth = Xhhh - xhh
                    If Xhhh < xhh Then
                       Xtoth = Xtoth + 24
                    End If
                 End If
                 Xtotm = Xmmh - xmm
                 If Xmmh < xmm Then
                    Xtotm = Xtotm + 60
                    Xtoth = Xtoth - 1
                 End If
                 If Xtoth > 0 Then
                    Xtotgral = Xtoth * 60
                    Xtotgral = Xtotgral + Xtotm
                 Else
                    Xtotgral = Xtotm
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                    Xtotgral = 0
                    Xtotm = 0
                    Xtoth = 0
                 Else
                    If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                       Xtotgral = 0
                       Xtotm = 0
                       Xtoth = 0
                    End If
                 End If
                 If Xtotgral > 120 Then
                    data_inftur.Recordset.AddNew
                    data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                    data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                    data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                    data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                    data_inftur.Recordset("cl_codced") = 1
                    data_inftur.Recordset("cl_nomvend") = "VERDES"
                    data_inftur.Recordset("cl_atrasoa") = Xnumera
                    data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                    data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                    data_inftur.Recordset("cl_fecing") = mf.Text
                    data_inftur.Recordset.Update
                    Xnumera = Xnumera + 1
                 End If
                 Xtotgral = 0
                 data_llamtur.Recordset.MoveNext
              Loop
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "VERDES" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL VERDES: " & data_llamtur.Recordset.RecordCount
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 1
                 data_inftur.Recordset("cl_nomvend") = "VERDES"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           Else
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "VERDES" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 1
                 data_inftur.Recordset("cl_nomvend") = "VERDES"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           
           End If
           
        ' amarillos
           data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "A" & "' and movilpas <>" & 99
           data_llamtur.Refresh
           If data_llamtur.Recordset.RecordCount > 0 Then
              data_llamtur.Recordset.MoveFirst
              Xnumera = 1
              Do While Not data_llamtur.Recordset.EOF
                 If IsNull(data_llamtur.Recordset("hora")) = False Then
                    xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                    xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                 Else
                    xhh = 0
                    xmm = 0
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                    If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                       Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                       Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                    Else
                       Xhhh = 0
                       Xmmh = 0
                    End If
                 Else
                    Xhhh = 0
                    Xmmh = 0
                 End If
                 If xhh = Xhhh Then
                    Xtoth = 0
                 Else
                    Xtoth = Xhhh - xhh
                    If Xhhh < xhh Then
                       Xtoth = Xtoth + 24
                    End If
                 End If
                 Xtotm = Xmmh - xmm
                 If Xmmh < xmm Then
                    Xtotm = Xtotm + 60
                    Xtoth = Xtoth - 1
                 End If
                 If Xtoth > 0 Then
                    Xtotgral = Xtoth * 60
                    Xtotgral = Xtotgral + Xtotm
                 Else
                    Xtotgral = Xtotm
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                    Xtotgral = 0
                    Xtotm = 0
                    Xtoth = 0
                 Else
                    If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                       Xtotgral = 0
                       Xtotm = 0
                       Xtoth = 0
                    End If
                 End If
                 If Xtotgral > 30 Then
                    data_inftur.Recordset.AddNew
                    data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                    data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                    data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                    data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                    data_inftur.Recordset("cl_nomvend") = "AMARILLOS"
                    data_inftur.Recordset("cl_codced") = 2
                    data_inftur.Recordset("cl_atrasoa") = Xnumera
                    data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                    data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                    data_inftur.Recordset("cl_fecing") = mf.Text
                    data_inftur.Recordset.Update
                    Xnumera = Xnumera + 1
                 End If
                 data_llamtur.Recordset.MoveNext
                 Xtotgral = 0
              Loop
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "AMARILLOS" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL DE AMARILLOS: " & data_llamtur.Recordset.RecordCount
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 2
                 data_inftur.Recordset("cl_nomvend") = "AMARILLOS"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           Else
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "AMARILLOS" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 2
                 data_inftur.Recordset("cl_nomvend") = "AMARILLOS"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           End If
        
        'celestes
           data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "C" & "' and movilpas <>" & 99
           data_llamtur.Refresh
           If data_llamtur.Recordset.RecordCount > 0 Then
              data_llamtur.Recordset.MoveFirst
              Xnumera = 1
              Do While Not data_llamtur.Recordset.EOF
                 If IsNull(data_llamtur.Recordset("hora")) = False Then
                    xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                    xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                 Else
                    xhh = 0
                    xmm = 0
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                    If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                       Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                       Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                    Else
                       Xhhh = 0
                       Xmmh = 0
                    End If
                 Else
                    Xhhh = 0
                    Xmmh = 0
                 End If
                 If xhh = Xhhh Then
                    Xtoth = 0
                 Else
                    Xtoth = Xhhh - xhh
                    If Xhhh < xhh Then
                       Xtoth = Xtoth + 24
                    End If
                 End If
                 Xtotm = Xmmh - xmm
                 If Xmmh < xmm Then
                    Xtotm = Xtotm + 60
                    Xtoth = Xtoth - 1
                 End If
                 If Xtoth > 0 Then
                    Xtotgral = Xtoth * 60
                    Xtotgral = Xtotgral + Xtotm
                 Else
                    Xtotgral = Xtotm
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                    Xtotgral = 0
                    Xtotm = 0
                    Xtoth = 0
                 Else
                    If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                       Xtotgral = 0
                       Xtotm = 0
                       Xtoth = 0
                    End If
                 End If
                 If Xtotgral > 30 Then
                    data_inftur.Recordset.AddNew
                    data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                    data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                    data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                    data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                    data_inftur.Recordset("cl_nomvend") = "CELESTES"
                    data_inftur.Recordset("cl_codced") = 3
                    data_inftur.Recordset("cl_atrasoa") = Xnumera
                    data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                    data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                    data_inftur.Recordset("cl_fecing") = mf.Text
                    data_inftur.Recordset.Update
                    Xnumera = Xnumera + 1
                 End If
                 data_llamtur.Recordset.MoveNext
                 Xtotgral = 0
              Loop
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "CELESTES" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL DE CELESTES: " & data_llamtur.Recordset.RecordCount
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 3
                 data_inftur.Recordset("cl_nomvend") = "CELESTES"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           Else
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "CELESTES" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 3
                 data_inftur.Recordset("cl_nomvend") = "CELESTES"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           End If
        
        'rojos
           data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "R" & "' and movilpas <>" & 99
           data_llamtur.Refresh
           If data_llamtur.Recordset.RecordCount > 0 Then
              data_llamtur.Recordset.MoveFirst
              Xnumera = 1
              Do While Not data_llamtur.Recordset.EOF
                 If IsNull(data_llamtur.Recordset("hora")) = False Then
                    xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                    xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                 Else
                    xhh = 0
                    xmm = 0
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                    If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                       Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                       Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                    Else
                       Xhhh = 0
                       Xmmh = 0
                    End If
                 Else
                    Xhhh = 0
                    Xmmh = 0
                 End If
                 If xhh = Xhhh Then
                    Xtoth = 0
                 Else
                    Xtoth = Xhhh - xhh
                    If Xhhh < xhh Then
                       Xtoth = Xtoth + 24
                    End If
                 End If
                 Xtotm = Xmmh - xmm
                 
                 If Xmmh < xmm Then
                    Xtotm = Xtotm + 60
                    Xtoth = Xtoth - 1
                 End If
                 If Xtoth <= 0 Then
                    Xtotgral = Xtotm
                 Else
                    Xtotgral = Xtoth * 60
                    Xtotgral = Xtotgral + Xtotm
                 End If
                 If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                    Xtotgral = 0
                    Xtotm = 0
                    Xtoth = 0
                 Else
                    If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                       Xtotgral = 0
                       Xtotm = 0
                       Xtoth = 0
                    End If
                 End If
                 If Xtotgral > 15 Then
                    data_inftur.Recordset.AddNew
                    data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                    data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                    data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                    data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                    data_inftur.Recordset("cl_nomvend") = "ROJOS"
                    data_inftur.Recordset("cl_codced") = 4
                    data_inftur.Recordset("cl_atrasoa") = Xnumera
                    data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                    data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                    data_inftur.Recordset("cl_fecing") = mf.Text
                    data_inftur.Recordset.Update
                    Xnumera = Xnumera + 1
                 End If
                 data_llamtur.Recordset.MoveNext
                 Xtotgral = 0
              Loop
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "ROJOS" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL DE ROJOS: " & data_llamtur.Recordset.RecordCount
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 4
                 data_inftur.Recordset("cl_nomvend") = "ROJOS"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           Else
              data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "ROJOS" & "'"
              data_inftur.Refresh
              If data_inftur.Recordset.RecordCount > 0 Then
              Else
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = "---"
                 data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                 data_inftur.Recordset("cl_dpto") = "---"
                 data_inftur.Recordset("cl_codced") = 4
                 data_inftur.Recordset("cl_nomvend") = "ROJOS"
                 data_inftur.Recordset("cl_atrasoa") = 0
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
              End If
           End If
           Xnumera = 1
           data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "N" & "'"
           data_llamtur.Refresh
           If data_llamtur.Recordset.RecordCount > 0 Then
              data_llamtur.Recordset.MoveFirst
              Do While Not data_llamtur.Recordset.EOF
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                 data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                 data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                 data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                 data_inftur.Recordset("cl_codced") = 5
                 data_inftur.Recordset("cl_nomvend") = "FALLECIDOS"
                 data_inftur.Recordset("cl_atrasoa") = Xnumera
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
                 data_llamtur.Recordset.MoveNext
                 Xnumera = Xnumera + 1
              Loop
           End If
           
           data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and colormot ='" & "N" & "'"
           data_llamtur.Refresh
           If data_llamtur.Recordset.RecordCount > 0 Then
              data_llamtur.Recordset.MoveFirst
              Do While Not data_llamtur.Recordset.EOF
                 data_inftur.Recordset.AddNew
                 data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                 data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                 data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                 data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                 data_inftur.Recordset("cl_nomvend") = "FALLECIDOS"
                 data_inftur.Recordset("cl_codced") = 5
                 data_inftur.Recordset("cl_atrasoa") = Xnumera
                 data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                 data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
                 data_inftur.Recordset("cl_fecing") = mf.Text
                 data_inftur.Recordset.Update
                 data_llamtur.Recordset.MoveNext
                 Xnumera = Xnumera + 1
              Loop
           End If
           data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "FALLECIDOS" & "'"
           data_inftur.Refresh
           If data_inftur.Recordset.RecordCount > 0 Then
           Else
              data_inftur.Recordset.AddNew
              data_inftur.Recordset("cl_codconv") = "---"
              data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
              data_inftur.Recordset("cl_dpto") = "---"
              data_inftur.Recordset("cl_codced") = 5
              data_inftur.Recordset("cl_nomvend") = "FALLECIDOS"
              data_inftur.Recordset("cl_atrasoa") = 0
              data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
              data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
              data_inftur.Recordset("cl_fecing") = mf.Text
              data_inftur.Recordset.Update
           End If
           
           Xnumera = 1
           data_llamtur.RecordSource = "Select * from mant_sol_hc where cl_codigo =" & labidtur.Caption
           data_llamtur.Refresh
           If data_llamtur.Recordset.RecordCount > 0 Then
              data_inftur.Recordset.AddNew
              data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("cl_ruc")
              data_inftur.Recordset("info_debit") = data_llamtur.Recordset("info_debit")
              data_inftur.Recordset("cl_telefon") = Mid(data_llamtur.Recordset("cl_descpag"), 1, 20)
              data_inftur.Recordset("cl_nomvend") = "OBSERVACIONES"
              data_inftur.Recordset("cl_codced") = 6
              data_inftur.Recordset("cl_atrasoa") = Xnumera
              data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
              data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text
              data_inftur.Recordset("cl_fecing") = mf.Text
              data_inftur.Recordset.Update
           End If
           If Check3.Value = 1 Then
              crturno.Destination = crptToPrinter
           End If
           
           data_inftur.Refresh
           crturno.DiscardSavedData = True
           crturno.ReportFileName = App.path & "\infdespacha.rpt"
           crturno.Action = 1
           MsgBox "Archivo Generado"
           If Check3.Value = 1 Then
              frm_opsdesp.MousePointer = 99
              frm_opsdesp.Enabled = False
              Timer1.Enabled = True
           Else
              frm_opsdesp.MousePointer = 0
              MsgBox "Proceso terminado"
           End If
        
        End If
      Else
        Xsionocerrado = MsgBox("El TURNO ESTA SIN CERRAR, DESEA VER EL INFORME?", vbInformation + vbYesNo)
        If Xsionocerrado = vbYes Then
                
            If mf.Text <> "__/__/____" And mffin.Text <> "__/__/____" And labidtur.Caption <> "" Then
               data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "'" & _
               " and categ not in ('UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and movilpas <>" & 99
               data_llamtur.Refresh
               If data_llamtur.Recordset.RecordCount > 0 Then
                  data_llamtur.Recordset.MoveFirst
                  Xnumera = 1
                  Do While Not data_llamtur.Recordset.EOF
                     If IsNull(data_llamtur.Recordset("hora")) = False Then
                        xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                        xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                     Else
                        xhh = 0
                        xmm = 0
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                        If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                           Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                           Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                        Else
                           Xhhh = 0
                           Xmmh = 0
                        End If
                     Else
                        Xhhh = 0
                        Xmmh = 0
                     End If
                     If xhh = Xhhh Then
                        Xtoth = 0
                     Else
                        Xtoth = Xhhh - xhh
                        If Xhhh < xhh Then
                           Xtoth = Xtoth + 24
                        End If
                     End If
                     Xtotm = Xmmh - xmm
                     If Xmmh < xmm Then
                        Xtotm = Xtotm + 60
                        Xtoth = Xtoth - 1
                     End If
                     If Xtoth > 0 Then
                        Xtotgral = Xtoth * 60
                        Xtotgral = Xtotgral + Xtotm
                     Else
                        Xtotgral = Xtotm
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                        Xtotgral = 0
                        Xtotm = 0
                        Xtoth = 0
                     Else
                        If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                           Xtotgral = 0
                           Xtotm = 0
                           Xtoth = 0
                        End If
                     End If
                     If Xtotgral > 120 Then
                        data_inftur.Recordset.AddNew
                        data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                        data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                        data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                        data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                        data_inftur.Recordset("cl_codced") = 1
                        data_inftur.Recordset("cl_nomvend") = "VERDES"
                        data_inftur.Recordset("cl_atrasoa") = Xnumera
                        data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                        data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                        data_inftur.Recordset("cl_fecing") = mf.Text
                        data_inftur.Recordset.Update
                        Xnumera = Xnumera + 1
                     End If
                     Xtotgral = 0
                     data_llamtur.Recordset.MoveNext
                  Loop
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "VERDES" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL VERDES: " & data_llamtur.Recordset.RecordCount
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 1
                     data_inftur.Recordset("cl_nomvend") = "VERDES"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               Else
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "VERDES" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 1
                     data_inftur.Recordset("cl_nomvend") = "VERDES"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               
               End If
               
            ' amarillos
               data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "A" & "' and movilpas <>" & 99
               data_llamtur.Refresh
               If data_llamtur.Recordset.RecordCount > 0 Then
                  data_llamtur.Recordset.MoveFirst
                  Xnumera = 1
                  Do While Not data_llamtur.Recordset.EOF
                     If IsNull(data_llamtur.Recordset("hora")) = False Then
                        xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                        xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                     Else
                        xhh = 0
                        xmm = 0
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                        If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                           Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                           Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                        Else
                           Xhhh = 0
                           Xmmh = 0
                        End If
                     Else
                        Xhhh = 0
                        Xmmh = 0
                     End If
                     If xhh = Xhhh Then
                        Xtoth = 0
                     Else
                        Xtoth = Xhhh - xhh
                        If Xhhh < xhh Then
                           Xtoth = Xtoth + 24
                        End If
                     End If
                     Xtotm = Xmmh - xmm
                     If Xmmh < xmm Then
                        Xtotm = Xtotm + 60
                        Xtoth = Xtoth - 1
                     End If
                     If Xtoth > 0 Then
                        Xtotgral = Xtoth * 60
                        Xtotgral = Xtotgral + Xtotm
                     Else
                        Xtotgral = Xtotm
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                        Xtotgral = 0
                        Xtotm = 0
                        Xtoth = 0
                     Else
                        If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                           Xtotgral = 0
                           Xtotm = 0
                           Xtoth = 0
                        End If
                     End If
                     
                     If Xtotgral > 30 Then
                        data_inftur.Recordset.AddNew
                        data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                        data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                        data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                        data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                        data_inftur.Recordset("cl_nomvend") = "AMARILLOS"
                        data_inftur.Recordset("cl_codced") = 2
                        data_inftur.Recordset("cl_atrasoa") = Xnumera
                        data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                        data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                        data_inftur.Recordset("cl_fecing") = mf.Text
                        data_inftur.Recordset.Update
                        Xnumera = Xnumera + 1
                     End If
                     data_llamtur.Recordset.MoveNext
                     Xtotgral = 0
                  Loop
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "AMARILLOS" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL DE AMARILLOS: " & data_llamtur.Recordset.RecordCount
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 2
                     data_inftur.Recordset("cl_nomvend") = "AMARILLOS"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               Else
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "AMARILLOS" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 2
                     data_inftur.Recordset("cl_nomvend") = "AMARILLOS"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               End If
            
            'celestes
               data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "C" & "' and movilpas <>" & 99
               data_llamtur.Refresh
               If data_llamtur.Recordset.RecordCount > 0 Then
                  data_llamtur.Recordset.MoveFirst
                  Xnumera = 1
                  Do While Not data_llamtur.Recordset.EOF
                     If IsNull(data_llamtur.Recordset("hora")) = False Then
                        xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                        xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                     Else
                        xhh = 0
                        xmm = 0
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                        If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                           Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                           Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                        Else
                           Xhhh = 0
                           Xmmh = 0
                        End If
                     Else
                        Xhhh = 0
                        Xmmh = 0
                     End If
                     If xhh = Xhhh Then
                        Xtoth = 0
                     Else
                        Xtoth = Xhhh - xhh
                        If Xhhh < xhh Then
                           Xtoth = Xtoth + 24
                        End If
                     End If
                     Xtotm = Xmmh - xmm
                     If Xmmh < xmm Then
                        Xtotm = Xtotm + 60
                        Xtoth = Xtoth - 1
                     End If
                     If Xtoth > 0 Then
                        Xtotgral = Xtoth * 60
                        Xtotgral = Xtotgral + Xtotm
                     Else
                        Xtotgral = Xtotm
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                        Xtotgral = 0
                        Xtotm = 0
                        Xtoth = 0
                     Else
                        If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                           Xtotgral = 0
                           Xtotm = 0
                           Xtoth = 0
                        End If
                     End If
                     
                     If Xtotgral > 30 Then
                        data_inftur.Recordset.AddNew
                        data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                        data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                        data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                        data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                        data_inftur.Recordset("cl_nomvend") = "CELESTES"
                        data_inftur.Recordset("cl_codced") = 3
                        data_inftur.Recordset("cl_atrasoa") = Xnumera
                        data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                        data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                        data_inftur.Recordset("cl_fecing") = mf.Text
                        data_inftur.Recordset.Update
                        Xnumera = Xnumera + 1
                     End If
                     data_llamtur.Recordset.MoveNext
                     Xtotgral = 0
                  Loop
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "CELESTES" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL DE CELESTES: " & data_llamtur.Recordset.RecordCount
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 3
                     data_inftur.Recordset("cl_nomvend") = "CELESTES"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               Else
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "CELESTES" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 3
                     data_inftur.Recordset("cl_nomvend") = "CELESTES"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               End If
            
            'rojos
               data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "R" & "' and movilpas <>" & 99
               data_llamtur.Refresh
               If data_llamtur.Recordset.RecordCount > 0 Then
                  data_llamtur.Recordset.MoveFirst
                  Xnumera = 1
                  Do While Not data_llamtur.Recordset.EOF
                     If IsNull(data_llamtur.Recordset("hora")) = False Then
                        xhh = Val(Mid(data_llamtur.Recordset("hora"), 1, 2))
                        xmm = Val(Mid(data_llamtur.Recordset("hora"), 4, 2))
                     Else
                        xhh = 0
                        xmm = 0
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = False Then
                        If Trim(data_llamtur.Recordset("hor_llega")) <> "" Then
                           Xhhh = Val(Mid(data_llamtur.Recordset("hor_llega"), 1, 2))
                           Xmmh = Val(Mid(data_llamtur.Recordset("hor_llega"), 4, 2))
                        Else
                           Xhhh = 0
                           Xmmh = 0
                        End If
                     Else
                        Xhhh = 0
                        Xmmh = 0
                     End If
                     If xhh = Xhhh Then
                        Xtoth = 0
                     Else
                        Xtoth = Xhhh - xhh
                        If Xhhh < xhh Then
                           Xtoth = Xtoth + 24
                        End If
                     End If
                     Xtotm = Xmmh - xmm
                     If Xmmh < xmm Then
                        Xtotm = Xtotm + 60
                        Xtoth = Xtoth - 1
                     End If
                     If Xtoth > 0 Then
                        Xtotgral = Xtoth * 60
                        Xtotgral = Xtotgral + Xtotm
                     Else
                        Xtotgral = Xtotm
                     End If
                     If IsNull(data_llamtur.Recordset("hor_llega")) = True Then
                        Xtotgral = 0
                        Xtotm = 0
                        Xtoth = 0
                     Else
                        If Trim(data_llamtur.Recordset("hor_llega")) = "" Then
                           Xtotgral = 0
                           Xtotm = 0
                           Xtoth = 0
                        End If
                     End If
                     
                     If Xtotgral > 15 Then
                        
                        data_inftur.Recordset.AddNew
                        data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                        data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                        data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                        data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                        data_inftur.Recordset("cl_nomvend") = "ROJOS"
                        data_inftur.Recordset("cl_codced") = 4
                        data_inftur.Recordset("cl_atrasoa") = Xnumera
                        data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                        data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                        data_inftur.Recordset("cl_fecing") = mf.Text
                        data_inftur.Recordset.Update
                        Xnumera = Xnumera + 1
                     End If
                     data_llamtur.Recordset.MoveNext
                     Xtotgral = 0
                  Loop
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "ROJOS" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD - TOTAL DE ROJOS: " & data_llamtur.Recordset.RecordCount
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 4
                     data_inftur.Recordset("cl_nomvend") = "ROJOS"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               Else
                  data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "ROJOS" & "'"
                  data_inftur.Refresh
                  If data_inftur.Recordset.RecordCount > 0 Then
                  Else
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = "---"
                     data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                     data_inftur.Recordset("cl_dpto") = "---"
                     data_inftur.Recordset("cl_codced") = 4
                     data_inftur.Recordset("cl_nomvend") = "ROJOS"
                     data_inftur.Recordset("cl_atrasoa") = 0
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                  End If
               End If
               Xnumera = 1
               data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and codmot ='" & "N" & "'"
               data_llamtur.Refresh
               If data_llamtur.Recordset.RecordCount > 0 Then
                  data_llamtur.Recordset.MoveFirst
                  Do While Not data_llamtur.Recordset.EOF
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                     data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                     data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                     data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                     data_inftur.Recordset("cl_codced") = 5
                     data_inftur.Recordset("cl_nomvend") = "FALLECIDOS"
                     data_inftur.Recordset("cl_atrasoa") = Xnumera
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                     data_llamtur.Recordset.MoveNext
                     Xnumera = Xnumera + 1
                  Loop
               End If
               
               data_llamtur.RecordSource = "Select * from llamado where timdes ='" & labu.Caption & "' and fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mffin.Text, "yyyy/mm/dd") & "# and colormot ='" & "N" & "'"
               data_llamtur.Refresh
               If data_llamtur.Recordset.RecordCount > 0 Then
                  data_llamtur.Recordset.MoveFirst
                  Do While Not data_llamtur.Recordset.EOF
                     data_inftur.Recordset.AddNew
                     data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("hora")
                     data_inftur.Recordset("cl_apellid") = Mid(data_llamtur.Recordset("nombre"), 1, 60)
                     data_inftur.Recordset("cl_telefon") = data_llamtur.Recordset("codmot")
                     data_inftur.Recordset("cl_dpto") = Trim(str(Xtotgral))
                     data_inftur.Recordset("cl_nomvend") = "FALLECIDOS"
                     data_inftur.Recordset("cl_codced") = 5
                     data_inftur.Recordset("cl_atrasoa") = Xnumera
                     data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                     data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                     data_inftur.Recordset("cl_fecing") = mf.Text
                     data_inftur.Recordset.Update
                     data_llamtur.Recordset.MoveNext
                     Xnumera = Xnumera + 1
                  Loop
               End If
               data_inftur.RecordSource = "Select * from infcli where cl_nomvend ='" & "FALLECIDOS" & "'"
               data_inftur.Refresh
               If data_inftur.Recordset.RecordCount > 0 Then
               Else
                  data_inftur.Recordset.AddNew
                  data_inftur.Recordset("cl_codconv") = "---"
                  data_inftur.Recordset("cl_apellid") = "SIN NOVEDAD"
                  data_inftur.Recordset("cl_dpto") = "---"
                  data_inftur.Recordset("cl_codced") = 5
                  data_inftur.Recordset("cl_nomvend") = "FALLECIDOS"
                  data_inftur.Recordset("cl_atrasoa") = 0
                  data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                  data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                  data_inftur.Recordset("cl_fecing") = mf.Text
                  data_inftur.Recordset.Update
               End If
               
               Xnumera = 1
               data_llamtur.RecordSource = "Select * from mant_sol_hc where cl_codigo =" & labidtur.Caption
               data_llamtur.Refresh
               If data_llamtur.Recordset.RecordCount > 0 Then
                  data_inftur.Recordset.AddNew
                  data_inftur.Recordset("cl_codconv") = data_llamtur.Recordset("cl_ruc")
                  data_inftur.Recordset("info_debit") = data_llamtur.Recordset("info_debit")
                  data_inftur.Recordset("cl_telefon") = Mid(data_llamtur.Recordset("cl_descpag"), 1, 20)
                  data_inftur.Recordset("cl_nomvend") = "OBSERVACIONES"
                  data_inftur.Recordset("cl_codced") = 6
                  data_inftur.Recordset("cl_atrasoa") = Xnumera
                  data_inftur.Recordset("cl_email") = Mid(labnomu.Caption, 1, 30)
                  data_inftur.Recordset("cl_nomcobr") = mhor.Text & "-" & mhfin.Text & "**SIN CERRAR**"
                  data_inftur.Recordset("cl_fecing") = mf.Text
                  data_inftur.Recordset.Update
               End If
               frm_opsdesp.MousePointer = 0
               If Check3.Value = 1 Then
                  
                  crturno.Destination = crptToPrinter
               End If
               data_inftur.Refresh
               crturno.DiscardSavedData = True
               crturno.ReportFileName = App.path & "\infdespacha.rpt"
               crturno.Action = 1
               
               If Check3.Value = 1 Then
                  frm_opsdesp.MousePointer = 99
                  frm_opsdesp.Enabled = False
                  Timer1.Enabled = True
               Else
                  frm_opsdesp.MousePointer = 0
                  MsgBox "Proceso terminado"
               End If
                              
            End If
        End If
      End If
   Else
      MsgBox "No hay registro de turno"
   End If
Else
   MsgBox "No hay registro de turno"
End If
frm_opsdesp.MousePointer = 0
b_inftur.Enabled = True

'Exit Sub

'Quepasainf:
'           If Err.Number = 3155 Then
'              MsgBox "Error al grabar"
'           Else
'              MsgBox "Error en el informe, cierre el programa y vuelva a intentar"
'           End If
           

End Sub

Private Sub b_modif_Click(index As Integer)

deshabmov
XAlta = 2
Frame1.Enabled = True
txt_nro.SetFocus
Command1(0).Enabled = True
Command2(0).Enabled = True


End Sub

Private Sub b_modiff_Click(index As Integer)
XAlta = 0
deshab
limpia

'data_graba.Recordset.FindFirst "nrolla =" & data_busca.Recordset("nrolla")
data_graba.RecordSource = "Select * from movil where nromov =" & 999 & " and nrolla =" & data_busca.Recordset("nrolla")
data_graba.Refresh
'data_graba.RecordSource = "Select * from movil where nromov =" & 999 & " order by nrolla"
If data_graba.Recordset.RecordCount > 0 Then
   iguala
   md.SetFocus
Else
   habilita
End If

End Sub

Private Sub b_nuetur_Click()
Dim Xhor2, Xhor As Date

mf.Text = "__/__/____"
mhor.Text = "__:__"
mffin.Text = "__/__/____"
mhfin.Text = "__:__"
labu.Caption = ""
t_obsturno.Text = ""
labnomu.Caption = ""
labidtur.Caption = ""

data_grabatur.RecordSource = "Select * from mant_sol_hc order by cl_codigo DESC"
data_grabatur.Refresh
If data_grabatur.Recordset.RecordCount > 0 Then
   labidtur.Caption = data_grabatur.Recordset("cl_codigo") + 1
Else
   labidtur.Caption = 1
End If
data_grabatur.Recordset.AddNew

mf.Text = Format(Date, "dd/mm/yyyy")
mhor.Text = Format(Time, "HH:mm")
mffin.Text = Format(Date, "dd/mm/yyyy")
Xhor2 = Time
Xhor = DateAdd("h", 6, Xhor2)

mhfin.Text = Format(Xhor, "HH:mm")
labu.Caption = WElusuario
labnomu.Caption = Welnombredu
t_obsturno.SetFocus
XAlta = 1
dbturno.Enabled = False
b_nuetur.Enabled = False
b_cantur.Enabled = True
b_inftur.Enabled = False
b_busturn.Enabled = False


End Sub

Private Sub b_samc_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xarchexel23 As New Excel.Worksheet

Dim Xlin, XCol As Integer
Dim Xtotreg, XtotregV, XtotregA, XtotregR, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim XcumpleV, XcumpleA, XcumpleR, XcumpleVl, XcumpleAl, XcumpleRl As Double
Dim Xdias As Integer
Dim Cuentactrol As Integer
Dim Xfdesde, Xfhasta As Date
Dim Xlabrir3 As New Excel.Application
Dim Letra As String
Dim Xhh1, Xmm1, Xhh2, Xmm2, Xhh3, Xmm3, XtotminR, Xths, Xtot3, Xtot3tot As Integer
Dim Xdemr, Xdema, Xdemv As Integer
Dim Xdemrl, Xdemal, Xdemvl As Integer
Dim XtotmmVerde, XtiempoVerde As Double

XtotmmVerde = 0
XtiempoVerde = 0

Xdemr = 0
Xdema = 0
Xdemv = 0
Xdemrl = 0
Xdemal = 0
Xdemvl = 0
XcumpleV = 0
XcumpleA = 0
XcumpleR = 0
XcumpleVl = 0
XcumpleAl = 0
XcumpleRl = 0

Xtot3tot = 0
Xhh1 = 0
Xmm1 = 0
Xhh2 = 0
Xmm2 = 0
Xhh3 = 0
Xmm3 = 0
XtotminR = 0
Xths = 0
Xtot3 = 0
XtotregA = 0
XtotregV = 0
XtotregR = 0
data_ctrf.DatabaseName = App.path & "\ctradm.mdb"
data_ctrf.RecordSource = "ctradm"
data_ctrf.Refresh
data_llam.Connect = "odbc;dsn=sappnew;"
data_covid.Connect = "odbc;dsn=sappnew;"
data_covidT.Connect = "odbc;dsn=sappnew;"

Xdias = DateDiff("d", data_ctrf.Recordset("fecha"), Date)

If Xdias >= 2 Then
   Xfdesde = data_ctrf.Recordset("fecha") + 1
   
   Xfhasta = Date - 1
   
   data_covidT.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfhasta, "yyyy/mm/dd") & "# and cancela is null and movilpas in (415,315,215) and fec_rea is not null and categ not in ('SAMCB') order by fecha,hora"
   data_covidT.Refresh
   
   If data_covidT.Recordset.RecordCount > 0 Then
      
      data_covidT.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfhasta, "yyyy/mm/dd") & "# and cancela is null and movilpas in (415,315,215) and fec_rea is not null order by fecha,hora"
      data_covidT.Refresh
      data_covidT.Recordset.MoveFirst
      
      Textofecha = Trim(str(Day(Xfhasta))) & Trim(str(Month(Xfhasta))) & Trim(str(Year(Xfhasta)))
      Xtotreg = 0
      Cuentactrol = 0
      Xlin = 1
      XCol = 1
      Xtotreg = 0
      Xsub = 0
'''desde aca
'''      Set Xobjexel22 = New Excel.Application
'''      Set Xlibexel22 = Xobjexel22.Workbooks.Add
'''      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      
'''hasta acá
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      
      Xlibexel22.SaveAs ("C:\planillas\Samc-llamados" & Trim(Textofecha) & ".xls")
      Xarchtex = "C:\planillas\Samc-Llamados" & Trim(Textofecha) & ".xls"
           
      Set Xarchexel23 = Xlibexel22.Worksheets.Add
      Xarchexel23.Name = Trim("Todos")
      Xlin = 1
      XCol = 1
      Xarchexel23.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel23.Range("A1", "C3").Font.Size = 16
      Xarchexel23.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
    
      Xarchexel23.Cells(Xlin, XCol) = "LLAMADOS TOTALES DE MÓVILES DE SAMC DESDE: " & Xfdesde & " HASTA: " & Xfhasta
      
      
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
                
      Xarchexel23.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel23.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel23.Cells(Xlin, XCol) = "MATRICULA"
      XCol = XCol + 1
      Xarchexel23.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
      Xarchexel23.Cells(Xlin, XCol) = "NOMBRES"
      XCol = XCol + 1
      Xarchexel23.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "CATEG."
      XCol = XCol + 1
      Xarchexel23.Range("D" & Trim(str(Xlin))).ColumnWidth = 25
      Xarchexel23.Cells(Xlin, XCol) = "FECHA_REC"
      XCol = XCol + 1
      Xarchexel23.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "HORA_REC"
      XCol = XCol + 1
      Xarchexel23.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "MOVIL"
      XCol = XCol + 1
      Xarchexel23.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel23.Cells(Xlin, XCol) = "ZONA"
      XCol = XCol + 1
      Xarchexel23.Range("H" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel23.Cells(Xlin, XCol) = "MOTIVO LLAMADO"
      XCol = XCol + 1
      Xarchexel23.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "CLAVE"
      XCol = XCol + 1
      Xarchexel23.Range("J" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "HORA_PASADO"
      XCol = XCol + 1
      Xarchexel23.Range("K" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "HORA_SALIDA"
      XCol = XCol + 1
      Xarchexel23.Range("L" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "HORA_LLEGA"
      XCol = XCol + 1
      Xarchexel23.Range("M" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel23.Cells(Xlin, XCol) = "HORA_REA"
      XCol = XCol + 1
      Xarchexel23.Range("N" & Trim(str(Xlin))).ColumnWidth = 17
      Xarchexel23.Cells(Xlin, XCol) = "RESPUESTA SAMC"
      XCol = XCol + 1
      Xarchexel23.Range("O" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel23.Cells(Xlin, XCol) = "TRASLADO"
        
      Xlin = Xlin + 1
      XCol = 1
      Do While Not data_covidT.Recordset.EOF
         If IsNull(data_covidT.Recordset("matric")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("matric")
         Else
            Xarchexel23.Cells(Xlin, XCol) = 0
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("nombre")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("nombre")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "NN"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("categ")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("categ")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("fecha")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = "'" & Format(data_covidT.Recordset("fecha"), "dd/mm/yyyy")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("hora")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("hora")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("movilpas")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("movilpas")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("motmov")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("motmov")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("obsmot")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("obsmot")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("descol")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("descol")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("horpas")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("horpas")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("horsali")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("horsali")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("hor_llega")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("hor_llega")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("hor_rea")) = False Then
            Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("hor_rea")
         Else
            Xarchexel23.Cells(Xlin, XCol) = "S/D"
         End If
         If IsNull(data_covidT.Recordset("hor_llega")) = False Then
            Xhh1 = Val(Mid(data_covidT.Recordset("hor_llega"), 1, 2))
            Xmm1 = Val(Mid(data_covidT.Recordset("hor_llega"), 4, 2))
         End If
         If IsNull(data_covidT.Recordset("horpas")) = False Then
            Xhh2 = Val(Mid(data_covidT.Recordset("horpas"), 1, 2))
            Xmm2 = Val(Mid(data_covidT.Recordset("horpas"), 4, 2))
         End If
         If Xhh1 = Xhh2 Then
            Xtmm = Xmm1 - Xmm2
         Else
            Xths = Xhh1 - Xhh2
            If IsNull(data_covidT.Recordset("fec_llega")) = False Then
               If data_covidT.Recordset("fec_llega") > data_covidT.Recordset("fecha") Then
                  Xths = Xths + 24
               End If
            End If
            Xtmm = Xmm1 - Xmm2 + 60
            If Xths = 2 Then
               Xtmm = Xtmm + 60
            End If
            If Xths = 3 Then
               Xtmm = Xtmm + 120
            End If
            If Xths = 4 Then
               Xtmm = Xtmm + 180
            End If
         End If
         XCol = XCol + 1
         Xarchexel23.Cells(Xlin, XCol) = Xtmm
'''-indicadores
         Xhh1 = 0
         Xmm1 = 0
         If IsNull(data_covidT.Recordset("horsali")) = False Then
            Xhh1 = Val(Mid(data_covidT.Recordset("horsali"), 1, 2))
            Xmm1 = Val(Mid(data_covidT.Recordset("horsali"), 4, 2))
         End If
         If IsNull(data_covidT.Recordset("horpas")) = False Then
            Xhh2 = Val(Mid(data_covidT.Recordset("horpas"), 1, 2))
            Xmm2 = Val(Mid(data_covidT.Recordset("horpas"), 4, 2))
         End If
         If Xhh1 = Xhh2 Then
            Xtmm = Xmm1 - Xmm2
         Else
            Xths = Xhh1 - Xhh2
            If IsNull(data_covidT.Recordset("fecsali")) = False Then
               If data_covidT.Recordset("fecsali") > data_covidT.Recordset("fecpas") Then
                  Xths = Xths + 24
               End If
            End If
            Xtmm = Xmm1 - Xmm2 + 60
            If Xths = 2 Then
               Xtmm = Xtmm + 60
            End If
            If Xths = 3 Then
               Xtmm = Xtmm + 120
            End If
            If Xths = 4 Then
               Xtmm = Xtmm + 180
            End If
         End If
         If data_covidT.Recordset("descol") = "VERDE" Or data_covidT.Recordset("descol") = "AZUL" Then
            If Xtmm > 5 Then
               Xdemv = Xdemv + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "AMARILLO" Or data_covidT.Recordset("descol") = "CELESTE" Then
            If Xtmm > 5 Then
               Xdema = Xdema + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "ROJO" Then
            If Xtmm > 3 Then
               Xdemr = Xdemr + 1
            End If
         End If
'''--demoras en llegada
         Xhh1 = 0
         Xmm1 = 0
         If IsNull(data_covidT.Recordset("hor_llega")) = False Then
            Xhh1 = Val(Mid(data_covidT.Recordset("hor_llega"), 1, 2))
            Xmm1 = Val(Mid(data_covidT.Recordset("hor_llega"), 4, 2))
         End If
         If IsNull(data_covidT.Recordset("horsali")) = False Then
            Xhh2 = Val(Mid(data_covidT.Recordset("horsali"), 1, 2))
            Xmm2 = Val(Mid(data_covidT.Recordset("horsali"), 4, 2))
         End If
         If Xhh1 = Xhh2 Then
            Xtmm = Xmm1 - Xmm2
         Else
            Xths = Xhh1 - Xhh2
            If IsNull(data_covidT.Recordset("fec_llega")) = False Then
               If data_covidT.Recordset("fec_llega") > data_covidT.Recordset("fecsali") Then
                  Xths = Xths + 24
               End If
            End If
            Xtmm = Xmm1 - Xmm2 + 60
            If Xths = 2 Then
               Xtmm = Xtmm + 60
            End If
            If Xths = 3 Then
               Xtmm = Xtmm + 120
            End If
            If Xths = 4 Then
               Xtmm = Xtmm + 180
            End If
         End If
         If data_covidT.Recordset("descol") = "VERDE" Or data_covidT.Recordset("descol") = "AZUL" Then
            XtotregV = XtotregV + 1
            If Xtmm > 180 Then
               Xdemvl = Xdemvl + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "AMARILLO" Or data_covidT.Recordset("descol") = "CELESTE" Then
            XtotregA = XtotregA + 1
            If Xtmm > 30 Then
               Xdemal = Xdemal + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "ROJO" Then
            XtotregR = XtotregR + 1
            If Xtmm > 15 Then
               Xdemrl = Xdemrl + 1
            End If
         End If
         ''' demoras verdes en domicilio
         If data_covidT.Recordset("descol") = "VERDE" Then
            Xhh1 = 0
            Xmm1 = 0
            If IsNull(data_covidT.Recordset("hor_rea")) = False Then
               Xhh1 = Val(Mid(data_covidT.Recordset("hor_rea"), 1, 2))
               Xmm1 = Val(Mid(data_covidT.Recordset("hor_rea"), 4, 2))
            End If
            If IsNull(data_covidT.Recordset("hor_llega")) = False Then
               Xhh2 = Val(Mid(data_covidT.Recordset("hor_llega"), 1, 2))
               Xmm2 = Val(Mid(data_covidT.Recordset("hor_llega"), 4, 2))
            End If
            If Xhh1 = Xhh2 Then
               Xtmm = Xmm1 - Xmm2
            Else
               Xths = Xhh1 - Xhh2
               If IsNull(data_covidT.Recordset("fec_llega")) = False Then
                  If data_covidT.Recordset("fec_llega") > data_covidT.Recordset("fecsali") Then
                     Xths = Xths + 24
                  End If
               End If
               Xtmm = Xmm1 - Xmm2 + 60
               If Xths = 2 Then
                  Xtmm = Xtmm + 60
               End If
               If Xths = 3 Then
                  Xtmm = Xtmm + 120
               End If
               If Xths = 4 Then
                  Xtmm = Xtmm + 180
               End If
            End If
            XtotmmVerde = XtotmmVerde + Xtmm
         
         End If
'''-fin indicadores
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("lugar")) = False Then
            If Trim(data_covidT.Recordset("lugar")) <> "" Then
               Xarchexel23.Cells(Xlin, XCol) = data_covidT.Recordset("lugar")
            End If
         End If
         Xlin = Xlin + 1
         XCol = 1
         Xtotreg = Xtotreg + 1
         data_covidT.Recordset.MoveNext
      Loop
      Xarchexel23.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS:" & Xtotreg
            
      XtotmmVerde = 0
      XtiempoVerde = 0
      Xdemr = 0
      Xdema = 0
      Xdemv = 0
      Xdemrl = 0
      Xdemal = 0
      Xdemvl = 0
      XcumpleV = 0
      XcumpleA = 0
      XcumpleR = 0
      XcumpleVl = 0
      XcumpleAl = 0
      XcumpleRl = 0
      Xtot3tot = 0
      Xhh1 = 0
      Xmm1 = 0
      Xhh2 = 0
      Xmm2 = 0
      Xhh3 = 0
      Xmm3 = 0
      XtotminR = 0
      Xths = 0
      Xtot3 = 0
      XtotregA = 0
      XtotregV = 0
      XtotregR = 0
      Xtotreg = 0
      Cuentactrol = 0
      Xlin = 1
      XCol = 1
      Xtotreg = 0
      Xsub = 0
      Xlin = 1
      XCol = 1
      data_covidT.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfhasta, "yyyy/mm/dd") & "# and cancela is null and movilpas in (415,315,215) and fec_rea is not null and categ not in ('SAMCB') order by fecha,hora"
      data_covidT.Refresh
      
      Set Xarchexel22 = Xlibexel22.Worksheets.Add

      Xarchexel22.Name = Trim("SAMC")
      Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
    
      Xarchexel22.Cells(Xlin, XCol) = "LLAMADOS MÓVILES DE SAMC DESDE: " & Xfdesde & " HASTA: " & Xfhasta
        
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
                
      Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "MATRICULA"
      XCol = XCol + 1
      Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
      XCol = XCol + 1
      Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "CATEG."
      XCol = XCol + 1
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 25
      Xarchexel22.Cells(Xlin, XCol) = "FECHA_REC"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "HORA_REC"
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "MOVIL"
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "ZONA"
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "MOTIVO LLAMADO"
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "CLAVE"
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "HORA_PASADO"
      XCol = XCol + 1
      Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "HORA_SALIDA"
      XCol = XCol + 1
      Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "HORA_LLEGA"
      XCol = XCol + 1
      Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "HORA_REA"
      XCol = XCol + 1
      Xarchexel22.Range("N" & Trim(str(Xlin))).ColumnWidth = 17
      Xarchexel22.Cells(Xlin, XCol) = "RESPUESTA SAMC"
      XCol = XCol + 1
      Xarchexel22.Range("O" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "TRASLADO"
        
      Xlin = Xlin + 1
      XCol = 1
      data_covidT.Recordset.MoveFirst
      Do While Not data_covidT.Recordset.EOF
         If IsNull(data_covidT.Recordset("matric")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("matric")
         Else
            Xarchexel22.Cells(Xlin, XCol) = 0
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("nombre")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("nombre")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "NN"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("categ")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("categ")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("fecha")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_covidT.Recordset("fecha"), "dd/mm/yyyy")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("hora")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("hora")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("movilpas")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("movilpas")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("motmov")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("motmov")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("obsmot")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("obsmot")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("descol")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("descol")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("horpas")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("horpas")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("horsali")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("horsali")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("hor_llega")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("hor_llega")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("hor_rea")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("hor_rea")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         If IsNull(data_covidT.Recordset("hor_llega")) = False Then
            Xhh1 = Val(Mid(data_covidT.Recordset("hor_llega"), 1, 2))
            Xmm1 = Val(Mid(data_covidT.Recordset("hor_llega"), 4, 2))
         End If
         If IsNull(data_covidT.Recordset("horpas")) = False Then
            Xhh2 = Val(Mid(data_covidT.Recordset("horpas"), 1, 2))
            Xmm2 = Val(Mid(data_covidT.Recordset("horpas"), 4, 2))
         End If
         If Xhh1 = Xhh2 Then
            Xtmm = Xmm1 - Xmm2
         Else
            Xths = Xhh1 - Xhh2
            If IsNull(data_covidT.Recordset("fec_llega")) = False Then
               If data_covidT.Recordset("fec_llega") > data_covidT.Recordset("fecha") Then
                  Xths = Xths + 24
               End If
            End If
            Xtmm = Xmm1 - Xmm2 + 60
            If Xths = 2 Then
               Xtmm = Xtmm + 60
            End If
            If Xths = 3 Then
               Xtmm = Xtmm + 120
            End If
            If Xths = 4 Then
               Xtmm = Xtmm + 180
            End If
         End If
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Xtmm
'''-indicadores
         Xhh1 = 0
         Xmm1 = 0
         If IsNull(data_covidT.Recordset("horsali")) = False Then
            Xhh1 = Val(Mid(data_covidT.Recordset("horsali"), 1, 2))
            Xmm1 = Val(Mid(data_covidT.Recordset("horsali"), 4, 2))
         End If
         If IsNull(data_covidT.Recordset("horpas")) = False Then
            Xhh2 = Val(Mid(data_covidT.Recordset("horpas"), 1, 2))
            Xmm2 = Val(Mid(data_covidT.Recordset("horpas"), 4, 2))
         End If
         If Xhh1 = Xhh2 Then
            Xtmm = Xmm1 - Xmm2
         Else
            Xths = Xhh1 - Xhh2
            If IsNull(data_covidT.Recordset("fecsali")) = False Then
               If data_covidT.Recordset("fecsali") > data_covidT.Recordset("fecpas") Then
                  Xths = Xths + 24
               End If
            End If
            Xtmm = Xmm1 - Xmm2 + 60
            If Xths = 2 Then
               Xtmm = Xtmm + 60
            End If
            If Xths = 3 Then
               Xtmm = Xtmm + 120
            End If
            If Xths = 4 Then
               Xtmm = Xtmm + 180
            End If
         End If
         If data_covidT.Recordset("descol") = "VERDE" Or data_covidT.Recordset("descol") = "AZUL" Then
            If Xtmm > 5 Then
               Xdemv = Xdemv + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "AMARILLO" Or data_covidT.Recordset("descol") = "CELESTE" Then
            If Xtmm > 5 Then
               Xdema = Xdema + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "ROJO" Then
            If Xtmm > 3 Then
               Xdemr = Xdemr + 1
            End If
         End If
'''--demoras en llegada
         Xhh1 = 0
         Xmm1 = 0
         If IsNull(data_covidT.Recordset("hor_llega")) = False Then
            Xhh1 = Val(Mid(data_covidT.Recordset("hor_llega"), 1, 2))
            Xmm1 = Val(Mid(data_covidT.Recordset("hor_llega"), 4, 2))
         End If
         If IsNull(data_covidT.Recordset("horsali")) = False Then
            Xhh2 = Val(Mid(data_covidT.Recordset("horsali"), 1, 2))
            Xmm2 = Val(Mid(data_covidT.Recordset("horsali"), 4, 2))
         End If
         If Xhh1 = Xhh2 Then
            Xtmm = Xmm1 - Xmm2
         Else
            Xths = Xhh1 - Xhh2
            If IsNull(data_covidT.Recordset("fec_llega")) = False Then
               If data_covidT.Recordset("fec_llega") > data_covidT.Recordset("fecsali") Then
                  Xths = Xths + 24
               End If
            End If
            Xtmm = Xmm1 - Xmm2 + 60
            If Xths = 2 Then
               Xtmm = Xtmm + 60
            End If
            If Xths = 3 Then
               Xtmm = Xtmm + 120
            End If
            If Xths = 4 Then
               Xtmm = Xtmm + 180
            End If
         End If
         If data_covidT.Recordset("descol") = "VERDE" Or data_covidT.Recordset("descol") = "AZUL" Then
            XtotregV = XtotregV + 1
            If Xtmm > 180 Then
               Xdemvl = Xdemvl + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "AMARILLO" Or data_covidT.Recordset("descol") = "CELESTE" Then
            XtotregA = XtotregA + 1
            If Xtmm > 30 Then
               Xdemal = Xdemal + 1
            End If
         End If
         If data_covidT.Recordset("descol") = "ROJO" Then
            XtotregR = XtotregR + 1
            If Xtmm > 15 Then
               Xdemrl = Xdemrl + 1
            End If
         End If
         ''' demoras verdes en domicilio
         If data_covidT.Recordset("descol") = "VERDE" Then
            Xhh1 = 0
            Xmm1 = 0
            If IsNull(data_covidT.Recordset("hor_rea")) = False Then
               Xhh1 = Val(Mid(data_covidT.Recordset("hor_rea"), 1, 2))
               Xmm1 = Val(Mid(data_covidT.Recordset("hor_rea"), 4, 2))
            End If
            If IsNull(data_covidT.Recordset("hor_llega")) = False Then
               Xhh2 = Val(Mid(data_covidT.Recordset("hor_llega"), 1, 2))
               Xmm2 = Val(Mid(data_covidT.Recordset("hor_llega"), 4, 2))
            End If
            If Xhh1 = Xhh2 Then
               Xtmm = Xmm1 - Xmm2
            Else
               Xths = Xhh1 - Xhh2
               If IsNull(data_covidT.Recordset("fec_llega")) = False Then
                  If data_covidT.Recordset("fec_llega") > data_covidT.Recordset("fecsali") Then
                     Xths = Xths + 24
                  End If
               End If
               Xtmm = Xmm1 - Xmm2 + 60
               If Xths = 2 Then
                  Xtmm = Xtmm + 60
               End If
               If Xths = 3 Then
                  Xtmm = Xtmm + 120
               End If
               If Xths = 4 Then
                  Xtmm = Xtmm + 180
               End If
            End If
            XtotmmVerde = XtotmmVerde + Xtmm
         
         End If
'''-fin indicadores
         XCol = XCol + 1
         If IsNull(data_covidT.Recordset("lugar")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("lugar")
         End If
         Xlin = Xlin + 1
         XCol = 1
         Xtotreg = Xtotreg + 1
         data_covidT.Recordset.MoveNext
      Loop
      Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS:" & Xtotreg
      Xlin = Xlin + 1
      Xlin = Xlin + 1
      XCol = 2
      Xarchexel22.Cells(Xlin, XCol) = "RESUMEN DE CUMPLIMIENTO INDICADORES DESDE: " & Xfdesde & " HASTA: " & Xfhasta
      Xlin = Xlin + 1
      Xarchexel22.Range("B" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(65, 105, 225)
      
      Xarchexel22.Cells(Xlin, XCol) = "DEMORAS EN SALIDA: "
      XCol = 3
      Xarchexel22.Cells(Xlin, XCol) = "No Cumpl."
      XCol = 4
      Xarchexel22.Cells(Xlin, XCol) = "DEMORAS EN LLEGADA: "
      XCol = 5
      Xarchexel22.Cells(Xlin, XCol) = "No Cumpl."
      If Xdemr = 0 Then
         XcumpleR = 100
      Else
         XcumpleR = Xdemr / XtotregR * 100
      End If
      If Xdemrl = 0 Then
         XcumpleRl = 100
      Else
         XcumpleRl = Xdemrl / XtotregR * 100
      End If
      
      If Xdema = 0 Then
         XcumpleA = 100
      Else
         XcumpleA = Xdema / XtotregA * 100
      End If
      If Xdemal = 0 Then
         XcumpleAl = 100
      Else
         XcumpleAl = Xdemal / XtotregA * 100
      End If
      
      If Xdemv = 0 Then
         XcumpleV = 100
      Else
         XcumpleV = Xdemv / XtotregV * 100
      End If
      If Xdemvl = 0 Then
         XcumpleVl = 100
      Else
         XcumpleVl = Xdemvl / XtotregV * 100
      End If
''rojo
      Xlin = Xlin + 1
      XCol = 2
      Xarchexel22.Range("B" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(255, 0, 0)
      If Val(XcumpleR) >= 100 Then
         Xarchexel22.Cells(Xlin, XCol) = "ROJOS >3': " & Trim(str(Val(XcumpleR))) & "% CUMPLIDO"
      Else
         Xarchexel22.Cells(Xlin, XCol) = "ROJOS >3': " & Format(XcumpleR, "Standard") & "% CUMPLIDO"
      End If
      XCol = 3
      Xarchexel22.Cells(Xlin, XCol) = Xdemr
      XCol = 4
      If Val(XcumpleRl) >= 100 Then
         Xarchexel22.Cells(Xlin, XCol) = "ROJOS >15': " & Trim(str(Val(XcumpleRl))) & "% CUMPLIDO"
      Else
         Xarchexel22.Cells(Xlin, XCol) = "ROJOS >15': " & Format(XcumpleRl, "Standard") & "% CUMPLIDO"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Xdemrl
'amarillos
      Xlin = Xlin + 1
      XCol = 2
      Xarchexel22.Range("B" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(255, 255, 0)
      If Val(XcumpleA) >= 100 Then
         Xarchexel22.Cells(Xlin, XCol) = "AMARILLO >5': " & Trim(str(Val(XcumpleA))) & "% CUMPLIDO"
      Else
         Xarchexel22.Cells(Xlin, XCol) = "AMARILLO >5': " & Format(XcumpleA, "Standard") & "% CUMPLIDO"
      End If
      XCol = 3
      Xarchexel22.Cells(Xlin, XCol) = Xdema
      XCol = 4
      If Val(XcumpleAl) >= 100 Then
         Xarchexel22.Cells(Xlin, XCol) = "AMARILLO >30': " & Trim(str(Val(XcumpleAl))) & "% CUMPLIDO"
      Else
         Xarchexel22.Cells(Xlin, XCol) = "AMARILLO >30': " & Format(XcumpleAl, "Standard") & "% CUMPLIDO"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Xdemal
''verdes
      Xlin = Xlin + 1
      XCol = 2
      Xarchexel22.Range("B" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(0, 128, 0)
      If Val(XcumpleV) >= 100 Then
         Xarchexel22.Cells(Xlin, XCol) = "VERDES >5': " & Trim(str(Val(XcumpleV))) & "% CUMPLIDO"
      Else
         Xarchexel22.Cells(Xlin, XCol) = "VERDES >5': " & Format(XcumpleV, "Standard") & "% CUMPLIDO"
      End If
      XCol = 3
      Xarchexel22.Cells(Xlin, XCol) = Xdemv
      XCol = 4
      If Val(XcumpleVl) >= 100 Then
         Xarchexel22.Cells(Xlin, XCol) = "VERDES >3hs': " & Trim(str(Val(XcumpleVl))) & "% CUMPLIDO"
      Else
         Xarchexel22.Cells(Xlin, XCol) = "VERDES >3hs': " & Format(XcumpleVl, "Standard") & "% CUMPLIDO"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Xdemvl
      
      Xlin = Xlin + 1
      XCol = 2
      Xarchexel22.Range("B" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(0, 128, 0)
      Xarchexel22.Cells(Xlin, XCol) = "PROMEDIO DEMORAS EN DOMICILIO VERDES"
      XCol = 3
      XtiempoVerde = XtotmmVerde / XtotregV
      Xarchexel22.Cells(Xlin, XCol) = Format(XtiempoVerde, "Standard")
      
      XCol = 1
      Xlin = Xlin + 1
      Xarchexel22.Cells(Xlin, XCol) = "DETALLE DE NEGATIVAS SAMC DESDE: " & Xfdesde & " HASTA: " & Xfhasta
      Xlin = Xlin + 1
      Xarchexel22.Cells(Xlin, XCol) = "FECHA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "HORA_NEGATIVA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "MOVIL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "MOTIVO"
      XCol = XCol + 1
      
      data_covidT.RecordSource = "select * from resplla where trasla in (215,315,415) and fecpas >=#" & Format(Xfdesde, "yyyy/mm/dd") & "# and fecpas <=#" & Format(Xfhasta, "yyyy/mm/dd") & "#"
      data_covidT.Refresh
      Xtotreg = 0
      Xlin = Xlin + 1
      XCol = 1
      
      If data_covidT.Recordset.RecordCount > 0 Then
         data_covidT.Recordset.MoveFirst
         Do While Not data_covidT.Recordset.EOF
            If IsNull(data_covidT.Recordset("fecpas")) = False Then
               Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_covidT.Recordset("fecpas"), "dd/mm/yyyy")
            Else
               Xarchexel22.Cells(Xlin, XCol) = "S/D"
            End If
            XCol = XCol + 1
            If IsNull(data_covidT.Recordset("horpas")) = False Then
               Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("horpas")
            Else
               Xarchexel22.Cells(Xlin, XCol) = "S/D"
            End If
            XCol = XCol + 1
            If IsNull(data_covidT.Recordset("trasla")) = False Then
               Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("trasla")
               XCol = XCol + 1
               If IsNull(data_covidT.Recordset("realiza")) = False Then
                  If data_covidT.Recordset("realiza") = 0 Then
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Disponibilidad"
                  Else
                     If data_covidT.Recordset("realiza") = 1 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Otros"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                     End If
                  End If
               Else
                  Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
               End If
            Else
               Xarchexel22.Cells(Xlin, XCol) = "S/D"
            End If
            Xlin = Xlin + 1
            XCol = 1
            Xtotreg = Xtotreg + 1
            data_covidT.Recordset.MoveNext
         Loop
      Else
         Xarchexel22.Cells(Xlin, XCol) = "Sin Negativas SAMC "
         Xlin = Xlin + 1
         XCol = 1
      End If
      Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
      
      
      Xlibexel22.Save
      Xlibexel22.Close
      Xobjexel22.Quit
         
      Dim MenCorreo As String
      Dim oMail As Class1
      Set oMail = New Class1
         With oMail
''             .servidor = "smtp.gmail.com"
             .servidor = "smtp.office365.com"
             .puerto = 25
             .UseAuntentificacion = True
             .ssl = True
             .Usuario = "despacho@sapp.com.uy"
             .PassWord = "Salinas1987"
             .Asunto = "Informe Llamados SAMC " & Xfdesde & "-- " & Xfhasta
             .de = "despacho@sapp.com.uy"
             .para = "jefedepartamentoti@sapp.com.uy; subdirectortecnico@sapp.com.uy; jefedespacho@sapp.com.uy; directoratecnica@sapp.com.uy; cd.samc@adinet.com.uy"
'''             .para = "jdanfer@gmail.com; sappjorge@hotmail.com"

'''cd.samc@adinet.com.uy
   '         .para = "sappjorge@hotmail.com; despachosapp@hotmail.com; sappsusanadominguez@hotmail.com; sappdirecciontecnica@hotmail.com; sappenrique@hotmail.com"
             .Adjunto = Xarchtex
             .Mensaje = "Informes de llamados realizados por SAMC."
             .Enviar_Backup ' manda el mail
         End With
         Set oMail = Nothing
   
         data_ctrf.Recordset.Edit
         data_ctrf.Recordset("fecha") = Date - 1
         data_ctrf.Recordset.Update
   
   Else
'      MsgBox "No hay registros seguimiento COVID-19 CCOU"
   End If
Else
'   MsgBox "No hay registros seguimiento COVID-19 CCOU"
End If

End Sub

Private Sub b_vermedcmt_Click()
Consultar_medicosCMT

End Sub

Private Sub bbusca_Click()
txt_bcob.Enabled = True
'DBGrid112(1).Enabled = True
'txt_bcob.SetFocus

End Sub

Private Sub bcance_Click()
If XAcnv = 1 Then
   data_cob.Recordset.CancelUpdate
   igualcob
   XAcnv = 0
   desh
Else
   igualcob
   XAcnv = 0
   desh
End If
bgraba.Enabled = False
bcance.Enabled = False
bmodif.Enabled = True
bbusca.Enabled = True
bimp.Enabled = True
bnuevo.Enabled = True

End Sub

Private Sub bgraba_Click()
Dim Ced As String

If txt_nrocob.Text <> "" And txt_nomcob.Text <> "" And t_ced.Text <> "" And t_codced.Text <> "" Then
   If txt_nrocob.Text <> 0 Then
         If XAcnv = 1 Then
            data_cob.Recordset("med_cod") = txt_nrocob.Text
            data_cob.Recordset("med_nombre") = txt_nomcob.Text
            data_cob.Recordset("med_esp") = txt_espec.Text
            data_cob.Recordset("med_socnom") = txt_tel.Text
            data_cob.Recordset.Update
            
            data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
            data_medhc.Refresh
            If data_medhc.Recordset.RecordCount > 0 Then
               data_medhc.Recordset.Edit
               data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
               data_medhc.Recordset("m_codmed") = Check4.Value
               data_medhc.Recordset.Update
            Else
               data_par.Recordset.Edit
               data_par.Recordset("nro_reg") = data_par.Recordset("nro_reg") + 1
               data_par.Recordset.Update
               data_par.Refresh
               data_medhc.Recordset.AddNew
               data_medhc.Recordset("id") = data_par.Recordset("nro_reg")
               data_medhc.Recordset("m_fecha") = Date
               data_medhc.Recordset("m_mat") = txt_nrocob.Text
               data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
               data_medhc.Recordset("m_codmed") = Check4.Value
               data_medhc.Recordset.Update
            End If
            
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            desh
         Else
            If data_cob.Recordset("med_nombre") <> txt_nomcob.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("med_nombre") = txt_nomcob.Text
               data_cob.Recordset.Update
            End If
            If data_cob.Recordset("med_esp") <> txt_espec.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("med_esp") = txt_espec.Text
               data_cob.Recordset.Update
            End If
            If data_cob.Recordset("med_socnom") <> txt_tel.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("med_socnom") = txt_tel.Text
               data_cob.Recordset.Update
            End If
            data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
            data_medhc.Refresh
            If data_medhc.Recordset.RecordCount > 0 Then
               Ced = Trim(t_ced.Text) & Trim(t_codced.Text)
               If data_medhc.Recordset("m_nrofrm") = Ced Then
               Else
                  data_medhc.Recordset.Edit
                  data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
                  data_medhc.Recordset("m_codmed") = Check4.Value
                  data_medhc.Recordset.Update
               End If
            Else
               data_par.Recordset.Edit
               data_par.Recordset("nro_reg") = data_par.Recordset("nro_reg") + 1
               data_par.Recordset.Update
               data_par.Refresh
               data_medhc.Recordset.AddNew
               data_medhc.Recordset("id") = data_par.Recordset("nro_reg")
               data_medhc.Recordset("m_fecha") = Date
               data_medhc.Recordset("m_mat") = txt_nrocob.Text
               data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
               data_medhc.Recordset("m_codmed") = Check4.Value
               data_medhc.Recordset.Update
            End If
            
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            txt_nrocob.Enabled = True
            desh
         End If
   Else
      MsgBox "No ingresó médico", vbCritical, "Médicos"
      txt_nrocob.SetFocus
   End If
Else
   MsgBox "Faltan ingresar datos, verifique!!", vbCritical, "Médicos"
   txt_nrocob.SetFocus
End If

End Sub

Private Sub bmodif_Click()
If XWeltipoU = "ADMINISTRADOR" Then
    XAcnv = 0
    hab
    txt_nrocob.Enabled = False
    txt_nomcob.SetFocus
    bgraba.Enabled = True
    bcance.Enabled = True
    bmodif.Enabled = False
    bbusca.Enabled = False
    bimp.Enabled = False
    bnuevo.Enabled = False
Else
    MsgBox "No se puede modificar datos de médico.", vbInformation
End If

End Sub

Private Sub bnuevo_Click()
XAcnv = 1
hab
txt_nrocob.Text = ""
txt_nomcob.Text = ""
txt_tel.Text = ""
txt_espec.Text = ""
t_ced.Text = ""
t_codced.Text = ""
Check4.Value = 0
txt_nomcob.SetFocus
txt_nrocob.Enabled = False
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
Data1.RecordSource = "Select * from medicos order by med_cod"
Data1.Refresh
Data1.Recordset.MoveLast
txt_nrocob.Text = Data1.Recordset("med_cod") + 1
data_cob.Recordset.AddNew

End Sub

Private Sub cboes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nromov.SetFocus
End If

End Sub

Private Sub Check1_Click()
Xfmov = Date - 365


If Check1.Value = 1 Then
   If t_movb.Text <> "" Then
      data_busca.RecordSource = "Select * from movil where nromov =" & 999 & " and codmed =" & t_movb.Text & " order by fecha DESC,nrolla"
      data_busca.Refresh
   Else
      data_busca.RecordSource = "Select * from movil where nromov =" & 999 & " order by fecha DESC,nrolla"
      data_busca.Refresh
   End If
Else
   Xfmov = Date - 2
   data_busca.RecordSource = "Select * from movil where nromov =" & 999 & " and fecha >=#" & Format(Xfmov, "yyyy/mm/dd") & "# order by fecha DESC,nrolla"
   data_busca.Refresh
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If b_gr.Enabled = True Then
      b_gr.SetFocus
   End If
End If

End Sub

Private Sub Combo2_LostFocus()
If Combo2.Text <> "" Then
   data_m.RecordSource = "Select * from medicos where med_nombre ='" & Combo2.Text & "'"
   data_m.Refresh
   If data_m.Recordset.RecordCount > 0 Then
      t_codmedm.Text = data_m.Recordset("med_cod")
      Combo2.Text = data_m.Recordset("med_nombre")
   Else
      MsgBox "Verifique nombre del médico", vbCritical
   End If
End If

      
End Sub

Private Sub Command1_Click(index As Integer)

frm_chofer.Show vbModal

End Sub

Private Sub Command10_Click()
Dim Deseaborrar As String
Deseaborrar = MsgBox("Seguro que desea borrar toda la lista?", vbInformation + vbYesNo, "Médicos")
If Deseaborrar = vbYes Then
   data_listamedicos.RecordSource = "select * from medicos_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
   data_listamedicos.Refresh
   If data_listamedicos.Recordset.RecordCount > 0 Then
      data_listamedicos.Recordset.MoveFirst
      Do While Not data_listamedicos.Recordset.EOF
         data_listamedicos.Recordset.Delete
         data_listamedicos.Recordset.MoveNext
      Loop
      MsgBox "Terminado."
      List4.Clear
      
   End If
End If

End Sub


Private Sub Command111_Click(index As Integer)
Xdeb = 22

frm_chofer.Show vbModal

End Sub

Private Sub Command2_Click(index As Integer)

frm_enferm.Show vbModal

End Sub

Private Sub Command222_Click(index As Integer)
frm_chofer.Show vbModal

End Sub




Private Sub Command3_Click()
Dim Xdiferenmsp As Integer
Dim Xlafecmsp As Date
Dim Xresultamsp As String

Xresultamsp = Dir$("c:\planillas\inftr.pdf")

Xdiferenmsp = 0
t_envmsp.Text = 0
Text1.Text = ""
Xdiferenmsp = DateDiff("d", data_fecmsp.Recordset("fecha"), Date)
If Xdiferenmsp > 1 Then
   t_envmsp.Text = Xdiferenmsp
   Xlafecmsp = data_fecmsp.Recordset("fecha") + 1
   Text1.Text = Xlafecmsp
   data_llamtur.RecordSource = "Select * from llamado where categ ='" & "MSP" & "' and fecha =#" & Format(Xlafecmsp, "yyyy/mm/dd") & "#"
   data_llamtur.Refresh
   If data_llamtur.Recordset.RecordCount > 0 Then
      data_llamtur.Recordset.MoveFirst
      Do While Not data_llamtur.Recordset.EOF
         If IsNull(data_llamtur.Recordset("cancela")) = True Then
            data_inftr.Recordset.AddNew
            data_inftr.Recordset("fecha") = data_llamtur.Recordset("fecha")
            data_inftr.Recordset("hora") = data_llamtur.Recordset("hora")
            data_inftr.Recordset("nombre") = data_llamtur.Recordset("nombre")
            If IsNull(data_llamtur.Recordset("ci")) = False Then
               If data_llamtur.Recordset("ci") > 0 Then
                  data_llam2.RecordSource = "Select * from resplla where nro =" & data_llamtur.Recordset("nrolla")
                  data_llam2.Refresh
                  If data_llam2.Recordset.RecordCount > 0 Then
                     If IsNull(data_llam2.Recordset("mes")) = False Then
                        data_inftr.Recordset("telef") = Trim(str(data_llamtur.Recordset("ci"))) & "-" & Trim(str(data_llam2.Recordset("mes")))
                     Else
                        data_inftr.Recordset("telef") = Trim(str(data_llamtur.Recordset("ci"))) & "-0"
                     End If
                  End If
               End If
            End If
            data_inftr.Recordset("nomcat") = Trim(Mid(data_llamtur.Recordset("obs"), 1, 50))
            data_inftr.Recordset("nommed") = Trim(Mid(data_llamtur.Recordset("referen"), 1, 45))
            data_inftr.Recordset("referen") = Trim(data_llamtur.Recordset("obsmot"))
            data_inftr.Recordset("codmot") = data_llamtur.Recordset("codmot")
            data_inftr.Recordset("lugar") = data_llamtur.Recordset("lugar")
            data_inftr.Recordset.Update
         Else
            If data_llamtur.Recordset("cancela") = 1 Then
            Else
               data_inftr.Recordset.AddNew
               data_inftr.Recordset("fecha") = data_llamtur.Recordset("fecha")
               data_inftr.Recordset("hora") = data_llamtur.Recordset("hora")
               data_inftr.Recordset("nombre") = data_llamtur.Recordset("nombre")
               If IsNull(data_llamtur.Recordset("ci")) = False Then
                  If data_llamtur.Recordset("ci") > 0 Then
                     data_llam2.RecordSource = "Select * from resplla where nro =" & data_llamtur.Recordset("nrolla")
                     data_llam2.Refresh
                     If data_llam2.Recordset.RecordCount > 0 Then
                        If IsNull(data_llam2.Recordset("mes")) = False Then
                           data_inftr.Recordset("telef") = Trim(str(data_llamtur.Recordset("ci"))) & "-" & Trim(str(data_llam2.Recordset("mes")))
                        Else
                           data_inftr.Recordset("telef") = Trim(str(data_llamtur.Recordset("ci"))) & "-0"
                        End If
                     End If
                  End If
               End If
               data_inftr.Recordset("nomcat") = Trim(Mid(data_llamtur.Recordset("obs"), 1, 50))
               data_inftr.Recordset("nommed") = Trim(Mid(data_llamtur.Recordset("referen"), 1, 45))
               data_inftr.Recordset("referen") = Trim(data_llamtur.Recordset("motcon"))
               data_inftr.Recordset("codmot") = data_llamtur.Recordset("codmot")
               data_inftr.Recordset("lugar") = data_llamtur.Recordset("lugar")
               data_inftr.Recordset.Update
            End If
         End If
         data_llamtur.Recordset.MoveNext
      Loop
      data_inftr.RecordSource = "Select * from inflla"
      data_inftr.Refresh
      If Xresultamsp <> "" Then
         Kill ("c:\planillas\inftr.pdf")
      End If
      If Check3.Value = 1 Then
         If Xdiferenmsp >= 1 Then
            MsgBox "Se prepara archivo para enviar a MSP, Aguarde!!", vbInformation
            crmsp.Destination = crptToPrinter
            crmsp.ReportFileName = App.path & "\inftrasmsp.rpt"
            crmsp.Action = 1
            Timer2.Enabled = True
            MsgBox "Aguarde que se enviará el correo al MSP", vbInformation
         End If
      Else
         t_envmsp.Text = 0
      End If
   Else
      t_envmsp.Text = 0
      data_fecmsp.Recordset.Edit
      data_fecmsp.Recordset("fecha") = data_fecmsp.Recordset("fecha") + 1
      data_fecmsp.Recordset.Update
      MsgBox "Proceso TERMINADO"
      frm_opsdesp.Enabled = True
   End If
Else
   MsgBox "Proceso TERMINADO"
   frm_opsdesp.Enabled = True
End If

End Sub

Private Sub Command4_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Xdias As Integer
Dim Cuentactrol As Integer
Dim Xfdesde, Xfhasta As Date
Dim Xlabrir3 As New Excel.Application
Dim Letra As String

data_ctrf.DatabaseName = App.path & "\ctradm.mdb"
data_ctrf.RecordSource = "ctradm"
data_ctrf.Refresh
data_llam.Connect = "odbc;dsn=sappnew;"
data_covid.Connect = "odbc;dsn=sappnew;"
data_covidT.Connect = "odbc;dsn=sappnew;"

Xdias = DateDiff("d", data_ctrf.Recordset("fecha"), Date)

If Xdias >= 2 Then
   Xfdesde = data_ctrf.Recordset("fecha") + 1
   Xfhasta = Date - 1
   
   data_covidT.RecordSource = "Select * from seguimiento_covid where fecha >=#" & Format(Xfdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfhasta, "yyyy/mm/dd") & "# order by fecha"
   data_covidT.Refresh
   
   If data_covidT.Recordset.RecordCount > 0 Then
      
      data_covidT.Recordset.MoveFirst
      Textofecha = Trim(str(Day(Xfhasta))) & Trim(str(Month(Xfhasta))) & Trim(str(Year(Xfhasta)))
      
      Cuentactrol = 0
      Xlin = 1
      XCol = 1
      Xtotreg = 0
      Xsub = 0
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      Xarchexel22.Name = Trim("CCOU")
      Xlibexel22.SaveAs ("C:\planillas\CCOU-COVID" & Trim(Textofecha) & ".xls")
      Xarchtex = "C:\planillas\CCOU-COVID" & Trim(Textofecha) & ".xls"

      Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
    
      Xarchexel22.Cells(Xlin, XCol) = "RELEVAMIENTO PACIENTES SOSPECHOSOS de CORONAVIRUS MUTUALISTA: CCOU DESDE: " & Xfdesde & " HASTA: " & Xfhasta
        
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
                
      Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "ID Paciente"
      XCol = XCol + 1
      Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
      XCol = XCol + 1
      Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
      XCol = XCol + 1
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "SEXO"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "EDAD"
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Dpto.Residencia"
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "INSTITUCIÓN MED."
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Personal de Salud?"
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Contacto ?"
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "FECHA CONT."
      XCol = XCol + 1
      Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "SINTOMAS?"
      XCol = XCol + 1
      Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Estuvo en cuarentena?"
      XCol = XCol + 1
      Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Fecha inicio Sint."
      XCol = XCol + 1
      Xarchexel22.Range("N" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Fiebre"
      XCol = XCol + 1
      Xarchexel22.Range("O" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Tos"
      XCol = XCol + 1
      Xarchexel22.Range("P" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Resfrío"
      XCol = XCol + 1
      Xarchexel22.Range("Q" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Insuf.Resp"
      XCol = XCol + 1
      Xarchexel22.Range("R" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Diarrea"
      XCol = XCol + 1
      Xarchexel22.Range("S" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Se realizó Test?"
      XCol = XCol + 1
      Xarchexel22.Range("T" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Result.Test?"
      XCol = XCol + 1
      Xarchexel22.Range("U" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Fec.Result"
      XCol = XCol + 1
      Xarchexel22.Range("V" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Fecha ALTA"
        
      Xlin = Xlin + 1
      XCol = 1
      Do While Not data_covidT.Recordset.EOF
         Cuentactrol = 0
         Letra = ""
         data_llam.RecordSource = "Select * from llamado where nrolla =" & data_covidT.Recordset("id_llamado")
         data_llam.Refresh
         If data_llam.Recordset.RecordCount > 0 Then
            data_covid.RecordSource = "select * from convenio where cnv_codigo ='" & data_llam.Recordset("categ") & "'"
            data_covid.Refresh
            If data_covid.Recordset.RecordCount > 0 Then
               If IsNull(data_covid.Recordset("cnv_grupo")) = False Then
                  If data_covid.Recordset("cnv_grupo") = "CCOU" Then
                     For i = 1 To Len(data_llam.Recordset("nombre"))
                         If Mid(data_llam.Recordset("nombre"), i, 1) = " " Then
                            Cuentactrol = Cuentactrol + 1
                         End If
                         If Cuentactrol = 2 Then
                            If Trim(Letra) = "" Then
                               Letra = Mid(data_llam.Recordset("nombre"), i + 1, 1)
                            End If
                         End If
                     Next
                     If Cuentactrol = 3 Then
                        Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 1)
                     Else
                        If Cuentactrol = 2 Then
                           Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 1)
                        Else
                           Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 2)
                        End If
                     End If
                     If Len(data_llam.Recordset("ci")) = 6 Then
                        Letra = Letra + Mid(Trim(str(data_llam.Recordset("ci"))), 4, 3)
                     Else
                        Letra = Letra + Mid(Trim(str(data_llam.Recordset("ci"))), 5, 3)
                     End If
                     Xarchexel22.Cells(Xlin, XCol) = Trim(Letra)
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("nombre")
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("ci")
                     XCol = XCol + 1
                     If data_llam.Recordset("hh") = 0 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Masculino"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Femenino"
                     End If
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("edad")
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = "Canelones"
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = "CÍRCULO CATÓLICO"
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("grupo_covid")) = False Then
                        If data_llam.Recordset("grupo_covid") = "Personal de Salud" Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("contacto")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("contacto")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fec_contac")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("contacto"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'---"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("inicio_sint")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("cuarent_ant")) = False Then
                        If data_llam.Recordset("cuarent_ant") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fec_sint")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("fec_sint"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'---"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fiebre")) = False Then
                        If data_llam.Recordset("fiebre") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("tos")) = False Then
                        If data_llam.Recordset("tos") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("resfrio")) = False Then
                        If data_llam.Recordset("resfrio") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("insuf")) = False Then
                        If data_llam.Recordset("insuf") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("diarrea")) = False Then
                        If data_llam.Recordset("diarrea") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_rea")) = False Then
                        If data_llam.Recordset("isopa_rea") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_result")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("isopa_result")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_fecrea")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("isopa_fecrea"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("cierre_fec")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("cierre_fec"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     Xlin = Xlin + 1
                     XCol = 1
                     Xtotreg = Xtotreg + 1
                  End If
               End If
            End If
         End If
         data_covidT.Recordset.MoveNext
      Loop
      Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
      Xlibexel22.Save
      Xlibexel22.Close
      Xobjexel22.Quit
      
      
      Dim MenCorreo As String
      Dim oMail As Class1
      Set oMail = New Class1
         With oMail
             .servidor = "smtp.gmail.com"
             .puerto = 465
             .UseAuntentificacion = True
             .ssl = True
             .Usuario = "despachosapp@gmail.com"
             .PassWord = "sapp1987"
             .Asunto = "Informe Relevamiento pacientes sospechosos de Coronavirus MUT: CCOU " & Xfdesde & "-- " & Xfhasta
             .de = "despachosapp@gmail.com"
             .para = "registrocovid@circulocatolico.com.uy; jefedepartamentoti@sapp.com.uy; directortecnico@sapp.com.uy"
   '         .para = "sappjorge@hotmail.com; despachosapp@hotmail.com; sappsusanadominguez@hotmail.com; sappdirecciontecnica@hotmail.com; sappenrique@hotmail.com"
             .Adjunto = Xarchtex
             .Mensaje = "Informe Relevamiento pacientes sospechosos de Coronavirus MUT: CCOU - Desde SAPP"
             .Enviar_Backup ' manda el mail
         End With
         Set oMail = Nothing
'         data_ctrf.Recordset.Edit
'         data_ctrf.Recordset("fecha") = Xfhasta
'         data_ctrf.Recordset.Update
   
   Else
'      MsgBox "No hay registros seguimiento COVID-19 CCOU"
   End If
Else
'   MsgBox "No hay registros seguimiento COVID-19 CCOU"
End If
   


End Sub

Private Sub Command5_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Xdias As Integer
Dim Cuentactrol As Integer
Dim Xfdesde, Xfhasta As Date
Dim Xlabrir3 As New Excel.Application
Dim Letra As String

data_ctrf.DatabaseName = App.path & "\ctradm.mdb"
data_ctrf.RecordSource = "ctradm"
data_ctrf.Refresh
data_llam.Connect = "odbc;dsn=sappnew;"
data_covid.Connect = "odbc;dsn=sappnew;"
data_covidT.Connect = "odbc;dsn=sappnew;"
data_paraci.Connect = "odbc;dsn=sappnew;"

Xdias = DateDiff("d", data_ctrf.Recordset("fecha"), Date)
         
If Xdias >= 2 Then
   Xfdesde = data_ctrf.Recordset("fecha") + 1
   Xfhasta = Date - 1
   
   data_covidT.RecordSource = "Select * from seguimiento_covid where fecha >=#" & Format(Xfdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfhasta, "yyyy/mm/dd") & "# order by fecha"
   data_covidT.Refresh
   
   If data_covidT.Recordset.RecordCount > 0 Then
      
      data_covidT.Recordset.MoveFirst
      Textofecha = Trim(str(Day(Xfhasta))) & Trim(str(Month(Xfhasta))) & Trim(str(Year(Xfhasta)))
      
      Cuentactrol = 0
      Xlin = 1
      XCol = 1
      Xtotreg = 0
      Xsub = 0
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      Xarchexel22.Name = Trim("Evangelico")
      Xlibexel22.SaveAs ("C:\planillas\Evang-COVID" & Trim(Textofecha) & ".xls")
      Xarchtex = "C:\planillas\Evang-COVID" & Trim(Textofecha) & ".xls"

      Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
    
      Xarchexel22.Cells(Xlin, XCol) = "RELEVAMIENTO PACIENTES SOSPECHOSOS de CORONAVIRUS MUTUALISTA: H.EVANGELICO DESDE: " & Xfdesde & " HASTA: " & Xfhasta
        
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
                
      Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "ID Paciente"
      XCol = XCol + 1
      Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
      XCol = XCol + 1
      Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
      XCol = XCol + 1
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "SEXO"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "EDAD"
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Dpto.Residencia"
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "INSTITUCIÓN MED."
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Personal de Salud?"
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Contacto ?"
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "FECHA CONT."
      XCol = XCol + 1
      Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "SINTOMAS?"
      XCol = XCol + 1
      Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Estuvo en cuarentena?"
      XCol = XCol + 1
      Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Fecha inicio Sint."
      XCol = XCol + 1
      Xarchexel22.Range("N" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Fiebre"
      XCol = XCol + 1
      Xarchexel22.Range("O" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Tos"
      XCol = XCol + 1
      Xarchexel22.Range("P" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Resfrío"
      XCol = XCol + 1
      Xarchexel22.Range("Q" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Insuf.Resp"
      XCol = XCol + 1
      Xarchexel22.Range("R" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Diarrea"
      XCol = XCol + 1
      Xarchexel22.Range("S" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Se realizó Test?"
      XCol = XCol + 1
      Xarchexel22.Range("T" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Result.Test?"
      XCol = XCol + 1
      Xarchexel22.Range("U" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Fec.Result"
      XCol = XCol + 1
      Xarchexel22.Range("V" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Fecha ALTA"
        
      Xlin = Xlin + 1
      XCol = 1
      Do While Not data_covidT.Recordset.EOF
         Cuentactrol = 0
         Letra = ""
         data_llam.RecordSource = "Select * from llamado where nrolla =" & data_covidT.Recordset("id_llamado")
         data_llam.Refresh
         If data_llam.Recordset.RecordCount > 0 Then
            data_covid.RecordSource = "select * from convenio where cnv_codigo ='" & data_llam.Recordset("categ") & "'"
            data_covid.Refresh
            If data_covid.Recordset.RecordCount > 0 Then
               If IsNull(data_covid.Recordset("cnv_grupo")) = False Then
                  If data_covid.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                     For i = 1 To Len(data_llam.Recordset("nombre"))
                         If Mid(data_llam.Recordset("nombre"), i, 1) = " " Then
                            Cuentactrol = Cuentactrol + 1
                         End If
                         If Cuentactrol = 2 Then
                            If Trim(Letra) = "" Then
                               Letra = Mid(data_llam.Recordset("nombre"), i + 1, 1)
                            End If
                         End If
                     Next
                     If Cuentactrol = 3 Then
                        Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 1)
                     Else
                        If Cuentactrol = 2 Then
                           Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 1)
                        Else
                           Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 2)
                        End If
                     End If
                     If Len(data_llam.Recordset("ci")) = 6 Then
                        Letra = Letra + Mid(Trim(str(data_llam.Recordset("ci"))), 4, 3)
                     Else
                        Letra = Letra + Mid(Trim(str(data_llam.Recordset("ci"))), 5, 3)
                     End If
                     Xarchexel22.Cells(Xlin, XCol) = Trim(Letra)
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("nombre")
                     XCol = XCol + 1
'                     data_paraci.RecordSource = "Select * from resplla where nro =" & data_llam.Recordset("nrolla")
'                     data_paraci.Refresh
'                     If data_paraci.Recordset.RecordCount > 0 Then
'                        If IsNull(data_paraci.Recordset("mes")) = False Then
'                           If IsNull(data_llam.Recordset("ci")) = False Then
'                              Xarchexel22.Cells(Xlin, XCol) = Trim(str(data_llam.Recordset("ci"))) & "-" & Trim(str(data_paraci.Recordset("mes")))
'                           Else
'                              Xarchexel22.Cells(Xlin, XCol) = "0"
'                           End If
'                        Else
'                           If IsNull(data_llam.Recordset("ci")) = False Then
'                              Xarchexel22.Cells(Xlin, XCol) = Trim(str(data_llam.Recordset("ci"))) & "-0"
'                           Else
'                              Xarchexel22.Cells(Xlin, XCol) = "0"
'                           End If
'                        End If
'                     End If
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("ci")
                     XCol = XCol + 1
                     If data_llam.Recordset("hh") = 0 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Masculino"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Femenino"
                     End If
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("edad")
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = "Canelones"
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = "HOSPITAL EVANGÉLICO"
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("grupo_covid")) = False Then
                        If data_llam.Recordset("grupo_covid") = "Personal de Salud" Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("contacto")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("contacto")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fec_contac")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("contacto"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'---"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("inicio_sint")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("cuarent_ant")) = False Then
                        If data_llam.Recordset("cuarent_ant") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fec_sint")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("fec_sint"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'---"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fiebre")) = False Then
                        If data_llam.Recordset("fiebre") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("tos")) = False Then
                        If data_llam.Recordset("tos") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("resfrio")) = False Then
                        If data_llam.Recordset("resfrio") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("insuf")) = False Then
                        If data_llam.Recordset("insuf") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("diarrea")) = False Then
                        If data_llam.Recordset("diarrea") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_rea")) = False Then
                        If data_llam.Recordset("isopa_rea") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_result")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("isopa_result")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_fecrea")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("isopa_fecrea"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("cierre_fec")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("cierre_fec"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     Xlin = Xlin + 1
                     XCol = 1
                     Xtotreg = Xtotreg + 1
                  End If
               End If
            End If
         End If
         data_covidT.Recordset.MoveNext
      Loop
      Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
      Xlibexel22.Save
      Xlibexel22.Close
      Xobjexel22.Quit
            
      Dim MenCorreo As String
      Dim oMail As Class1
      Set oMail = New Class1
         With oMail
             .servidor = "smtp.gmail.com"
             .puerto = 465
             .UseAuntentificacion = True
             .ssl = True
             .Usuario = "despachosapp@gmail.com"
             .PassWord = "sapp1987"
             .Asunto = "Informe Relevamiento pacientes sospechosos de Coronavirus MUT: H.Evangélico " & Xfdesde & "-- " & Xfhasta
             .de = "despachosapp@gmail.com"
             .para = "direccioncolonia@hospitalevangelico.com; directortecnico@sapp.com.uy; jefedepartamentoti@sapp.com.uy"
             .Adjunto = Xarchtex
             .Mensaje = "Informe Relevamiento pacientes sospechosos de Coronavirus MUT: H.Evangélico - Desde SAPP"
             .Enviar_Backup
         End With
         Set oMail = Nothing
         
'''         MsgBox "TERMINADO"

   Else
   End If
Else
End If

data_ctrf.Recordset.Edit
data_ctrf.Recordset("fecha") = Date - 1
data_ctrf.Recordset.Update

End Sub

Private Sub Command6_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Xdias As Integer
Dim Cuentactrol As Integer
Dim Xfdesde, Xfhasta As Date
Dim Xlabrir3 As New Excel.Application
Dim Letra As String

data_ctrf.DatabaseName = App.path & "\ctradm.mdb"
data_ctrf.RecordSource = "ctradm"
data_ctrf.Refresh
data_llam.Connect = "odbc;dsn=sappnew;"
data_covid.Connect = "odbc;dsn=sappnew;"
data_covidT.Connect = "odbc;dsn=sappnew;"

Xdias = DateDiff("d", data_ctrf.Recordset("fecha"), Date)
         
If Xdias >= 2 Then
   Xfdesde = data_ctrf.Recordset("fecha") + 1
   Xfhasta = Date - 1
   
   data_covidT.RecordSource = "Select * from seguimiento_covid where fecha >=#" & Format(Xfdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfhasta, "yyyy/mm/dd") & "# order by fecha"
   data_covidT.Refresh
   
   If data_covidT.Recordset.RecordCount > 0 Then
      
      data_covidT.Recordset.MoveFirst
      Textofecha = Trim(str(Day(Xfhasta))) & Trim(str(Month(Xfhasta))) & Trim(str(Year(Xfhasta)))
      
      Cuentactrol = 0
      Xlin = 1
      XCol = 1
      Xtotreg = 0
      Xsub = 0
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      Xarchexel22.Name = Trim("SMI")
      Xlibexel22.SaveAs ("C:\planillas\SMI-COVID" & Trim(Textofecha) & ".xls")
      Xarchtex = "C:\planillas\SMI-COVID" & Trim(Textofecha) & ".xls"

      Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
    
      Xarchexel22.Cells(Xlin, XCol) = "RELEVAMIENTO PACIENTES SOSPECHOSOS de CORONAVIRUS MUTUALISTA: SMI DESDE: " & Xfdesde & " HASTA: " & Xfhasta
        
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
                
      Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "ID Paciente"
      XCol = XCol + 1
      Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
      XCol = XCol + 1
      Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
      XCol = XCol + 1
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "SEXO"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "EDAD"
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Dpto.Residencia"
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "INSTITUCIÓN MED."
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Personal de Salud?"
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Contacto ?"
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "FECHA CONT."
      XCol = XCol + 1
      Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "SINTOMAS?"
      XCol = XCol + 1
      Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Estuvo en cuarentena?"
      XCol = XCol + 1
      Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "Fecha inicio Sint."
      XCol = XCol + 1
      Xarchexel22.Range("N" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Fiebre"
      XCol = XCol + 1
      Xarchexel22.Range("O" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Tos"
      XCol = XCol + 1
      Xarchexel22.Range("P" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Resfrío"
      XCol = XCol + 1
      Xarchexel22.Range("Q" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Insuf.Resp"
      XCol = XCol + 1
      Xarchexel22.Range("R" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "Diarrea"
      XCol = XCol + 1
      Xarchexel22.Range("S" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Se realizó Test?"
      XCol = XCol + 1
      Xarchexel22.Range("T" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Result.Test?"
      XCol = XCol + 1
      Xarchexel22.Range("U" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Fec.Result"
      XCol = XCol + 1
      Xarchexel22.Range("V" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Fecha ALTA"
        
      Xlin = Xlin + 1
      XCol = 1
      Do While Not data_covidT.Recordset.EOF
         Cuentactrol = 0
         Letra = ""
         data_llam.RecordSource = "Select * from llamado where nrolla =" & data_covidT.Recordset("id_llamado")
         data_llam.Refresh
         If data_llam.Recordset.RecordCount > 0 Then
            data_covid.RecordSource = "select * from convenio where cnv_codigo ='" & data_llam.Recordset("categ") & "'"
            data_covid.Refresh
            If data_covid.Recordset.RecordCount > 0 Then
               If IsNull(data_covid.Recordset("cnv_grupo")) = False Then
                  If data_covid.Recordset("cnv_grupo") = "SMI" Then
                     For i = 1 To Len(data_llam.Recordset("nombre"))
                         If Mid(data_llam.Recordset("nombre"), i, 1) = " " Then
                            Cuentactrol = Cuentactrol + 1
                         End If
                         If Cuentactrol = 2 Then
                            If Trim(Letra) = "" Then
                               Letra = Mid(data_llam.Recordset("nombre"), i + 1, 1)
                            End If
                         End If
                     Next
                     If Cuentactrol = 3 Then
                        Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 1)
                     Else
                        If Cuentactrol = 2 Then
                           Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 1)
                        Else
                           Letra = Letra + Mid(data_llam.Recordset("nombre"), 1, 2)
                        End If
                     End If
                     If Len(data_llam.Recordset("ci")) = 6 Then
                        Letra = Letra + Mid(Trim(str(data_llam.Recordset("ci"))), 4, 3)
                     Else
                        Letra = Letra + Mid(Trim(str(data_llam.Recordset("ci"))), 5, 3)
                     End If
                     Xarchexel22.Cells(Xlin, XCol) = Trim(Letra)
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("nombre")
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("ci")
                     XCol = XCol + 1
                     If data_llam.Recordset("hh") = 0 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Masculino"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Femenino"
                     End If
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("edad")
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = "Canelones"
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = "S.M.I."
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("grupo_covid")) = False Then
                        If data_llam.Recordset("grupo_covid") = "Personal de Salud" Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("contacto")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("contacto")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fec_contac")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("contacto"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'---"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("inicio_sint")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("cuarent_ant")) = False Then
                        If data_llam.Recordset("cuarent_ant") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fec_sint")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("fec_sint"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'---"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("fiebre")) = False Then
                        If data_llam.Recordset("fiebre") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("tos")) = False Then
                        If data_llam.Recordset("tos") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("resfrio")) = False Then
                        If data_llam.Recordset("resfrio") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("insuf")) = False Then
                        If data_llam.Recordset("insuf") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("diarrea")) = False Then
                        If data_llam.Recordset("diarrea") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_rea")) = False Then
                        If data_llam.Recordset("isopa_rea") = 1 Then
                           Xarchexel22.Cells(Xlin, XCol) = "SI"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "NO"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_result")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("isopa_result")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("isopa_fecrea")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("isopa_fecrea"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_llam.Recordset("cierre_fec")) = False Then
                        Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("cierre_fec"), "dd/mm/yyyy")
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "'--"
                     End If
                     Xlin = Xlin + 1
                     XCol = 1
                     Xtotreg = Xtotreg + 1
                  End If
               End If
            End If
         End If
         data_covidT.Recordset.MoveNext
      Loop
      Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
      Xlibexel22.Save
      Xlibexel22.Close
      Xobjexel22.Quit
      
      
      Dim MenCorreo As String
      Dim oMail As Class1
      Set oMail = New Class1
         With oMail
             .servidor = "smtp.gmail.com"
             .puerto = 465
             .UseAuntentificacion = True
             .ssl = True
             .Usuario = "despachosapp@gmail.com"
             .PassWord = "sapp1987"
             .Asunto = "Informe Relevamiento pacientes sospechosos de Coronavirus MUT: SMI " & Xfdesde & "-- " & Xfhasta
             .de = "despachosapp@gmail.com"
             .para = "mbrotos@smi.com.uy; jefedepartamentoti@sapp.com.uy; directortecnico@sapp.com.uy"
             .Adjunto = Xarchtex
             .Mensaje = "Informe Relevamiento pacientes sospechosos de Coronavirus MUT: SMI - Desde SAPP"
             .Enviar_Backup ' manda el mail
         End With
         Set oMail = Nothing
   
   Else
   End If
Else
End If

End Sub

Private Sub Command7_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Cuentactrol As Integer
Dim Xfdesde, Xfhasta As Date
Dim Xlabrir3 As New Excel.Application
Dim Letra As String

data_ctrf.DatabaseName = App.path & "\ctradm.mdb"
data_ctrf.RecordSource = "ctradm"
data_ctrf.Refresh
data_llam.Connect = "odbc;dsn=sappnew;"
data_covid.Connect = "odbc;dsn=sappnew;"
data_covidT.Connect = "odbc;dsn=sappnew;"

Xfdesde = CDate("01/04/2021")
Xfhasta = Date - 1
   
data_covidT.RecordSource = "Select llamado.nro,llamado.fecha,llamado.nombre,llamado.matric,llamado.ci,llamado.categ,llamado.sintomas,llamado.fec_sint,llamado.telef," & _
"llamado.edad,llamado.grupo_covid,llamado.fec_contac,llamado.cancela,llamado.segui_covid,llamado.contacto,llamado.cierre_fec,llamado.isopa_result,convenio.cnv_codigo," & _
"convenio.cnv_grupo from llamado inner join convenio on llamado.categ=convenio.cnv_codigo where llamado.isopa_result in ('Positivo') " & _
"and llamado.segui_covid in (1) and (llamado.cierre_fec is null or llamado.cierre_fec >=#" & Format(Xfdesde, "yyyy/mm/dd") & "#) and llamado.cancela is null and convenio.cnv_grupo in ('CCOU')"
'select * from llamado where isopa_result ='Positivo' and segui_covid in (1) and cierre_fec is null and cancela is null order by fecha
data_covidT.Refresh
   
If data_covidT.Recordset.RecordCount > 0 Then
      
   data_covidT.Recordset.MoveFirst
   Textofecha = Trim(str(Day(Xfhasta))) & Trim(str(Month(Xfhasta))) & Trim(str(Year(Xfhasta)))
      
   Cuentactrol = 0
   Xlin = 1
   XCol = 1
   Xtotreg = 0
   Xsub = 0
   Set Xobjexel22 = New Excel.Application
   Set Xlibexel22 = Xobjexel22.Workbooks.Add
   Set Xarchexel22 = Xlibexel22.Worksheets.Add
   Xarchexel22.Name = Trim("CCOU")
   Xlibexel22.SaveAs ("C:\planillas\CCOU-Positivos" & Trim(Textofecha) & ".xls")
   Xarchtex = "C:\planillas\CCOU-Positivos" & Trim(Textofecha) & ".xls"

   Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
   Xlin = Xlin + 1
   XCol = XCol + 1
   Xarchexel22.Range("A1", "C3").Font.Size = 16
   Xarchexel22.Range("A" & Trim(str(Xlin)), "L" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
    
   Xarchexel22.Cells(Xlin, XCol) = "SEGUIMIENTO PACIENTES COVID-19 POSITIVOS MUTUALISTA: CCOU DESDE: " & Xfdesde & " HASTA: " & Xfhasta
        
   XCol = 1
   Xlin = Xlin + 2
   Xnrocan = Xnrocan + Xlin
              
   Xarchexel22.Range("A" & Trim(str(Xlin)), "L" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
   Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
   Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
   XCol = XCol + 1
   Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
   Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
   XCol = XCol + 1
   Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 20
   Xarchexel22.Cells(Xlin, XCol) = "TELEFONOS"
   XCol = XCol + 1
   Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 20
   Xarchexel22.Cells(Xlin, XCol) = "EDAD"
   XCol = XCol + 1
   Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 25
   Xarchexel22.Cells(Xlin, XCol) = "Grupo"
   XCol = XCol + 1
   Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 20
   Xarchexel22.Cells(Xlin, XCol) = "Contacto ?"
   XCol = XCol + 1
   Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "FECHA CONT."
   XCol = XCol + 1
   Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "SINTOMAS?"
   XCol = XCol + 1
   Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 20
   Xarchexel22.Cells(Xlin, XCol) = "Fecha inicio Sint."
   XCol = XCol + 1
   Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "Test Antígeno"
   XCol = XCol + 1
   Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "FECHA ALTA"
        
   Xlin = Xlin + 1
   XCol = 1
   
   Do While Not data_covidT.Recordset.EOF
      Cuentactrol = 0
      Letra = ""
      Xarchexel22.Cells(Xlin, XCol) = Trim(data_covidT.Recordset("ci"))
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("nombre")
      XCol = XCol + 1
      If IsNull(data_covidT.Recordset("telef")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("telef")
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("edad")
      XCol = XCol + 1
      If IsNull(data_covidT.Recordset("grupo_covid")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("grupo_covid")
      End If
      XCol = XCol + 1
      If IsNull(data_covidT.Recordset("contacto")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("contacto")
      End If
      XCol = XCol + 1
      If IsNull(data_covidT.Recordset("fec_contac")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_covidT.Recordset("fec_contac"), "dd/mm/yyyy")
      End If
      XCol = XCol + 1
      If IsNull(data_covidT.Recordset("sintomas")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_covidT.Recordset("sintomas")
      End If
      XCol = XCol + 1
      If IsNull(data_covidT.Recordset("fec_sint")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_covidT.Recordset("fec_sint"), "dd/mm/yyyy")
      End If
      XCol = XCol + 1
      If IsNull(data_covidT.Recordset("cierre_fec")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_covidT.Recordset("cierre_fec"), "dd/mm/yyyy")
      End If
      
      XCol = 1
      
      data_covid.RecordSource = "Select * from seguimiento_covid where id_llamado =" & data_covidT.Recordset("nro") & " order by dia"
      data_covid.Refresh
      If data_covid.Recordset.RecordCount > 0 Then
         data_covid.Recordset.MoveFirst
         Xlin = Xlin + 1
         Do While Not data_covid.Recordset.EOF
            XCol = XCol + 1
            Xarchexel22.Cells(Xlin, XCol) = "CONTROL Nro." & Trim(str(data_covid.Recordset("dia")))
            XCol = XCol + 1
            Xarchexel22.Cells(Xlin, XCol) = "FECHA: " & Format(data_covid.Recordset("fecha"), "dd/mm/yyyy")
            data_llam.RecordSource = "select * from usuarios where usuario ='" & data_covid.Recordset("nom_usu") & "'"
            data_llam.Refresh
            XCol = XCol + 1
            If data_llam.Recordset.RecordCount > 0 Then
               Xarchexel22.Cells(Xlin, XCol) = "MÉDICO: " & data_llam.Recordset("nombre")
            Else
               Xarchexel22.Cells(Xlin, XCol) = "MÉDICO: Sin Datos"
            End If
            XCol = XCol + 1
            Xarchexel22.Cells(Xlin, XCol) = "PROX.CONTROL: " & Format(data_covid.Recordset("fecha_control"), "dd/mm/yyyy")
            data_covid.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
      End If
      
      XCol = 1
      Xlin = Xlin + 1
                 
      data_covidT.Recordset.MoveNext
   Loop
   Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
   Xlibexel22.Save
   Xlibexel22.Close
   Xobjexel22.Quit
   MsgBox "Proceso terminado.", vbInformation
End If


End Sub

Private Sub Command8_Click()
      frm_opsdesp.MousePointer = 11
      Consultar_CMT_total
      Consultar_CMT_PendDesp
      Consultar_CMT_Realiza
      Consultar_CMT_poli
      Consultar_CMT_posit
      Consultar_CMT_poliPend
      Consultar_CMT_PolRealiza
      Consultar_MG_Poli
      Consultar_PolRealizaMG
      Consultar_poliPend_MG
      Consultar_Positivos_Realiza
      Consultar_Pend_posit
      frm_opsdesp.MousePointer = 0

End Sub

Private Sub Command9_Click()
Dim XusuarioMedico As String
XusuarioMedico = ""
If t_ced.Text <> "" And t_codced.Text <> "" Then
   If Val(t_ced.Text) > 0 Then
      XusuarioMedico = Devuelve_Usuario()
      If Trim(XusuarioMedico) <> "" Then
         If mdmed.Text = "__/__/____" Then
            data_listamedicos.RecordSource = "select * from medicos_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & txt_nrocob.Text
         Else
            data_listamedicos.RecordSource = "select * from medicos_cmt where fecha =#" & Format(mdmed.Text, "yyyy/mm/dd") & "# and cod_med =" & txt_nrocob.Text
         End If
         data_listamedicos.Refresh
         If data_listamedicos.Recordset.RecordCount > 0 Then
            MsgBox "Ya existe el médico seleccionado en la lista.", vbExclamation
         Else
            data_listamedicos.Recordset.AddNew
            data_listamedicos.Recordset("cod_med") = txt_nrocob.Text
            data_listamedicos.Recordset("nom_usuario") = XusuarioMedico
            data_listamedicos.Recordset("cedula") = t_ced.Text & t_codced.Text
            If mdmed.Text = "__/__/____" Then
               data_listamedicos.Recordset("fecha") = Date
            Else
               data_listamedicos.Recordset("fecha") = mdmed.Text
            End If
            data_listamedicos.Recordset.Update
            If mdmed.Text = "__/__/____" Then
               List4.AddItem txt_nrocob.Text & "--" & txt_nomcob.Text & " Ced." & t_ced.Text & "-" & t_codced.Text
            End If
         End If
      Else
         MsgBox "No se puede agregar médico.", vbCritical
      End If
   Else
      MsgBox "Debe registrar cédula del médico.", vbCritical
   End If
Else
   MsgBox "Debe registrar cédula del médico.", vbCritical
End If

End Sub

Private Sub dbensal_DblClick()
igualahs

End Sub

Private Sub DBGrid112_DblClick(index As Integer)
'If KeyAscii = 13 Then
    If IsNull(data_cob.Recordset("med_cod")) = False Then
       txt_nrocob.Text = data_cob.Recordset("med_cod")
    Else
       txt_nrocob.Text = ""
    End If
    If IsNull(data_cob.Recordset("med_nombre")) = False Then
       txt_nomcob.Text = data_cob.Recordset("med_nombre")
    Else
       txt_nomcob.Text = ""
    End If
    If IsNull(data_cob.Recordset("med_esp")) = False Then
       txt_espec.Text = data_cob.Recordset("med_esp")
    Else
       txt_espec.Text = ""
    End If
    If IsNull(data_cob.Recordset("med_socnom")) = False Then
       txt_tel.Text = data_cob.Recordset("med_socnom")
    Else
       txt_tel.Text = ""
    End If
    
    If txt_nrocob.Text <> "" Then
       data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
       data_medhc.Refresh
       If data_medhc.Recordset.RecordCount > 0 Then
          If IsNull(data_medhc.Recordset("m_nrofrm")) = False Then
             If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 7 Then
                t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 6)
                t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 7, 1)
             Else
                If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 8 Then
                   t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 7)
                   t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 8, 1)
                Else
                   t_ced.Text = ""
                   t_codced.Text = ""
                End If
             End If
          Else
             t_ced.Text = ""
             t_codced.Text = ""
          End If
          If IsNull(data_medhc.Recordset("m_codmed")) = False Then
             Check1.Value = data_medhc.Recordset("m_codmed")
          End If
       Else
          t_ced.Text = ""
          t_codced.Text = ""
       End If
    End If



txt_bcob.Enabled = False
'DBGrid112(1).Enabled = False
'bmodif.SetFocus

End Sub

Private Sub DBGrid13_DblClick(index As Integer)

iguala

End Sub

Private Sub dbturno_DblClick()
data_grabatur.RecordSource = "Select * from mant_sol_hc where cl_codigo =" & data_vertur.Recordset("cl_codigo")
data_grabatur.Refresh
If data_grabatur.Recordset.RecordCount > 0 Then
   igualaturno
Else
   mf.Text = "__/__/____"
   mhor.Text = "__:__"
   mffin.Text = "__/__/____"
   mhfin.Text = "__:__"
   labu.Caption = ""
   t_obsturno.Text = ""
   labnomu.Caption = ""
   labidtur.Caption = ""

End If

End Sub

Private Sub Form_Initialize()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("med_cod")) = False Then
   txt_nrocob.Text = data_cob.Recordset("med_cod")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("med_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("med_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("med_esp")) = False Then
   txt_espec.Text = data_cob.Recordset("med_esp")
Else
   txt_espec.Text = ""
End If
If IsNull(data_cob.Recordset("med_socnom")) = False Then
   txt_tel.Text = data_cob.Recordset("med_socnom")
Else
   txt_tel.Text = ""
End If
If txt_nrocob.Text <> "" Then
   data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
   data_medhc.Refresh
   If data_medhc.Recordset.RecordCount > 0 Then
      If IsNull(data_medhc.Recordset("m_nrofrm")) = False Then
         If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 7 Then
            t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 6)
            t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 7, 1)
         Else
            If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 8 Then
               t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 7)
               t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 8, 1)
            Else
               t_ced.Text = ""
               t_codced.Text = ""
            End If
         End If
      Else
         t_ced.Text = ""
         t_codced.Text = ""
      End If
      If IsNull(data_medhc.Recordset("m_codmed")) = False Then
         Check1.Value = data_medhc.Recordset("m_codmed")
      End If
   Else
      t_ced.Text = ""
      t_codced.Text = ""
   End If
End If


End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub List4_DblClick()
Dim Xsnlper As String
Dim TextoaBorrar As String
Dim XX, XbandMed As Integer
XX = 0
XbandMed = 0
TextoaBorrar = ""
If List4.ListCount > 0 Then
   Xsnlper = MsgBox("Desea borrar de la lista el médico seleccionado?", vbExclamation + vbYesNo)
   If Xsnlper = vbYes Then
      If List4.ListIndex >= 0 Then
         For XX = 1 To Len(List4.List(List4.ListIndex))
             If Mid(Trim(List4.List(List4.ListIndex)), XX, 1) = "-" Then
                XbandMed = 1
             Else
                If XbandMed = 0 Then
                   TextoaBorrar = TextoaBorrar & Mid(Trim(List4.List(List4.ListIndex)), XX, 1)
                End If
             End If
         Next
         If Trim(TextoaBorrar) <> "" Then
            data_listamedicos.RecordSource = "select * from medicos_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Val(TextoaBorrar)
            data_listamedicos.Refresh
            If data_listamedicos.Recordset.RecordCount > 0 Then
               data_listamedicos.Recordset.Delete
            End If
            List4.RemoveItem List4.ListIndex
         End If
      End If
   End If
End If

End Sub

Private Sub md_GotFocus()
md.Text = Date

End Sub
Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhord.SetFocus
End If

End Sub

Private Sub mh_GotFocus()
mh.Text = Date

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhorh.SetFocus
End If

End Sub

Private Sub mhord_GotFocus()
mhord.Text = Format(Time, "HH:mm:ss")

End Sub

Private Sub mhord_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mhorh_GotFocus()
mhorh.Text = Format(Time, "HH:mm:ss")

End Sub

Private Sub mhorh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub Form_Load()
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_grabatur.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_vertur.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hsmed2.ConnectionString = "dsn=" & Xconexrmt

data_inftur.DatabaseName = App.path & "\informes.mdb"
data_inftr.DatabaseName = App.path & "\informes.mdb"

data_listamedicos.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_llamtur.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_fecmsp.DatabaseName = App.path & "\env_msp.mdb"
data_fecmsp.RecordSource = "env_msp"
data_fecmsp.Refresh
data_llam2.Connect = "odbc;dsn=" & Xconexrmt & ";"

If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "HVENTRE" Then
   data_vertur.RecordSource = "Select * from mant_sol_hc order by cl_fnac DESC"
Else
   data_vertur.RecordSource = "Select * from mant_sol_hc where cl_descpag ='" & WElusuario & "' order by cl_fnac DESC"
End If
data_vertur.Refresh

data_mov.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_mov.RecordSource = "Select * from moviles order by nroreg"
data_mov.Refresh
data_med(0).Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med(0).RecordSource = "Select * from medicos order by med_nombre"
data_med(0).Refresh

data_m.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_m.RecordSource = "Select * from medicos order by med_nombre"
data_m.Refresh
If data_m.Recordset.RecordCount > 0 Then
   data_m.Recordset.MoveFirst
   Do While Not data_m.Recordset.EOF
      Combo2.AddItem data_m.Recordset("med_nombre")
      data_m.Recordset.MoveNext
   Loop
End If
data_infhs.DatabaseName = App.path & "\informes.mdb"
'data_infhs.RecordSource = "infvtas"
'data_infhs.Refresh

data_enf(0).DatabaseName = App.path & "\enferm.mdb"
data_enf(0).RecordSource = "enferm"
data_enf(0).Refresh
'If data_mov.Recordset.RecordCount > 0 Then
'   data_mov.Recordset.MoveLast
'   igualacuadros
'End If
data_inf(0).DatabaseName = App.path & "\informes.mdb"
data_inf(0).RecordSource = "infcli"
data_inf(0).Refresh

data_chof(0).Connect = "odbc;dsn=" & Xconexrmt & ";"
data_chof(0).RecordSource = "Select * from movil where nromov >=" & 14 & " and nromov <" & 999 & " order by nromov"
data_chof(0).Refresh

data_horasmed.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_medhc.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_medhcc.ConnectionString = "dsn=" & Xconexrmt

data_verhsmed.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_verhsmed.RecordSource = "Select * from hc_archotro order by id DESC"
data_verhsmed.Refresh
data_verhsme.ConnectionString = "dsn=" & Xconexrmt

Dim Xfmov As Date
Xfmov = Date - 2
data_graba.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_graba.RecordSource = "Select * from movil where nromov =" & 999 & " order by nrolla"
data_graba.Refresh
data_busca.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_busca.RecordSource = "Select * from movil where nromov =" & 999 & " and fecha >=#" & Format(Xfmov, "yyyy/mm/dd") & "# order by fecha DESC,nrolla"
data_busca.Refresh

data_med(1).Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med(1).RecordSource = "medicos"
data_med(1).Refresh

data_chof(1).Connect = "odbc;dsn=" & Xconexrmt & ";"
data_chof(1).RecordSource = "Select * from moviles"
data_chof(1).Refresh

SSTab1.Tab = 0

data_par.DatabaseName = App.path & "\paramhoras.mdb"
data_par.RecordSource = "parsec0"
data_par.Refresh

End Sub

Private Sub b_busca_Click()
frm_buscamovil.Show vbModal

End Sub



Private Sub b_elimina_Click()
Dim Xrespumov As String
Xrespumov = MsgBox("Desea eliminar el registro del móvil?", vbYesNo, "Móviles")
If Xrespumov = vbYes Then
   data_mov.Recordset.Delete
   data_mov.Refresh
   borra_mov
   If data_mov.Recordset.RecordCount > 0 Then
      igualacuadros
   End If
End If

End Sub


Private Sub b_imprime_Click()
If data_inf(0).Recordset.RecordCount > 0 Then
   data_inf(0).Recordset.MoveFirst
   Do While Not data_inf(0).Recordset.EOF
      data_inf(0).Recordset.Delete
      data_inf(0).Recordset.MoveNext
   Loop
   data_inf(0).Refresh
End If

If data_mov.Recordset.RecordCount > 0 Then
   data_mov.Recordset.MoveFirst
   Do While Not data_mov.Recordset.EOF
      data_inf(0).Recordset.AddNew
      data_inf(0).Recordset("cl_codigo") = data_mov.Recordset("movil")
      data_inf(0).Recordset("cl_apellid") = Mid(data_mov.Recordset("nommed"), 1, 30)
      data_inf(0).Recordset("cl_nombre") = Mid(data_mov.Recordset("nomchof"), 1, 30)
      data_inf(0).Recordset("cl_localid") = Mid(data_mov.Recordset("nomenf"), 1, 30)
      data_inf(0).Recordset("cl_fnac") = data_mov.Recordset("fecha_Act")
      data_inf(0).Recordset.Update
      data_mov.Recordset.MoveNext
   Loop
   data_inf(0).RecordSource = "Select * from infcli"
   data_inf(0).Refresh
   cr1.ReportFileName = App.path & "\infmovile.rpt"
   cr1.ReportTitle = "INFORME DE MOVILES ACTUALIZADOS"
   cr1.Action = 1
End If

borra_mov

End Sub


Private Sub b_nuevo_Click()
deshabmov
borra_mov
Frame1.Enabled = True
txt_nro.SetFocus
XAlta = 1
If data_mov.Recordset.RecordCount > 0 Then
   data_mov.Recordset.MoveLast
   Label6(0).Caption = data_mov.Recordset("nroreg") + 1
Else
   Label6(0).Caption = 1
End If
data_mov.Recordset.AddNew
Command1(0).Enabled = True
Command2(0).Enabled = True

End Sub





Private Sub dbmedic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_chof(0).SetFocus
End If

End Sub

Private Sub dbmedic_LostFocus()
If IsNumeric(dbmedic.Text) = True Then
   If Val(dbmedic.Text) > 0 Then
      data_med(0).Recordset.FindFirst "med_cod =" & dbmedic.Text
      If Not data_med(0).Recordset.NoMatch Then
         dbmedic.Text = data_med(0).Recordset("med_nombre")
         txt_codmed.Text = data_med(0).Recordset("med_cod")
      Else
         MsgBox "No encontrado, busque por nombre", vbInformation, "Médicos"
         dbmedic.SetFocus
      End If
   Else
      data_med(0).Recordset.FindFirst "med_nombre ='" & dbmedic.Text & "'"
      If Not data_med(0).Recordset.NoMatch Then
         dbmedic.Text = data_med(0).Recordset("med_nombre")
         txt_codmed.Text = data_med(0).Recordset("med_cod")
      Else
         MsgBox "No encontrado, VERIFIQUE NOMBRE", vbInformation, "Médicos"
         dbmedic.SetFocus
      End If
   End If
Else
   If Len(dbmedic.Text) > 0 Then
      data_med(0).Recordset.FindFirst "med_nombre ='" & dbmedic.Text & "'"
      If Not data_med(0).Recordset.NoMatch Then
         dbmedic.Text = data_med(0).Recordset("med_nombre")
         txt_codmed.Text = data_med(0).Recordset("med_cod")
      Else
         MsgBox "No encontrado, VERIFIQUE NOMBRE", vbInformation, "Médicos"
         dbmedic.SetFocus
      End If
   End If
End If

End Sub





Private Sub mfec_GotFocus()
mfec.Text = Date

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_base.SetFocus
End If

End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 4 Then
   If XWeltipoU = "USUARIOS RECEP" Or XWeltipoU = "USUARIOS" Then
      MsgBox "Solo usuarios Largador y Administradores"
      SSTab1.Tab = 0
   End If
End If
If SSTab1.Tab = 5 Then
'   MsgBox "Mostrar datos"
   frm_opsdesp.MousePointer = 11
   Consultar_CMT_total
   Consultar_CMT_PendDesp
   Consultar_CMT_Realiza
   Consultar_CMT_poli
   Consultar_CMT_posit
   Consultar_CMT_poliPend
   Consultar_CMT_PolRealiza
   Consultar_MG_Poli
   Consultar_PolRealizaMG
   Consultar_poliPend_MG
   Consultar_Positivos_Realiza
   Consultar_Pend_posit
   Consultar_Results
   Consultar_ResPend
   Consultar_Resulta_Realiza
   frm_opsdesp.MousePointer = 0

End If
If SSTab1.Tab = 2 Then
   List4.Clear
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
   data_listamedicos.Refresh
   If data_listamedicos.Recordset.RecordCount > 0 Then
      data_listamedicos.Recordset.MoveFirst
      Do While Not data_listamedicos.Recordset.EOF
         List4.AddItem data_listamedicos.Recordset("cod_med") & "--" & data_listamedicos.Recordset("med_nombre") & " Ced." & data_listamedicos.Recordset("cedula")
         data_listamedicos.Recordset.MoveNext
      Loop
   End If
End If

End Sub

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba(0).SetFocus
End If

End Sub



Private Sub t_chof_KeyPress(index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
   t_enf(0).SetFocus
End If



End Sub

Private Sub t_chof_LostFocus(index As Integer)

If t_chof(0).Text <> "" Then
   If IsNumeric(t_chof(0).Text) = True Then
      data_chof(0).Recordset.FindFirst "nromov =" & t_chof(0).Text
      If Not data_chof(0).Recordset.NoMatch Then
         t_chof(0).Text = data_chof(0).Recordset("chofer")
         labchof.Caption = data_chof(0).Recordset("nromov")
      Else
         t_chof(0).Text = ""
         labchof.Caption = ""
      End If
   Else
      data_chof(0).Recordset.FindFirst "chofer ='" & t_chof(0).Text & "'"
      If Not data_chof(0).Recordset.NoMatch Then
         t_chof(0).Text = data_chof(0).Recordset("chofer")
         labchof.Caption = data_chof(0).Recordset("nromov")
      Else
         t_chof(0).Text = ""
         labchof.Caption = ""
      End If
   End If
End If

End Sub

Private Sub t_choff_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   md.SetFocus
End If

End Sub



Private Sub t_codmedb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_codmedb.Text <> "" Then
      data_verhsmed.RecordSource = "Select * from hc_archotro where hc_mat =" & t_codmedb.Text & " order by id DESC"
      data_verhsmed.Refresh
   Else
      data_verhsmed.RecordSource = "Select * from hc_archotro order by id DESC"
      data_verhsmed.Refresh
   End If
   dbensal.SetFocus
End If

End Sub

Private Sub t_codmedm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub t_codmedm_LostFocus()
If t_codmedm.Text <> "" Then
   data_m.RecordSource = "Select * from medicos where med_cod =" & t_codmedm.Text
   data_m.Refresh
   If data_m.Recordset.RecordCount > 0 Then
      Combo2.Text = data_m.Recordset("med_nombre")
   Else
      MsgBox "Médico no econtrado"
      Combo2.SetFocus
   End If
End If

End Sub

Private Sub t_enf_KeyPress(index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
   mfec.SetFocus
End If

End Sub



Private Sub t_enf_LostFocus(index As Integer)

If t_enf(0).Text <> "" Then
   If IsNumeric(t_enf(0).Text) = True Then
      data_enf(0).Recordset.FindFirst "id =" & t_enf(0).Text
      If Not data_enf(0).Recordset.NoMatch Then
         t_enf(0).Text = data_enf(0).Recordset("nomb")
         labenf.Caption = data_enf(0).Recordset("id")
      Else
         t_enf(0).Text = ""
         labenf.Caption = ""
      End If
   Else
      data_enf(0).Recordset.FindFirst "nomb ='" & t_enf(0).Text & "'"
      If Not data_enf(0).Recordset.NoMatch Then
         t_enf(0).Text = data_enf(0).Recordset("nomb")
         labenf.Caption = data_enf(0).Recordset("id")
      Else
         t_enf(0).Text = ""
         labenf.Caption = ""
      End If
   End If
End If

End Sub

Private Sub t_enff_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   t_med.SetFocus
End If

End Sub

Private Sub t_mov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   md.SetFocus
End If

End Sub

Private Sub t_mov_LostFocus()
If t_mov.Text <> "" Then
   data_chof(1).RecordSource = "Select * from moviles where movil =" & t_mov.Text
   data_chof(1).Refresh
   If data_chof(1).Recordset.RecordCount > 0 Then
      If IsNull(data_chof(1).Recordset("nomchof")) = False Then
         t_choff(1).Text = data_chof(1).Recordset("nomchof")
      Else
         t_choff(1).Text = ""
      End If
      If IsNull(data_chof(1).Recordset("codchof")) = False Then
         Label9.Caption = data_chof(1).Recordset("codchof")
      Else
         Label9.Caption = ""
      End If
      If IsNull(data_chof(1).Recordset("codenf")) = False Then
         laben.Caption = data_chof(1).Recordset("codenf")
      Else
         laben.Caption = ""
      End If
      If IsNull(data_chof(1).Recordset("codmed")) = False Then
         labmed.Caption = data_chof(1).Recordset("codmed")
      Else
         labmed.Caption = ""
      End If
      If IsNull(data_chof(1).Recordset("nommed")) = False Then
         t_med.Text = data_chof(1).Recordset("nommed")
      Else
         t_med.Text = ""
      End If
      If IsNull(data_chof(1).Recordset("nomenf")) = False Then
         t_enff(1).Text = data_chof(1).Recordset("nomenf")
      Else
         t_enff(1).Text = ""
      End If
   Else
      MsgBox "NUMERO DE MOVIL NO ENCONTRADO", vbCritical, "SAPP"
      t_choff(1).Text = ""
      Label9.Caption = ""
      laben.Caption = ""
      labmed.Caption = ""
      t_med.Text = ""
      t_enff(1).Text = ""
   End If
End If

End Sub

Private Sub t_movb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_movb.Text <> "" Then
      data_busca.RecordSource = "Select * from movil where codmed =" & t_movb.Text & " and nromov =" & 999 & " order by nrolla"
      data_busca.Refresh
   Else
      data_busca.RecordSource = "Select * from movil where nromov =" & 999 & " order by nrolla"
      data_busca.Refresh
   End If
End If

End Sub

Private Sub t_nromov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codmedm.SetFocus
End If

End Sub

Private Sub t_obsturno_Change()
If XAlta = 1 Then
   t_obsturno.ToolTipText = "DOCUMENTO SIN GRABAR"
Else
   t_obsturno.ToolTipText = ""
End If

End Sub

Private Sub Timer1_Timer()

frm_opsdesp.MousePointer = 0
MsgBox "Archivo Generado, presione Aceptar para enviar correo"
b_envcor_Click
Timer1.Enabled = False
'MsgBox "Proceso TERMINADO"

'frm_opsdesp.Enabled = True

End Sub

Private Sub Timer2_Timer()
                 Timer2.Enabled = False

Dim lineatexto, textocorreo As String

Open App.path & "\correostr.txt" For Input As #1
Do While Not EOF(1)
   Line Input #1, lineatexto
Loop
textocorreo = lineatexto
Close #1
            
            Dim oMail2 As Class1
                 Set oMail2 = New Class1
                 With oMail2
                     .servidor = "smtp.office365.com"
                     .puerto = 25
                     .UseAuntentificacion = True
                     .ssl = True
                     .Usuario = "despacho@sapp.com.uy"
                     .PassWord = "Salinas1987"
                     .Asunto = "SAPP - Informe Traslados realizados a MSP " & Text1.Text
                     .de = "despacho@sapp.com.uy"
                     .para = textocorreo
'                     .para = "caneloneseste.rapcanelones@asse.com.uy; maira.castro.carli@gmail.com; sappjorge@hotmail.com; sappsusanadominguez@hotmail.com; sappdirecciontecnica@hotmail.com; sappsergioperez@hotmail.com; ventas@sapp.com.uy; direccion.pando@asse.com.uy"
'                     .para = "sappjorge@hotmail.com"
                     .Adjunto = "c:\planillas\inftr.pdf"
                     .Mensaje = "Informe de traslados realizados a MSP"
                     .Enviar_Backup ' manda el mail
                 End With
                 t_envmsp.Text = 0
                 Set oMail2 = Nothing
                 MsgBox "Correo a MSP Enviado!!"
                 data_fecmsp.Recordset.Edit
                 data_fecmsp.Recordset("fecha") = data_fecmsp.Recordset("fecha") + 1
                 data_fecmsp.Recordset.Update
                 MsgBox "Proceso TERMINADO"
                 frm_opsdesp.Enabled = True

End Sub

Private Sub txt_bcob_Change()
data_cob.RecordSource = "select * from medicos where med_nombre >='" & txt_bcob.Text & "' order by med_nombre"
data_cob.Refresh

End Sub

Private Sub txt_espec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_tel.SetFocus
End If

End Sub

Private Sub txt_nomcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_espec.SetFocus
End If

End Sub

Private Sub txt_nro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   dbmedic.SetFocus
End If

End Sub

Private Sub txt_nro_LostFocus()
If txt_nro.Text <> "" Then
Else
   MsgBox "No Ingresó móvil"
   txt_nro.SetFocus
End If

End Sub



Public Sub habilitamov()
b_nuevo.Enabled = True
b_modif(0).Enabled = True
b_graba(0).Enabled = False
b_imprime.Enabled = True
b_busca.Enabled = True
b_cance(0).Enabled = False
b_elimina.Enabled = True

End Sub

Public Sub deshabmov()
b_nuevo.Enabled = False
b_modif(0).Enabled = False
b_graba(0).Enabled = True
b_imprime.Enabled = False
b_busca.Enabled = False
b_cance(0).Enabled = True
b_elimina.Enabled = False

End Sub

Public Sub igualacuadros()
txt_nro.Text = data_mov.Recordset("movil")
If IsNull(data_mov.Recordset("codmed")) = False Then
   txt_codmed.Text = data_mov.Recordset("codmed")
Else
   txt_codmed.Text = 0
End If
If IsNull(data_mov.Recordset("nommed")) = False Then
   dbmedic.Text = data_mov.Recordset("nommed")
Else
   dbmedic.Text = ""
End If
If IsNull(data_mov.Recordset("fecha_act")) = False Then
   mfec.Text = Format(data_mov.Recordset("fecha_act"), "dd/mm/yyyy")
Else
   mfec.Text = "__/__/____"
End If
If IsNull(data_mov.Recordset("ano")) = False Then
   t_base.Text = data_mov.Recordset("ano")
Else
   t_base.Text = 0
End If
If IsNull(data_mov.Recordset("codchof")) = False Then
   labchof.Caption = data_mov.Recordset("codchof")
Else
   labchof.Caption = 0
End If
If IsNull(data_mov.Recordset("nomchof")) = False Then
   t_chof(0).Text = data_mov.Recordset("nomchof")
Else
   t_chof(0).Text = ""
End If
If IsNull(data_mov.Recordset("codenf")) = False Then
   labenf.Caption = data_mov.Recordset("codenf")
Else
   labenf.Caption = 0
End If
If IsNull(data_mov.Recordset("nomenf")) = False Then
   t_enf(0).Text = data_mov.Recordset("nomenf")
Else
   t_enf(0).Text = ""
End If


End Sub

Public Function borra_mov()
txt_nro.Text = ""
txt_codmed.Text = ""
dbmedic.Text = ""
mfec.Text = "__/__/____"
t_base.Text = ""
labchof.Caption = ""
t_chof(0).Text = ""
labenf.Caption = ""
t_enf(0).Text = ""


End Function

Public Function deshab()
b_alta.Enabled = False
b_modiff(1).Enabled = False
b_grabaa(1).Enabled = True
b_cancee(1).Enabled = True
b_eli.Enabled = False
b_imp.Enabled = False
DBGrid13(0).Enabled = False
Frame1.Enabled = True
'Command2.Enabled = False

End Function

Public Function limpia()
t_mov.Text = ""
t_choff(1).Text = ""
md.Text = "__/__/____"
mhord.Text = "__:__:__"
mh.Text = "__/__/____"
mhorh.Text = "__:__:__"
t_obs.Text = ""
Label6(1).Caption = ""
Combo1.ListIndex = -1
Label9.Caption = 0
laben.Caption = 0
labmed.Caption = 0
t_enff(1).Text = ""
t_med.Text = ""

End Function

Public Function habilita()
b_alta.Enabled = True
b_modiff(1).Enabled = True
b_grabaa(1).Enabled = False
b_cancee(1).Enabled = False
b_eli.Enabled = True
b_imp.Enabled = True
DBGrid13(0).Enabled = True
Frame1.Enabled = False
'Command2.Enabled = True

End Function

Public Function iguala()
If IsNull(data_busca.Recordset("codmed")) = False Then
   t_mov.Text = data_busca.Recordset("codmed")
Else
   t_mov.Text = ""
End If
If IsNull(data_busca.Recordset("medico")) = False Then
   t_choff(1).Text = data_busca.Recordset("medico")
Else
   t_choff(1).Text = ""
End If
If IsNull(data_busca.Recordset("zona")) = False Then
   Label9.Caption = data_busca.Recordset("zona")
Else
   Label9.Caption = ""
End If
If IsNull(data_busca.Recordset("fecha")) = False Then
   md.Text = data_busca.Recordset("fecha")
Else
   md.Text = "__/__/____"
End If
If IsNull(data_busca.Recordset("usuario")) = False Then
   mhord.Text = data_busca.Recordset("usuario")
Else
   mhord.Text = "__:__:__"
End If
If IsNull(data_busca.Recordset("fecmod")) = False Then
   mh.Text = data_busca.Recordset("fecmod")
Else
   mh.Text = "__/__/____"
End If
If IsNull(data_busca.Recordset("matricm")) = False Then
   mhorh.Text = data_busca.Recordset("matricm")
Else
   mhorh.Text = "__:__:__"
End If
If IsNull(data_busca.Recordset("nroseg")) = False Then
   Label6(1).Caption = data_busca.Recordset("nroseg")
Else
   Label6(1).Caption = ""
End If
If IsNull(data_busca.Recordset("kmactu")) = False Then
   Combo1.ListIndex = data_busca.Recordset("kmactu")
Else
   Combo1.ListIndex = -1
End If
If IsNull(data_busca.Recordset("motivo")) = False Then
   t_obs.Text = data_busca.Recordset("motivo")
Else
   t_obs.Text = ""
End If
If IsNull(data_busca.Recordset("ult_kms")) = False Then
   laben.Caption = data_busca.Recordset("ult_kms")
Else
   laben.Caption = 0
End If
If IsNull(data_busca.Recordset("pro_kms")) = False Then
   labmed.Caption = data_busca.Recordset("pro_kms")
   data_med(1).RecordSource = "Select * from medicos where med_cod =" & labmed.Caption
   data_med(1).Refresh
   If data_med(1).Recordset.RecordCount > 0 Then
      t_med.Text = data_med(1).Recordset("med_nombre")
   Else
      t_med.Text = ""
   End If
Else
   labmed.Caption = 0
   t_med.Text = ""
End If


End Function

Public Function hab()
txt_nrocob.Enabled = True
txt_nomcob.Enabled = True
txt_tel.Enabled = True
txt_espec.Enabled = True
t_ced.Enabled = True
t_codced.Enabled = True

End Function

Public Function desh()
txt_nrocob.Enabled = False
txt_nomcob.Enabled = False
txt_tel.Enabled = False
txt_espec.Enabled = False
t_ced.Enabled = False
t_codced.Enabled = False

End Function

Public Function igualcob()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("med_cod")) = False Then
   txt_nrocob.Text = data_cob.Recordset("med_cod")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("med_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("med_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("med_esp")) = False Then
   txt_espec.Text = data_cob.Recordset("med_esp")
Else
   txt_espec.Text = ""
End If
If IsNull(data_cob.Recordset("med_socnom")) = False Then
   txt_tel.Text = data_cob.Recordset("med_socnom")
Else
   txt_tel.Text = ""
End If

If txt_nrocob.Text <> "" Then
   data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
   data_medhc.Refresh
   If data_medhc.Recordset.RecordCount > 0 Then
      If IsNull(data_medhc.Recordset("m_nrofrm")) = False Then
         If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 7 Then
            t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 6)
            t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 7, 1)
         Else
            If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 8 Then
               t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 7)
               t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 8, 1)
            Else
               t_ced.Text = ""
               t_codced.Text = ""
            End If
         End If
      Else
         t_ced.Text = ""
         t_codced.Text = ""
      End If
      If IsNull(data_medhc.Recordset("m_codmed")) = False Then
         Check1.Value = data_medhc.Recordset("m_codmed")
      End If
   Else
      t_ced.Text = ""
      t_codced.Text = ""
   End If
End If

End Function

Private Sub txt_nrocob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomcob.SetFocus
End If

End Sub

Private Sub txt_nrocob_LostFocus()
If XAcnv = 1 Then
   Data1.Recordset.FindFirst "med_cod =" & txt_nrocob.Text
   If Not Data1.Recordset.NoMatch Then
      MsgBox "Ya existe este número de médico", vbCritical, "Médicos"
   End If
End If

End Sub

Public Function igualahs()
data_horasmed.RecordSource = "Select * from hc_archotro where id =" & data_verhsmed.Recordset("id")
data_horasmed.Refresh
If data_horasmed.Recordset.RecordCount > 0 Then
    If IsNull(data_horasmed.Recordset("id")) = False Then
       labid.Caption = data_horasmed.Recordset("id")
    Else
       labid.Caption = ""
    End If
    If IsNull(data_horasmed.Recordset("hc_fecha")) = False Then
       mfmed.Text = Format(data_horasmed.Recordset("hc_fecha"), "dd/mm/yyyy")
    Else
       mfmed.Text = "__/__/___"
    End If
    If IsNull(data_horasmed.Recordset("hc_hora")) = False Then
       mhmed.Text = Format(data_horasmed.Recordset("hc_hora"), "HH:mm")
    Else
       mhmed.Text = "__:__"
    End If
    If IsNull(data_horasmed.Recordset("hc_nro")) = False Then
       cboes.ListIndex = data_horasmed.Recordset("hc_nro")
    Else
       cboes.ListIndex = 0
    End If
    If IsNull(data_horasmed.Recordset("hc_mat")) = False Then
       t_codmedm.Text = data_horasmed.Recordset("hc_mat")
    Else
       t_codmedm.Text = ""
    End If
    If IsNull(data_horasmed.Recordset("hc_lugar")) = False Then
       Combo2.Text = data_horasmed.Recordset("hc_lugar")
    Else
       Combo2.ListIndex = -1
    End If
    data_hsmed2.RecordSource = "Select * from hc_viaae where id =" & data_horasmed.Recordset("id")
    data_hsmed2.Refresh
    If data_hsmed2.Recordset.RecordCount > 0 Then
       If IsNull(data_hsmed2.Recordset("hc_cod")) = False Then
          t_nromov.Text = data_hsmed2.Recordset("hc_cod")
       End If
    End If
End If

End Function

Public Function igualaturno()
If IsNull(data_grabatur.Recordset("cl_codigo")) = False Then
   labidtur.Caption = data_grabatur.Recordset("cl_codigo")
Else
   labidtur.Caption = ""
End If
If IsNull(data_grabatur.Recordset("cl_fnac")) = False Then
   mf.Text = Format(data_grabatur.Recordset("cl_fnac"), "dd/mm/yyyy")
Else
   mf.Text = "__/__/___"
End If
If IsNull(data_grabatur.Recordset("cl_ruc")) = False Then
   mhor.Text = Format(data_grabatur.Recordset("cl_ruc"), "HH:mm")
Else
   mhor.Text = "__:__"
End If
If IsNull(data_grabatur.Recordset("cl_fultmov")) = False Then
   mffin.Text = Format(data_grabatur.Recordset("cl_fultmov"), "dd/mm/yyyy")
Else
   mffin.Text = "__/__/____"
End If
If IsNull(data_grabatur.Recordset("cl_fax")) = False Then
   mhfin.Text = Format(data_grabatur.Recordset("cl_fax"), "HH:mm")
Else
   mf.Text = "__:__"
End If
If IsNull(data_grabatur.Recordset("cl_descpag")) = False Then
   labu.Caption = data_grabatur.Recordset("cl_descpag")
Else
   labu.Caption = ""
End If
If IsNull(data_grabatur.Recordset("info_debit")) = False Then
   t_obsturno.Text = data_grabatur.Recordset("info_debit")
Else
   t_obsturno.Text = ""
End If
If IsNull(data_grabatur.Recordset("cl_atrasoa")) = False Then
   Check2.Value = data_grabatur.Recordset("cl_atrasoa")
Else
   Check2.Value = 0
End If
If IsNull(data_grabatur.Recordset("cl_nom_sup")) = False Then
   labnomu.Caption = data_grabatur.Recordset("cl_nom_sup")
Else
   labnomu.Caption = ""
End If
If Check2.Value = 1 Then
   t_obsturno.Enabled = False
   b_gratur.Enabled = False
Else
   t_obsturno.Enabled = True
   b_gratur.Enabled = True
End If

End Function


Public Sub Consultar_CMT_total()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select * from llamado where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and (pend in (4) or movilpas in (2015)) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null)"
Else
   Xsqlpromo = "Select * from llamado where fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and (pend in (4) or movilpas in (2015)) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null)"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labcmtdesp.Caption = Xrecclii.RecordCount
Else
   labcmtdesp.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub Consultar_CMT_PendDesp()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open

If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select * from llamado where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and pend in (4) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null)"
Else
   Xsqlpromo = "Select * from llamado where fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and pend in (4) and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null)"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveLast
   labpenddesp.Caption = Xrecclii.RecordCount
Else
   labpenddesp.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub Consultar_CMT_Realiza()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xelmedico As String
Dim Xconteo As Integer
Xconteo = 0
Xelmedico = ""
ConectarBD
ConbdSapp.Open
List1.Clear

If mdcmt.Text = "__/__/____" Then
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
Else
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha >=#" & Format(mdcmt.Text, "yyyy/mm/dd") & "# and medicos_cmt.fecha <=#" & Format(mhcmt.Text, "yyyy/mm/dd") & "#"
End If
data_listamedicos.Refresh
If data_listamedicos.Recordset.RecordCount > 0 Then
   data_listamedicos.Recordset.MoveFirst
   Do While Not data_listamedicos.Recordset.EOF
      If mdcmt.Text = "__/__/____" Then
         Xsqlpromo = "Select llamado.nrolla,llamado.fecha,llamado.movilpas,llamado.pend,llamado.codmot,llamado.segui_covid,resplla.nro,resplla.timdes from llamado inner join " & _
         "resplla on llamado.nrolla=resplla.nro where llamado.fecha ='" & Format(Date, "yyyy-mm-dd") & "' and llamado.movilpas in (2015) and llamado.codmot <>'" & "Z" & "' and (llamado.segui_covid not in (1) or llamado.segui_covid is null) and resplla.timdes='" & data_listamedicos.Recordset("nom_usuario") & "'"
      Else
         Xsqlpromo = "Select llamado.nrolla,llamado.fecha,llamado.movilpas,llamado.pend,llamado.codmot,llamado.segui_covid,resplla.nro,resplla.timdes from llamado inner join " & _
         "resplla on llamado.nrolla=resplla.nro where llamado.fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and llamado.movilpas in (2015) and llamado.codmot <>'" & "Z" & "' and (llamado.segui_covid not in (1) or llamado.segui_covid is null) and resplla.timdes='" & data_listamedicos.Recordset("nom_usuario") & "'"
      End If
      With Xrecclii
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Xsqlpromo, ConbdSapp, , , adCmdText
      End With
      If Xrecclii.RecordCount > 0 Then
         Xrecclii.MoveLast
         List1.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & Xrecclii.RecordCount
      Else
         List1.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & 0
      End If
      data_listamedicos.Recordset.MoveNext
      Xrecclii.Close
   Loop
End If

ConbdSapp.Close

End Sub


Public Function Devuelve_Usuario() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim t_cedula As String

t_cedula = t_ced.Text & t_codced.Text

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from cap_ciap where cod_cap ='" & Trim(t_cedula) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_Usuario = Xrecclii("des_cap")
Else
   Devuelve_Usuario = ""
   MsgBox "No figura documento del médico en tabla de usuarios. No sincronizará datos.", vbCritical
End If

Xrecclii.Close
ConbdSapp.Close


End Function

Public Sub Consultar_CMT_poli()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select * from linmmdd where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cod_prod in (10050)"
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
Else
   Xsqlpromo = "Select * from linmmdd where fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and cod_prod in (10050)"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labcmtpol.Caption = Xrecclii.RecordCount
Else
   labcmtpol.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub Consultar_CMT_posit()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open

If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select * from llamado where fecha >='" & Format("01/07/2019", "yyyy-mm-dd") & "' and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <='" & Format(Date, "yyyy-mm-dd") & "' or prox_control is null)"
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
Else
   Xsqlpromo = "Select llamado.nrolla,llamado.segui_covid,llamado.cierre_hora,llamado.isopa_result,llamado.resuliso2,llamado.prox_control,seguimiento_covid.id_llamado,seguimiento_covid.fecha,seguimiento_covid.nom_usu from seguimiento_covid inner join llamado on seguimiento_covid.id_llamado=llamado.nrolla where seguimiento_covid.fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and seguimiento_covid.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and llamado.segui_covid in (1) and (isopa_result in ('Positivo') or resuliso2 in ('Positivo'))"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labseg.Caption = Xrecclii.RecordCount
Else
   labseg.Caption = "0"
End If

Xrecclii.Close

If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "select llamado.nrolla,llamado.fecha,llamado.segui_covid,llamado.cierre_hora,llamado.isopa_result,llamado.resuliso2,llamado.prox_control,seguimiento_covid.id_llamado,seguimiento_covid.fecha,seguimiento_covid.nom_usu from llamado inner join seguimiento_covid on llamado.nrolla=seguimiento_covid.id_llamado where llamado.segui_covid in (1) and llamado.cierre_hora is null and (llamado.isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and seguimiento_covid.fecha ='" & Format(Date, "yyyy-mm-dd") & "'"
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
Else
   Xsqlpromo = "select llamado.nrolla,llamado.fecha,llamado.segui_covid,llamado.cierre_hora,llamado.isopa_result,llamado.resuliso2,llamado.prox_control,seguimiento_covid.id_llamado,seguimiento_covid.fecha,seguimiento_covid.nom_usu from llamado inner join seguimiento_covid on llamado.nrolla=seguimiento_covid.id_llamado where llamado.segui_covid in (1) and llamado.cierre_hora is null and (llamado.isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and seguimiento_covid.fecha ='" & Format(Date, "yyyy-mm-dd") & "'"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labseg.Caption = Val(labseg.Caption) + Xrecclii.RecordCount
End If

Xrecclii.Close

ConbdSapp.Close

End Sub

Public Sub Consultar_CMT_poliPend()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pendientes As Integer
Pendientes = 0
ConectarBD
ConbdSapp.Open

If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select linmmdd.cod_cli,linmmdd.fecha,linmmdd.cod_prod,cabezal_hcdig.mat,cabezal_hcdig.fecha from linmmdd inner join cabezal_hcdig on linmmdd.cod_cli=cabezal_hcdig.mat where linmmdd.fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cabezal_hcdig.fecha ='" & Format(Date, "yyyy-mm-dd") & "' and linmmdd.cod_prod in (10050)"
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
Else
   Xsqlpromo = "Select linmmdd.cod_cli,linmmdd.fecha,linmmdd.cod_prod,cabezal_hcdig.mat,cabezal_hcdig.fecha from linmmdd inner join cabezal_hcdig on linmmdd.cod_cli=cabezal_hcdig.mat where linmmdd.fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and cabezal_hcdig.fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and cabezal_hcdig.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and linmmdd.cod_prod in (10050)"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Pendientes = Val(labcmtpol.Caption) - Xrecclii.RecordCount
   labpendpol.Caption = Pendientes
Else
   labpendpol.Caption = labcmtpol.Caption
End If
If mdcmt.Text <> "__/__/____" Then
   labpendpol.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consultar_CMT_PolRealiza()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xelmedico As String
Dim Xconteo As Integer
Xconteo = 0
Xelmedico = ""
ConectarBD
ConbdSapp.Open
List2.Clear

If mdcmt.Text = "__/__/____" Then
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
Else
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha >=#" & Format(mdcmt.Text, "yyyy/mm/dd") & "# and medicos_cmt.fecha <=#" & Format(mhcmt.Text, "yyyy/mm/dd") & "#"
End If
data_listamedicos.Refresh
If data_listamedicos.Recordset.RecordCount > 0 Then
   data_listamedicos.Recordset.MoveFirst
   Do While Not data_listamedicos.Recordset.EOF
      Xsqlpromo = "Select * from linmmdd where nro_med_a =" & data_listamedicos.Recordset("cod_med") & " and fecha ='" & Format(data_listamedicos.Recordset("fecha"), "yyyy-mm-dd") & "' and cod_prod in (10050)"
      With Xrecclii
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Xsqlpromo, ConbdSapp, , , adCmdText
      End With
      If Xrecclii.RecordCount > 0 Then
         Xrecclii.MoveLast
         List2.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & Xrecclii.RecordCount
      Else
         List2.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & 0
      End If
      data_listamedicos.Recordset.MoveNext
      Xrecclii.Close
   Loop
End If

ConbdSapp.Close

End Sub


Public Sub Consultar_MG_Poli()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pendientes As Integer
Pendientes = 0
ConectarBD
ConbdSapp.Open

If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select linmmdd.fecha,linmmdd.cod_prod from linmmdd where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cod_prod in (10001)"
Else
   Xsqlpromo = "Select linmmdd.fecha,linmmdd.cod_prod from linmmdd where fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and cod_prod in (10001)"
End If
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labtotmg.Caption = Xrecclii.RecordCount
Else
   labtotmg.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consultar_PolRealizaMG()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xelmedico As String
Dim Xconteo As Integer
Xconteo = 0
Xelmedico = ""
ConectarBD
ConbdSapp.Open
List5.Clear

If mdcmt.Text = "__/__/____" Then
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
Else
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha >=#" & Format(mdcmt.Text, "yyyy/mm/dd") & "# and medicos_cmt.fecha <=#" & Format(mhcmt.Text, "yyyy/mm/dd") & "#"
End If
data_listamedicos.Refresh
If data_listamedicos.Recordset.RecordCount > 0 Then
   data_listamedicos.Recordset.MoveFirst
   Do While Not data_listamedicos.Recordset.EOF
      Xsqlpromo = "Select * from linmmdd where nro_med_a =" & data_listamedicos.Recordset("cod_med") & " and fecha ='" & Format(data_listamedicos.Recordset("fecha"), "yyyy-mm-dd") & "' and cod_prod in (10001)"
      With Xrecclii
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Xsqlpromo, ConbdSapp, , , adCmdText
      End With
      If Xrecclii.RecordCount > 0 Then
         Xrecclii.MoveLast
         List5.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & Xrecclii.RecordCount
      Else
         List5.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & 0
      End If
      data_listamedicos.Recordset.MoveNext
      Xrecclii.Close
   Loop
End If

ConbdSapp.Close

End Sub

Public Sub Consultar_poliPend_MG()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pendientes As Integer
Pendientes = 0
ConectarBD
ConbdSapp.Open
If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select linmmdd.cod_cli,linmmdd.fecha,linmmdd.cod_prod,cabezal_hcdig.mat,cabezal_hcdig.fecha from linmmdd inner join cabezal_hcdig on linmmdd.cod_cli=cabezal_hcdig.mat where linmmdd.fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cabezal_hcdig.fecha ='" & Format(Date, "yyyy-mm-dd") & "' and linmmdd.cod_prod in (10001)"
Else
   Xsqlpromo = "Select linmmdd.cod_cli,linmmdd.fecha,linmmdd.cod_prod,cabezal_hcdig.mat,cabezal_hcdig.fecha from linmmdd inner join cabezal_hcdig on linmmdd.cod_cli=cabezal_hcdig.mat where linmmdd.fecha >='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and cabezal_hcdig.fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and cabezal_hcdig.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and linmmdd.cod_prod in (10001)"
End If
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Pendientes = Val(labtotmg.Caption) - Xrecclii.RecordCount
   labpendmg.Caption = Pendientes
Else
   labpendmg.Caption = labtotmg.Caption
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consultar_Positivos_Realiza()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xelmedico As String
Dim Xconteo As Integer
Xconteo = 0
Xelmedico = ""
ConectarBD
ConbdSapp.Open
List3.Clear

If mdcmt.Text = "__/__/____" Then
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
Else
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha >=#" & Format(mdcmt.Text, "yyyy/mm/dd") & "# and medicos_cmt.fecha <=#" & Format(mhcmt.Text, "yyyy/mm/dd") & "#"
End If
data_listamedicos.Refresh
If data_listamedicos.Recordset.RecordCount > 0 Then
   data_listamedicos.Recordset.MoveFirst
   Do While Not data_listamedicos.Recordset.EOF
'      Xsqlpromo = "Select * from llamado where fecha >='" & Format("01/07/2019", "yyyy-mm-dd") & "' and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <='" & Format(Date, "yyyy-mm-dd") & "' or prox_control is null)"
      If mdcmt.Text = "__/__/____" Then
         Xsqlpromo = "Select llamado.nrolla,llamado.fecha,llamado.movilpas,llamado.pend,llamado.codmot,llamado.segui_covid,llamado.cierre_hora,llamado.isopa_result,llamado.prox_control,seguimiento_covid.id_llamado,seguimiento_covid.fecha,seguimiento_covid.nom_usu from llamado inner join " & _
         "seguimiento_covid on llamado.nrolla=seguimiento_covid.id_llamado where seguimiento_covid.fecha ='" & Format(Date, "yyyy-mm-dd") & "' and llamado.segui_covid in (1) and (llamado.isopa_result in ('Positivo') or llamado.resuliso2 in ('Positivo')) and seguimiento_covid.nom_usu='" & data_listamedicos.Recordset("nom_usuario") & "'"
      Else
         Xsqlpromo = "Select llamado.nrolla,llamado.segui_covid,llamado.cierre_hora,llamado.isopa_result,llamado.resuliso2,llamado.prox_control,seguimiento_covid.id_llamado,seguimiento_covid.fecha,seguimiento_covid.nom_usu from seguimiento_covid inner join llamado on seguimiento_covid.id_llamado=llamado.nrolla where seguimiento_covid.fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and seguimiento_covid.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and llamado.segui_covid in (1) and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and seguimiento_covid.nom_usu='" & data_listamedicos.Recordset("nom_usuario") & "'"
'         Xsqlpromo = "Select llamado.nrolla,llamado.fecha,llamado.movilpas,llamado.pend,llamado.codmot,llamado.segui_covid,llamado.cierre_hora,llamado.isopa_result,llamado.prox_control,seguimiento_covid.id_llamado,seguimiento_covid.fecha,seguimiento_covid.nom_usu from llamado inner join " & _
'         "seguimiento_covid on llamado.nrolla=seguimiento_covid.id_llamado where seguimiento_covid.fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and seguimiento_covid.fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and llamado.segui_covid in (1) and (llamado.isopa_result in ('Positivo') or llamado.resuliso2 in ('Positivo')) and seguimiento_covid.nom_usu='" & data_listamedicos.Recordset("nom_usuario") & "'"
      End If
      With Xrecclii
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Xsqlpromo, ConbdSapp, , , adCmdText
      End With
      If Xrecclii.RecordCount > 0 Then
         Xrecclii.MoveLast
         List3.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & Xrecclii.RecordCount
      Else
         List3.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & 0
      End If
      data_listamedicos.Recordset.MoveNext
      Xrecclii.Close
   Loop
End If

ConbdSapp.Close

End Sub


Public Sub Consultar_Pend_posit()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pendientes As Integer
Pendientes = 0
ConectarBD
ConbdSapp.Open

   Xsqlpromo = "Select * from llamado where fecha >='" & Format("01/07/2019", "yyyy/mm/dd") & "' and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <='" & Format(Date, "yyyy/mm/dd") & "' or prox_control is null)"
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labpendposi.Caption = Xrecclii.RecordCount
Else
   labpendposi.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consultar_Results()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pendientes As Integer
Pendientes = 0
ConectarBD
ConbdSapp.Open

If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "select * from sol_hisopos where deriva in (1) and (fecha_cierre is null or fecha_cierre ='" & Format(Date, "yyyy-mm-dd") & "')"
Else
   Xsqlpromo = "Select * from sol_hisopos where fecha_cierre >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and fecha_cierre <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "'"
End If
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labtotres.Caption = Xrecclii.RecordCount
Else
   labtotres.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consultar_ResPend()
         
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
ConectarBD
ConbdSapp.Open
If mdcmt.Text = "__/__/____" Then
   Xsqlpromo = "Select * from sol_hisopos where fecha_cierre is null and deriva in (1)"
Else
   Xsqlpromo = "select * from sol_hisopos where fecha_cierre >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and fecha_cierre <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "'"
End If
'                     Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2019", "yyyy/mm/dd") & "# and segui_covid in (1) and cierre_hora is null and (isopa_result in ('Positivo') or resuliso2 in ('Positivo')) and (prox_control <=#" & Format(Date, "yyyy/mm/dd") & "# or prox_control is null) order by fecha,hora"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labpendres.Caption = Xrecclii.RecordCount
Else
   labpendres.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consultar_Resulta_Realiza()
         

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xelmedico As String
Dim Xconteo As Integer
Xconteo = 0
Xelmedico = ""
ConectarBD
ConbdSapp.Open
List6.Clear

If mdcmt.Text = "__/__/____" Then
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
Else
   data_listamedicos.RecordSource = "select medicos_cmt.cod_med,medicos_cmt.nom_usuario,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha >=#" & Format(mdcmt.Text, "yyyy/mm/dd") & "# and medicos_cmt.fecha <=#" & Format(mhcmt.Text, "yyyy/mm/dd") & "#"
End If
data_listamedicos.Refresh
If data_listamedicos.Recordset.RecordCount > 0 Then
   data_listamedicos.Recordset.MoveFirst
   Do While Not data_listamedicos.Recordset.EOF
      If mdcmt.Text = "__/__/____" Then
         Xsqlpromo = "Select * from sol_hisopos where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and nom_usu ='" & data_listamedicos.Recordset("nom_usuario") & "'"
      Else
         Xsqlpromo = "Select * from sol_hisopos where fecha >='" & Format(mdcmt.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mhcmt.Text, "yyyy-mm-dd") & "' and nom_usu ='" & data_listamedicos.Recordset("nom_usuario") & "'"
      End If
      With Xrecclii
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Xsqlpromo, ConbdSapp, , , adCmdText
      End With
      If Xrecclii.RecordCount > 0 Then
         Xrecclii.MoveLast
         List6.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & Xrecclii.RecordCount
      Else
         List6.AddItem data_listamedicos.Recordset("med_nombre") & " Total:" & 0
      End If
      data_listamedicos.Recordset.MoveNext
      Xrecclii.Close
   Loop
End If

ConbdSapp.Close

End Sub



Public Sub Consultar_medicosCMT()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

List4.Clear
ConectarBD
ConbdSapp.Open

If mdmed.Text = "__/__/____" Then
   Xsqlpromo = "select medicos_cmt.cod_med,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha ='" & Format(Date, "yyyy-mm-dd") & "'"
Else
   Xsqlpromo = "select medicos_cmt.cod_med,medicos_cmt.cedula,medicos_cmt.fecha,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha ='" & Format(mdmed.Text, "yyyy-mm-dd") & "'"
End If

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      List4.AddItem Xrecclii("cod_med") & "--" & Xrecclii("med_nombre") & " Ced." & Xrecclii("cedula")
      Xrecclii.MoveNext
   Loop
End If
Xrecclii.Close
ConbdSapp.Close


End Sub
