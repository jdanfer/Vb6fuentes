VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   Caption         =   "Sistema de gestión para la salud"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   11910
   Icon            =   "frm_menu.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox t_info 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1815
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      ToolTipText     =   "Doble click sobre el botón de info para ocultar"
      Top             =   1560
      Visible         =   0   'False
      Width           =   7815
   End
   Begin Crystal.CrystalReport crcontrol 
      Left            =   5160
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_simulafin 
      Caption         =   "data_simulafin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_simulaini 
      Caption         =   "data_simulaini"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_convrutas 
      Caption         =   "data_convrutas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_conscliprom 
      Caption         =   "data_conscliprom"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_promos 
      Caption         =   "data_promos"
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
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_redpagos 
      Caption         =   "data_redpagos"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin Crystal.CrystalReport crpro 
      Left            =   5160
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc adoconvpromo 
      Height          =   330
      Left            =   240
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
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
      Caption         =   "adoconvpromo"
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
   Begin MSAdodcLib.Adodc adoclipromo 
      Height          =   330
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
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
      Caption         =   "adoclipromo"
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
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc adoctrrutas 
      Height          =   330
      Left            =   8760
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      Caption         =   "adoctrrutas"
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
   Begin MSAdodcLib.Adodc adologueo 
      Height          =   375
      Left            =   4560
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "adologueo"
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
   Begin VB.Data data_ctrabre 
      Caption         =   "data_ctrabre"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data_cab 
      Height          =   375
      Left            =   3240
      Top             =   1680
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
      Caption         =   "data_cab"
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
   Begin MSAdodcLib.Adodc data_cnv 
      Height          =   495
      Left            =   5160
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Caption         =   "data_cnv"
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
   Begin MSMask.MaskEdBox mfh 
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mfd 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   8280
      Top             =   1560
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
      Caption         =   "data_lin"
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
   Begin VB.CommandButton env_fac 
      Caption         =   "env_fac"
      Height          =   495
      Left            =   9960
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data data_ctrenvfac 
      Caption         =   "data_ctrenvfac"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc data_sql 
      Height          =   375
      Left            =   8880
      Top             =   2760
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
      Caption         =   "data_sql"
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
   Begin MSAdodcLib.Adodc data_mejor 
      Height          =   615
      Left            =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
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
      Caption         =   "data_mejor"
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
   Begin VB.Data data_ctrlfact 
      Caption         =   "data_ctrlfact"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5160
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_mdb 
      Caption         =   "data_mdb"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_ucod 
      Caption         =   "data_ucod"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_mejo2 
      Caption         =   "data_mejo2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_cargo 
      Caption         =   "data_cargo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_cargos 
      Caption         =   "data_cargos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_fecha 
      Caption         =   "data_fecha"
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
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Data data_parse 
      Caption         =   "data_parse"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PARSEC0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_usuac 
      Caption         =   "data_usuac"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   6000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10080
      Top             =   3840
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   9720
      MouseIcon       =   "frm_menu.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frm_menu.frx":1254
      Stretch         =   -1  'True
      ToolTipText     =   "Doble click para ver modificaciones en última versión"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aguarde, procesando..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aguarde, procesando..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   3120
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   6840
      Width           =   4095
   End
   Begin VB.Label labempre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label labusua 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Usuario Actual"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label labhora 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   8760
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   5640
      Picture         =   "frm_menu.frx":1996
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Menu mnusocios 
      Caption         =   "Socios"
      Begin VB.Menu mnumant 
         Caption         =   "Mantenimiento"
      End
      Begin VB.Menu mnuafilia 
         Caption         =   "Afiliación nueva"
      End
      Begin VB.Menu mnuctrles 
         Caption         =   "Controles"
         Begin VB.Menu menupeddomi 
            Caption         =   "Medicación a domicilio"
         End
         Begin VB.Menu mnuctrolb 
            Caption         =   "Control medicación base"
         End
         Begin VB.Menu mnuctrocons 
            Caption         =   "Control de consultas"
         End
         Begin VB.Menu mnuctrolen 
            Caption         =   "Control Actos enfermería"
         End
         Begin VB.Menu mnuctrolme 
            Caption         =   "Control de medicación"
         End
         Begin VB.Menu mnuinfmedxmut 
            Caption         =   "Informes de medicación"
         End
         Begin VB.Menu mnuregmedctr 
            Caption         =   "Registros Médicos"
            Begin VB.Menu mnusolhcctr 
               Caption         =   "Solicitudes de HC"
            End
            Begin VB.Menu mnuctrolenvhcctr 
               Caption         =   "Control Envío Sol.HC"
            End
            Begin VB.Menu mnuctrolenthcctr 
               Caption         =   "Control Entrega HC"
            End
         End
      End
      Begin VB.Menu mnuadm 
         Caption         =   "Administración"
         Begin VB.Menu mnutarj 
            Caption         =   "Personal p/tarjeta BROU"
         End
         Begin VB.Menu mnusuel 
            Caption         =   "Sueldos al BROU"
         End
         Begin VB.Menu mnupresta 
            Caption         =   "Préstamos BROU"
         End
      End
      Begin VB.Menu mnuctacte 
         Caption         =   "Cuentas corrientes"
         Begin VB.Menu mnuproc 
            Caption         =   "Procesos"
            Begin VB.Menu mnuemisim 
               Caption         =   "Emisión simulada"
            End
            Begin VB.Menu mnuemimen 
               Caption         =   "Emisión mensual"
            End
            Begin VB.Menu mnuctrrutas 
               Caption         =   "Control de Rutas"
            End
            Begin VB.Menu mnuentnew 
               Caption         =   "Nuevas entregas"
            End
            Begin VB.Menu mnuprocdebweb 
               Caption         =   "Proceso débitos automáticos"
            End
            Begin VB.Menu mnuactdeu 
               Caption         =   "Actualizar deuda individual"
            End
            Begin VB.Menu mnudeuemi 
               Caption         =   "Pasar deudas emisión"
            End
            Begin VB.Menu mnubajbas 
               Caption         =   "Procesar Bajas de Base"
            End
            Begin VB.Menu proctimbres 
               Caption         =   "Procesar timbres para emisión"
            End
            Begin VB.Menu mnuprocred 
               Caption         =   "Procesar ventas CREDITO a emisión"
            End
            Begin VB.Menu mnucaremdeu 
               Caption         =   "Cargar emisión a Deudas"
            End
            Begin VB.Menu mnumodbases 
               Caption         =   "Afiliaciones y modificaciones en base"
            End
            Begin VB.Menu mnuconsemisocnew 
               Caption         =   "Consultar emisión por socio"
            End
            Begin VB.Menu mnuprocmutnew 
               Caption         =   "Procesar padrones mutuales"
            End
         End
         Begin VB.Menu mnuarq 
            Caption         =   "Arqueo"
            Begin VB.Menu mnugenarq 
               Caption         =   "Generar arqueo"
            End
            Begin VB.Menu mnupaspen 
               Caption         =   "Pasar COBRADOS"
            End
            Begin VB.Menu menupasbaj 
               Caption         =   "Pasar Bajas"
            End
            Begin VB.Menu mnupasdev 
               Caption         =   "Pasar Devoluciones"
            End
            Begin VB.Menu mnupaspendos 
               Caption         =   "Pasar Pendientes"
            End
            Begin VB.Menu mnuentregas 
               Caption         =   "Cargar Entregas"
            End
            Begin VB.Menu mnucobrarq 
               Caption         =   "Cobradores"
            End
            Begin VB.Menu mnucarfac 
               Caption         =   "Cargar nuevas entregas"
            End
            Begin VB.Menu mnuconshisarq 
               Caption         =   "Consultar historial"
            End
            Begin VB.Menu mnuinfarq 
               Caption         =   "Informes arqueos"
            End
            Begin VB.Menu mnucerrarar 
               Caption         =   "CERRAR Arqueo"
            End
            Begin VB.Menu mnurepag 
               Caption         =   "Procesar archivo RedPagos"
            End
            Begin VB.Menu mnucboredp 
               Caption         =   "Generar archivo para RedPagos"
            End
         End
      End
      Begin VB.Menu mnuatsoc 
         Caption         =   "Atención al socio"
         Begin VB.Menu mnucrm 
            Caption         =   "CRM"
         End
         Begin VB.Menu mnuencue 
            Caption         =   "Encuestas"
         End
      End
   End
   Begin VB.Menu mnuest 
      Caption         =   "Estudios"
      Begin VB.Menu mnumane 
         Caption         =   "ABM Estudios"
      End
   End
   Begin VB.Menu mnucaja 
      Caption         =   "Caja"
      Begin VB.Menu mnuingcaj 
         Caption         =   "Ingreso de caja"
      End
      Begin VB.Menu mnureimp 
         Caption         =   "Re imprimir Factura"
      End
      Begin VB.Menu mnurubcaj 
         Caption         =   "Rubros de caja"
      End
      Begin VB.Menu mnurubgral 
         Caption         =   "Rubros generales"
      End
      Begin VB.Menu mnutes 
         Caption         =   "Caja Tesorería"
      End
      Begin VB.Menu mnucomer 
         Caption         =   "Mantenimiento Comercios"
      End
   End
   Begin VB.Menu mnuenfer 
      Caption         =   "Enfermería"
      Begin VB.Menu mnumatestenf 
         Caption         =   "Material para esterilizar"
      End
      Begin VB.Menu mnuctractenf 
         Caption         =   "Control actos de enfermería"
      End
      Begin VB.Menu mnuverllaenf 
         Caption         =   "Ver llamados a domicilio"
      End
      Begin VB.Menu mnuelechce 
         Caption         =   "Cargar Electros a HCE"
      End
      Begin VB.Menu mnuvences 
         Caption         =   "Vencimientos"
      End
      Begin VB.Menu mnusrvenfdom 
         Caption         =   "Servicios Enfermeria a domicilio"
      End
      Begin VB.Menu mnucmtpend 
         Caption         =   "Pendientes CMT y Polic MG"
      End
      Begin VB.Menu mnuresconstot 
         Caption         =   "Resumen de consultas"
      End
   End
   Begin VB.Menu mnumetasreg 
      Caption         =   "Metas"
      Begin VB.Menu mnuregactmetas 
         Caption         =   "Registro de actos"
      End
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "Informes"
      Begin VB.Menu mnuinfadm 
         Caption         =   "Administración (FP08)"
         Begin VB.Menu mnucompadm 
            Caption         =   "Cómputos"
            Begin VB.Menu mnucompsina 
               Caption         =   "Informes para SINADI"
            End
            Begin VB.Menu mnuinfsmsnew 
               Caption         =   "Informes para Envíos de SMS"
            End
            Begin VB.Menu mnurucaf 
               Caption         =   "Informes al RUCAF"
            End
         End
         Begin VB.Menu menupsocnew 
            Caption         =   "Padrón Social"
            Begin VB.Menu mnuinfabmpsnew 
               Caption         =   "Altas, Bajas, Modificaciones de socios"
            End
            Begin VB.Menu mnuinfemisnew 
               Caption         =   "Informes de emisión"
            End
            Begin VB.Menu mnuimpeminew 
               Caption         =   "Imprimir emisión"
            End
            Begin VB.Menu mnuinfprodcnvnew 
               Caption         =   "Producción por Convenios/Zonas"
            End
         End
         Begin VB.Menu menuvtaservnew 
            Caption         =   "Ventas/Servicios"
            Begin VB.Menu mnuvtaxsernew 
               Caption         =   "Ventas por Servicio"
            End
            Begin VB.Menu mnuvtasxflianew 
               Caption         =   "Ventas por Familia"
            End
            Begin VB.Menu mnuvtasxmednew 
               Caption         =   "Ventas por Médico"
            End
            Begin VB.Menu mnuvtasxconvnew 
               Caption         =   "Ventas por Convenios"
            End
            Begin VB.Menu mnuvtasxmutnew 
               Caption         =   "Ventas por Mutualista"
            End
            Begin VB.Menu mnuvtasxtipfacnew 
               Caption         =   "Ventas por tipo de factura"
            End
            Begin VB.Menu mnuvtascredadm 
               Caption         =   "Ventas a crédito (Administración)"
            End
            Begin VB.Menu mnullaconbmnew 
               Caption         =   "Llamados c/costo (Boleta manual)"
            End
            Begin VB.Menu mnuliqespnew 
               Caption         =   "Liquidación de especialistas"
            End
            Begin VB.Menu mnufaccnvnew2 
               Caption         =   "Facturación de convenios"
            End
            Begin VB.Menu mnuinfgestion 
               Caption         =   "Informes Gestión cobranza"
            End
            Begin VB.Menu mnuplacmt 
               Caption         =   "Planilla mensual CMT"
            End
         End
      End
      Begin VB.Menu mnuinfdesp33 
         Caption         =   "Despacho (FP10)"
         Begin VB.Menu mnuinfdesp1 
            Caption         =   "Informes Despacho (predefinidos)"
         End
         Begin VB.Menu infdespnuevo 
            Caption         =   "Informes Despacho (Selección)"
         End
         Begin VB.Menu mnuinfazuln 
            Caption         =   "Informes CODIGO AZUL"
         End
         Begin VB.Menu mnuinfdemorarl 
            Caption         =   "Demoras Receptor/Largador"
         End
         Begin VB.Menu mnuinfsca 
            Caption         =   "Informes SCA"
         End
      End
      Begin VB.Menu mnueconom 
         Caption         =   "Compras, Almacenamiento, Entrega (FP06)"
         Begin VB.Menu mnuinfmednew 
            Caption         =   "Informes de medicación"
         End
         Begin VB.Menu mnuinfstocknew 
            Caption         =   "Informes de Stock"
         End
      End
      Begin VB.Menu mnumark 
         Caption         =   "Marketing (FP09)"
         Begin VB.Menu mnusocpromnew 
            Caption         =   "Socios por promotor (con UMP)"
         End
         Begin VB.Menu mnusocxmutnew 
            Caption         =   "Socios ACTIVOS por mutualista"
         End
         Begin VB.Menu mnuinfabnew 
            Caption         =   "Socios (Altas/Bajas)"
         End
         Begin VB.Menu mnuinfsocssnew 
            Caption         =   "Socios sin servicios"
         End
         Begin VB.Menu mnuinfdeudanew 
            Caption         =   "Deudas por socio"
         End
         Begin VB.Menu mnuinfestudnew 
            Caption         =   "Precios de estudios"
         End
         Begin VB.Menu mnumodyafnew 
            Caption         =   "Afiliación y modificaciones en BASE"
         End
         Begin VB.Menu mnuservmutnew 
            Caption         =   "Servicios por mutualista"
         End
         Begin VB.Menu mnucirepe 
            Caption         =   "Socios con CI repetidos"
         End
         Begin VB.Menu mnuconssoc 
            Caption         =   "Consultas por socio"
         End
         Begin VB.Menu mnucartasm 
            Caption         =   "Cartas mutuales"
         End
         Begin VB.Menu mnuvtaserusu 
            Caption         =   "Ventas por servicio con datos usuario"
         End
      End
      Begin VB.Menu mnuinfrecenf 
         Caption         =   "Enfermería (FP13)"
         Begin VB.Menu mnuvtaslab 
            Caption         =   "Laboratorios"
         End
         Begin VB.Menu mnuinfctrhcemov 
            Caption         =   "Informes control HC MOVILES"
         End
      End
      Begin VB.Menu mnumetasis 
         Caption         =   "Metas Asisteciales (FP16)"
         Begin VB.Menu mnuinfedades 
            Caption         =   "Informe por rangos de edad"
         End
         Begin VB.Menu mnuinfhcemet 
            Caption         =   "Informes desde HCE"
         End
      End
      Begin VB.Menu mnuinfcali22 
         Caption         =   "Calidad"
         Begin VB.Menu mnuinfdemnew 
            Caption         =   "Demoras de llamados por clave"
         End
         Begin VB.Menu mnuinfdemmnew 
            Caption         =   "Demoras de llamados (por médico)"
         End
         Begin VB.Menu mnuinfctroltpnew 
            Caption         =   "Control de tiempos en policlínica"
         End
         Begin VB.Menu mnuinfprmednew 
            Caption         =   "Productividad médicos"
         End
         Begin VB.Menu mnuselconalenew 
            Caption         =   "Selección de consultas aleatorias"
         End
         Begin VB.Menu mnuconsres 
            Caption         =   "Resumen de consultas"
         End
      End
   End
   Begin VB.Menu mnures 
      Caption         =   "Reservas"
      Begin VB.Menu mnuespecnew 
         Caption         =   "Fechas Especialistas"
      End
      Begin VB.Menu mnureshnf 
         Caption         =   "Reserva realización HNF"
      End
   End
   Begin VB.Menu mnuecon 
      Caption         =   "Economato"
      Begin VB.Menu mnuregpro 
         Caption         =   "Registro de productos"
      End
      Begin VB.Menu mnuingcomp 
         Caption         =   "Ingreso de compras"
      End
      Begin VB.Menu mnureggas 
         Caption         =   "Registro de entregas-gastos"
      End
      Begin VB.Menu mnulabcom 
         Caption         =   "Laboratorios/Comercios"
      End
      Begin VB.Menu mnuclieco 
         Caption         =   "Clientes (ABM)"
      End
      Begin VB.Menu mnuinfeco 
         Caption         =   "Informes Economato"
      End
   End
   Begin VB.Menu mnuconta 
      Caption         =   "Contabilidad"
      Begin VB.Menu mnuprocasi 
         Caption         =   "Procesar asientos"
      End
      Begin VB.Menu mnurubcont 
         Caption         =   "Rubros Contabilidad"
      End
      Begin VB.Menu mnuUruwa 
         Caption         =   "Consultar en Uruware"
      End
      Begin VB.Menu manufacmail 
         Caption         =   "Enviar facturas por mail"
      End
      Begin VB.Menu mnusappuruw 
         Caption         =   "Controles SAPP-URUWARE"
      End
   End
   Begin VB.Menu mnudesp 
      Caption         =   "Despacho"
      Begin VB.Menu mnularg 
         Caption         =   "Receptor/Largador"
      End
      Begin VB.Menu mnuutildesp 
         Caption         =   "Utilitarios despacho"
      End
      Begin VB.Menu mnuinf 
         Caption         =   "Informes del Despacho"
      End
      Begin VB.Menu mnudemrecla 
         Caption         =   "Demoras Receptor/Largador"
      End
      Begin VB.Menu mnuinfdespp 
         Caption         =   "Informes despacho"
      End
      Begin VB.Menu mnuinfaz 
         Caption         =   "Informes CODIGO AZUL"
      End
      Begin VB.Menu mnuverlladesp 
         Caption         =   "Ver llamados en domicilio"
      End
      Begin VB.Menu mnuctrosasis 
         Caption         =   "Centros Asistenciales"
      End
      Begin VB.Menu mnuscap 
         Caption         =   "Pendientes SCA"
      End
   End
   Begin VB.Menu mnuuti 
      Caption         =   "Utilitarios"
      Begin VB.Menu mnusolhisop 
         Caption         =   "Solicitud de hisopados"
      End
      Begin VB.Menu mnucapta 
         Caption         =   "Captación de socios"
      End
      Begin VB.Menu mnuautorizacod 
         Caption         =   "Generar código autorización"
      End
      Begin VB.Menu mnuage 
         Caption         =   "Solicitud Insumos Informáticos"
      End
      Begin VB.Menu menuevalper 
         Caption         =   "Evaluación de personal"
         Begin VB.Menu mnuficper 
            Caption         =   "Ficha Personal"
         End
         Begin VB.Menu mnuinfeval 
            Caption         =   "Informes de evaluación"
         End
      End
      Begin VB.Menu mnuservap 
         Caption         =   "Servicios A.P."
      End
      Begin VB.Menu mnucamcon 
         Caption         =   "Cambiar contraseña"
      End
      Begin VB.Menu mnusolmark 
         Caption         =   "Solicitud a Marketing"
      End
      Begin VB.Menu mnusolasis 
         Caption         =   "Solicitud asistencia técnica informática"
      End
      Begin VB.Menu mnusolasiman 
         Caption         =   "Solicitud asistencia a mantenimiento"
      End
      Begin VB.Menu mnusolrrhh 
         Caption         =   "Solicitud a RRHH"
      End
      Begin VB.Menu mnusolpsoc 
         Caption         =   "Solicitud a Padrón social"
      End
      Begin VB.Menu mnusolibaja 
         Caption         =   "Solicitudes de Baja"
      End
      Begin VB.Menu mnuregoperac 
         Caption         =   "Registro Operaciones"
      End
      Begin VB.Menu mnumejor 
         Caption         =   "Mejora Continua"
      End
      Begin VB.Menu mamind 
         Caption         =   "Registro MAM Jefes"
      End
      Begin VB.Menu mnusolmejo 
         Caption         =   "Iniciativas del Personal"
      End
      Begin VB.Menu mnuaum 
         Caption         =   "Aumentos"
      End
      Begin VB.Menu mnutabsis 
         Caption         =   "Tablas del sistema"
         Begin VB.Menu mnutabconv 
            Caption         =   "Convenios ABM"
         End
         Begin VB.Menu mnupromos 
            Caption         =   "Promociones"
         End
         Begin VB.Menu mnutabcob 
            Caption         =   "Cobradores ABM"
         End
         Begin VB.Menu mnutabpro 
            Caption         =   "Promotores ABM"
         End
         Begin VB.Menu mnupromfunc 
            Caption         =   "Promotor-Funcionarios"
         End
         Begin VB.Menu mnumedic 
            Caption         =   "Médicos ABM"
         End
         Begin VB.Menu mnuzona 
            Caption         =   "Zonas ABM"
         End
         Begin VB.Menu mnufamili 
            Caption         =   "Familias ABM"
         End
         Begin VB.Menu mnuarancon 
            Caption         =   "Aranceles de convenio"
         End
         Begin VB.Menu mnumutual 
            Caption         =   "Mutualistas ABM"
         End
         Begin VB.Menu mnumutadm 
            Caption         =   "Mutualistas Adm"
         End
         Begin VB.Menu mnuusu 
            Caption         =   "Usuarios"
         End
         Begin VB.Menu mnuparamemp 
            Caption         =   "Parámetros de la empresa"
         End
      End
   End
   Begin VB.Menu mnusalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frm_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      Dim MyForm As FRMSIZE
      Dim DesignX As Integer
      Dim DesignY As Integer
      
      Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1


      

Private Sub abmprom_Click()
frm_infabmpromo.Show vbModal

End Sub

Private Sub demora_Click()
frm_infdemoras.Show vbModal

End Sub







Private Sub env_fac_Click()

Dim Xlibro As String
Dim Arch As String
Dim XImp As Double
Dim Xiva, Xtimbre, XIVA2 As Double
Dim Xusu, Xtexobs As String
Dim Ctacaja As String
Dim mes, ano, dia, xbase, xnumrub, XNfac As Long
Dim Xdeb, Xhab, Xqm, Xqa As String
Dim Xbander As Integer
Dim Xelstrin, Xelrut As String
Xelstrin = ""
Xbander = 0
Xelrut = ""
XNfac = 0

'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in (101,102) order by realizada,factura"
data_lin.Refresh

Arch = "IM"
mes = Month(mfd.Text)
ano = Year(mfh.Text)
If mes < 10 Then
   Arch = Arch + Mid(Trim(str(ano)), 3, 2) + Trim("0") + Trim(str(mes)) + "01.txt"
   Xqm = Trim("0") + Trim(str(mes))
   Xqa = Mid(Trim(str(ano)), 3, 2)
Else
   Arch = Arch + Mid(Trim(str(ano)), 3, 2) + Trim(str(mes)) + "01.txt"
   Xqm = Trim(str(mes))
   Xqa = Mid(Trim(str(ano)), 3, 2)
End If

If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      If Dir("C:\Cajas Memory" & "\" & Trim(Arch)) <> "" Then
         Kill ("C:\Cajas Memory" & "\" & Trim(Arch))
      End If
      Open "C:\Cajas Memory" & "\" & Trim(Arch) For Output As #1
      If data_lin.Recordset.RecordCount > 0 Then
         MsgBox "ATENCION! Hay registros de facturación de convenios para enviar a Administración. Aguarde..", vbInformation
         data_lin.Recordset.MoveFirst
         Xusu = data_lin.Recordset("operador")
         dia = Day(data_lin.Recordset("fecha"))
         xnumrub = data_lin.Recordset("rub_cont")
         XNfac = data_lin.Recordset("factura")
         Print #1, "Dia, Debe, Haber,Concepto,Ruc, Moneda,  Total, CodigoIVA, IVA, Cotizacion, Libro"
         Do While Not data_lin.Recordset.EOF
            If XNfac = data_lin.Recordset("factura") Then
               If IsNull(data_lin.Recordset("pendiente")) = False Then
                  If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                     XImp = XImp - data_lin.Recordset("tot_lin") - data_lin.Recordset("valor_iva")
                     Xiva = Xiva - data_lin.Recordset("valor_iva")
                  Else
                     XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                     Xiva = Xiva + data_lin.Recordset("valor_iva")
                  End If
               Else
                  data_cab.RecordSource = "Select * from clirespl where cl_numero =" & data_lin.Recordset("factura") & " and cl_codigo =" & data_lin.Recordset("cod_cli")
                  data_cab.Refresh
                  If data_cab.Recordset.RecordCount > 0 Then
                     If IsNull(data_cab.Recordset("cl_telefon")) = False Then
                        If data_cab.Recordset("cl_telefon") = "NC E-TICKET" Or data_cab.Recordset("cl_telefon") = "NC E-FACTURA" Then
                           XImp = XImp - data_lin.Recordset("tot_lin") - data_lin.Recordset("valor_iva")
                           Xiva = Xiva - data_lin.Recordset("valor_iva")
                        Else
                           XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                           Xiva = Xiva + data_lin.Recordset("valor_iva")
                        End If
                     Else
                        XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                        Xiva = Xiva + data_lin.Recordset("valor_iva")
                     End If
                  Else
                     XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                     Xiva = Xiva + data_lin.Recordset("valor_iva")
                  End If
               End If
               If data_lin.Recordset("costo_prod") = 1 Then
                  If IsNull(data_lin.Recordset("pendiente")) = False Then
                     If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                        Xiva = data_lin.Recordset("tot_lin") * 0.22
                        XImp = XImp - data_lin.Recordset("tot_lin") - Xiva
                        Xiva = Xiva
                     Else
                        Xiva = data_lin.Recordset("tot_lin") * 0.22
                        XImp = data_lin.Recordset("tot_lin") + Xiva
                     End If
                  Else
                     Xiva = data_lin.Recordset("tot_lin") * 0.22
                     XImp = data_lin.Recordset("tot_lin") + Xiva
                  End If
               End If
               xnumrub = data_lin.Recordset("rub_cont")
               Xusu = data_lin.Recordset("operador")
               dia = Day(data_lin.Recordset("fecha"))
               XNfac = data_lin.Recordset("factura")
               data_lin.Recordset.MoveNext
            Else
                data_lin.Recordset.MovePrevious
                If IsNull(data_lin.Recordset("ruc")) = False Then
                   Xelrut = data_lin.Recordset("ruc")
                Else
                   Xelrut = ""
                End If
                xnumrub = data_lin.Recordset("rub_cont")
                Xusu = data_lin.Recordset("operador")
                dia = Day(data_lin.Recordset("fecha"))
                data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                data_cnv.Refresh
                If data_cnv.Recordset.RecordCount > 0 Then
                   Xlibro = "E"
                   If dia < 10 Then
                      Xelstrin = Trim(str(dia)) & ","
                   Else
                      Xelstrin = Trim(str(dia)) & ","
                   End If
                   If IsNull(data_cnv.Recordset("cnv_uapago")) = False Then
                      Xelstrin = Xelstrin & data_cnv.Recordset("cnv_uapago") & ","
                   Else
                      Xelstrin = Xelstrin & "0" & ","
                   End If
                   Xelstrin = Xelstrin & data_lin.Recordset("rub_cont") & ","
                   Xelstrin = Xelstrin & "F." & Trim(str(data_lin.Recordset("factura"))) & " " & data_lin.Recordset("nom_cli") & ","
                   If Len(Trim(Xelrut)) > 2 Then
                      Xelstrin = Xelstrin & Trim(Xelrut) & ","
                   Else
                      Xelstrin = Xelstrin & ","
                   End If
                   Xelstrin = Xelstrin & "0,"
                   Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
                   If IsNull(data_lin.Recordset("costo_prod")) = False Then
                      XIVA2 = data_lin.Recordset("costo_prod")
                      If XIVA2 = 0 Then
                         XIVA2 = 3
                      Else
                         If XIVA2 = 1 Then
                            XIVA2 = 4
                         Else
                            If XIVA2 = 2 Then
                               XIVA2 = 0
                            Else
                               XIVA2 = 3
                            End If
                         End If
                      End If
                   Else
                       XIVA2 = 3
                   End If
                   Xelstrin = Xelstrin & Trim(str(XIVA2)) & ","
                   Xelstrin = Xelstrin & Format(Xiva, "######0.00") & "," & "0.000" & ","
                   Xelstrin = Xelstrin & Trim(Xlibro)
'                   Print #1, Trim(Xlibro)
                   Print #1, Xelstrin
    '                   data_caja.Recordset.MoveNext
                
                Else
                   MsgBox "No se encontró CONVENIO", vbInformation, "Mensaje"
                   Unload Me
                End If
                XImp = 0
                Xiva = 0
                data_lin.Recordset.MoveNext
                xnumrub = data_lin.Recordset("rub_cont")
                Xusu = data_lin.Recordset("operador")
                dia = Day(data_lin.Recordset("fecha"))
                XNfac = data_lin.Recordset("factura")
            End If
         Loop
         data_lin.Recordset.MovePrevious
         If IsNull(data_lin.Recordset("ruc")) = False Then
            Xelrut = data_lin.Recordset("ruc")
         Else
            Xelrut = ""
         End If
         xnumrub = data_lin.Recordset("rub_cont")
         Xusu = data_lin.Recordset("operador")
         dia = Day(data_lin.Recordset("fecha"))
         data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
         data_cnv.Refresh
         If data_cnv.Recordset.RecordCount > 0 Then
            Xlibro = "E"
            If dia < 10 Then
               Xelstrin = Trim(str(dia)) & ","
            Else
               Xelstrin = Trim(str(dia)) & ","
            End If
            If IsNull(data_cnv.Recordset("cnv_uapago")) = False Then
               Xelstrin = Xelstrin & data_cnv.Recordset("cnv_uapago") & ","
            Else
               Xelstrin = Xelstrin & "0" & ","
            End If
            Xelstrin = Xelstrin & data_lin.Recordset("rub_cont") & ","
            Xelstrin = Xelstrin & "F." & Trim(str(data_lin.Recordset("factura"))) & " " & data_lin.Recordset("nom_cli") & ","
            If Len(Trim(Xelrut)) > 2 Then
               Xelstrin = Xelstrin & Trim(Xelrut) & ","
            Else
               Xelstrin = Xelstrin & ","
            End If
            Xelstrin = Xelstrin & "0,"
            Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
            If IsNull(data_lin.Recordset("costo_prod")) = False Then
               XIVA2 = data_lin.Recordset("costo_prod")
               If XIVA2 = 0 Then
                  XIVA2 = 3
               Else
                  If XIVA2 = 1 Then
                     XIVA2 = 4
                  Else
                     If XIVA2 = 2 Then
                        XIVA2 = 0
                     Else
                        XIVA2 = 3
                     End If
                  End If
               End If
            Else
                XIVA2 = 3
            End If
            Xelstrin = Xelstrin & Trim(str(XIVA2)) & ","
            Xelstrin = Xelstrin & Format(Xiva, "######0.00") & "," & "0.000" & ","
            Xelstrin = Xelstrin & Trim(Xlibro)
'                   Print #1, Trim(Xlibro)
            Print #1, Xelstrin
         Else
            MsgBox "No se encontró convenio.", vbInformation, "Mensaje"
            Unload Me
         End If
         XImp = 0
         Xiva = 0
         Close #1
         MsgBox "Se ha generado el archivo. Aguarde a que se envíe por correo a Administración.", vbInformation, "Mensaje"
         Dim MenCorreo As String
         Dim oMail As Class1
              Set oMail = New Class1
              With oMail
                  .servidor = "smtp.gmail.com"
                  .puerto = 465
                  .UseAuntentificacion = True
                  .ssl = True
                  .Usuario = "sappfacturacion@gmail.com"
                  .PassWord = "sapp1987"
                  .Asunto = "Facturación Convenios " & mfd.Text & " A " & mfh.Text
                  .de = "sappfacturacion@gmail.com"
                  .para = "jefeadministracion@sapp.com.uy; contaduria@sapp.com.uy; jefedepartamentoti@sapp.com.uy"
                  .Adjunto = "C:\Cajas Memory" & "\" & Trim(Arch)
                  .Mensaje = "Archivo para procesar en Conty."
                  .Enviar_Backup ' manda el mail
              End With
              Set oMail = Nothing
              data_ctrenvfac.Recordset.Edit
              data_ctrenvfac.Recordset("fecha") = mfh.Text
              data_ctrenvfac.Recordset.Update
              
              MsgBox "Correo enviado.", vbInformation
      
      Else
         Close #1
         data_ctrenvfac.Recordset.Edit
         data_ctrenvfac.Recordset("fecha") = mfh.Text
         data_ctrenvfac.Recordset.Update
      
      End If
   End If
End If

End Sub

Private Sub Form_Initialize()
labusua.Caption = UCase(data_usuac.Recordset("nombre"))
labempre.Caption = UCase(data_parse.Recordset("empresa"))
Label3.Caption = Welnombredu

End Sub
      Private Sub Command1_Click()
      Dim ScaleFactorX As Single, ScaleFactorY As Single

      DesignX = Xpixels
      DesignY = Ypixels
      RePosForm = True
      DoResize = False
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      ScaleFactorX = (Xpixels / DesignX)
      ScaleFactorY = (Ypixels / DesignY)
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'      Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
'       "  by " + Str$(Ypixels)
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
      End Sub

Private Sub Form_Load()
Dim Xdiasenv As Integer
Dim Xlafecaenv As Date
      Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      DesignX = 800
      DesignY = 600
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      ScaleFactorX = (Xpixels / DesignX)
      ScaleFactorY = (Ypixels / DesignY)
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
 '     Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
  '     "  by " + Str$(Ypixels)
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width

With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

data_parse.DatabaseName = App.path & "\PARSE.mdb"
data_parse.RecordSource = "parsec0"
data_parse.Refresh
data_ctrlfact.DatabaseName = App.path & "\ctrf.mdb"
data_ctrlfact.RecordSource = "ctrf"
data_ctrlfact.Refresh
data_ctrabre.DatabaseName = App.path & "\ctrabre.mdb"
data_ctrabre.RecordSource = "ctrabre"
data_ctrabre.Refresh

data_promos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conscliprom.Connect = "odbc;dsn=" & Xconexrmt & ";"

adologueo.ConnectionString = "dsn=" & Xconexrmt
adologueo.RecordSource = "select * from version"
adologueo.Refresh
If adologueo.Recordset.RecordCount > 0 Then
   adologueo.Recordset.MoveFirst
   If IsNull(adologueo.Recordset("obs")) = False Then
      t_info.Text = adologueo.Recordset("obs")
   Else
      t_info.Text = ""
   End If
Else
   t_info.Text = ""
End If

adologueo.RecordSource = "select * from ctrdesp where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and usuario ='" & WElusuario & "' and respuesta ='" & "SI" & "'"
adologueo.Refresh
If data_ctrabre.Recordset("base") = 60 Then
    If adologueo.Recordset.RecordCount > 0 Then
    Else
       frm_mensajelog.Show vbModal
    End If
End If
adologueo.Recordset.Close

data_ctrenvfac.DatabaseName = App.path & "\ctradmc.mdb"
data_ctrenvfac.RecordSource = "ctrabre"
data_ctrenvfac.Refresh

data_usuac.DatabaseName = "C:\WINDOWS\usapp.mdb"
data_usuac.RecordSource = "usuarioact"
data_usuac.Refresh
Data1.DatabaseName = App.path & "\parse.mdb"
Data1.RecordSource = "parsec0"
Data1.Refresh

data_mejor.ConnectionString = "dsn=" & Xconexrmt
data_cnv.ConnectionString = "dsn=" & Xconexrmt
data_cab.ConnectionString = "dsn=" & Xconexrmt

data_lin.ConnectionString = "dsn=" & Xconexrmt

Data2.DatabaseName = App.path & "\informes.mdb"
Data2.RecordSource = "infvtas"
Data2.Refresh
If Data2.Recordset.RecordCount > 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      Data2.Recordset.Delete
      Data2.Recordset.MoveNext
   Loop
End If

On Error GoTo Hayerror

Label4.Caption = 0
Label5.Caption = 0

   
If WElusuario = "BDD" Or WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Or WElusuario = "MARTINC" Or _
   WElusuario = "NROCHINOTTI" Or WElusuario = "MCOSTA" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "BRUNO" Or WElusuario = "ARIZZO" Or _
   WElusuario = "MARCELOM" Or WElusuario = "ENRIQUE" Or WElusuario = "AGUILLEN" Or WElusuario = "DARIOH" Or WElusuario = "MARIAJOSE" Then
   data_mejor.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 2 & " and cl_val1 <" & 0 & " and cl_descpag ='" & WElusuario & "'"
   data_mejor.Refresh
   If data_mejor.Recordset.RecordCount > 0 Then
      data_mejor.Recordset.MoveLast
      MsgBox "Tiene un total de: " & data_mejor.Recordset.RecordCount & " REGISTROS EN EL MAM JEFES SIN CERRAR", vbInformation, "REGISTROS MAM INDIVIDUAL"
   End If
   Dim Xelvencmam As Date
   Xelvencmam = Date
   data_mejor.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_nomcobr =" & 1 & " and cl_nom_sup ='" & WElusuario & "' and cl_fec1 <='" & Format(Xelvencmam, "yyyy-mm-dd") & "' and cl_val1 <" & 0 & " order by cl_fnac DESC"
   data_mejor.Refresh
   If data_mejor.Recordset.RecordCount > 0 Then
      data_mejor.Recordset.MoveLast
      MsgBox "TIENE " & data_mejor.Recordset.RecordCount & " TAREAS VENCIDAS EN EL MAM GENERAL, VERIFIQUE!!", vbInformation
   End If
   data_mejor.Recordset.Close
Else
End If

If WElusuario = "MCOSTA" Or WElusuario = "PPONS" Or WElusuario = "FOSORIO" Or WElusuario = "NROCHINOTTI" Or WElusuario = "CDEMORAES" Or WElusuario = "CHEQUES" Then
'''''   Control_vence
End If
If data_ctrenvfac.Recordset("base") = 85 Then
   Xdiasenv = DateDiff("d", data_ctrenvfac.Recordset("fecha"), Date)
   Xdiasenv = Xdiasenv - 1
   Dim Xlafecaenvh As Date
   If Xdiasenv > 0 Then
      Xlafecaenv = data_ctrenvfac.Recordset("fecha") + 1
      Xlafecaenvh = data_ctrenvfac.Recordset("fecha") + Xdiasenv
      mfd.Text = Format(Xlafecaenv, "dd/mm/yyyy")
      mfh.Text = Format(Xlafecaenvh, "dd/mm/yyyy")
      frm_menu.Enabled = False
      env_fac_Click
      frm_menu.Enabled = True
   End If
End If

   
Exit Sub

Hayerror:
         If Err.Number = 3155 Then
           MsgBox "Hubo un error al actualizar, comunique a Informática", vbCritical, "Mensaje"
         Else
            MsgBox "Hubo un error " & Trim(str(Err.Number)) & " ", vbCritical, "Mensaje"
         End If
        

End Sub

Private Sub Form_Resize()
      Dim ScaleFactorX As Single, ScaleFactorY As Single

      If Not DoResize Then  ' To avoid infinite loop
         DoResize = False
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width



End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub infdemdes_Click()
frm_infdemoras.Show vbModal

End Sub

Private Sub Image2_DblClick()
If t_info.Visible = True Then
   t_info.Visible = False
Else
   t_info.Visible = True
End If

End Sub

Private Sub infdespnuevo_Click()
'If ControlUsuario(infdespnuevo.Caption) = 1 Then
'   frm_infdesp3.Show vbModal
'End If

End Sub

Private Sub infemixcob_Click()
frm_infemis.Show vbModal

End Sub

Private Sub mactre_Click()
'If ControlUsuario(mactre.Caption) = 1 Then
'   frm_actasrr.Show vbModal
'End If

End Sub

Private Sub mamind_Click()
If ControlUsuario(mamind.Caption) = 1 Then
   frm_mejorai.Show vbModal
End If

End Sub

Private Sub manufacmail_Click()
If ControlUsuario(manufacmail.Caption) = 1 Then
   frm_envfaccnv.Show vbModal
End If

End Sub



Private Sub menupasbaj_Click()
If ControlUsuario(menupasbaj.Caption) = 1 Then
   frm_pasbaj.Show vbModal
End If

End Sub

Private Sub mnuaumest_Click()

End Sub

Private Sub menusocser_Click()
frm_infconsultas.Show vbModal

End Sub

Private Sub menuvtasxsernewr_Click()
If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
   frm_vtasservcp.Show vbModal
Else
   frm_vtasserv.Show vbModal
End If

End Sub

Private Sub mnuabmvar_Click()
frm_infabmvar.Show vbModal

End Sub




Private Sub mnuacccli_Click()

End Sub

Private Sub menupeddomi_Click()
If frm_pedidomedic.Visible = True Then
   MsgBox "Ya está abierto."
Else
   frm_pedidomedic.Show
End If

End Sub

Private Sub mnuactdeu_Click()
If ControlUsuario(mnuactdeu.Caption) = 1 Then
    If frm_actdeuda.Visible = True Then
       MsgBox "Ya está abierto"
    Else
    '   If WElusuario = "JFERNAN" Or WElusuario = "MCOSTA" Or WElusuario = "PPONS" Or WElusuario = "CDEMORAES" Or WElusuario = "PPATRON" Or WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Then
          frm_actdeuda.Show
    '   Else
    '      MsgBox "Usuario no autorizado", vbInformation
    '   End If
    End If
End If

End Sub

Private Sub mnuagend_Click()
'frm_agenda.Show vbModal

End Sub


Private Sub mnuafilia_Click()
If frm_afilia.Visible = True Then
   MsgBox "El formulario ya está abierto", vbInformation
Else
   frm_afilia.Show
End If

End Sub

Private Sub mnuage_Click()
'Dim ejecutar As Long
'On Error GoTo Quepasaalabrir

'ejecutar = ShellExecute(Me.hwnd, "Open", App.path & "\Agenda.jar", "", "", 1)

'Exit Sub

'Quepasaalabrir:
'                If Err.Number = 53 Then
'                   MsgBox "No existe archivo"
'                Else
''                   MsgBox "No se puede abrir la agenda"
'                End If
If ControlUsuario(mnuage.Caption) = 1 Then
   frm_solinsumos.Show vbModal
End If

End Sub

Private Sub mnuarancon_Click()
'frm_aran.Show vbModal
If ControlUsuario(mnuarancon.Caption) = 1 Then
   frm_abmgrupos.Show vbModal
End If

End Sub

'Private Sub mnuaudit_Click()
'data_cargo.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_cargo.RecordSource = "movil"
'data_cargo.Refresh
'data_cargo.Recordset.FindFirst "medico ='" & WElusuario & "'"
'If Not data_cargo.Recordset.NoMatch Then
'   frm_solaudito.Show vbModal
'Else
'   MsgBox "Usuario sin permisos"
'End If

'End Sub

Private Sub mnuaum_Click()
Dim VerAfiliaciones As String
VerAfiliaciones = ""
If ControlUsuario(mnuaum.Caption) = 1 Then
   VerAfiliaciones = MsgBox("Desea modificar valores de Afiliaciones?", vbYesNo + vbInformation, "Aumentos")
   If VerAfiliaciones = vbYes Then
      frm_afiliamodivalor.Show vbModal
   Else
      frm_aumentos.Show vbModal
   End If
End If

End Sub


Private Sub mnuautorizacod_Click()
If ControlUsuario(mnuautorizacod.Caption) = 1 Then
   Genera_codigo
End If

End Sub

Private Sub mnubajbas_Click()
If ControlUsuario(mnubajbas.Caption) = 1 Then
   frm_bajabase.Show vbModal
End If

End Sub


Private Sub mnucamcon_Click()
frm_cambiocont.Show vbModal

End Sub

Private Sub mnucamcont_Click()

frm_cambiocont.Show vbModal

End Sub


Private Sub mnucarnuev_Click()

End Sub

Private Sub mnucapta_Click()
frm_bloqueos.Show vbModal

End Sub

Private Sub mnucarfac_Click()
If ControlUsuario(mnucarfac.Caption) = 1 Then
   frm_carfaccnv.Show vbModal
End If

End Sub

Private Sub mnucartasm_Click()
If ControlUsuario(mnucartasm.Caption) = 1 Then
   frm_infcartasm.Show vbModal
End If

End Sub

Private Sub mnucboredp_Click()

'frm_pasaraRP.Show vbModal
Dim Xlineat As String
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook
Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir3 As New Excel.Application
Dim Xmat As Long
Dim TextodeFecha As String
TextodeFecha = ""
If Month(Date) > 9 Then
   If Day(Date) > 9 Then
      TextodeFecha = Trim(str(Year(Date))) & Trim(str(Month(Date))) & Trim(str(Day(Date)))
   Else
      TextodeFecha = Trim(str(Year(Date))) & Trim(str(Month(Date))) & "0" & Trim(str(Day(Date)))
   End If
Else
   If Day(Date) > 9 Then
      TextodeFecha = Trim(str(Year(Date))) & "0" & Trim(str(Month(Date))) & Trim(str(Day(Date)))
   Else
      TextodeFecha = Trim(str(Year(Date))) & "0" & Trim(str(Month(Date))) & "0" & Trim(str(Day(Date)))
   End If
End If
Xmat = 0
frm_menu.MousePointer = 11
Xlineat = ""
Xtotreg = 0
data_arq.Connect = "odbc;dsn=sappnew;"
data_arq.RecordSource = "select deudas.nro_cobr,deudas.nombre,deudas.fecha_pago,deudas.ano,deudas.mes,deudas.total,deudas.fecha,deudas.documento,deudas.servi,deudas.cliente," & _
"deudas.fecha_pago,clientes.cl_codigo,clientes.cl_cedula,clientes.cl_codced,clientes.estado from deudas inner join clientes on deudas.cliente=clientes.cl_codigo" & _
" where clientes.estado in (1) and deudas.fecha_pago is null and deudas.nro_cobr in (221) order by deudas.cliente,deudas.fecha"
data_arq.Refresh
If data_arq.Recordset.RecordCount > 0 Then
   data_arq.Recordset.MoveFirst
   Xlin = 1
   XCol = 1
   MsgBox "El archivo se guardará en la carpeta planillas del disco C", vbInformation
   Set Xobjexel22 = New Excel.Application
   Set Xlibexel22 = Xobjexel22.Workbooks.Add
   Set Xarchexel22 = Xlibexel22.Worksheets.Add
   Xarchexel22.Name = Trim("Sapp")
   Xlibexel22.SaveAs ("C:\planillas\RedPagos-" & TextodeFecha & ".csv")
   Xarchtex = "C:\planillas\RedPagos-" & TextodeFecha & ".csv"
   Xmat = data_arq.Recordset("cliente")
   Do While Not data_arq.Recordset.EOF
      Xlineat = Trim(str(data_arq.Recordset("ano"))) & "," & Trim(str(data_arq.Recordset("mes"))) & "," & Trim(str(Xtotreg)) & "," & Trim(str(data_arq.Recordset("cl_cedula"))) & Trim(str(data_arq.Recordset("cl_codced"))) & "," & _
      data_arq.Recordset("nombre") & ",0," & Trim(str(Val(data_arq.Recordset("total")))) & "," & Format(data_arq.Recordset("fecha"), "dd/mm/yyyy") & "," & _
      Format(data_arq.Recordset("fecha"), "dd/mm/yyyy") & "," & Trim(str(data_arq.Recordset("documento"))) & "," & Trim(str(1)) & "," & Format(data_arq.Recordset("servi"), "###0.00")
      Xmat = data_arq.Recordset("cliente")
      data_arq.Recordset.MoveNext
      If data_arq.Recordset.EOF = True Then
      Else
         If Xmat = data_arq.Recordset("cliente") Then
            Xtotreg = Xtotreg + 1
         Else
            Xtotreg = 0
         End If
      End If
      Xarchexel22.Cells(Xlin, XCol) = Xlineat
      Xlin = Xlin + 1
      Xlineat = ""
   Loop
   Xlibexel22.Save
   Xlibexel22.Close
   Xobjexel22.Quit
   Xlabrir3.Workbooks.Open Xarchtex, , False
   Xlabrir3.Visible = True
   Xlabrir3.WindowState = xlMaximized
End If

frm_menu.MousePointer = 0
MsgBox "Proceso terminado"

End Sub

Private Sub mnucerrarar_Click()
If ControlUsuario(mnucerrarar.Caption) = 1 Then

    On Error GoTo Xelerr
    
        Dim Xelmensacerra As String
        MsgBox "LUEGO DE CERRADO EL ARQUEO NO SE PERMITIRA AGREGAR O MODIFICAR DATOS.", vbCritical
        
        Xelmensacerra = InputBox("Ingrese número de cobrador para cerrar el arqueo")
        If Trim(Xelmensacerra) <> "" Then
           If Val(Xelmensacerra) > 0 Then
              frm_menu.MousePointer = 11
              ConectarBD
              ConbdSapp.Open
              ConbdSapp.Execute "Update arqueo set codpro =" & 98 & " where cob=" & Val(Xelmensacerra)
              ConbdSapp.Close
              frm_menu.MousePointer = 0
              MsgBox "Proceso terminado"
           Else
              MsgBox "El cobrador no puede ser cero."
           End If
        Else
           MsgBox "No se ingreso cobrador."
        End If
    
    Exit Sub
    
Xelerr:
           If Err.Number = 3051 Then
              frm_menu.MousePointer = 0
              MsgBox "Error al procesar arqueo. COB:" & Xelmensacerra, vbInformation
           Else
              frm_menu.MousePointer = 0
              MsgBox "Error al cerrar el arqueo. Verifique datos. COB:" & Xelmensacerra, vbInformation
           End If
End If

End Sub

Private Sub mnuclavesiso_Click()
frm_calidadiso.Show vbModal

End Sub

Private Sub mnucirepe_Click()
If ControlUsuario(mnucirepe.Caption) = 1 Then
    Dim Xlacee As Double
    Dim Xcanc As Integer
    Xcanc = 0
    frm_menu.MousePointer = 11
    Label6.Visible = True
    
    
    Dim MiBaseact As Database
    Dim Unasesact As Workspace
    Set Unasesact = Workspaces(0)
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
    
    MiBaseact.Execute "Delete * from infcli"
    
    data_mdb.DatabaseName = App.path & "\informes.mdb"
    data_mdb.RecordSource = "infcli"
    data_mdb.Refresh
    
    DoEvents
    
    data_sql.ConnectionString = "dsn=" & Xconexrmt
    data_sql.RecordSource = "Select * from clientes where cl_cedula not in (0) and estado =" & 1 & " order by cl_cedula"
    data_sql.Refresh
    If data_sql.Recordset.RecordCount > 0 Then
       data_sql.Recordset.MoveNext
       Xlacee = Int(data_sql.Recordset("cl_cedula"))
       Do While Not data_sql.Recordset.EOF
          If IsNull(data_sql.Recordset("cl_cedula")) = False Then
             If data_sql.Recordset("cl_Cedula") <> 0 Then
                If Xlacee = Int(data_sql.Recordset("cl_cedula")) Then
                   Xcanc = Xcanc + 1
                Else
                   If Xcanc >= 1 Then
                      data_sql.Recordset.MovePrevious
                      data_mdb.Recordset.AddNew
                      data_mdb.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
                      data_mdb.Recordset("cl_apellid") = data_sql.Recordset("cl_apellid")
                      data_mdb.Recordset("cl_cedula") = data_sql.Recordset("cl_cedula")
                      data_mdb.Recordset("cl_codced") = data_sql.Recordset("cl_codced")
                      data_mdb.Recordset("cl_codconv") = data_sql.Recordset("cl_codconv")
                      data_mdb.Recordset.Update
                      data_sql.Recordset.MoveNext
                      Xcanc = 0
                   End If
                End If
             End If
          End If
          Xlacee = Int(data_sql.Recordset("cl_cedula"))
          data_sql.Recordset.MoveNext
       Loop
    End If
    frm_menu.MousePointer = 0
    Label6.Visible = False
    MsgBox "Proceso terminado"
    
    cr1.ReportFileName = App.path & "\infclirep.rpt"
    cr1.Action = 1
End If


End Sub

Private Sub mnuclieco_Click()
If ControlUsuario(mnuclieco.Caption) = 1 Then
   frm_abmcli.Show vbModal
End If

End Sub

Private Sub mnucmtpend_Click()
If frm_pendcmtpol.Visible = True Then
   MsgBox "Ya está abierto"
Else
   frm_pendcmtpol.Show
End If

End Sub

Private Sub mnucobrarq_Click()
If ControlUsuario(mnucobrarq.Caption) = 1 Then
   frm_cobr.Show vbModal
End If

End Sub

Private Sub mnucomer_Click()

If ControlUsuario(mnucomer.Caption) = 1 Then
   frm_labo.Show vbModal
End If

End Sub

Private Sub mnuconemi_Click()
frm_consemi.Show vbModal

End Sub

Private Sub mnucompsina_Click()
If ControlUsuario(mnucompsina.Caption) = 1 Then
   frm_infmsp.Show vbModal
End If

End Sub

Private Sub mnuconentr_Click()

End Sub

Private Sub mnuconse_Click()
frm_infmsp.Show

End Sub

Private Sub mnuconsemisocnew_Click()
If ControlUsuario(mnuconsemisocnew.Caption) = 1 Then
   frm_consemi.Show vbModal
End If

End Sub

Private Sub mnucontroltpol_Click()
If WElusuario = "MCURBELO" Then
   frm_infctrolcons.Show vbModal
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub mnuconshisarq_Click()
If ControlUsuario(mnuconshisarq.Caption) = 1 Then
   frm_consarq.Show vbModal
End If

End Sub

Private Sub mnuconsres_Click()
If frm_pendcons.Visible = True Then
   MsgBox "Ya está abierto"
Else
   frm_pendcons.Show
End If

End Sub

Private Sub mnuconssoc_Click()
If ControlUsuario(mnuconssoc.Caption) = 1 Then
   frm_infconssoc.Show vbModal
End If

End Sub

Private Sub mnucrm_Click()
If ControlUsuario(mnucrm.Caption) = 1 Then
   frm_atsocio.Show vbModal
End If

End Sub

Private Sub mnuctractenf_Click()
'MsgBox "Realice el control desde la opción CONTROL DE CONSULTAS...", vbInformation
If ControlUsuario(mnuctractenf.Caption) = 1 Then
    If frm_ctrlactenf.Visible = True Then
       MsgBox "Ya está abierto!", vbCritical
    Else
       frm_ctrlactenf.Show
    End If
End If

End Sub

Private Sub mnuctrocons_Click()
If ControlUsuario(mnuctrocons.Caption) = 1 Then
    If frm_ctrolcons.Visible = True Then
       MsgBox "Ya está abierto!", vbCritical
    Else
       frm_ctrolcons.Show
    End If
End If

End Sub

Private Sub mnuctrolb_Click()
If ControlUsuario(mnuctrolb.Caption) = 1 Then
   frm_ctrlmedb.Show vbModal
End If

End Sub

Private Sub mnuctrolconsinf_Click()
frm_infctrolcons.Show vbModal

End Sub

Private Sub mnuctrolen_Click()
'MsgBox "Realice el control desde la opción CONTROL DE CONSULTAS...", vbInformation
If ControlUsuario(mnuctrolen.Caption) = 1 Then
    If frm_ctrlactenf.Visible = True Then
       MsgBox "Ya está abierto!", vbCritical
    Else
       frm_ctrlactenf.Show
    End If
End If

End Sub

Private Sub mnuctrolenthcctr_Click()
If ControlUsuario(mnuctrolenthcctr.Caption) = 1 Then
   frm_ctrolsolhc.Show vbModal
End If

End Sub

Private Sub mnuctrolenvh_Click()

End Sub

Private Sub mnuctrolenvhcctr_Click()
If ControlUsuario(mnuctrolenvhcctr.Caption) = 1 Then
   frm_ctrolsoldt.Show vbModal
End If

End Sub

Private Sub mnuctrolme_Click()
If ControlUsuario(mnuctrolme.Caption) = 1 Then
       If frm_ctrlfarm.Visible = True Then
          MsgBox "Ya está abierto!", vbCritical
       Else
          frm_ctrlfarm.Show
       End If
End If

End Sub

Private Sub mnuctrolmut_Click()
frm_ctrolmut.Show vbModal

End Sub

Private Sub mnuctrolsolhc_Click()
frm_ctrolsolhc.Show vbModal

End Sub

Private Sub mnuctrosasis_Click()
If ControlUsuario(mnuctrosasis.Caption) = 1 Then
   frm_mutu.Show vbModal
End If

End Sub

Private Sub mnuctrrutas_Click()
Dim XIdpromo As Integer
XIdpromo = 0
Dim XCedStr As String
Dim ControlEntreSim As String
ControlEntreSim = ""

XCedStr = ""

If ControlUsuario(mnuctrrutas.Caption) = 1 Then
    Dim Idpromos As Integer
    Dim CedPromo As String
    
    Dim MiBaseact As Database
    Dim Unasesact As Workspace
    Set Unasesact = Workspaces(0)
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
    Idpromos = 0
    CedPromo = ""
    MiBaseact.Execute "Delete * from infcli"
    
    data_mdb.DatabaseName = App.path & "\informes.mdb"
    data_mdb.RecordSource = "infcli"
    data_mdb.Refresh
       
    ControlEntreSim = MsgBox("Desea realizar control entre SIMULACIONES de los que NO están generados?", vbInformation + vbYesNo, "Controles Emisión")
    If ControlEntreSim = vbYes Then
       Dim Xmessim, Xaniosim As Integer
       
       frm_menu.MousePointer = 11
       data_simulaini.DatabaseName = App.path & "\simula.mdb"
       data_simulafin.DatabaseName = App.path & "\simula.mdb"
       data_simulaini.RecordSource = "emisim"
       data_simulaini.Refresh
       adoclipromo.ConnectionString = "dsn=" & Xconexrmt
       
       If data_simulaini.Recordset.RecordCount > 0 Then
          data_simulaini.Recordset.MoveFirst
          Xmessim = data_simulaini.Recordset("mes")
          Xaniosim = data_simulaini.Recordset("ano")
          If Xmessim = 1 Then
             Xmessim = 12
             Xaniosim = data_simulaini.Recordset("ano") - 1
          Else
             Xmessim = data_simulaini.Recordset("mes") - 1
             Xaniosim = data_simulaini.Recordset("ano")
          End If
          data_simulaini.DatabaseName = App.path & "\simulan.mdb"
          data_simulaini.RecordSource = "select * from emisim where mes =" & Xmessim & " and ano =" & Xaniosim
          data_simulaini.Refresh
          If data_simulaini.Recordset.RecordCount > 0 Then
             data_simulaini.Recordset.MoveFirst
             Do While Not data_simulaini.Recordset.EOF
                data_simulafin.RecordSource = "select * from emisim where cliente =" & data_simulaini.Recordset("cliente")
                data_simulafin.Refresh
                If data_simulafin.Recordset.RecordCount > 0 Then
                Else
                   data_mdb.Recordset.AddNew
                   data_mdb.Recordset("cl_codigo") = data_simulaini.Recordset("cliente")
                   data_mdb.Recordset("cl_apellid") = data_simulaini.Recordset("apellidos")
                   data_mdb.Recordset("cl_codconv") = data_simulaini.Recordset("cod_cnv")
                   data_mdb.Recordset.Update
                End If
                data_simulaini.Recordset.MoveNext
             Loop
          End If
       End If
       If data_mdb.Recordset.RecordCount > 0 Then
          data_mdb.Recordset.MoveFirst
          Do While Not data_mdb.Recordset.EOF
             adoclipromo.RecordSource = "select * from clientes where cl_codigo =" & data_mdb.Recordset("cl_codigo")
             adoclipromo.Refresh
             If adoclipromo.Recordset.RecordCount > 0 Then
                data_mdb.Recordset.Edit
                If IsNull(adoclipromo.Recordset("fecha_baja")) = False Then
                   data_mdb.Recordset("fecha_baja") = adoclipromo.Recordset("fecha_baja")
                   data_mdb.Recordset("info_debit") = "BAJA"
                Else
                   If IsNull(adoclipromo.Recordset("mesproxemi")) = False Then
                      data_mdb.Recordset("info_debit") = adoclipromo.Recordset("cl_codconv") & " EMI:" & Trim(str(adoclipromo.Recordset("mesproxemi"))) & "/" & Trim(str(adoclipromo.Recordset("anoproxemi"))) & " COB:" & adoclipromo.Recordset("cl_nrocobr")
                   Else
                      data_mdb.Recordset("info_debit") = adoclipromo.Recordset("cl_codconv") & " EMI:" & " COB:" & adoclipromo.Recordset("cl_nrocobr")
                   End If
                End If
                data_mdb.Recordset.Update
             End If
             data_mdb.Recordset.MoveNext
          Loop
          frm_menu.MousePointer = 0
          MsgBox "Terminado"
          data_mdb.RecordSource = "select * from infcli"
          data_mdb.Refresh
          crcontrol.ReportFileName = App.path & "\infctrolsim.rpt"
          crcontrol.Action = 1
       Else
          MsgBox "No hay datos para controlar."
       End If
    
    Else
        Label8.Visible = True
        pb1.Visible = True
        adoclipromo.ConnectionString = "dsn=" & Xconexrmt
        adoconvpromo.ConnectionString = "dsn=" & Xconexrmt
        adoctrrutas.ConnectionString = "dsn=" & Xconexrmt
        data_convrutas.Connect = "odbc;dsn=" & Xconexrmt & ";"
        adoctrrutas.RecordSource = "select * from clientes where estado =" & 1 & " and cl_codruta is not null and idpromos is null"
        adoctrrutas.Refresh
        
        DoEvents
        If adoctrrutas.Recordset.RecordCount > 0 Then
           adoctrrutas.Recordset.MoveLast
           pb1.Max = adoctrrutas.Recordset.RecordCount
           pb1.Value = 0
           adoctrrutas.Recordset.MoveFirst
           Do While Not adoctrrutas.Recordset.EOF
              adoclipromo.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codigo =" & adoctrrutas.Recordset("cl_codruta")
              adoclipromo.Refresh
              If adoclipromo.Recordset.RecordCount > 0 Then
                 adoconvpromo.RecordSource = "Select * from convenio where cnv_codigo ='" & adoclipromo.Recordset("cl_codconv") & "'"
                 adoconvpromo.Refresh
                 If adoconvpromo.Recordset.RecordCount > 0 Then
                    If IsNull(adoconvpromo.Recordset("cnv_grupo")) = False Then
                       If adoconvpromo.Recordset("cnv_grupo") <> "" Then
                          If adoconvpromo.Recordset("cnv_grupo") = "SEMM" Or adoconvpromo.Recordset("cnv_grupo") = "CASH" Or _
                             adoconvpromo.Recordset("cnv_grupo") = "CPS" Or adoconvpromo.Recordset("cnv_grupo") = "CASMU" Then
        '                     MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                             data_mdb.Recordset.AddNew
                             data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                             data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                             data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                             data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                             data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                             data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                             data_mdb.Recordset.Update
                          Else
                             If adoconvpromo.Recordset("cnv_codigo") = "CCNOS" Or adoconvpromo.Recordset("cnv_codigo") = "SMIN" Or adoconvpromo.Recordset("cnv_codigo") = "HEVANO" Or adoconvpromo.Recordset("cnv_codigo") = "CASANO" Or adoconvpromo.Recordset("cnv_codigo") = "GANOS" Or _
                                adoconvpromo.Recordset("cnv_codigo") = "UNIVS" Or adoconvpromo.Recordset("cnv_codigo") = "SMINR" Then
                                data_mdb.Recordset.AddNew
                                data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                                data_mdb.Recordset.Update
                                'MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                             End If
                          End If
                       Else
                          If adoconvpromo.Recordset("cnv_codigo") = "EMEFES" Or adoconvpromo.Recordset("cnv_codigo") = "SAFES" Or adoconvpromo.Recordset("cnv_codigo") = "SPFES" Or adoconvpromo.Recordset("cnv_codigo") = "SPFFES" Then
                             data_mdb.Recordset.AddNew
                             data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                             data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                             data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                             data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                             data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                             data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                             data_mdb.Recordset.Update
        '                     MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                          Else
                             If IsNull(adoconvpromo.Recordset("cnv_precio")) = True Then
        '                        MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                                data_mdb.Recordset.AddNew
                                data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                                data_mdb.Recordset.Update
                             Else
                                If adoconvpromo.Recordset("cnv_precio") <= 0 Then
                                   data_mdb.Recordset.AddNew
                                   data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                   data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                   data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                   data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                   data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                   data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                                   data_mdb.Recordset.Update
        '                           MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                                End If
                             End If
                          End If
                       End If
                    Else
                       If adoconvpromo.Recordset("cnv_codigo") = "EMEFES" Or adoconvpromo.Recordset("cnv_codigo") = "SAFES" Or adoconvpromo.Recordset("cnv_codigo") = "SPFES" Or adoconvpromo.Recordset("cnv_codigo") = "SPFFES" Then
                          data_mdb.Recordset.AddNew
                          data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                          data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                          data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                          data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                          data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                          data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                          data_mdb.Recordset.Update
        '                  MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                       Else
                          If IsNull(adoconvpromo.Recordset("cnv_precio")) = True Then
                             data_mdb.Recordset.AddNew
                             data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                             data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                             data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                             data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                             data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                             data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                             data_mdb.Recordset.Update
        '                     MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                          Else
                            If adoconvpromo.Recordset("cnv_precio") <= 0 Then
                               data_mdb.Recordset.AddNew
                               data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                               data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                               data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                               data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                               data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                               data_mdb.Recordset("cl_nom_sup") = "VERIFICAR CONVENIO"
                               data_mdb.Recordset.Update
        '                       MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
                            End If
                          End If
                       End If
                    End If
                 End If
              Else
                 data_mdb.Recordset.AddNew
                 data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                 data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                 data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                 data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                 data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                 data_mdb.Recordset("cl_nom_sup") = "SIN PROMOCION"
                 data_mdb.Recordset.Update
              End If
              adoctrrutas.Recordset.MoveNext
              pb1.Value = pb1.Value + 1
           Loop
           adoctrrutas.RecordSource = "select * from clientes where estado =" & 1 & " and idpromos in (1)"
           adoctrrutas.Refresh
           DoEvents
           If adoctrrutas.Recordset.RecordCount > 0 Then
              adoctrrutas.Recordset.MoveLast
              pb1.Max = pb1.Max + adoctrrutas.Recordset.RecordCount
              adoctrrutas.Recordset.MoveFirst
              Do While Not adoctrrutas.Recordset.EOF
                CedPromo = ""
                Idpromos = 0
    '            If IsNull(adoctrrutas.Recordset("idpromos")) = False Then
    '               Idpromos = adoctrrutas.Recordset("idpromos")
    '               If Idpromos > 0 Then
    '                  data_promos.RecordSource = "select * from promocion_gpo where id =" & Idpromos
    '                  data_promos.Refresh
    '                  If data_promos.Recordset.RecordCount > 0 Then
    '                     If data_promos.Recordset("descrip") = "Grupo de 3 o más" Then
                            If IsNull(adoctrrutas.Recordset("cl_codruta")) = False Then
                               data_conscliprom.RecordSource = "select * from clientes where cl_codruta =" & adoctrrutas.Recordset("cl_codruta") & " and estado in (1)"
                               data_conscliprom.Refresh
                               If data_conscliprom.Recordset.RecordCount > 0 Then
                                  data_conscliprom.Recordset.MoveLast
                                  If data_conscliprom.Recordset.RecordCount >= 2 Then
                                     If data_conscliprom.Recordset.RecordCount = 2 Then
                                        If Len(data_conscliprom.Recordset("cl_codruta")) = 7 Then
                                           XCedStr = Mid(Trim(data_conscliprom.Recordset("cl_codruta")), 1, 6)
                                        Else
                                           XCedStr = Mid(Trim(data_conscliprom.Recordset("cl_codruta")), 1, 7)
                                        End If
                                        data_conscliprom.RecordSource = "select * from clientes where cl_cedula =" & Val(Trim(XCedStr)) & " and estado in (1) and idpromos in (1)"
                                        data_conscliprom.Refresh
                                        If data_conscliprom.Recordset.RecordCount > 0 Then
                                           data_convrutas.RecordSource = "select * from convenio where cnv_codigo ='" & data_conscliprom.Recordset("cl_codconv") & "' and cnv_emite='" & "SI" & "' and cnv_cant_r =" & 2
                                           data_convrutas.Refresh
                                           If data_convrutas.Recordset.RecordCount > 0 Then
                                           Else
                                              data_mdb.Recordset.AddNew
                                              data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                              data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                              data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                              data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                              data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                              data_mdb.Recordset("cl_nom_sup") = "FALTA RUTA >=GPO.3"
                                              data_mdb.Recordset.Update
                                           End If
                                        Else
                                           data_mdb.Recordset.AddNew
                                           data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                           data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                           data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                           data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                           data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                           data_mdb.Recordset("cl_nom_sup") = "FALTA RUTA >=GPO.3"
                                           data_mdb.Recordset.Update
                                        End If
                                     End If
                                  Else
                                     data_mdb.Recordset.AddNew
                                     data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                     data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                     data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                     data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                     data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                     data_mdb.Recordset("cl_nom_sup") = "FALTA RUTA >=GPO.3"
                                     data_mdb.Recordset.Update
                                  End If
                               Else
                                  data_mdb.Recordset.AddNew
                                  data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                  data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                  data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                  data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                  data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                  data_mdb.Recordset("cl_nom_sup") = "FALTA RUTA >=GPO.3"
                                  data_mdb.Recordset.Update
                               End If
                            Else 'si la ruta es nula, ES TITULAR
                               If IsNull(adoctrrutas.Recordset("cl_cedula")) = False Then
                                  CedPromo = Trim(str(adoctrrutas.Recordset("cl_cedula"))) & Trim(str(adoctrrutas.Recordset("cl_codced")))
                                  data_conscliprom.RecordSource = "select * from clientes where cl_codruta =" & Val(CedPromo) & " and estado in (1) and idpromos in (1)"
                                  data_conscliprom.Refresh
                                  If data_conscliprom.Recordset.RecordCount > 0 Then
                                     data_conscliprom.Recordset.MoveLast
                                     If data_conscliprom.Recordset.RecordCount < 2 Then
                                        data_mdb.Recordset.AddNew
                                        data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                        data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                        data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                        data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                        data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                        data_mdb.Recordset("cl_nom_sup") = "FALTA RUTA >=GPO.3"
                                        data_mdb.Recordset.Update
                                     End If
                                  Else
                                     data_mdb.Recordset.AddNew
                                     data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                                     data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                                     data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                                     data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                                     data_mdb.Recordset("cl_nro_sup") = adoctrrutas.Recordset("cl_codruta")
                                     data_mdb.Recordset("cl_nom_sup") = "FALTA RUTA >=GPO.3"
                                     data_mdb.Recordset.Update
                                  End If
                               Else
                                  CedPromo = "0"
                               End If
                            End If
                         'End If
                      'End If
                   'End If
                'End If
                adoctrrutas.Recordset.MoveNext
                pb1.Value = pb1.Value + 1
              Loop
           End If
           
        
        Else
           MsgBox "No hay registros para controlar", vbInformation
        End If
        
        MsgBox "Terminado control de rutas, comienza control de próxima emisión.", vbInformation
        Dim mes, anio As Integer
        If Month(Date) = 12 Then
           mes = 1
           anio = Year(Date) + 1
        Else
           mes = Month(Date) + 1
           anio = Year(Date)
        End If
        frm_menu.MousePointer = 11
        adoctrrutas.RecordSource = "select clientes.cl_codigo,clientes.cl_apellid,clientes.cl_codconv,clientes.idpromos," & _
        "clientes.estado,clientes.cl_nrocobr,clientes.mesproxemi,clientes.anoproxemi,convenio.cnv_codigo," & _
        "convenio.cnv_fbaja,convenio.cnv_cant_r,convenio.cnv_emite,convenio.cnv_colrec,convenio.cnv_hasta,convenio.cnv_precio from clientes inner join " & _
        "convenio on clientes.cl_codconv=convenio.cnv_codigo where clientes.estado=" & 1 & " and convenio.cnv_cant_r in (2) " & _
        "and convenio.cnv_colrec in ('R','M','V','A') and convenio.cnv_fbaja is null and convenio.cnv_precio >" & 0 & " and cnv_emite in ('SI') and clientes.cl_nrocobr not in (333,14,511,513,515,518,516,512,110,15,517)"
        adoctrrutas.Refresh
        If adoctrrutas.Recordset.RecordCount > 0 Then
           adoctrrutas.Recordset.MoveLast
           pb1.Max = pb1.Max + adoctrrutas.Recordset.RecordCount
           adoctrrutas.Recordset.MoveFirst
           Do While Not adoctrrutas.Recordset.EOF
              If IsNull(adoctrrutas.Recordset("mesproxemi")) = False Then
                 If adoctrrutas.Recordset("mesproxemi") = mes And adoctrrutas.Recordset("anoproxemi") = anio Then
                 Else
                    data_mdb.Recordset.AddNew
                    data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                    data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                    data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                    data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                    If IsNull(adoctrrutas.Recordset("idpromos")) = False Then
                       XIdpromo = adoctrrutas.Recordset("idpromos")
                       If XIdpromo = 2 Then
                          data_mdb.Recordset("cl_nom_sup") = "PROX.EMI.-ANUAL"
                       Else
                          data_mdb.Recordset("cl_nom_sup") = "PROX.EMISION"
                       End If
                    Else
                       data_mdb.Recordset("cl_nom_sup") = "PROX.EMISION"
                    End If
                    data_mdb.Recordset.Update
                 End If
              Else
                 data_mdb.Recordset.AddNew
                 data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cl_codigo")
                 data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cl_apellid")
                 data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cl_codconv")
                 data_mdb.Recordset("cl_nrocobr") = adoctrrutas.Recordset("cl_nrocobr")
                 If IsNull(adoctrrutas.Recordset("idpromos")) = False Then
                    XIdpromo = adoctrrutas.Recordset("idpromos")
                    If XIdpromo = 2 Then
                       data_mdb.Recordset("cl_nom_sup") = "PROX.EMI.-ANUAL"
                    Else
                       data_mdb.Recordset("cl_nom_sup") = "PROX.EMISION"
                    End If
                 Else
                    data_mdb.Recordset("cl_nom_sup") = "PROX.EMISION"
                 End If
                 data_mdb.Recordset.Update
              End If
              adoctrrutas.Recordset.MoveNext
              pb1.Value = pb1.Value + 1
           Loop
           frm_menu.MousePointer = 0
        End If
        
        frm_menu.MousePointer = 11
        
        adoctrrutas.RecordSource = "select * from convenio where cnv_cant_r in (1) and cnv_emite='" & "SI" & "' and cnv_hasta >='" & Format(Date, "yyyy-mm-dd") & "' and cnv_colrec in ('M','A','R','V') and cnv_fbaja is null"
        adoctrrutas.Refresh
        
        If adoctrrutas.Recordset.RecordCount > 0 Then
           adoctrrutas.Recordset.MoveLast
           pb1.Max = pb1.Max + adoctrrutas.Recordset.RecordCount
           adoctrrutas.Recordset.MoveFirst
           Do While Not adoctrrutas.Recordset.EOF
              If IsNull(adoctrrutas.Recordset("cnv_cuenta")) = False Then
                 If adoctrrutas.Recordset("cnv_cuenta") > 0 Then
                    data_conscliprom.RecordSource = "select * from clientes where cl_codigo =" & adoctrrutas.Recordset("cnv_cuenta") & " and estado in (1)"
                    data_conscliprom.Refresh
                    If data_conscliprom.Recordset.RecordCount > 0 Then
                        If IsNull(data_conscliprom.Recordset("mesproxemi")) = False Then
                           If data_conscliprom.Recordset("mesproxemi") = mes And data_conscliprom.Recordset("anoproxemi") = anio Then
                           Else
                              data_mdb.Recordset.AddNew
                              data_mdb.Recordset("cl_codigo") = data_conscliprom.Recordset("cl_codigo")
                              data_mdb.Recordset("cl_apellid") = data_conscliprom.Recordset("cl_apellid")
                              data_mdb.Recordset("cl_codconv") = data_conscliprom.Recordset("cl_codconv")
                              data_mdb.Recordset("cl_nrocobr") = data_conscliprom.Recordset("cl_nrocobr")
                              If IsNull(data_conscliprom.Recordset("idpromos")) = False Then
                                 XIdpromo = data_conscliprom.Recordset("idpromos")
                                 If XIdpromo = 2 Then
                                    data_mdb.Recordset("cl_nom_sup") = "PROX.EMI.-ANUAL"
                                 Else
                                    data_mdb.Recordset("cl_nom_sup") = "PROX.EMISION"
                                 End If
                              Else
                                 data_mdb.Recordset("cl_nom_sup") = "PROX.EMISION"
                              End If
                              data_mdb.Recordset.Update
                           End If
                        Else
                           data_mdb.Recordset.AddNew
                           data_mdb.Recordset("cl_codigo") = data_conscliprom.Recordset("cl_codigo")
                           data_mdb.Recordset("cl_apellid") = data_conscliprom.Recordset("cl_apellid")
                           data_mdb.Recordset("cl_codconv") = data_conscliprom.Recordset("cl_codconv")
                           data_mdb.Recordset("cl_nrocobr") = data_conscliprom.Recordset("cl_nrocobr")
                           If IsNull(data_conscliprom.Recordset("idpromos")) = False Then
                              XIdpromo = data_conscliprom.Recordset("idpromos")
                              If XIdpromo = 2 Then
                                 data_mdb.Recordset("cl_nom_sup") = "FALTA-ANUAL"
                              Else
                                 data_mdb.Recordset("cl_nom_sup") = "FALTA EMISION"
                              End If
                           Else
                              data_mdb.Recordset("cl_nom_sup") = "FALTA EMISION"
                           End If
                           data_mdb.Recordset.Update
                        End If
                    Else
                        data_mdb.Recordset.AddNew
                        data_mdb.Recordset("cl_codigo") = adoctrrutas.Recordset("cnv_cuenta")
                        data_mdb.Recordset("cl_apellid") = adoctrrutas.Recordset("cnv_desc")
                        data_mdb.Recordset("cl_codconv") = adoctrrutas.Recordset("cnv_codigo")
                        data_mdb.Recordset("cl_nom_sup") = "NO ENCONTRADO"
                        data_mdb.Recordset.Update
                    
                    End If
                 End If
              End If
              adoctrrutas.Recordset.MoveNext
              pb1.Value = pb1.Value + 1
                   
           Loop
        End If
        data_conscliprom.RecordSource = "select * from clientes where idpromos in (4) and estado in (1)"
        data_conscliprom.Refresh
        If data_conscliprom.Recordset.RecordCount > 0 Then
           data_conscliprom.Recordset.MoveFirst
           Do While Not data_conscliprom.Recordset.EOF
              data_mdb.Recordset.AddNew
              data_mdb.Recordset("cl_codigo") = data_conscliprom.Recordset("cl_codigo")
              data_mdb.Recordset("cl_apellid") = data_conscliprom.Recordset("cl_apellid")
              data_mdb.Recordset("cl_codconv") = data_conscliprom.Recordset("cl_codconv")
              data_mdb.Recordset("cl_nom_sup") = "EMERG 20%"
              data_mdb.Recordset.Update
           
              data_conscliprom.Recordset.MoveNext
           Loop
        End If
        
        frm_menu.MousePointer = 0
        data_mdb.RecordSource = "select * from infcli"
        data_mdb.Refresh
           
        crpro.ReportFileName = App.path & "\infctrpromoami.rpt"
        crpro.ReportTitle = "Informe de socios para Verificar"
        crpro.Action = 1
        MsgBox "Proceso terminado", vbInformation
        
        Label8.Visible = False
        pb1.Visible = False
    End If
End If

   
End Sub


Private Sub mnudebaut_Click()
frm_debitos.Show vbModal

End Sub

Private Sub mnudemcalilla_Click()
frm_calidadiso.Show vbModal

End Sub

Private Sub mnudemrecla_Click()
If ControlUsuario(mnudemrecla.Caption) = 1 Then
   frm_calidadiso.Show vbModal
End If

End Sub

Private Sub mnudeuemi_Click()
If ControlUsuario(mnudeuemi.Caption) = 1 Then
   frm_pasdeuda.Show vbModal
End If

End Sub

Private Sub mnudeuxsoc_Click()
frm_infsaldos.Show vbModal

End Sub


Private Sub mnuelechce_Click()
If ControlUsuario(mnuelechce.Caption) = 1 Then
   Shell App.path & "\cargadatos.exe", vbNormalFocus
End If

End Sub

Private Sub mnuemimen_Click()
If ControlUsuario(mnuemimen.Caption) = 1 Then
   frm_emision.Show vbModal
End If

End Sub

Private Sub mnuemisim_Click()
If ControlUsuario(mnuemisim.Caption) = 1 Then
   frm_emisim.Show vbModal
End If

End Sub

Private Sub mnuencue_Click()
If ControlUsuario(mnuencue.Caption) = 1 Then
   frm_encuestas.Show vbModal
End If

End Sub

Private Sub mnuentnew_Click()
If ControlUsuario(mnuentnew.Caption) = 1 Then
   frm_nuevas.Show vbModal
End If

End Sub

Private Sub mnuentregas_Click()
If ControlUsuario(mnuentregas.Caption) = 1 Then
   frm_entrega.Show vbModal
End If

End Sub

Private Sub mnuespecnew_Click()
If ControlUsuario(mnuespecnew.Caption) = 1 Then
    If frm_especialistas.Visible = True Then
       MsgBox "Ya está abierto"
    Else
       frm_especialistas.Show
    End If
End If

End Sub

Private Sub mnufaccnvnew1_Click()
If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
   frm_vtasconvcp.Show vbModal
Else
   frm_inffaccnv.Show vbModal
End If

End Sub

Private Sub mnufaccnvnew2_Click()
If ControlUsuario(mnufaccnvnew2.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasconvcp.Show vbModal
    Else
       frm_inffaccnv.Show vbModal
    End If
End If

End Sub

Private Sub mnufacconv_Click()
frm_inffaccnv.Show vbModal

End Sub



Private Sub mnufamili_Click()
If ControlUsuario(mnufamili.Caption) = 1 Then
   frm_famili.Show vbModal
End If

End Sub

Private Sub mnuficper_Click()
If ControlUsuario(mnuficper.Caption) = 1 Then
   frm_abmper.Show vbModal
End If

End Sub

Private Sub mnugenarq_Click()

If ControlUsuario(mnugenarq.Caption) = 1 Then

    Dim Clav As String
    Clav = InputBox("Ingrese su clave de SAPP para confirmar")
    If WxclaveU = Clav Then
       frm_genarq.Show vbModal
    Else
       MsgBox "Error en la clave"
    End If
Else
   MsgBox "Usuario no habilitado", vbCritical, "Arqueos"

End If

End Sub

Private Sub mnuimpemi_Click()
frm_impemi.Show vbModal

End Sub

Private Sub mnuimpeminew_Click()
If ControlUsuario(mnuimpeminew.Caption) = 1 Then
   frm_impemi.Show vbModal
End If

End Sub

Private Sub mnuinf_Click()
If ControlUsuario(mnuinf.Caption) = 1 Then
   frm_infdesp.Show vbModal
End If

End Sub

Private Sub mnuinfabm_Click()
frm_infabm.Show vbModal

End Sub

Private Sub mnuinfabmpsnew_Click()
If ControlUsuario(mnuinfabmpsnew.Caption) = 1 Then
   frm_infabm.Show vbModal
End If

End Sub

Private Sub mnuinfabnew_Click()
If ControlUsuario(mnuinfabnew.Caption) = 1 Then
   frm_infabmvar.Show vbModal
End If

End Sub

Private Sub mnuinfarq_Click()
If ControlUsuario(mnuinfarq.Caption) = 1 Then
   frm_infarq.Show vbModal
End If

End Sub

Private Sub mnuinfmsp_Click()
frm_infmsp.Show vbModal

End Sub

Private Sub mnuinfaz_Click()
If ControlUsuario(mnuinfaz.Caption) = 1 Then
   frm_infazul.Show vbModal
End If

End Sub

Private Sub mnuinfazuln_Click()
If ControlUsuario(mnuinfazuln.Caption) = 1 Then
   frm_infazul.Show vbModal
End If

End Sub

Private Sub mnuinfctrhcemov_Click()
If ControlUsuario(mnuinfctrhcemov.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       MsgBox "Opción no habilitada"
    Else
       frm_infhce.Show vbModal
    End If
End If

End Sub

Private Sub mnuinfctroltpnew_Click()
If ControlUsuario(mnuinfctroltpnew.Caption) = 1 Then
   frm_infctrolcons.Show vbModal
End If

End Sub

Private Sub mnuinfdemmnew_Click()
If ControlUsuario(mnuinfdemnew.Caption) = 1 Then
   frm_infdemoras.Show vbModal
End If

End Sub

Private Sub mnuinfdemnew_Click()
If ControlUsuario(mnuinfdemnew.Caption) = 1 Then
   frm_calidadiso.Show vbModal
End If

End Sub

Private Sub mnuinfdemorarl_Click()
If ControlUsuario(mnuinfdemorarl.Caption) = 1 Then
   frm_calidadiso.Show vbModal
End If

End Sub

Private Sub mnuinfdesp1_Click()
If ControlUsuario(mnuinfdesp1.Caption) = 1 Then
    frm_infdesp2.Show vbModal
End If

End Sub

Private Sub mnuinfdesp3_Click()
'frm_infdesp3.Show vbModal

End Sub

Private Sub mnuinfdespp_Click()
If ControlUsuario(mnuinfdespp.Caption) = 1 Then
   frm_infdesp2.Show vbModal
End If

End Sub

Private Sub mnuinfdeudanew_Click()
If ControlUsuario(mnuinfdeudanew.Caption) = 1 Then
   frm_infsaldos.Show vbModal
End If

End Sub

Private Sub mnuinfeco_Click()
If ControlUsuario(mnuinfeco.Caption) = 1 Then
   frm_infstock.Show vbModal
End If

End Sub

Private Sub mnuinfedades_Click()
'If ControlUsuario(mnuinfedades.Caption) = 1 Then
   frm_infedades.Show vbModal
'End If

End Sub

Private Sub mnuinfemisnew_Click()
If ControlUsuario(mnuinfemisnew.Caption) = 1 Then
   frm_infemis.Show vbModal
End If

End Sub

Private Sub mnuinfestudnew_Click()
If ControlUsuario(mnuinfestudnew.Caption) = 1 Then
   frm_infestud.Show vbModal
End If

End Sub

Private Sub mnuinfeval_Click()
If ControlUsuario(mnuinfeval.Caption) = 1 Then
   frm_infeval.Show vbModal
End If

End Sub

Private Sub mnuinfgestion_Click()
If ControlUsuario(mnuinfgestion.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       MsgBox "Se emite desde historial adm. del socio"
    Else
       frm_infaccadm.Show vbModal
    End If
End If

End Sub

Private Sub mnuinfhcemet_Click()
If ControlUsuario(mnuinfhcemet.Caption) = 1 Then
   frm_infhcemet.Show vbModal
End If

End Sub

Private Sub mnuinfmednew_Click()
If ControlUsuario(mnuinfmednew.Caption) = 1 Then
   frm_infmedxmut.Show vbModal
End If

End Sub

Private Sub mnuinfmedxmut_Click()
If ControlUsuario(mnuinfmedxmut.Caption) = 1 Then
   frm_infmedxmut.Show vbModal
End If

End Sub

Private Sub mnuinfprmednew_Click()
If ControlUsuario(mnuinfprmednew.Caption) = 1 Then
   frm_prodmed.Show vbModal
End If

End Sub

Private Sub mnuinfprod_Click()
frm_infprod.Show vbModal

End Sub

Private Sub mnuinfprodcnvnew_Click()
If ControlUsuario(mnuinfprodcnvnew.Caption) = 1 Then
   frm_infprod.Show vbModal
End If

End Sub

Private Sub mnuinfsinadinew_Click()


End Sub

Private Sub mnuinfsca_Click()
'frm_infsca.Show vbModal

End Sub

Private Sub mnuinfsmsnew_Click()
'If ControlUsuario(mnuinfsmsnew.Caption) = 1 Then
'   frm_sms.Show vbModal
'End If

End Sub

Private Sub mnuinfsocmut_Click()
frm_infsocmut.Show vbModal

End Sub

Private Sub mnuinfsocssnew_Click()
If ControlUsuario(mnuinfsocssnew.Caption) = 1 Then
   frm_infconsultas.Show vbModal
End If

End Sub

Private Sub mnuinfstocknew_Click()
If ControlUsuario(mnuinfstocknew.Caption) = 1 Then
   frm_infstock.Show vbModal
End If

End Sub

Private Sub mnuingcaj_Click()
If ControlUsuario(mnuingcaj.Caption) = 1 Then
   frm_caja.Show vbModal
End If

End Sub

Private Sub mnuingcomp_Click()
If ControlUsuario(mnuingcomp.Caption) = 1 Then
   frm_compsto.Show vbModal
End If

End Sub

Private Sub mnuinsolhc_Click()


End Sub

Private Sub mnulabcom_Click()
If ControlUsuario(mnulabcom.Caption) = 1 Then
   frm_labo.Show vbModal
End If

End Sub

Private Sub mnularg_Click()
If ControlUsuario(mnularg.Caption) = 1 Then
    WDespa = 0
    If frm_largador.Visible = True Then
       MsgBox "Ya está abierto"
    Else
       frm_largador.Show
    End If
End If

End Sub

Private Sub mnuliqesp_Click()
frm_liqesp.Show vbModal

End Sub

Private Sub mnuliqespnew_Click()
If ControlUsuario(mnuliqespnew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       MsgBox "Opción en construcción"
    Else
       frm_liqesp.Show vbModal
    End If
End If

End Sub

Private Sub mnulisest_Click()
frm_infestud.Show vbModal

End Sub

Private Sub mnullaconbmnew_Click()
If ControlUsuario(mnullaconbmnew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       MsgBox "Opción obsoleta"
    Else
       frm_vtallama.Show vbModal
    End If
End If

End Sub

Private Sub mnumane_Click()
If ControlUsuario(mnumane.Caption) = 1 Then
   frm_estudios.Show vbModal
End If

End Sub

Private Sub mnumant_Click()
'Dim XrecUsuario As New ADODB.Recordset
'Dim Xsqlusua As String

'ConectarBD
'ConbdSapp.Open

'Xsqlusua = "Select * from usua_permisos where id_usuario =" & Welnrou & " and opcion ='" & mnumant.Caption & "'"
'With XrecUsuario
'    .CursorLocation = adUseClient
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open Xsqlusua, ConbdSapp, , , adCmdText
'End With
If ControlUsuario(mnumant.Caption) = 1 Then
   If frmabm.Visible = True Then
      MsgBox "Ya está abierto!", vbCritical
   Else
      frmabm.Show
   End If
End If


End Sub





Private Sub mnumatestenf_Click()
If ControlUsuario(mnumatestenf.Caption) = 1 Then
   frm_matester.Show vbModal
End If

End Sub



Private Sub mnumedic_Click()
If ControlUsuario(mnumedic.Caption) = 1 Then
   frm_medicos.Show vbModal
End If

End Sub

Private Sub mnumejor_Click()
If ControlUsuario(mnumejor.Caption) = 1 Then

'    data_cargos.Connect = "odbc;dsn=" & Xconexrmt & ";"
'    data_cargos.RecordSource = "movil"
'    data_cargos.Refresh
'    If WElusuario = "MEUGENIA" Then
       frm_mejora.Show vbModal
'    Else
'        data_cargos.RecordSource = "select * from movil where medico ='" & WElusuario & "'"
'        data_cargos.Refresh
'        If data_cargos.Recordset.RecordCount > 0 Then
'           frm_mejora.Show vbModal
'        Else
'           MsgBox "Usuario no registrado en MAM.", vbCritical
'
'        End If
'    End If
End If

End Sub

Private Sub mnumodbases_Click()
If ControlUsuario(mnumodbases.Caption) = 1 Then
   frm_infmodpad.Show vbModal
End If

End Sub



Private Sub mnumodyafnew_Click()
If ControlUsuario(mnumodyafnew.Caption) = 1 Then
   frm_infmodpad.Show vbModal
End If

End Sub

Private Sub mnumutadm_Click()
If ControlUsuario(mnumutadm.Caption) = 1 Then
   frm_mutAdm.Show vbModal
End If

End Sub

Private Sub mnumutual_Click()
If ControlUsuario(mnumutual.Caption) = 1 Then
   frm_mutu.Show vbModal
End If

End Sub



Private Sub mnuparamemp_Click()
If ControlUsuario(mnuparamemp.Caption) = 1 Then
   frm_param.Show vbModal
End If

End Sub

Private Sub mnupasdev_Click()
If ControlUsuario(mnupasdev.Caption) = 1 Then
   frm_pasdev.Show vbModal
End If

End Sub

Private Sub mnupaspen_Click()
If ControlUsuario(mnupaspen.Caption) = 1 Then
   frm_paspen.Show vbModal
End If

End Sub

Private Sub mnupaspendos_Click()
If ControlUsuario(mnupaspendos.Caption) = 1 Then
   frm_paspendos.Show vbModal
End If

End Sub

Private Sub mnuplacmt_Click()
frm_planicmt.Show vbModal

End Sub

Private Sub mnupresta_Click()
If ControlUsuario(mnupresta.Caption) = 1 Then
   frm_prestamo.Show vbModal
End If

End Sub

Private Sub mnuprocasi_Click()
If ControlUsuario(mnuprocasi.Caption) = 1 Then
   frm_pasamem.Show vbModal
End If

End Sub

Private Sub mnuprocdebweb_Click()
Dim Procesacabal As String
If ControlUsuario(mnuprocdebweb.Caption) = 1 Then
'   Procesacabal = MsgBox("Desea procesar débitos de CABAL?", vbInformation + vbYesNo, "Débitos")
'   If Procesacabal = vbYes Then
      frm_debitos.Show vbModal
'   Else
'      Dim x
'      x = ShellExecute(Me.hwnd, "Open", "http://192.168.10.25:3000/debitos/envio", &O0, &O0, SW_NORMAL)
'   End If
End If
'frm_debitos.Show vbModal

End Sub

Private Sub mnuprocmutnew_Click()
If ControlUsuario(mnuprocmutnew.Caption) = 1 Then
   frm_ctrolmut.Show vbModal
End If

End Sub

Private Sub mnuprocred_Click()
'frm_proccred.Show vbModal

End Sub

Private Sub mnuprodmed_Click()
frm_prodmed.Show vbModal

End Sub

Private Sub mnupromfunc_Click()
If ControlUsuario(mnupromfunc.Caption) = 1 Then
   frm_abmvendefun.Show vbModal
End If

End Sub


Private Sub mnurecarq_Click()
'frm_recarq.Show vbModal

End Sub



Private Sub mnupromos_Click()
frm_promos.Show vbModal

End Sub

Private Sub mnuregactmetas_Click()
If ControlUsuario(mnuregactmetas.Caption) = 1 Then
   frm_metas.Show vbModal
End If

End Sub

Private Sub mnureggas_Click()
If ControlUsuario(mnureggas.Caption) = 1 Then
   frm_reggasto.Show vbModal
End If

End Sub

Private Sub mnuregoperac_Click()

If WElusuario = "ENRIQUE" Or WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Then
   Xmodulomant = 19
   frm_asismant.Show vbModal
Else
   Xmodulomant = 0
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub mnuregpro_Click()
If ControlUsuario(mnuregpro.Caption) = 1 Then
'If WElusuario = "AGUILLEN" Or WElusuario = "AGUSTINAC" Or WElusuario = "NROCHINOTTI" Then
   frm_ctrolstock.Show vbModal
'Else
 '  MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub mnureimp_Click()
If ControlUsuario(mnureimp.Caption) = 1 Then
   frm_reimpfact.Show vbModal
End If

End Sub

'Private Sub mnurespd_Click()
'frm_borrahist.Show vbModal

'End Sub


Private Sub mnurepag_Click()
If ControlUsuario(mnurepag.Caption) = 1 Then

    Dim Xdesea As String
    
    On Error GoTo RedP
    
    Xdesea = MsgBox("Desea procesar los pagos?", vbInformation + vbYesNo)
    
    If Xdesea = vbYes Then
        MsgBox "Verifique que esté guardado el archivo redpagos.xls en la carpeta PLANILLAS.", vbInformation
        
        'Ernesto
        frm_menu.MousePointer = 11
        
        data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
        
        data_redpagos.DatabaseName = "C:\planillas\redpagos.xls"
        data_redpagos.RecordSource = "redpagos$"
        data_redpagos.Refresh
        If data_redpagos.Recordset.RecordCount > 0 Then
           data_redpagos.Recordset.MoveFirst
           Do While Not data_redpagos.Recordset.EOF
              data_arq.RecordSource = "select * from arqueo where nrorec =" & data_redpagos.Recordset("factura") & " and cob =" & 221
              data_arq.Refresh
              If data_arq.Recordset.RecordCount > 0 Then
                 If data_arq.Recordset("arqueo") <> "C" Then
                    data_arq.Recordset.Edit
                    data_arq.Recordset("arqueo") = "C"
                    data_arq.Recordset("fecha") = data_redpagos.Recordset("fecha")
                    data_arq.Recordset("usuar") = WElusuario
                    data_arq.Recordset.Update
                 End If
              End If
              data_arq.RecordSource = "select * from deudas where documento =" & data_redpagos.Recordset("factura") & " and nro_cobr =" & 221
              data_arq.Refresh
              If data_arq.Recordset.RecordCount > 0 Then
                 If IsNull(data_arq.Recordset("fecha_pago")) = False Then
                 Else
                    data_arq.Recordset.Edit
                    data_arq.Recordset("fecha_pago") = data_redpagos.Recordset("fecha")
                    data_arq.Recordset.Update
                 End If
              End If
              data_redpagos.Recordset.MoveNext
           Loop
           frm_menu.MousePointer = 0
           MsgBox "Proceso terminado. Se actualizó arqueo y deudas.", vbInformation
           End
        End If
        'Al generar el arqueo ingresar cómo pendiente los cobradores redpagos y tarjetas
    End If
    
    Exit Sub
    
RedP:
         If Err.Number = 53 Then
            MsgBox "ERROR: Verifique si existe archivo", vbInformation
         Else
            MsgBox "ERROR al procesar, verifique si hay archivo", vbInformation
         End If
End If

End Sub



Private Sub mnuresconstot_Click()
If frm_pendcons.Visible = True Then
   MsgBox "Ya está abierto"
Else
   frm_pendcons.Show
End If

End Sub

Private Sub mnureshnf_Click()
If frm_especcovid.Visible = True Then
   MsgBox "Ya está abierto."
Else
   frm_especcovid.Show
End If

End Sub

Private Sub mnurubcaj_Click()
If ControlUsuario(mnurubcaj.Caption) = 1 Then
   frm_rubrec.Show vbModal
End If

End Sub

Private Sub mnurubcont_Click()
If ControlUsuario(mnurubcont.Caption) = 1 Then
   frm_rubcontab.Show vbModal
End If

End Sub

Private Sub mnurubgral_Click()
If ControlUsuario(mnurubgral.Caption) = 1 Then
   frm_rubteso.Show vbModal
End If

End Sub

Private Sub mnurucaf_Click()
If ControlUsuario(mnurucaf.Caption) = 1 Then
   frm_rucaf.Show vbModal
End If

End Sub

Private Sub mnusalir_Click()
End

End Sub

Private Sub mnuselconale_Click()
frm_infselcons.Show vbModal


End Sub

Private Sub mnusappuruw_Click()
If ControlUsuario(mnusappuruw.Caption) = 1 Then
   frm_uruware.Show vbModal
End If

End Sub

Private Sub mnuscap_Click()
'frm_pendsca.Show


End Sub

Private Sub mnusegunda_Click()
'If ControlUsuario(mnusegunda.Caption) = 1 Then
'   frm_serv2da.Show vbModal
'End If

End Sub

Private Sub mnuselconalenew_Click()
If ControlUsuario(mnuselconalenew.Caption) = 1 Then
   frm_infselcons.Show vbModal
End If

End Sub

Private Sub mnuservap_Click()
If ControlUsuario(mnuservap.Caption) = 1 Then
   If frm_servap.Visible = True Then
      MsgBox "Ya está abierto"
   Else
      frm_servap.Show
   End If
End If

End Sub

Private Sub mnuservmutnew_Click()
If ControlUsuario(mnuservmutnew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasxgpocp.Show vbModal
    Else
       frm_vtasxgpo.Show vbModal
    End If
End If

End Sub

Private Sub mnusocpromnew_Click()
If ControlUsuario(mnusocpromnew.Caption) = 1 Then
   frm_infabmpromo.Show vbModal
End If

End Sub

Private Sub mnusocxmutnew_Click()
If ControlUsuario(mnusocxmutnew.Caption) = 1 Then
   frm_infsocmut.Show vbModal
End If

End Sub

Private Sub mnusolasiman_Click()
Xmodulomant = 0
frm_asismant.Show vbModal

End Sub

Private Sub mnusolasis_Click()

frm_regasis.Show vbModal

End Sub

Private Sub mnusolhcu_Click()
frm_infsolhc.Show vbModal

End Sub

Private Sub mnusolhcctr_Click()
If ControlUsuario(mnusolhcctr.Caption) = 1 Then
   frm_infsolhc.Show vbModal
End If

End Sub

Private Sub mnusolhisop_Click()
If frm_solhisopa.Visible = True Then
   MsgBox "Ya está abierto"
Else
   frm_solhisopa.Show
   
End If

End Sub

Private Sub mnusolibaja_Click()
If ControlUsuario(mnusolibaja.Caption) = 1 Then
   frm_solicitudbaja.Show vbModal
End If

End Sub

Private Sub mnusolmark_Click()
frm_asismark.Show vbModal

End Sub

Private Sub mnusolmejo_Click()
frm_solmejoras.Show vbModal

End Sub

Private Sub mnusolpsoc_Click()
frm_asispadron.Show vbModal

End Sub

Private Sub mnusolrrhh_Click()
frm_asisrrhh.Show vbModal

End Sub

Private Sub mnusrvenfdom_Click()
If ControlUsuario(mnusrvenfdom.Caption) = 1 Then
    If frm_srvenferm.Visible = True Then
       MsgBox "Ya está abierto"
    Else
       frm_srvenferm.Show
    End If
End If

End Sub

Private Sub mnusuel_Click()
If ControlUsuario(mnusuel.Caption) = 1 Then
   frm_sueldos.Show vbModal
End If

End Sub

Private Sub mnutabcob_Click()
If ControlUsuario(mnutabcob.Caption) = 1 Then
   frm_cobr.Show vbModal
End If

End Sub

Private Sub mnutabconv_Click()
If ControlUsuario(mnutabconv.Caption) = 1 Then
   frm_convenios.Show vbModal
End If

End Sub

Private Sub mnutabpro_Click()
If ControlUsuario(mnutabpro.Caption) = 1 Then
   frm_prom.Show vbModal
End If

End Sub

Private Sub mnutarj_Click()
If ControlUsuario(mnutarj.Caption) = 1 Then
   frm_personal.Show vbModal
End If

End Sub

Private Sub mnutes_Click()
If ControlUsuario(mnutes.Caption) = 1 Then
   frm_teso.Show vbModal
End If

End Sub

Private Sub mnuUruwa_Click()
If ControlUsuario(mnuUruwa.Caption) = 1 Then
    Dim X
    On Error GoTo Vererralabrir
    
    X = ShellExecute(Me.hwnd, "Open", "https://prod1974.ucfe.com.uy/Gestion/", "", "", 1)
    
    Exit Sub
    
Vererralabrir:
                  If Err.Number = 53 Then
                     MsgBox "No se encontró el archivo"
                  Else
                     MsgBox "Error al abrir la página"
                  End If
End If

End Sub

Private Sub mnuusu_Click()

If WElusuario = "COMPUTOS" Then
   frm_usuarios.Show vbModal
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub mnuutildesp_Click()
If ControlUsuario(mnuutildesp.Caption) = 1 Then
    If frm_opsdesp.Visible = True Then
       MsgBox "Ya está abierto"
    Else
       frm_opsdesp.Show
    End If
Else
    MsgBox "Usuario sin permisos"
End If

End Sub



Private Sub mnuvences_Click()
If ControlUsuario(mnuvences.Caption) = 1 Then
   frm_matvence.Show vbModal
End If


End Sub

Private Sub mnuverll_Click()
If frm_llamadotot.Visible = True Then
   MsgBox "Ya está abierto"
Else
   frm_llamadotot.Show vbModal
End If

End Sub


Private Sub mnuverlladesp_Click()
If ControlUsuario(mnuverlladesp.Caption) = 1 Then
    If frm_llamadotot.Visible = True Then
       MsgBox "Ya está abierto"
    Else
       frm_llamadotot.Show vbModal
    End If
End If

End Sub

Private Sub mnuverllaenf_Click()
If ControlUsuario(mnuverllaenf.Caption) = 1 Then
    If frm_llamadotot.Visible = True Then
       MsgBox "Ya está abierto"
    Else
       frm_llamadotot.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtagpo_Click()
frm_vtasxgpo.Show vbModal

End Sub

Private Sub mnuvtallam_Click()
frm_vtallama.Show vbModal

End Sub

Private Sub mnuvtascre_Click()
frm_infvtascre.Show vbModal

End Sub

Private Sub mnuvtascredadm_Click()
If ControlUsuario(mnuvtascredadm.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasservcp.Show vbModal
    Else
       frm_infvtascre.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtaserusu_Click()
If ControlUsuario(mnuvtaserusu.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasservcp.Show vbModal
    Else
       frm_vtasservjv.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtaslab_Click()
If ControlUsuario(mnuvtaslab.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       MsgBox "Opción para laboratorios"
    Else
       frm_inflab.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtasmednewr_Click()
If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
   frm_vtasmedcp2.Show vbModal
Else
   frm_vtasmed.Show vbModal
End If

End Sub

Private Sub mnuvtasxconvnew_Click()
If ControlUsuario(mnuvtasxconvnew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasconvcp.Show vbModal
    Else
       frm_vtasconv.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtasxflianew_Click()
If ControlUsuario(mnuvtasxflianew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasfliacp.Show vbModal
    Else
       frm_vtasflia.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtasxflianewr_Click()
If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
   frm_vtasfliacp.Show vbModal
Else
   frm_vtasflia.Show vbModal
End If

End Sub

Private Sub mnuvtasxmednew_Click()
If ControlUsuario(mnuvtasxmednew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasmedcp2.Show vbModal
    Else
       frm_vtasmed.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtasxmutnew_Click()
If ControlUsuario(mnuvtasxmutnew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasxgpocp.Show vbModal
    Else
       frm_vtasxgpo.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtasxtipfacnew_Click()
If ControlUsuario(mnuvtasxtipfacnew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasxfaccp.Show vbModal
    Else
       frm_vtasxfac.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtaxmed_Click()
frm_vtasmed.Show vbModal

End Sub

Private Sub mnuvtaxser_Click()
frm_vtasserv.Show vbModal

End Sub

Private Sub mnuvtaxsernew_Click()
If ControlUsuario(mnuvtaxsernew.Caption) = 1 Then
    If WElusuario = "CONTADURIA" Or WElusuario = "JORDAN" Or WElusuario = "CONTADOR" Then
       frm_vtasservcp.Show vbModal
    Else
       frm_vtasserv.Show vbModal
    End If
End If

End Sub

Private Sub mnuvtipof_Click()
frm_vtasxfac.Show vbModal

End Sub

Private Sub mnuvtllama_Click()
frm_vtallama.Show vbModal

End Sub

Private Sub mnuvxcon_Click()
frm_vtasconv.Show vbModal

End Sub

Private Sub mnuvxfam_Click()
frm_vtasflia.Show vbModal

End Sub

Private Sub mnuzona_Click()
If ControlUsuario(mnuzona.Caption) = 1 Then
   frm_zonas.Show vbModal
End If

End Sub



Private Sub porconvenio_Click()
frm_vtasconv.Show vbModal

End Sub

Private Sub porfamilia_Click()
frm_vtasflia.Show vbModal

End Sub

Private Sub pormedico_Click()
frm_vtasmed.Show vbModal

End Sub

Private Sub pormedico2_Click()
frm_liqesp.Show vbModal

End Sub

Private Sub pormutual_Click()
frm_vtasxgpo.Show vbModal

End Sub

Private Sub porserv_Click()
frm_vtasserv.Show vbModal

End Sub

Private Sub proctimbres_Click()
If ControlUsuario(proctimbres.Caption) = 1 Then
   frm_proctimbre.Show vbModal
End If

End Sub


Private Sub Timer1_Timer()
labhora.Caption = Format(Time, "HH:mm:ss")
End Sub

Private Sub varioslla_Click()
frm_infdesp.Show vbModal

End Sub
Public Sub Control_vence()
Dim Xsqlpromo, XsqlCons, XelGrupo As String
Dim Xrecclii As New ADODB.Recordset
Dim XrecGraba As New ADODB.Recordset
Dim XfechaChe As Date
XfechaChe = Date + 10

XelGrupo = ""

ConectarBD
ConbdSapp.Open

Xsqlpromo = "Select * from tesorero where vence_cheq >='" & Format(Date, "yyyy-mm-dd") & "' and vence_cheq <='" & Format(XfechaChe, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      MsgBox "Cheque vence el " & Format(Xrecclii("vence_cheq"), "dd/mm/yyyy") & " Obs:" & Xrecclii("obs")
      Xrecclii.MoveNext
   Loop
Else
   MsgBox "No hay cheques a vencer en los 10 días.", vbInformation
End If
Xrecclii.Close

ConbdSapp.Close

End Sub


Public Sub Genera_codigo()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim cod As Integer
Dim Valorminuto As Integer


Valorminuto = Val(Mid(Format(Time, "HH:mm"), 4, 2))

cod = Int(Rnd * 10000)
cod = cod + Valorminuto

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
                          
Xsqlpromo = "Select * from codaut_devol where codigo =" & cod

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
   

Xrecclii.AddNew
Xrecclii("fecha") = Date
Xrecclii("hora") = Format(Time, "HH:mm")
Xrecclii("usuario") = WElusuario
Xrecclii("base") = data_parse.Recordset("base")
Xrecclii("codigo") = cod
Xrecclii("usado") = 0
Xrecclii.Update

MsgBox "Código de autorización: " & Trim(str(cod))

Xrecclii.Close
ConbdSapp.Close

End Sub
