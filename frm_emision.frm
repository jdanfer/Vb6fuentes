VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_emision 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Emisión Mensual (Generación)"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   Icon            =   "frm_emision.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cabanual 
      Caption         =   "data_cabanual"
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
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_emianual 
      Caption         =   "data_emianual"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_clipromo 
      Caption         =   "data_clipromo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Data data_promos 
      Caption         =   "data_promos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_refinan 
      Caption         =   "data_refinan"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_informe 
      Caption         =   "data_informe"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_emicopi 
      Caption         =   "data_emicopi"
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
      Top             =   960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_emision 
      Caption         =   "data_emision"
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc data_estud 
      Height          =   375
      Left            =   2160
      Top             =   2040
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=sappnew"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_estud"
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
   Begin MSAdodcLib.Adodc data_deu 
      Height          =   375
      Left            =   5280
      Top             =   3960
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_deu"
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
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   2880
      Top             =   3240
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_cli"
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
      Height          =   375
      Left            =   -240
      Top             =   1800
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=sappnew"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "sappnew"
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
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   7080
      TabIndex        =   25
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Renumerar"
      Height          =   495
      Left            =   4800
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data data_eror 
      Caption         =   "data_eror"
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
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restaurar"
      Height          =   495
      Left            =   3120
      TabIndex        =   23
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      DataField       =   "qr"
      DataSource      =   "data_emicopi"
      Height          =   1575
      Left            =   5760
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data data_noemite 
      Caption         =   "data_noemite"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton b_env 
      Caption         =   "Enviar 11.2020"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton b_efct 
      Caption         =   "efct"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   495
      Left            =   5040
      Picture         =   "frm_emision.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton b_etck 
      Caption         =   "etck"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton b_fin 
      Caption         =   "FIN"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "odbc;dsn=sappfact;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "paramsapp"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_ctrolmes 
      Caption         =   "data_ctrolmes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "ODBC;DSN=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data data_rectiq 
      Caption         =   "data_rectiq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EMITIQ"
      Top             =   1560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin Crystal.CrystalReport crem 
      Left            =   5640
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_ultrec2 
      Caption         =   "data_ultrec2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "NROSREC"
      Top             =   0
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Data data_ultrec 
      Caption         =   "data_ultrec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ULTNRO"
      Top             =   0
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data data_emitiq 
      Caption         =   "data_emitiq"
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
      RecordSource    =   "EMITIQ"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Data data_usua 
      Caption         =   "data_usua"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_ultemiemi 
      Caption         =   "data_ultemiemi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ULTEMI"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_ultemisim 
      Caption         =   "data_ultemisim"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ULTSIM"
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar"
      Height          =   495
      Left            =   1200
      Picture         =   "frm_emision.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label labadenda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "El pago de este recibo no cancela deudas anteriores."
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Procesando..."
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
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label labvenceok 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labcodseg 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label labcae 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label labautoriza 
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label labvence 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labcabemi 
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labnomemi 
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labnrofact 
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labserie 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label labfec 
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "FECHA:"
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
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label labano 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label labmes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "EMISION:"
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
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   240
      Picture         =   "frm_emision.frx":0F56
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
End
Attribute VB_Name = "frm_emision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objPosCfe As PosCfe
Dim Miscr As Scripting.FileSystemObject
Dim xmlsc As TextStream
Dim Xlafac, Xlaserieref As String

Dim objUltimaSerieNumero As SerieNumeroCfe

Dim strUltimoGuid As String

Dim strIdTransaccionPos2000 As String

Private Sub b_efct_Click()
Dim strIdTransac As String

Dim Xindi, Xlalinea As Integer
Dim Ximpposi As Double
Ximpposi = 0

Dim Xnograva2 As Double
Xnograva2 = 0

Set objPosCfe = New PosCfe
Dim objresultado As Resultado
'Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
   Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
Else
'    If frm_menu.data_parse.Recordset("base") = 38 Then
'       Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
'    Else
       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)
'    End If
End If

Dim strMensaje As String
strMensaje = "No se pudo inicializar el POS"
If objresultado Is Nothing Then
   MsgBox strMensaje
   Exit Sub
End If
If Not objresultado.OperacionExitosa Then
   If objresultado.Mensaje <> vbNullString Then strMensaje = strMensaje & ": " & objresultado.Mensaje
      MsgBox strMensaje
      Exit Sub
End If
strIdTransac = objPosCfe.CrearGuid
If Not EstaInicializado() Then Exit Sub
   Dim objresultado22 As ResultadoConsultaConexion
   Set objresultado22 = objPosCfe.ObtenerEstadoConexion
   Dim strMensaje22 As String
   strMensaje22 = "No se pudo consultar el estado de la conexión"
   If objresultado22 Is Nothing Then
      MsgBox strMensaje22
      Exit Sub
   End If
   If Not objresultado22.OperacionExitosa Then
      If objresultado22.Mensaje <> vbNullString Then strMensaje22 = strMensaje22 & ": " & objresultado22.Mensaje
         MsgBox strMensaje22
         Exit Sub
   End If
'  MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'       "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
  If Not EstaInicializado() Then Exit Sub
     Dim objCfe As CFE
     Dim objCf As ClassFactory
    
     data_emicopi.RecordSource = "Select * from " & labnomemi.Caption & " where debe_haber =" & 111 & " order by nro_cobr, importe"
     data_emicopi.Refresh
     If data_emicopi.Recordset.RecordCount > 0 Then
        data_emicopi.Recordset.MoveFirst
        Do While Not data_emicopi.Recordset.EOF
           Set objCfe = New CFE
           Set objCf = New ClassFactory
           Set objCfe.EFact = New EFact
'           data_emision.RecordSource = "Select * from " & labcabemi.Caption & " where cliente2 =" & data_emicopi.Recordset("cliente") & " and nro_linea not in (12) order by nro_linea"
           data_emision.RecordSource = "Select * from " & labcabemi.Caption & " where cliente2 =" & data_emicopi.Recordset("cliente") & " and nro_linea not in (12) and mesc =" & data_emicopi.Recordset("mes") & " and anioc =" & data_emicopi.Recordset("ano") & " order by nro_linea"
           data_emision.Refresh
           If data_emision.Recordset.RecordCount > 0 Then
              data_emision.Recordset.MoveFirst
              With objCfe.EFact.Encabezado.IdDoc
                  .TipoCFE = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(data_emicopi.Recordset("debe_haber"))))
                  .FchEmis.SetDate Year(data_emicopi.Recordset("fecha")), Month(data_emicopi.Recordset("fecha")), Day(data_emicopi.Recordset("fecha"))
                  .IsValidMntBruto = True
                  .MntBruto = IdDoc_Tck_MntBruto_1
                  .FmaPago = IdDoc_Fact_FmaPago_2
              End With
              With objCfe.EFact.Encabezado.Emisor
                  .RUCEmisor = data_par.Recordset("ruc")
                  .RznSoc = data_par.Recordset("nomc")
                  .CdgDGISucur.FromString Trim(str(data_par.Recordset("codsuc")))
                  .DomFiscal = data_par.Recordset("domic")
                  .Ciudad = data_par.Recordset("ciudad")
                  .Departamento = data_par.Recordset("dpto")
              End With
              With objCfe.EFact.Encabezado.Receptor
                  .TipoDocRecep = DocType_2
                  .CodPaisRecep = CodPaisType_UY
                  .DocRecep = data_emicopi.Recordset("ruc")
                  .RznSocRecep = data_emicopi.Recordset("apellidos")
                  .DirRecep = data_emicopi.Recordset("dir_cli")
                  .CiudadRecep = data_emicopi.Recordset("zona")
              End With
              With objCfe.EFact.Encabezado.Totales
                   If IsNull(data_emicopi.Recordset("tiquet")) = False Then
                      Xnograva2 = data_emicopi.Recordset("tiquet")
                   End If
                   If IsNull(data_emicopi.Recordset("deudas")) = False Then
                      Xnograva2 = Xnograva2 + data_emicopi.Recordset("deudas")
                   End If
                   .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_emicopi.Recordset("tipodoc"))
                   .IsValidTpoCambio = True
                   .TpoCambio.FromString "1"
                   .IsValidMntNetoIvaTasaMin = True
                   .IsValidMntIVATasaMin = True
                   .IsValidMntNoGrv = True
                   .MntNoGrv.FromString Format(Xnograva2, "0.00")
                   .MntNetoIvaTasaMin.FromString Format(data_emicopi.Recordset("servi"), "0.00")
                   .IVATasaMin = TasaIVAType_10FullStop000
                   .MntIVATasaMin.FromString Format(data_emicopi.Recordset("iva"), "0.00")
                   .CantLinDet.FromString data_emicopi.Recordset("numero")
                   .MntTotal.FromString Format(data_emicopi.Recordset("total"), "0.00")
                   .MntPagar.FromString Format(data_emicopi.Recordset("total"), "0.00")
              End With
              Do While Not data_emision.Recordset.EOF
                 With objCfe.EFact.Detalle.Item.AddNew
                       If data_emision.Recordset("serie") = "DS" Then
                          Ximpposi = data_emision.Recordset("monto") - data_emision.Recordset("nro_doc")
                         .NroLinDet.FromString Trim(str(data_emision.Recordset("nro_linea")))
                         .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_emision.Recordset("indic_fact"))))
                         .NomItem = data_emision.Recordset("descrip")
                         .cantidad.FromString Trim(str(data_emision.Recordset("cantidad")))
                         .UniMed = "N/A"
                         .PrecioUnitario.FromString Format(data_emision.Recordset("imp_srv"), "0.00")
                         .IsValidDescuentoMonto = True
                         .IsValidDescuentoPct = True
                         .DescuentoPct.FromString Format(data_emicopi.Recordset("descpor"), "0")
                         .DescuentoMonto.FromString Format(data_emision.Recordset("nro_doc"), "0.00")
                         .MontoItem.FromString Format(data_emision.Recordset("monto"), "0.00")
                       Else
                         .NroLinDet.FromString Trim(str(data_emision.Recordset("nro_linea")))
                         .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_emision.Recordset("indic_fact"))))
                         .NomItem = data_emision.Recordset("descrip")
                         .cantidad.FromString Trim(str(data_emision.Recordset("cantidad")))
                         .UniMed = "N/A"
                         .PrecioUnitario.FromString Format(data_emision.Recordset("imp_srv"), "0.00")
                         .MontoItem.FromString Format(data_emision.Recordset("monto"), "0.00")
                      End If
                 End With
                 data_emision.Recordset.MoveNext
              Loop
              Dim s As String
              s = objCfe.ToXml(True, XmlFormatting_Indented)
            '       Text1.Text = s
              Dim strGuid As String
              strGuid = objPosCfe.CrearGuid()
              Dim objResultadoCfe As ResultadoCfe
              Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
              Set objUltimaSerieNumero = Nothing
              DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
              If Not objUltimaSerieNumero Is Nothing Then _
                     ' cmdFirmarNc.Enabled = True
            '          MsgBox "firmar NC"
              End If
              If labserie.Caption <> "" And labnrofact.Caption <> "" Then
                 data_emicopi.Recordset.Edit
                 data_emicopi.Recordset("tipocta") = labserie.Caption
                 data_emicopi.Recordset("documento") = labnrofact.Caption
                 data_emicopi.Recordset.Update
              End If
              labserie.Caption = ""
              labnrofact.Caption = ""
              Set objCfe = Nothing
              Set objCf = Nothing
              data_emicopi.Recordset.MoveNext
           Else
              labserie.Caption = ""
              labnrofact.Caption = ""
              Set objCfe = Nothing
              Set objCf = Nothing
              data_emicopi.Recordset.MoveNext
           End If
        Loop
     End If
     frm_emision.MousePointer = 0
     Label1.Caption = "Procesando tareas finales..."
     b_fin_Click


End Sub

Private Sub b_env_Click()
'borrar los que ya están numerados
'en boton e_fact dar finalización
'después ejecutar únicamente el FIN

b_etck_Click '''no borrar


'data_emicopi.RecordSource = "Select * from " & labnomemi.Caption & " where debe_haber =" & 101 & " order by nro_cobr, importe"
'data_emicopi.Refresh
'If data_emicopi.Recordset.RecordCount > 0 Then
'   data_emicopi.Recordset.MoveFirst
'   Do While Not data_emicopi.Recordset.EOF
'      b_etck_Click
'      data_emicopi.Recordset.MoveNext
'   Loop
'   data_emicopi.RecordSource = "Select * from " & labnomemi.Caption & " where debe_haber =" & 111 & " order by nro_cobr, importe"
'   data_emicopi.Refresh
'   If data_emicopi.Recordset.RecordCount > 0 Then
'      data_emicopi.Recordset.MoveFirst
'      Do While Not data_emicopi.Recordset.EOF
'         b_efct_Click
'         data_emicopi.Recordset.MoveNext
'      Loop
'   End If
'End If
'MsgBox "Proceso de numeración terminado. Se copiará la emisión definitiva y se crearán los informes.", vbInformation
'b_fin_Click



End Sub

Private Sub b_etck_Click()
Dim strIdTransac As String
Dim Xnograva, XimpCuota As Double

Dim Xindi, Xlalinea As Integer
Dim Ximpposi As Double
Ximpposi = 0


Xnograva = 0
XimpCuota = 0
Label1.Visible = True
Label1.Caption = "Aguarde, numerando..."
frm_emision.MousePointer = 11
Set objPosCfe = New PosCfe
Dim objresultado As Resultado
'''''''Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
   Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
'   Set objresultado = objPosCfe.Inicializar("SAPP-105", frm_menu.data_parse.Recordset("contacto"), vbNullString)
Else
   Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)
'   Set objresultado = objPosCfe.Inicializar("SAPP-105", "SAPP-206", vbNullString)
End If
Dim strMensaje As String
strMensaje = "No se pudo inicializar el POS"
If objresultado Is Nothing Then
   MsgBox strMensaje
   Exit Sub
End If
If Not objresultado.OperacionExitosa Then
   If objresultado.Mensaje <> vbNullString Then strMensaje = strMensaje & ": " & objresultado.Mensaje
      MsgBox strMensaje
      Exit Sub
End If
strIdTransac = objPosCfe.CrearGuid
 'estado de la conexión
If Not EstaInicializado() Then Exit Sub
   Dim objresultado22 As ResultadoConsultaConexion
   Set objresultado22 = objPosCfe.ObtenerEstadoConexion
   Dim strMensaje22 As String
   strMensaje22 = "No se pudo consultar el estado de la conexión"
   If objresultado22 Is Nothing Then
      MsgBox strMensaje22
      Exit Sub
   End If
   If Not objresultado22.OperacionExitosa Then
      If objresultado22.Mensaje <> vbNullString Then strMensaje22 = strMensaje22 & ": " & objresultado22.Mensaje
         MsgBox strMensaje22
        Exit Sub
   End If
'   MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'         "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
 
   If Not EstaInicializado() Then Exit Sub
      Dim objCfe As CFE
      Dim objCf As ClassFactory
      data_emicopi.RecordSource = "Select * from " & labnomemi.Caption & " where debe_haber =" & 101 & " order by nro_cobr, importe"
      data_emicopi.Refresh
      If data_emicopi.Recordset.RecordCount > 0 Then
         data_emicopi.Recordset.MoveFirst
         Do While Not data_emicopi.Recordset.EOF
            Set objCfe = New CFE
            Set objCf = New ClassFactory
            Set objCfe.ETck = New ETck
            data_emision.RecordSource = "Select * from " & labcabemi.Caption & " where cliente2 =" & data_emicopi.Recordset("cliente") & " and nro_linea not in (12) and mesc =" & data_emicopi.Recordset("mes") & " and anioc =" & data_emicopi.Recordset("ano") & " order by nro_linea"
            data_emision.Refresh
            If data_emision.Recordset.RecordCount > 0 Then
                data_emision.Recordset.MoveFirst
                With objCfe.ETck.Encabezado.IdDoc
                  .TipoCFE = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_emicopi.Recordset("debe_haber"))))
                  .FchEmis.SetDate Year(data_emicopi.Recordset("fecha")), Month(data_emicopi.Recordset("fecha")), Day(data_emicopi.Recordset("fecha"))
                  .IsValidMntBruto = True
                  .MntBruto = IdDoc_Tck_MntBruto_1
                  .FmaPago = IdDoc_Tck_FmaPago_2
                End With
                With objCfe.ETck.Encabezado.Emisor
                  .RUCEmisor = data_par.Recordset("ruc")
                  .RznSoc = data_par.Recordset("nomc")
                  .CdgDGISucur.FromString Trim(str(data_par.Recordset("codsuc")))
                  .DomFiscal = data_par.Recordset("domic")
                  .Ciudad = data_par.Recordset("ciudad")
                  .Departamento = data_par.Recordset("dpto")
                End With
                Set objCfe.ETck.Encabezado.Receptor = New Receptor_Tck
                Set objCfe.ETck.Encabezado.Receptor.Receptor_Tck_Choice = New Receptor_Tck_Choice
                With objCfe.ETck.Encabezado.Receptor
                  .TipoDocRecep = DocType_4
                  .CodPaisRecep = CodPaisType_UY
                  .Receptor_Tck_Choice.DocRecepExt = Trim(str(data_emicopi.Recordset("cliente")))
                  .RznSocRecep = data_emicopi.Recordset("apellidos")
                  .DirRecep = data_emicopi.Recordset("dir_cli")
                  .CiudadRecep = data_emicopi.Recordset("zona")
                End With
                With objCfe.ETck.Encabezado.Totales
                  If IsNull(data_emicopi.Recordset("tiquet")) = False Then
                     Xnograva = data_emicopi.Recordset("tiquet")
                  End If
                  If IsNull(data_emicopi.Recordset("deudas")) = False Then
                     Xnograva = Xnograva + data_emicopi.Recordset("deudas")
                  End If
                  .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_emicopi.Recordset("tipodoc"))
                  .IsValidTpoCambio = True
                  .TpoCambio.FromString "1"
                  .IsValidMntNetoIvaTasaMin = True
                  .IsValidMntIVATasaMin = True
                  .IsValidMntNoGrv = True
                  .MntNoGrv.FromString Format(Xnograva, "0.00")
                  .MntNetoIvaTasaMin.FromString Format(data_emicopi.Recordset("servi"), "0.00")
                  .IVATasaMin = TasaIVAType_10FullStop000
                  .MntIVATasaMin.FromString Format(data_emicopi.Recordset("iva"), "0.00")
                  .CantLinDet.FromString Trim(str(data_emicopi.Recordset("numero")))
                  .MntTotal.FromString Format(data_emicopi.Recordset("total"), "0.00")
                  .MntPagar.FromString Format(data_emicopi.Recordset("total"), "0.00")
                End With
                
                Do While Not data_emision.Recordset.EOF
                   With objCfe.ETck.Detalle.Item.AddNew
                       If data_emision.Recordset("serie") = "DS" Then
                          Ximpposi = data_emision.Recordset("monto") - data_emision.Recordset("nro_doc")
                          .NroLinDet.FromString Trim(str(data_emision.Recordset("nro_linea")))
                          .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_emision.Recordset("indic_fact"))))
                          .NomItem = data_emision.Recordset("descrip")
                          .cantidad.FromString Trim(str(data_emision.Recordset("cantidad")))
                          .UniMed = "N/A"
                          .PrecioUnitario.FromString Format(data_emision.Recordset("imp_srv"), "0.00")
                          .IsValidDescuentoMonto = True
                          .IsValidDescuentoPct = True
                          .DescuentoPct.FromString Format(data_emicopi.Recordset("descpor"), "0")
                          .DescuentoMonto.FromString Format(data_emision.Recordset("nro_doc"), "0.00")
                          .MontoItem.FromString Format(Ximpposi, "0.00")
                       Else
                          .NroLinDet.FromString Trim(str(data_emision.Recordset("nro_linea")))
                          .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_emision.Recordset("indic_fact"))))
                          .NomItem = data_emision.Recordset("descrip")
                          .cantidad.FromString Trim(str(data_emision.Recordset("cantidad")))
                          .UniMed = "N/A"
                          .PrecioUnitario.FromString Format(data_emision.Recordset("imp_srv"), "0.00")
                          .MontoItem.FromString Format(data_emision.Recordset("monto"), "0.00")
                       End If
                   End With
                   data_emision.Recordset.MoveNext
                Loop
                Dim s As String
                s = objCfe.ToXml(True, XmlFormatting_Indented)
                Dim strGuid As String
                strGuid = objPosCfe.CrearGuid()
                Dim objResultadoCfe As ResultadoCfe
'                If IsNull(data_cabeza2.Recordset("obsp")) = False Then
'                   Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))
'                Else
'                   Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
'                End If
                
                Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
                Set objUltimaSerieNumero = Nothing
                DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
                If Not objUltimaSerieNumero Is Nothing Then _
                    ' cmdFirmarNc.Enabled = True
            '       MsgBox "firmar NC"
                End If
                
                If labserie.Caption <> "" And labnrofact.Caption <> "" Then
                   data_emicopi.Recordset.Edit
                   data_emicopi.Recordset("tipocta") = labserie.Caption
                   data_emicopi.Recordset("documento") = labnrofact.Caption
                   data_emicopi.Recordset.Update
                End If
                data_emicopi.Recordset.MoveNext
                
                labserie.Caption = ""
                labnrofact.Caption = ""
                Set objCfe = Nothing
                Set objCf = Nothing
            Else
                labserie.Caption = ""
                labnrofact.Caption = ""
                Set objCfe = Nothing
                Set objCf = Nothing
                data_emicopi.Recordset.MoveNext
            End If
         Loop
      Else
         MsgBox "No hay e-ticket para numerar" & data_emicopi.Recordset("cliente")
         data_emicopi.Recordset.MoveNext
      End If
      b_efct_Click
''     b_fin_Click
'   b_efct_Click


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub b_fin_Click()

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim Xnumeroprov As Integer
Xnumeroprov = 100
MiBaseact.Execute "Delete * from infemis"

Label1.Visible = True
''''numeracion provisoria, después borrar

''''borrar hasta acá
data_emicopi.RecordSource = "Select * from " & labnomemi.Caption & " order by documento"
data_emicopi.Refresh
'''Realiza control de las facturas q no fueron numeradas por algún error y lo guarda en la tabla de no emitidos para verificar
If data_emicopi.Recordset.RecordCount > 0 Then
   data_emicopi.Recordset.MoveFirst
   
   Do While Not data_emicopi.Recordset.EOF
      If IsNull(data_emicopi.Recordset("documento")) = True Then
         data_noemite.Recordset.AddNew
         data_noemite.Recordset("cliente") = data_emicopi.Recordset("cliente")
         data_noemite.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
         data_noemite.Recordset("apellidos") = data_emicopi.Recordset("apellidos")
         data_noemite.Recordset("importe") = data_emicopi.Recordset("importe")
         data_noemite.Recordset("mes") = data_emicopi.Recordset("mes")
         data_noemite.Recordset("ano") = data_emicopi.Recordset("ano")
         data_noemite.Recordset("color_rec") = data_emicopi.Recordset("color_rec")
         data_noemite.Recordset("total") = data_emicopi.Recordset("total")
         data_noemite.Recordset.Update
         data_emicopi.Recordset.Delete
      Else
         If data_emicopi.Recordset("documento") = 0 Then
            data_noemite.Recordset.AddNew
            data_noemite.Recordset("cliente") = data_emicopi.Recordset("cliente")
            data_noemite.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
            data_noemite.Recordset("apellidos") = data_emicopi.Recordset("apellidos")
            data_noemite.Recordset("importe") = data_emicopi.Recordset("importe")
            data_noemite.Recordset("mes") = data_emicopi.Recordset("mes")
            data_noemite.Recordset("ano") = data_emicopi.Recordset("ano")
            data_noemite.Recordset("color_rec") = data_emicopi.Recordset("color_rec")
            data_noemite.Recordset("total") = data_emicopi.Recordset("total")
            data_noemite.Recordset.Update

            data_emicopi.Recordset.Delete
         End If
      End If
      data_emicopi.Recordset.MoveNext
   Loop
End If

Label1.Caption = "Procesando numeración de facturas..."
DoEvents
''''Realiza la actualización de la numeración en la tabla cabezal
If data_emicopi.Recordset.RecordCount > 0 Then
   data_emicopi.Recordset.MoveFirst
   Do While Not data_emicopi.Recordset.EOF
      data_emision.RecordSource = "Select * from " & labcabemi.Caption & " where cliente2 =" & data_emicopi.Recordset("cliente") & " and mesc =" & data_emicopi.Recordset("mes") & " and anioc =" & data_emicopi.Recordset("ano")
      data_emision.Refresh
      If data_emision.Recordset.RecordCount > 0 Then
         data_emision.Recordset.MoveFirst
         Do While Not data_emision.Recordset.EOF
            data_emision.Recordset.Edit
            data_emision.Recordset("serie") = data_emicopi.Recordset("tipocta")
            data_emision.Recordset("nro_doc") = data_emicopi.Recordset("documento")
            data_emision.Recordset.Update
            data_emision.Recordset.MoveNext
         Loop
      End If
      data_emicopi.Recordset.MoveNext
   Loop
'''   MsgBox "Se continúa con la carga de las facturas a la deuda del socio.", vbInformation
   data_emision.RecordSource = "Select * from " & labcabemi.Caption & " where nro_doc is null"
   data_emision.Refresh
   If data_emision.Recordset.RecordCount > 0 Then
      data_emision.Recordset.MoveLast
      If data_emision.Recordset.RecordCount <= 100 Then 'Si hay más de 100 registros sin numerar, enviar aviso
         data_emision.Recordset.MoveFirst
         Do While Not data_emision.Recordset.EOF
            If IsNull(data_emision.Recordset("nro_doc")) = True Then
               data_emision.Recordset.Delete
            Else
               If data_emision.Recordset("nro_doc") = 0 Then
                  data_emision.Recordset.Delete
               End If
            End If
            data_emision.Recordset.MoveNext
         Loop
      Else
         MsgBox "ATENCION!! se detectaron varios registros sin numeración, avise a informática!", vbCritical, "Emisión"
      End If
   End If
   data_emicopi.Recordset.MoveFirst

'   MsgBox "Terminado eliminación de no aceptados."
   Label1.Caption = "Cargando emisión definitiva a la base de datos..."
   Data1.RecordSource = labnomemi.Caption
   Data1.Refresh
   DoEvents
'   MsgBox "Se Comienza a cargar la emisión definitiva."
   Dim Xcontarregsemi As Long
   Xcontarregsemi = 0
   Do While Not data_emicopi.Recordset.EOF
'      Data1.RecordSource = "Select * from emi1016 where cliente =" & data_emicopi.Recordset("cliente") & " and documento =" & data_emicopi.Recordset("documento")
'      Data1.Refresh
'      If Data1.Recordset.RecordCount > 0 Then
'      Else
      Data1.Recordset.AddNew
      Data1.Recordset("cliente") = data_emicopi.Recordset("cliente")
      Data1.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
      Data1.Recordset("nom_cnv") = data_emicopi.Recordset("nom_cnv")
      Data1.Recordset("ruc") = data_emicopi.Recordset("ruc")
      Data1.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
      Data1.Recordset("apellidos") = data_emicopi.Recordset("apellidos")
      Data1.Recordset("cedula") = data_emicopi.Recordset("cedula")
      Data1.Recordset("cod") = data_emicopi.Recordset("cod")
      Data1.Recordset("fecha") = data_emicopi.Recordset("fecha")
      Data1.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
      Data1.Recordset("documento") = data_emicopi.Recordset("documento")
      Data1.Recordset("tipo") = data_emicopi.Recordset("tipo")
      Data1.Recordset("importe") = data_emicopi.Recordset("importe")
      Data1.Recordset("debe_haber") = data_emicopi.Recordset("debe_haber")
      Data1.Recordset("moneda") = data_emicopi.Recordset("moneda")
      Data1.Recordset("origen") = data_emicopi.Recordset("origen")
      Data1.Recordset("operador") = data_emicopi.Recordset("operador")
      Data1.Recordset("hora") = data_emicopi.Recordset("hora")
      Data1.Recordset("dir_cli") = data_emicopi.Recordset("dir_cli")
      Data1.Recordset("loc_cli") = data_emicopi.Recordset("loc_cli")
      Data1.Recordset("tel_cli") = data_emicopi.Recordset("tel_cli")
      Data1.Recordset("nro_superv") = data_emicopi.Recordset("nro_superv")
      Data1.Recordset("nom_superv") = data_emicopi.Recordset("nom_superv")
      Data1.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
      Data1.Recordset("nom_vende") = data_emicopi.Recordset("nom_vende")
      Data1.Recordset("grupo") = data_emicopi.Recordset("grupo")
      Data1.Recordset("numero") = data_emicopi.Recordset("numero")
      Data1.Recordset("zona") = data_emicopi.Recordset("zona")
      Data1.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
      Data1.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
      Data1.Recordset("mes") = data_emicopi.Recordset("mes")
      Data1.Recordset("ano") = data_emicopi.Recordset("ano")
      Data1.Recordset("color_rec") = data_emicopi.Recordset("color_rec")
      Data1.Recordset("fecha_ing") = data_emicopi.Recordset("fecha_ing")
      Data1.Recordset("fecha_nac") = data_emicopi.Recordset("fecha_nac")
      Data1.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
      Data1.Recordset("promo") = data_emicopi.Recordset("promo")
      Data1.Recordset("descimp") = data_emicopi.Recordset("descimp")
      Data1.Recordset("descpor") = data_emicopi.Recordset("descpor")
      Data1.Recordset("deudas") = data_emicopi.Recordset("deudas")
      Data1.Recordset("servi") = data_emicopi.Recordset("servi")
      Data1.Recordset("iva") = data_emicopi.Recordset("iva")
      Data1.Recordset("total") = data_emicopi.Recordset("total")
      Data1.Recordset("deudaap") = data_emicopi.Recordset("deudaap")
      Data1.Recordset.Update
      data_emicopi.Recordset.MoveNext
      Xcontarregsemi = Xcontarregsemi + 1
      If Xcontarregsemi > 1000 Then
         DoEvents
         Xcontarregsemi = 0
      End If
   Loop
   
'   MsgBox "Se cargó la emisión a la tabla definitiva del sistema", vbInformation
   
   data_deu.RecordSource = "deudas"
   data_deu.Refresh
   
   Label1.Caption = "Cargando emisión a la deuda del socio...Aguarde!"
   DoEvents
   Xcontarregsemi = 0
   data_emicopi.Recordset.MoveFirst
   Do While Not data_emicopi.Recordset.EOF
      data_deu.Recordset.AddNew
      data_deu.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
      data_deu.Recordset("nom_cnv") = Mid(data_emicopi.Recordset("nom_cnv"), 1, 20)
      data_deu.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
      data_deu.Recordset("cliente") = data_emicopi.Recordset("cliente")
      data_deu.Recordset("nombre") = data_emicopi.Recordset("apellidos")
      data_deu.Recordset("fecha") = data_emicopi.Recordset("fecha")
      data_deu.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
      data_deu.Recordset("documento") = data_emicopi.Recordset("documento")
      data_deu.Recordset("importe") = data_emicopi.Recordset("importe")
      data_deu.Recordset("moneda") = data_emicopi.Recordset("moneda")
      data_deu.Recordset("origen") = "EMISION..." & Trim(str(data_emicopi.Recordset("mes"))) & "/" & Trim(str(data_emicopi.Recordset("ano")))
      data_deu.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
      data_deu.Recordset("grupo") = data_emicopi.Recordset("grupo")
      data_deu.Recordset("saldo_cc") = 0
      data_deu.Recordset("mes") = data_emicopi.Recordset("mes")
      data_deu.Recordset("ano") = data_emicopi.Recordset("ano")
      data_deu.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
      data_deu.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
      data_deu.Recordset("estado_cta") = 1
      data_deu.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
      data_deu.Recordset("deudas") = data_emicopi.Recordset("deudas")
      data_deu.Recordset("total") = data_emicopi.Recordset("total")
      data_deu.Recordset("servi") = data_emicopi.Recordset("servi")
      data_deu.Recordset("iva") = data_emicopi.Recordset("iva")
      data_deu.Recordset("promo") = data_emicopi.Recordset("promo")
      data_deu.Recordset("descimp") = data_emicopi.Recordset("descimp")
      data_deu.Recordset("descpor") = data_emicopi.Recordset("descpor")
      data_deu.Recordset("deudaap") = data_emicopi.Recordset("deudaap")
      
      data_deu.Recordset("nro_superv") = 50
      data_deu.Recordset.Update
      data_emicopi.Recordset.MoveNext
      Xcontarregsemi = Xcontarregsemi + 1
      If Xcontarregsemi > 1000 Then
         DoEvents
         Xcontarregsemi = 0
      End If
   Loop
'   MsgBox "Se cargó la emisión a la deuda de los socios", vbInformation
   DoEvents
'   data_ctrolmes.Refresh
   data_ctrolmes.Recordset.Edit
   data_ctrolmes.Recordset("salidas") = Val(labmes.Caption)
   data_ctrolmes.Recordset("entradas") = Val(labano.Caption)
   data_ctrolmes.Recordset.Update
   
   Label1.Caption = "Proceso de EMISION TERMINADO!"
   MsgBox "Proceso de emisión terminado!! Puede imprimir el reporte desde la opción INFORMES.", vbInformation


End If
Command2.Enabled = True

'Unload Me



End Sub

Private Sub Command1_Click()
Dim MiBase As Database
Dim UnaSesion As Workspace
Set UnaSesion = Workspaces(0)
Dim Archivo As String
Dim Xfec As Date
Dim Tabla1 As TableDef
Dim Tabla2 As TableDef
Dim Recemi As Recordset
Dim RecCab As Recordset
Dim Xcount, Xcountemi As Long
Dim Xbaseemi As Database
Dim Xsesioemi As Workspace
Dim Xfechasta As Date
Dim Xivanuevo As Double
Dim Xmes As Integer
Dim Xano As Integer
Dim ParaelIva As Double
Dim Cuantaslineas As Integer

Dim MiBaseDatos As Database
Dim Misesion As Workspace
Dim Nombaseemi As String
Dim Regeminew As New ADODB.Recordset
Dim Sqlregnew As String

 Dim Xsindescuento As Integer
 Xsindescuento = 0
Cuantaslineas = 0

Dim Generaranual As Integer
Generaranual = 0

'On Error GoTo Yaesta
ParaelIva = 0

Xmes = Val(labmes.Caption)
Xano = Val(labano.Caption)

Nombaseemi = "db"
If Xmes < 10 Then
   Nombaseemi = Nombaseemi + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
Else
   Nombaseemi = Nombaseemi + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
End If

ConectarBD
ConbdSapp.Open

'Sqlcli = "Select * from t_meta1 order by fecha DESC"
'With Regcli
'     .CursorLocation = adUseClient
'     .CursorType = adOpenKeyset
'     .LockType = adLockOptimistic
'     .Open Sqlcli, ConbdSappM, , , adCmdText
'End With

Set Misesion = Workspaces(0)
Set MiBaseDatos = Misesion.CreateDatabase(App.path & "\" & Nombaseemi, dbLangGeneral) 'Crea la base temporal de la emisiòn

data_emision.DatabaseName = App.path & "\" & Nombaseemi & ".mdb"
'''data_emision.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\" & Nombaseemi & ".mdb"
'data_emision.RecordSource = ""
'data_emision.Refresh
data_emicopi.DatabaseName = App.path & "\" & Nombaseemi & ".mdb"
'''data_emicopi.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\" & Nombaseemi & ".mdb"

''' Set Xsesioemi = Workspaces(0)
Dim Totdescuento As Double
Dim Descustr As String

Dim XRecemisql As Recordset
Xcountemi = 0
 Archivo = App.path & "\" & Nombaseemi & ".mdb"
' Archivo = App.Path & "\emisnueva.mdb"
 Set MiBase = UnaSesion.OpenDatabase(Archivo) 'abre la base temporal local para la emisiòn
 Dim Nomemi, Cabemi As String
 Dim Xfecentexto As String
' If data_ctrolmes.Recordset("salidas") = 12 Then
 Xfec = CDate(labfec.Caption)
 Command1.Enabled = False
 Command2.Enabled = False
 Label1.Visible = True
 If Xmes < 10 Then
    If Day(Xfec) < 10 Then
       Xfecentexto = "0" & Trim(str(Day(Xfec))) & "/0" & Trim(str(Xmes)) & "/" & Trim(str(Xano))
    Else
       Xfecentexto = Trim(str(Day(Xfec))) & "/0" & Trim(str(Xmes)) & "/" & Trim(str(Xano))
    End If
'    Xfechasta = "30/0" & Trim(Str(Xmes)) & "/" & Trim(Str(Xano))
 Else
    If Day(Xfec) < 10 Then
       Xfecentexto = "0" & Trim(str(Day(Xfec))) & "/" & Trim(str(Xmes)) & "/" & Trim(str(Xano))
    Else
       Xfecentexto = Trim(str(Day(Xfec))) & "/" & Trim(str(Xmes)) & "/" & Trim(str(Xano))
    End If
'    Xfechasta = "30/" & Trim(Str(Xmes)) & "/" & Trim(Str(Xano))
 End If
' labfec.Caption = Xfecentexto
Dim Idpromos As Integer
Dim CedPromo As String
Dim CancelaUpdate As Integer
CancelaUpdate = 0

CedPromo = ""
Idpromos = 0

 Dim Xvenctext As String
 If Xmes > 9 Then
    Xvenctext = "20/" & Trim(str(Xmes)) & "/" & Trim(str(Xano))
 Else
    Xvenctext = "20/0" & Trim(str(Xmes)) & "/" & Trim(str(Xano))
 End If
' XFec = CDate(Xfecentexto)
 Xfechasta = CDate(Xvenctext)
 frm_emision.MousePointer = 11
 Nomemi = "emi"
 Cabemi = "cab"
 Command2.Enabled = False
 
 If Xmes < 10 Then
    Nomemi = Nomemi + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
    Cabemi = Cabemi + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
 Else
    Nomemi = Nomemi + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
    Cabemi = Cabemi + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
 End If
 If data_ctrolmes.Recordset("salidas") = Xmes And data_ctrolmes.Recordset("entradas") = Xano Then
    MsgBox "La emisión que intenta generar YA EXISTE, Verifique!!", vbCritical, "Emisión"
    End
 Else
''''''    createTableEmision (Nomemi) 'ESTE NO HAY QUE BORRAR (Función anterior para crear tabla emisión en mysql
    ConbdSapp.Execute "create table " & Nomemi & " (cod_cnv varchar(6) default NULL, nom_cnv varchar(40) default NULL," & _
    " ruc varchar(20) default NULL, tipocta varchar(6) default NULL, cliente int(10) default NULL," & _
    " apellidos varchar(60) default NULL, cedula int(10) default NULL, cod smallint(5) default NULL," & _
    " fecha datetime default NULL, tipodoc varchar(10) default NULL, documento int(10) default NULL," & _
    " tipo varchar(10) default NULL, importe double(15,5) default NULL, debe_haber smallint(5) default NULL," & _
    " moneda smallint(5) default NULL, origen varchar(40) default NULL, operador varchar(15) default NULL," & _
    " hora varchar(5) default NULL, dir_cli varchar(80) default NULL, loc_cli varchar(30) default NULL," & _
    " tel_cli varchar(30) default NULL, nro_superv smallint(5) default NULL, nom_superv varchar(255) default NULL," & _
    " nro_vende smallint(5) default NULL, nom_vende varchar(255) default NULL, fecha_cobr datetime default NULL," & _
    " grupo smallint(5) default NULL, numero smallint(5) default NULL, zona varchar(35) default NULL," & _
    " nro_cobr smallint(5) default NULL, nom_cobr varchar(35) default NULL, mes smallint(5) default NULL," & _
    " ano smallint(5) default NULL, color_rec varchar(1) default NULL, fecha_ing datetime default NULL," & _
    " fecha_nac datetime default NULL, tiquet double(15,5) default NULL, servi double(15,5) default NULL, promo varchar(50) default NULL," & _
    " deudas double(15,5) default NULL, descimp double(15,5) default NULL, descpor double(15,5) default NULL, iva double(15,5) default NULL, total double(15,5) default NULL, deudaap double(15,5) default NULL," & _
    " nro int(10) NOT NULL auto_increment, PRIMARY KEY (nro)) ENGINE=InnoDB DEFAULT CHARSET=latin1;"

'crea el cabezal de la emisión en base local mdb
     Set Tabla1 = MiBase.CreateTableDef(Cabemi)
     Dim serie As Field, nro_doc As Field
     Dim cod_srv As Field, descrip As Field, imp_srv As Field
     Dim nro_linea As Field, Fecha2 As Field, tipo_cod As Field
     Dim indic_fact As Field, cantidad As Field, monto As Field
     Dim cliente2 As Field, mesc As Field, anioc As Field
     Set serie = Tabla1.CreateField("serie", dbText, 3)
     Set nro_doc = Tabla1.CreateField("nro_doc", dbLong, 8)
     Set cod_srv = Tabla1.CreateField("cod_srv", dbText, 35)
     Set descrip = Tabla1.CreateField("descrip", dbText, 80)
     Set imp_srv = Tabla1.CreateField("imp_srv", dbDouble)
     Set nro_linea = Tabla1.CreateField("nro_linea", dbInteger)
     Set Fecha2 = Tabla1.CreateField("fecha2", dbDate)
     Set tipo_cod = Tabla1.CreateField("tipo_cod", dbText, 10)
     Set indic_fact = Tabla1.CreateField("indic_fact", dbInteger)
     Set cantidad = Tabla1.CreateField("cantidad", dbInteger)
     Set monto = Tabla1.CreateField("monto", dbDouble)
     Set cliente2 = Tabla1.CreateField("cliente2", dbLong)
     Set mesc = Tabla1.CreateField("mesc", dbInteger)
     Set anioc = Tabla1.CreateField("anioc", dbInteger)
     
     With Tabla1
       .Fields.Append serie
       .Fields.Append nro_doc
       .Fields.Append cod_srv
       .Fields.Append descrip
       .Fields.Append imp_srv
       .Fields.Append nro_linea
       .Fields.Append Fecha2
       .Fields.Append tipo_cod
       .Fields.Append indic_fact
       .Fields.Append cantidad
       .Fields.Append monto
       .Fields.Append cliente2
       .Fields.Append mesc
       .Fields.Append anioc
     End With
    With MiBase
       .TableDefs.Append Tabla1
    End With
    ConbdSapp.Close
'crea la tabla de emision en la base de datos mdb local
     Set Tabla2 = MiBase.CreateTableDef(Nomemi)
     Dim cod_cnv As Field, nom_cnv As Field, tipocta As Field
     Dim cliente As Field, apellidos As Field, cedula As Field
     Dim cod As Field, fecha As Field, tipodoc As Field
     Dim documento As Field, tipo As Field, importe As Field
     Dim debe_haber As Field, moneda As Field, origen As Field
     Dim operador As Field, hora As Field, dir_cli As Field
     Dim loc_cli As Field, tel_cli As Field, nro_superv As Field
     Dim nom_superv As Field, nro_vende As Field
     Dim nom_vende As Field, fecha_cobr As Field, grupo As Field
     Dim Numero As Field, zona As Field, nro_cobr As Field
     Dim nom_cobr As Field, mes As Field, ano As Field
     Dim color_rec As Field, fecha_ing As Field, ruc, qr As Field
     Dim fecha_nac As Field, tiquet As Field, deudas, iva, servi, total As Field
     Dim Promo As Field, descimp As Field, descpor As Field
     Dim Fvence, Autoriza, RangoCAE, CodSeg As Field, deudaap As Field
     Set cod_cnv = Tabla2.CreateField("cod_cnv", dbText, 6)
     Set nom_cnv = Tabla2.CreateField("nom_cnv", dbText, 40)
     Set ruc = Tabla2.CreateField("ruc", dbText, 20)
     Set tipocta = Tabla2.CreateField("tipocta", dbText, 6) 'serie
     Set cliente = Tabla2.CreateField("cliente", dbLong, 10)
     Set apellidos = Tabla2.CreateField("apellidos", dbText, 60)
     Set cedula = Tabla2.CreateField("cedula", dbLong)
     Set cod = Tabla2.CreateField("cod", dbInteger)
     Set fecha = Tabla2.CreateField("fecha", dbDate)
     Set tipodoc = Tabla2.CreateField("tipodoc", dbText, 10)
     Set documento = Tabla2.CreateField("documento", dbLong)
     Set tipo = Tabla2.CreateField("tipo", dbText, 10)
     Set importe = Tabla2.CreateField("importe", dbDouble)
     Set debe_haber = Tabla2.CreateField("debe_haber", dbInteger) 'tipo de cfe
     Set moneda = Tabla2.CreateField("moneda", dbInteger)
     Set origen = Tabla2.CreateField("origen", dbText, 40)
     Set operador = Tabla2.CreateField("operador", dbText, 15)
     Set hora = Tabla2.CreateField("hora", dbText, 5)
     Set dir_cli = Tabla2.CreateField("dir_cli", dbText, 80)
     Set loc_cli = Tabla2.CreateField("loc_cli", dbText, 30)
     Set tel_cli = Tabla2.CreateField("tel_cli", dbText, 30)
     Set nro_superv = Tabla2.CreateField("nro_superv", dbInteger)
     Set nom_superv = Tabla2.CreateField("nom_superv", dbText)
     Set nro_vende = Tabla2.CreateField("nro_vende", dbInteger)
     Set nom_vende = Tabla2.CreateField("nom_vende", dbText)
     Set fecha_cobr = Tabla2.CreateField("fecha_cobr", dbDate) 'fecha hasta de servicios
     Set grupo = Tabla2.CreateField("grupo", dbInteger)
     Set Numero = Tabla2.CreateField("numero", dbInteger)
     Set zona = Tabla2.CreateField("zona", dbText, 35)
     Set nro_cobr = Tabla2.CreateField("nro_cobr", dbInteger)
     Set nom_cobr = Tabla2.CreateField("nom_cobr", dbText, 35)
     Set mes = Tabla2.CreateField("mes", dbInteger)
     Set ano = Tabla2.CreateField("ano", dbInteger)
     Set color_rec = Tabla2.CreateField("color_rec", dbText, 1)
     Set Promo = Tabla2.CreateField("promo", dbText, 50)
     Set fecha_ing = Tabla2.CreateField("fecha_ing", dbDate)
     Set fecha_nac = Tabla2.CreateField("fecha_nac", dbDate)
     Set tiquet = Tabla2.CreateField("tiquet", dbDouble)
     Set servi = Tabla2.CreateField("servi", dbDouble)
     Set deudas = Tabla2.CreateField("deudas", dbDouble)
     Set descimp = Tabla2.CreateField("descimp", dbDouble)
     Set descpor = Tabla2.CreateField("descpor", dbDouble)
     Set iva = Tabla2.CreateField("iva", dbDouble)
     Set total = Tabla2.CreateField("total", dbDouble)
     Set serie = Tabla2.CreateField("serie", dbText, 3)
     Set qr = Tabla2.CreateField("qr", dbLongBinary)
     Set Fvence = Tabla2.CreateField("Fvence", dbDate)
     Set Autoriza = Tabla2.CreateField("Autoriza", dbText, 50)
     Set RangoCAE = Tabla2.CreateField("RangoCAE", dbText, 50)
     Set CodSeg = Tabla2.CreateField("CodSeg", dbText, 50)
     Set deudaap = Tabla2.CreateField("deudaap", dbDouble)
     
     With Tabla2
     .Fields.Append cod_cnv
     .Fields.Append nom_cnv
     .Fields.Append ruc
     .Fields.Append tipocta
     .Fields.Append cliente
     .Fields.Append apellidos
     .Fields.Append cedula
     .Fields.Append cod
     .Fields.Append fecha
     .Fields.Append tipodoc
     .Fields.Append documento
     .Fields.Append tipo
     .Fields.Append importe
     .Fields.Append debe_haber
     .Fields.Append moneda
     .Fields.Append origen
     .Fields.Append operador
     .Fields.Append hora
     .Fields.Append dir_cli
     .Fields.Append loc_cli
     .Fields.Append tel_cli
     .Fields.Append nro_superv
     .Fields.Append nom_superv
     .Fields.Append nro_vende
     .Fields.Append nom_vende
     .Fields.Append fecha_cobr
     .Fields.Append grupo
     .Fields.Append Numero
     .Fields.Append zona
     .Fields.Append nro_cobr
     .Fields.Append nom_cobr
     .Fields.Append mes
     .Fields.Append ano
     .Fields.Append Promo
     .Fields.Append color_rec
     .Fields.Append fecha_ing
     .Fields.Append fecha_nac
     .Fields.Append tiquet
     .Fields.Append descimp
     .Fields.Append descpor
     .Fields.Append servi
     .Fields.Append deudas
     .Fields.Append iva
     .Fields.Append total
     .Fields.Append qr
     .Fields.Append Fvence
     .Fields.Append Autoriza
     .Fields.Append RangoCAE
     .Fields.Append CodSeg
     .Fields.Append deudaap
     End With
    
    With MiBase
     .TableDefs.Append Tabla2
    End With
    
    Set Recemi = MiBase.OpenRecordset(Nomemi)
    Set RecCab = MiBase.OpenRecordset(Cabemi)
'344,6043782,6043742,92898
    data_cli.RecordSource = "Select * from clientes where cl_nrocobr <>" & 0 & " and cl_apellid <> '" & "" & "' And estado in (1,0) " & _
    "and cl_codconv not in ('CCNOS','UNIVS','CCSP','CCSD','UNIDI','CASH','SMI4','SMIN')"
'    data_cli.RecordSource = "Select * from clientes where cl_codigo in (344,6043782,6043742,92898,10131897)"
    data_cli.Refresh
    data_cli.Recordset.MoveLast
    Xcount = data_cli.Recordset.RecordCount
    data_cli.Recordset.MoveFirst
    ProgressBar1.Max = Xcount
    Xcount = 0
    ProgressBar1.Value = Xcount
'    MsgBox "Terminado selección clientes"
    Do While Not data_cli.Recordset.EOF
       If data_cli.Recordset("cl_nrocobr") = 4 Or data_cli.Recordset("cl_nrocobr") = 14 Or _
          data_cli.Recordset("cl_nrocobr") = 101 Or data_cli.Recordset("cl_nrocobr") = 102 Or _
          data_cli.Recordset("cl_nrocobr") = 110 Or data_cli.Recordset("cl_nrocobr") = 111 Or _
          data_cli.Recordset("cl_nrocobr") = 133 Or data_cli.Recordset("cl_nrocobr") = 144 Or _
          data_cli.Recordset("cl_nrocobr") = 222 Or data_cli.Recordset("cl_nrocobr") = 333 Or _
          data_cli.Recordset("cl_nrocobr") = 511 Or _
          data_cli.Recordset("cl_nrocobr") = 513 Or data_cli.Recordset("cl_nrocobr") = 515 Or _
          data_cli.Recordset("cl_nrocobr") = 516 Or _
          data_cli.Recordset("cl_nrocobr") = 555 Or data_cli.Recordset("cl_nrocobr") = 518 Or data_cli.Recordset("cl_nrocobr") = 15 Then
          data_cli.Recordset.MoveNext
       Else
''''''SEGUIR A PARTIR DE ACÁ PARA LA PROMOCION
            If data_cli.Recordset("fecha_baja") <> "" Then
               data_cli.Recordset.MoveNext
            Else
               If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                  data_cli.Recordset.MoveNext
               Else
                 If data_cli.Recordset("cl_codigo") <> 0 Then
                    If data_cli.Recordset("cl_nrocobr") <> "" Then
                       If data_cli.Recordset("cl_codigo") <> "" Then
'                          data_cnv.Recordset.FindFirst "cnv_codigo = '" & Trim(data_cli.Recordset("cl_codconv")) & "'"
                          data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & Trim(data_cli.Recordset("cl_codconv")) & "'"
                          data_cnv.Refresh
                          If data_cnv.Recordset.RecordCount > 0 Then
                             If data_cnv.Recordset("cnv_emite") = "SI" Then
                                If IsNull(data_cnv.Recordset("cnv_fbaja")) = True Then
                                   If data_cnv.Recordset("cnv_hasta") >= Date Then
                                      If data_cnv.Recordset("cnv_cant_r") = 2 Then
                                         If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
                                            CancelaUpdate = 0
                                            Recemi.AddNew
                                            Recemi("deudaap") = 0
                                            Recemi("cod_cnv") = data_cli.Recordset("cl_codconv")
                                            Recemi("nom_cnv") = Mid(data_cli.Recordset("cl_nomconv"), 1, 40)
                                            Recemi("debe_haber") = 101 'tipo de cfe
                                            If IsNull(data_cnv.Recordset("cnv_ruc")) = False Then
                                               If data_cnv.Recordset("cnv_ruc") <> "" Then
                                                  Recemi("ruc") = Mid(data_cnv.Recordset("cnv_ruc"), 1, 20)
                                                  Recemi("debe_haber") = 111 'tipo de cfe
                                                  If IsNull(data_cnv.Recordset("cnv_entre")) = False Then
                                                     If Trim(data_cnv.Recordset("cnv_entre")) <> "" Then
                                                        Recemi("apellidos") = Mid(data_cnv.Recordset("cnv_entre"), 1, 60)
                                                     Else
                                                        Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                                     End If
                                                  Else
                                                     Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                                  End If
                                               Else
                                                  Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                               End If
                                            Else
                                               Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                            End If
                                            Recemi("tipocta") = "SR" 'Serie
                                            Recemi("cliente") = data_cli.Recordset("cl_codigo")
'                                            Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                            Recemi("cedula") = Int(data_cli.Recordset("cl_cedula"))
                                            Recemi("cod") = data_cli.Recordset("cl_codced")
                                            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                                               CedPromo = Trim(str(data_cli.Recordset("cl_cedula"))) & Trim(str(data_cli.Recordset("cl_codced")))
                                            Else
                                               CedPromo = ""
                                            End If
                                            Recemi("fecha") = Xfec
                                            Recemi("tipodoc") = "UYU" 'moneda
                                            Recemi("documento") = 0
                                            Recemi("tipo") = "EMISION"
                                            If IsNull(data_cnv.Recordset("cnv_precio")) = False Then
                                               If data_cnv.Recordset("cnv_precio") > 0 Then
                                                  Xivanuevo = data_cnv.Recordset("cnv_precio") / 1.1 * 0.1
                                                  Recemi("servi") = data_cnv.Recordset("cnv_precio") - Xivanuevo
                                                  Recemi("servi") = Format(Recemi("servi"), "0.00")
                                                  Recemi("importe") = data_cnv.Recordset("cnv_precio")
                                               Else
                                                  Xivanuevo = 0
                                                  Recemi("importe") = 0
                                                  Recemi("servi") = 0
                                               End If
                                            Else
                                               Xivanuevo = 0
                                               Recemi("importe") = 0
                                               Recemi("servi") = 0
                                            End If
                                            Recemi("moneda") = 2 'fpago crédito
                                            Recemi("origen") = "Cuota " + Trim(str(Xmes)) + "/" + Trim(str(Xano))
                                            Recemi("operador") = data_usua.Recordset("nombre")
                                            Recemi("hora") = Format(Time, "HH:mm")
                                            If IsNull(data_cli.Recordset("cl_dircobr")) = True Then
                                               If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                                                  If data_cli.Recordset("cl_direcci") <> "" Then
                                                     Recemi("dir_cli") = Mid(data_cli.Recordset("cl_direcci"), 1, 70)
                                                  Else
                                                     Recemi("dir_cli") = "S/D"
                                                  End If
                                               Else
                                                  Recemi("dir_cli") = "S/D"
                                               End If
                                            Else
                                               If data_cli.Recordset("cl_dircobr") <> "" Then
                                                  Recemi("dir_cli") = Mid(data_cli.Recordset("cl_dircobr"), 1, 50)
                                               Else
                                                  If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                                                     If data_cli.Recordset("cl_direcci") <> "" Then
                                                        Recemi("dir_cli") = Mid(data_cli.Recordset("cl_direcci"), 1, 70)
                                                     Else
                                                        Recemi("dir_cli") = "S/D"
                                                     End If
                                                  Else
                                                     Recemi("dir_cli") = "S/D"
                                                  End If
                                               End If
                                            End If
                                            If IsNull(data_cli.Recordset("cl_entre")) = False Then
                                               If Len(data_cli.Recordset("cl_entre")) > 0 Then
                                                  Recemi("loc_cli") = Mid(data_cli.Recordset("cl_entre"), 1, 30)
                                               End If
                                            End If
                                            If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                                               If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                                                  If Trim(data_cli.Recordset("cl_telefon")) = "NO APLICA" Then
                                                     If Trim(data_cli.Recordset("cl_dpto")) = "NO APLICA" Then
                                                        Recemi("tel_cli") = "Sin Tel."
                                                     Else
                                                        Recemi("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 10)
                                                     End If
                                                  Else
                                                     If Trim(data_cli.Recordset("cl_dpto")) = "NO APLICA" Then
                                                        Recemi("tel_cli") = Mid(data_cli.Recordset("cl_telefon"), 1, 9)
                                                     Else
                                                        Recemi("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 10) & "/" & Mid(data_cli.Recordset("cl_telefon"), 1, 9)
                                                     End If
                                                  End If
                                               Else
                                                  Recemi("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 18)
                                               End If
                                            Else
                                               If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                                                  If data_cli.Recordset("cl_telefon") <> "" Then
                                                     Recemi("tel_cli") = Mid(data_cli.Recordset("cl_telefon"), 1, 20)
                                                  Else
                                                     Recemi("tel_cli") = "Sin Tel."
                                                  End If
                                               Else
                                                  Recemi("tel_cli") = "Sin Tel."
                                               End If
                                            End If
                                            Recemi("nro_superv") = 1
                                            Recemi("nom_superv") = "SUPERVISOR GENERAL"
                                            Recemi("fecha_cobr") = Xfechasta 'servicio hasta
                                            Recemi("nro_vende") = data_cli.Recordset("cl_nrovend")
                                            If IsNull(data_cli.Recordset("cl_nomvend")) = False Then
                                               Recemi("nom_vende") = data_cli.Recordset("cl_nomvend")
                                            End If
                                            Recemi("grupo") = data_cli.Recordset("cl_grupo")
                                            Recemi("numero") = 0
                                            If IsNull(data_cli.Recordset("cl_zona")) = False Then
                                               If data_cli.Recordset("cl_zona") = "" Then
                                                  Recemi("zona") = "Sin Zona"
                                               Else
                                                  Recemi("zona") = Mid(data_cli.Recordset("cl_zona"), 1, 30)
                                               End If
                                            Else
                                               Recemi("zona") = "Sin Zona"
                                            End If
                                            Recemi("nro_cobr") = data_cli.Recordset("cl_nrocobr")
                                            If IsNull(data_cli.Recordset("cl_nomcobr")) = False Then
                                               If Trim(data_cli.Recordset("cl_nomcobr")) = "" Then
                                                  Recemi("nom_cobr") = "Sin Cob"
                                               Else
                                                  Recemi("nom_cobr") = data_cli.Recordset("cl_nomcobr")
                                               End If
                                            Else
                                               Recemi("nom_cobr") = "Sin Cob"
                                            End If
                                            Recemi("mes") = Xmes
                                            Recemi("ano") = Xano
                                            If IsNull(data_cnv.Recordset("cnv_colrec")) = False Then
                                               If Trim(data_cnv.Recordset("cnv_colrec")) = "" Then
                                                  Recemi("color_rec") = "B"
                                               Else
                                                  Recemi("color_rec") = data_cnv.Recordset("cnv_colrec")
                                               End If
                                            Else
                                               Recemi("color_rec") = "B"
                                            End If
                                            If data_cli.Recordset("cl_fecing") <> "" Then
                                               Recemi("fecha_ing") = Format(data_cli.Recordset("cl_fecing"), "dd/mm/yyyy")
                                            End If
                                            If data_cli.Recordset("cl_fnac") <> "" Then
                                               Recemi("fecha_nac") = Format(data_cli.Recordset("cl_fnac"), "dd/mm/yyyy")
                                            End If
                                             
                                            If IsNull(data_cli.Recordset("idpromos")) = False Then
                                               Idpromos = data_cli.Recordset("idpromos")
                                               If Idpromos > 0 Then
                                                  data_promos.RecordSource = "select * from promocion_gpo where id =" & Idpromos
                                                  data_promos.Refresh
                                                  If data_promos.Recordset.RecordCount > 0 Then
                                                     data_promos.Recordset.MoveFirst
                                                     If data_promos.Recordset("descu_imp") > 0 Then
                                                        Totdescuento = data_promos.Recordset("descu_imp")
                                                     Else
                                                        Descustr = "0." & data_promos.Recordset("descu_por")
                                                        Totdescuento = data_cnv.Recordset("cnv_precio") * CDbl(Descustr)
                                                     End If
                                                  End If
                                               End If
                                            Else
                                               Idpromos = 0
                                               Totdescuento = 0
                                            End If
                                            If Totdescuento > 0 Then
                                               If data_promos.Recordset("descrip") = "Pago anual" Then
                                                  If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                     If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                        If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                           Generaranual = 1
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                               
                                               If data_promos.Recordset("descrip") = "Grupo de 3 o más" Then
                                                  If IsNull(data_cli.Recordset("cl_codruta")) = False Then
                                                     data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & data_cli.Recordset("cl_codruta") & " and estado in (1)"
                                                     data_clipromo.Refresh
                                                     If data_clipromo.Recordset.RecordCount > 0 Then
                                                        data_clipromo.Recordset.MoveLast
                                                        If data_clipromo.Recordset.RecordCount < 2 Then
                                                           Xsindescuento = 1
                                                        Else
                                                           Xsindescuento = 0
                                                        End If
                                                     Else
                                                        Xsindescuento = 1
                                                     End If
                                                  Else
                                                     data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & Val(CedPromo) & " and estado in (1)"
                                                     data_clipromo.Refresh
                                                     If data_clipromo.Recordset.RecordCount > 0 Then
                                                        data_clipromo.Recordset.MoveLast
                                                        If data_clipromo.Recordset.RecordCount < 2 Then
                                                           Xsindescuento = 1
                                                        Else
                                                           Xsindescuento = 0
                                                        End If
                                                     Else
                                                        Xsindescuento = 1
                                                     End If
                                                  End If
                                               End If
                                               If Xsindescuento = 1 Then
                                                  Totdescuento = 0
                                                  Recemi("descimp") = 0
                                                  Recemi("descpor") = 0
                                               Else
                                                  Recemi("descimp") = -Totdescuento
                                                  Recemi("descpor") = data_promos.Recordset("descu_por")
                                                  Recemi("promo") = data_promos.Recordset("descrip")
                                               End If
                                               If data_cnv.Recordset("cnv_precio") > 0 Then
                                                  Recemi("total") = data_cnv.Recordset("cnv_precio") - Totdescuento
                                                  ParaelIva = data_cnv.Recordset("cnv_precio") - Totdescuento
                                               Else
                                                  Recemi("total") = 0
                                                  ParaelIva = 0
                                               End If
                                            Else
                                               Recemi("descimp") = -Totdescuento
                                               Recemi("descpor") = 0
                                               Recemi("total") = data_cnv.Recordset("cnv_precio")
                                               ParaelIva = data_cnv.Recordset("cnv_precio")
                                            End If
                                            If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                               If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                  If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                     If data_cnv.Recordset("cnv_precio") > 0 Then
                                                        Xivanuevo = ParaelIva / 1.1 * 0.1
                                                        Recemi("servi") = ParaelIva - Xivanuevo
                                                        Recemi("servi") = Format(Recemi("servi"), "0.00")
                                                     Else
                                                        Xivanuevo = 0
                                                     End If
                                                     Recemi("tiquet") = 0
                                                     Recemi("deudas") = 0
                                                     Recemi("iva") = Format(Xivanuevo, "0.00")
                                                     Recemi.Update
                                                     If Xmes = 12 Then
                                                        data_cli.Recordset("mesproxemi") = 1
                                                        data_cli.Recordset("anoproxemi") = Xano + 1
                                                     Else
                                                        data_cli.Recordset("mesproxemi") = Xmes + 1
                                                        data_cli.Recordset("anoproxemi") = Xano
                                                     End If
                                                     data_cli.Recordset.Update
                                                  Else
                                                     CancelaUpdate = 9
                                                     Recemi.CancelUpdate
                                                  End If
                                               Else
                                                  If data_cnv.Recordset("cnv_precio") > 0 Then
                                                     Xivanuevo = ParaelIva / 1.1 * 0.1
                                                     Recemi("servi") = ParaelIva - Xivanuevo
                                                     Recemi("servi") = Format(Recemi("servi"), "0.00")
                                                  Else
                                                     Xivanuevo = 0
                                                  End If
                                                  Recemi("tiquet") = 0
                                                  Recemi("deudas") = 0
                                                  Recemi("iva") = Format(Xivanuevo, "0.00")
                                                  Recemi.Update
                                                  If Xmes = 12 Then
                                                     data_cli.Recordset("mesproxemi") = 1
                                                     data_cli.Recordset("anoproxemi") = Xano + 1
                                                  Else
                                                     data_cli.Recordset("mesproxemi") = Xmes + 1
                                                     data_cli.Recordset("anoproxemi") = Xano
                                                  End If
                                                  data_cli.Recordset.Update
                                               End If
                                            Else
                                               If data_cnv.Recordset("cnv_precio") > 0 Then
                                                  Xivanuevo = ParaelIva / 1.1 * 0.1
                                                  Recemi("servi") = ParaelIva - Xivanuevo
                                                  Recemi("servi") = Format(Recemi("servi"), "0.00")
                                               Else
                                                  Xivanuevo = 0
                                               End If
                                               Recemi("tiquet") = 0
                                               Recemi("deudas") = 0
                                               Recemi("iva") = Format(Xivanuevo, "0.00")
                                               Recemi.Update
                                               If Xmes = 12 Then
                                                  data_cli.Recordset("mesproxemi") = 1
                                                  data_cli.Recordset("anoproxemi") = Xano + 1
                                               Else
                                                  data_cli.Recordset("mesproxemi") = Xmes + 1
                                                  data_cli.Recordset("anoproxemi") = Xano
                                               End If
                                               data_cli.Recordset.Update
                                            End If
                                            If CancelaUpdate = 9 Then
                                            Else
                                               RecCab.AddNew
                                               If Totdescuento > 0 Then
                                                  RecCab("serie") = "DS"
                                                  RecCab("nro_doc") = Totdescuento
                                               Else
                                                  RecCab("serie") = "SR"
                                                  RecCab("nro_doc") = 0
                                               End If
                                               RecCab("cod_srv") = "881"
                                               RecCab("descrip") = "CUOTA MENSUAL"
                                               RecCab("imp_srv") = Format(data_cnv.Recordset("cnv_precio"), "Standard")
                                               RecCab("nro_linea") = 1
                                               RecCab("fecha2") = CDate(labfec.Caption)
                                               RecCab("mesc") = Xmes
                                               RecCab("anioc") = Xano
                                               RecCab("tipo_cod") = "INT1"
                                               If data_cnv.Recordset("cnv_precio") > 0 Then
                                                  RecCab("indic_fact") = 2
                                               Else
                                                  RecCab("indic_fact") = 5
                                               End If
                                               RecCab("cantidad") = 1
                                               RecCab("monto") = Format(data_cnv.Recordset("cnv_precio"), "Standard")
                                               RecCab("cliente2") = data_cli.Recordset("cl_codigo")
                                               RecCab.Update
                                               Xcountemi = Xcountemi + 1
                                               If Totdescuento > 0 Then
                                                  RecCab.AddNew
                                                  RecCab("serie") = "SR"
                                                  RecCab("nro_doc") = 0
                                                  RecCab("cod_srv") = "883"
                                                  If IsNull(data_promos.Recordset("descrip")) = False Then
                                                     RecCab("descrip") = "PROMOCION " & data_promos.Recordset("descu_por") & "% " & data_promos.Recordset("descrip")
                                                  Else
                                                     RecCab("descrip") = "PROMOCION "
                                                  End If
                                                  RecCab("imp_srv") = Format(-Totdescuento, "Standard")
                                                  RecCab("nro_linea") = 12
                                                  RecCab("fecha2") = CDate(labfec.Caption)
                                                  RecCab("mesc") = Xmes
                                                  RecCab("anioc") = Xano
                                                  RecCab("tipo_cod") = "INT1"
                                                  RecCab("indic_fact") = 2
                                                  RecCab("cantidad") = 1
                                                  RecCab("monto") = Format(-Totdescuento, "Standard")
                                                  RecCab("cliente2") = data_cli.Recordset("cl_codigo")
                                                  RecCab.Update
                                                  Xcountemi = Xcountemi + 1
                                               End If
                                               If Generaranual = 1 Then
                                                  Generar_anual (data_cli.Recordset("cl_codigo"))
                                                  Generaranual = 0
                                                  Xcountemi = Xcountemi + 11
                                               End If
                                            End If
                                         End If
                                         Xsindescuento = 0
                                         Idpromos = 0
                                         Totdescuento = 0
                                      Else
                                         If data_cli.Recordset("cl_codigo") = data_cnv.Recordset("cnv_cuenta") Then
                                            If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
                                               CancelaUpdate = 0
                                               Recemi.AddNew
                                               Recemi("deudaap") = 0
                                               Recemi("cod_cnv") = data_cli.Recordset("cl_codconv")
                                               Recemi("nom_cnv") = Mid(data_cli.Recordset("cl_nomconv"), 1, 40)
                                               Recemi("debe_haber") = 101 'tipo de cfe
                                               If IsNull(data_cnv.Recordset("cnv_ruc")) = False Then
                                                  If data_cnv.Recordset("cnv_ruc") <> "" Then
                                                     Recemi("ruc") = Mid(data_cnv.Recordset("cnv_ruc"), 1, 20)
                                                     Recemi("debe_haber") = 111 'tipo de cfe
                                                     If IsNull(data_cnv.Recordset("cnv_entre")) = False Then
                                                        If Trim(data_cnv.Recordset("cnv_entre")) <> "" Then
                                                           Recemi("apellidos") = Mid(data_cnv.Recordset("cnv_entre"), 1, 60)
                                                        Else
                                                           Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                                        End If
                                                     Else
                                                        Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                                     End If
                                                  Else
                                                     Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                                  
                                                  End If
                                               Else
                                                  Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                               End If
                                               Recemi("tipocta") = "SR" 'Serie
                                               Recemi("cliente") = data_cli.Recordset("cl_codigo")
'                                               Recemi("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                                               Recemi("cedula") = Int(data_cli.Recordset("cl_cedula"))
                                               Recemi("cod") = data_cli.Recordset("cl_codced")
                                               If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                                                  CedPromo = Trim(str(data_cli.Recordset("cl_cedula"))) & Trim(str(data_cli.Recordset("cl_codced")))
                                               End If
                                               Recemi("fecha") = Xfec
                                               Recemi("tipodoc") = "UYU" ' moneda
                                               Recemi("documento") = 0
                                               Recemi("tipo") = "EMISION"
                                               If IsNull(data_cnv.Recordset("cnv_precio")) = False Then
                                                  If data_cnv.Recordset("cnv_precio") > 0 Then
                                                     Xivanuevo = data_cnv.Recordset("cnv_precio") / 1.1 * 0.1
                                                     Recemi("servi") = data_cnv.Recordset("cnv_precio") - Xivanuevo
                                                     Recemi("servi") = Format(Recemi("servi"), "0.00")
                                                     Recemi("importe") = data_cnv.Recordset("cnv_precio")
                                                  Else
                                                     Xivanuevo = 0
                                                     Recemi("importe") = 0
                                                     Recemi("servi") = 0
                                                  End If
                                               Else
                                                  Xivanuevo = 0
                                                  Recemi("importe") = 0
                                                  Recemi("servi") = 0
                                               End If
                                               Recemi("moneda") = 2 'fpago crédito
                                               Recemi("origen") = "Cuota " + Trim(str(Xmes)) + "/" + Trim(str(Xano))
                                               Recemi("operador") = data_usua.Recordset("nombre")
                                               Recemi("hora") = Format(Time, "HH:mm")
                                               If IsNull(data_cli.Recordset("cl_dircobr")) = True Then
                                                   If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                                                      If data_cli.Recordset("cl_direcci") <> "" Then
                                                         Recemi("dir_cli") = Mid(data_cli.Recordset("cl_direcci"), 1, 70)
                                                      Else
                                                         Recemi("dir_cli") = "S/D"
                                                      End If
                                                   Else
                                                      Recemi("dir_cli") = "S/D"
                                                   End If
                                               Else
                                                   If data_cli.Recordset("cl_dircobr") <> "" Then
                                                      Recemi("dir_cli") = Mid(data_cli.Recordset("cl_dircobr"), 1, 50)
                                                   Else
                                                      If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                                                         If data_cli.Recordset("cl_direcci") <> "" Then
                                                            Recemi("dir_cli") = Mid(data_cli.Recordset("cl_direcci"), 1, 70)
                                                         Else
                                                            Recemi("dir_cli") = "S/D"
                                                         End If
                                                      Else
                                                         Recemi("dir_cli") = "S/D"
                                                      End If
                                                   End If
                                               End If
                                               If IsNull(data_cli.Recordset("cl_entre")) = False Then
                                                   If Len(data_cli.Recordset("cl_entre")) > 0 Then
                                                      Recemi("loc_cli") = Mid(data_cli.Recordset("cl_entre"), 1, 30)
                                                   End If
                                               End If
                                               If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                                                  If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                                                     If Trim(data_cli.Recordset("cl_telefon")) = "NO APLICA" Then
                                                        If Trim(data_cli.Recordset("cl_dpto")) = "NO APLICA" Then
                                                           Recemi("tel_cli") = "Sin Tel."
                                                        Else
                                                           Recemi("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 10)
                                                        End If
                                                     Else
                                                        If Trim(data_cli.Recordset("cl_dpto")) = "NO APLICA" Then
                                                           Recemi("tel_cli") = Mid(data_cli.Recordset("cl_telefon"), 1, 9)
                                                        Else
                                                           Recemi("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 10) & "/" & Mid(data_cli.Recordset("cl_telefon"), 1, 9)
                                                        End If
                                                     End If
                                                  Else
                                                     Recemi("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 18)
                                                  End If
                                               Else
                                                  If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                                                     If data_cli.Recordset("cl_telefon") <> "" Then
                                                        Recemi("tel_cli") = Mid(data_cli.Recordset("cl_telefon"), 1, 20)
                                                     Else
                                                        Recemi("tel_cli") = "Sin Tel."
                                                     End If
                                                  Else
                                                     Recemi("tel_cli") = "Sin Tel."
                                                  End If
                                               End If
                                               Recemi("nro_superv") = 1
                                               Recemi("fecha_cobr") = Xfechasta 'servicio hasta
                                               Recemi("nom_superv") = "SUPERVISOR GENERAL"
                                               Recemi("nro_vende") = data_cli.Recordset("cl_nrovend")
                                               If IsNull(data_cli.Recordset("cl_nomvend")) = False Then
                                                  Recemi("nom_vende") = data_cli.Recordset("cl_nomvend")
                                               End If
                                               Recemi("grupo") = data_cli.Recordset("cl_grupo")
                                               Recemi("numero") = 0
                                               If IsNull(data_cli.Recordset("cl_zona")) = False Then
                                                  If data_cli.Recordset("cl_zona") = "" Then
                                                     Recemi("zona") = "Sin Zona"
                                                  Else
                                                     Recemi("zona") = Mid(data_cli.Recordset("cl_zona"), 1, 30)
                                                  End If
                                               Else
                                                  Recemi("zona") = "Sin Zona"
                                               End If
                                               Recemi("nro_cobr") = data_cli.Recordset("cl_nrocobr")
                                               If IsNull(data_cli.Recordset("cl_nomcobr")) = False Then
                                                  If Trim(data_cli.Recordset("cl_nomcobr")) = "" Then
                                                     Recemi("nom_cobr") = "Sin Cob"
                                                  Else
                                                     Recemi("nom_cobr") = data_cli.Recordset("cl_nomcobr")
                                                  End If
                                               Else
                                                  Recemi("nom_cobr") = "Sin Cob"
                                               End If
                                               Recemi("mes") = Xmes
                                               Recemi("ano") = Xano
                                               If IsNull(data_cnv.Recordset("cnv_colrec")) = False Then
                                                  If Trim(data_cnv.Recordset("cnv_colrec")) = "" Then
                                                     Recemi("color_rec") = "B"
                                                  Else
                                                     Recemi("color_rec") = data_cnv.Recordset("cnv_colrec")
                                                  End If
                                               Else
                                                  Recemi("color_rec") = "B"
                                               End If
                                               If data_cli.Recordset("cl_fecing") <> "" Then
                                                  Recemi("fecha_ing") = Format(data_cli.Recordset("cl_fecing"), "dd/mm/yyyy")
                                               End If
                                               If data_cli.Recordset("cl_fnac") <> "" Then
                                                  Recemi("fecha_nac") = Format(data_cli.Recordset("cl_fnac"), "dd/mm/yyyy")
                                               End If
                                               
                                               If IsNull(data_cli.Recordset("idpromos")) = False Then
                                                  Idpromos = data_cli.Recordset("idpromos")
                                                  If Idpromos > 0 Then
                                                     data_promos.RecordSource = "select * from promocion_gpo where id =" & Idpromos
                                                     data_promos.Refresh
                                                     If data_promos.Recordset.RecordCount > 0 Then
                                                        data_promos.Recordset.MoveFirst
                                                        If data_promos.Recordset("descu_imp") > 0 Then
                                                           Totdescuento = data_promos.Recordset("descu_imp")
                                                        Else
                                                           Descustr = "0." & data_promos.Recordset("descu_por")
                                                           Totdescuento = data_cnv.Recordset("cnv_precio") * CDbl(Descustr)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  Idpromos = 0
                                                  Totdescuento = 0
                                               End If
                                               
                                               If Totdescuento > 0 Then
                                                  If data_promos.Recordset("descrip") = "Pago anual" Then
                                                     If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                        If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                           If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                              Generaranual = 1
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                                  
                                                  If data_promos.Recordset("descrip") = "Grupo de 3 o más" Then
                                                     If IsNull(data_cli.Recordset("cl_codruta")) = False Then
                                                        data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & data_cli.Recordset("cl_codruta") & " and estado in (1)"
                                                        data_clipromo.Refresh
                                                        If data_clipromo.Recordset.RecordCount > 0 Then
                                                           data_clipromo.Recordset.MoveLast
                                                           If data_clipromo.Recordset.RecordCount < 2 Then
                                                              Xsindescuento = 1
                                                           Else
                                                              Xsindescuento = 0
                                                           End If
                                                        Else
                                                           Xsindescuento = 1
                                                        End If
                                                     Else
                                                        data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & Val(CedPromo) & " and estado in (1)"
                                                        data_clipromo.Refresh
                                                        If data_clipromo.Recordset.RecordCount > 0 Then
                                                           data_clipromo.Recordset.MoveLast
                                                           If data_clipromo.Recordset.RecordCount < 2 Then
                                                              Xsindescuento = 1
                                                           Else
                                                              Xsindescuento = 0
                                                           End If
                                                        Else
                                                           Xsindescuento = 1
                                                        End If
                                                     End If
                                                  End If
                                                  If Xsindescuento = 1 Then
                                                     Totdescuento = 0
                                                     Recemi("descimp") = 0
                                                     Recemi("descpor") = 0
                                                  Else
                                                     Recemi("descimp") = -Totdescuento
                                                     Recemi("descpor") = data_promos.Recordset("descu_por")
                                                     Recemi("promo") = data_promos.Recordset("descrip")
                                                  End If
                                                  If data_cnv.Recordset("cnv_precio") > 0 Then
                                                     Recemi("total") = data_cnv.Recordset("cnv_precio") - Totdescuento
                                                     ParaelIva = data_cnv.Recordset("cnv_precio") - Totdescuento
                                                  Else
                                                     Recemi("total") = 0
                                                     ParaelIva = 0
                                                  End If
                                               Else
                                                  Recemi("descimp") = -Totdescuento
                                                  Recemi("descpor") = 0
                                                  Recemi("total") = data_cnv.Recordset("cnv_precio")
                                                  ParaelIva = data_cnv.Recordset("cnv_precio")
                                               End If
                                               If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                  If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                     If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                        If data_cnv.Recordset("cnv_precio") > 0 Then
                                                           Xivanuevo = ParaelIva / 1.1 * 0.1
                                                           Recemi("servi") = ParaelIva - Xivanuevo
                                                           Recemi("servi") = Format(Recemi("servi"), "0.00")
                                                        Else
                                                           Xivanuevo = 0
                                                        End If
                                                        Recemi("tiquet") = 0
                                                        Recemi("deudas") = 0
                                                        Recemi("iva") = Format(Xivanuevo, "0.00")
                                                        Recemi.Update
                                                        If Xmes = 12 Then
                                                           data_cli.Recordset("mesproxemi") = 1
                                                           data_cli.Recordset("anoproxemi") = Xano + 1
                                                        Else
                                                           data_cli.Recordset("mesproxemi") = Xmes + 1
                                                           data_cli.Recordset("anoproxemi") = Xano
                                                        End If
                                                        data_cli.Recordset.Update
                                                     Else
                                                        CancelaUpdate = 9
                                                        Recemi.CancelUpdate
                                                     End If
                                                  Else
                                                     If data_cnv.Recordset("cnv_precio") > 0 Then
                                                        Xivanuevo = ParaelIva / 1.1 * 0.1
                                                        Recemi("servi") = ParaelIva - Xivanuevo
                                                        Recemi("servi") = Format(Recemi("servi"), "0.00")
                                                     Else
                                                        Xivanuevo = 0
                                                     End If
                                                     Recemi("tiquet") = 0
                                                     Recemi("deudas") = 0
                                                     Recemi("iva") = Format(Xivanuevo, "0.00")
                                                     Recemi.Update
                                                     If Xmes = 12 Then
                                                        data_cli.Recordset("mesproxemi") = 1
                                                        data_cli.Recordset("anoproxemi") = Xano + 1
                                                     Else
                                                        data_cli.Recordset("mesproxemi") = Xmes + 1
                                                        data_cli.Recordset("anoproxemi") = Xano
                                                     End If
                                                     data_cli.Recordset.Update
                                                  End If
                                               Else
                                                  If data_cnv.Recordset("cnv_precio") > 0 Then
                                                     Xivanuevo = ParaelIva / 1.1 * 0.1
                                                     Recemi("servi") = ParaelIva - Xivanuevo
                                                     Recemi("servi") = Format(Recemi("servi"), "0.00")
                                                  Else
                                                     Xivanuevo = 0
                                                  End If
                                                  Recemi("tiquet") = 0
                                                  Recemi("deudas") = 0
                                                  Recemi("iva") = Format(Xivanuevo, "0.00")
                                                  Recemi.Update
                                                  If Xmes = 12 Then
                                                     data_cli.Recordset("mesproxemi") = 1
                                                     data_cli.Recordset("anoproxemi") = Xano + 1
                                                  Else
                                                     data_cli.Recordset("mesproxemi") = Xmes + 1
                                                     data_cli.Recordset("anoproxemi") = Xano
                                                  End If
                                                  data_cli.Recordset.Update
                                               End If
                                               If CancelaUpdate = 9 Then
                                               Else
                                                  RecCab.AddNew
                                                  If Totdescuento > 0 Then
                                                     RecCab("serie") = "DS"
                                                     RecCab("nro_doc") = Totdescuento
                                                  Else
                                                     RecCab("serie") = "SR"
                                                     RecCab("nro_doc") = 0
                                                  End If
                                                  RecCab("cod_srv") = "881"
                                                  RecCab("descrip") = "CUOTA MENSUAL"
                                                  RecCab("imp_srv") = Format(data_cnv.Recordset("cnv_precio"), "Standard")
                                                  RecCab("nro_linea") = 1
                                                  RecCab("fecha2") = CDate(labfec.Caption)
                                                  RecCab("mesc") = Xmes
                                                  RecCab("anioc") = Xano
                                                  RecCab("tipo_cod") = "INT1"
                                                  If data_cnv.Recordset("cnv_precio") > 0 Then
                                                     RecCab("indic_fact") = 2
                                                  Else
                                                     RecCab("indic_fact") = 5
                                                  End If
                                                  RecCab("cantidad") = 1
                                                  RecCab("monto") = Format(data_cnv.Recordset("cnv_precio"), "Standard")
                                                  RecCab("cliente2") = data_cli.Recordset("cl_codigo")
                                                  RecCab.Update
                                                  Xcountemi = Xcountemi + 1
                                                  If Totdescuento > 0 Then
                                                     RecCab.AddNew
                                                     RecCab("serie") = "SR"
                                                     RecCab("nro_doc") = 0
                                                     RecCab("cod_srv") = "883"
                                                     If IsNull(data_promos.Recordset("descrip")) = False Then
                                                        RecCab("descrip") = "PROMOCION " & data_promos.Recordset("descu_por") & "% " & data_promos.Recordset("descrip")
                                                     Else
                                                        RecCab("descrip") = "PROMOCION "
                                                     End If
                                                     RecCab("imp_srv") = Format(-Totdescuento, "Standard")
                                                     RecCab("nro_linea") = 12
                                                     RecCab("fecha2") = CDate(labfec.Caption)
                                                     RecCab("mesc") = Xmes
                                                     RecCab("anioc") = Xano
                                                     RecCab("tipo_cod") = "INT1"
                                                     RecCab("indic_fact") = 2
                                                     RecCab("cantidad") = 1
                                                     RecCab("monto") = Format(-Totdescuento, "Standard")
                                                     RecCab("cliente2") = data_cli.Recordset("cl_codigo")
                                                     RecCab.Update
                                                  End If
                                                  If Generaranual = 1 Then
                                                     Generar_anual (data_cli.Recordset("cl_codigo"))
                                                     Xcountemi = Xcountemi + 11
                                                     Generaranual = 0
                                                  End If
                                               End If
                                            End If
                                         End If
                                         Xsindescuento = 0
                                         Idpromos = 0
                                         Totdescuento = 0
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                 End If
                 data_cli.Recordset.MoveNext
               End If
            End If
       End If
       DoEvents
       Xcount = Xcount + 1
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
''Deudas AP
    data_emitiq.RecordSource = "select * from convenio_tiquets where fecha_pago is null"
    data_emitiq.Refresh
    Dim MesyAnio As String
    If data_emitiq.Recordset.RecordCount > 0 Then
       data_emitiq.Recordset.MoveLast
       Xcount = Xcount + data_emitiq.Recordset.RecordCount
       data_emitiq.Recordset.MoveFirst
    Else
       MsgBox "Atención!!!! no hay deudas de servicios AP.", vbInformation, "Emisión"
    '   ProgressBar1.Max = Xcount
    End If
    Dim Xdeu22, Xtotd22, XMatDeuda As Double
    Xdeu22 = 0
    Xtotd22 = 0
    XMatDeuda = 0
    ProgressBar1.Max = ProgressBar1.Max + Xcount
    Do While Not data_emitiq.Recordset.EOF
       data_emicopi.RecordSource = "Select * from " & Nomemi & " where cod_cnv ='" & data_emitiq.Recordset("nom_grupo") & "'"
       data_emicopi.Refresh
       If data_emicopi.Recordset.RecordCount > 0 Then
          XMatDeuda = data_emicopi.Recordset("cliente")
       Else
          XMatDeuda = 0
       End If
       data_emision.RecordSource = "Select * from " & Cabemi & " where cliente2 =" & XMatDeuda & " and nro_linea not in (12)"
       data_emision.Refresh
       If data_emision.Recordset.RecordCount > 0 Then
          data_emision.Recordset.MoveLast
          Cuantaslineas = data_emision.Recordset.RecordCount
          data_emision.RecordSource = "Select * from " & Cabemi & " where cliente2 =" & XMatDeuda & " and cod_srv ='" & "884" & "'"
          data_emision.Refresh
          If data_emision.Recordset.RecordCount > 0 Then
             data_emision.Recordset.Edit
             data_emision.Recordset("cantidad") = data_emision.Recordset("cantidad") + 1
             data_emision.Recordset("imp_srv") = data_emitiq.Recordset("importe")
             data_emision.Recordset("monto") = data_emision.Recordset("monto") + data_emitiq.Recordset("importe")
             data_emision.Recordset.Update
          Else
             data_emision.Recordset.AddNew
             data_emision.Recordset("cantidad") = 1
             data_emision.Recordset("serie") = "SR"
             data_emision.Recordset("nro_doc") = 0
             data_emision.Recordset("cod_srv") = "884"
             data_emision.Recordset("descrip") = "LLAMADO FUERA DE TOPE"
             data_emision.Recordset("imp_srv") = data_emitiq.Recordset("importe")
             data_emision.Recordset("nro_linea") = 2 ' 3 sería si tiene promoción
             data_emision.Recordset("fecha2") = CDate(labfec.Caption)
             data_emision.Recordset("tipo_cod") = "INT1"
             data_emision.Recordset("indic_fact") = 2
             data_emision.Recordset("monto") = data_emitiq.Recordset("importe")
             data_emision.Recordset("cliente2") = XMatDeuda
             data_emision.Recordset("mesc") = Xmes
             data_emision.Recordset("anioc") = Xano
             data_emision.Recordset.Update
          End If
       End If
'       Cuantaslineas = 0
       MesyAnio = Trim(Val(Xmes)) & Trim(Val(Xano))
       data_emitiq.Recordset.Edit
       data_emitiq.Recordset("fecha_pago") = Date
       data_emitiq.Recordset("nro_doc") = Val(MesyAnio)
       data_emitiq.Recordset.Update
       data_emitiq.Recordset.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
        
'deudas
    data_emitiq.RecordSource = "emitiq"
    data_emitiq.Refresh
    
    data_emision.RecordSource = Cabemi
    data_emision.Refresh
    '''data_emision.Recordset.MoveLast
    ProgressBar1.Max = ProgressBar1.Max + Xcountemi
    data_emision.Recordset.MoveFirst
    If data_emitiq.Recordset.RecordCount > 0 Then
       data_emitiq.Recordset.MoveLast
       ProgressBar1.Max = ProgressBar1.Max + data_emitiq.Recordset.RecordCount
       data_emitiq.Recordset.MoveFirst
    Else
       MsgBox "Atención!!!! no hay deudas para emisión, PUEDE CARGARLAS AHORA Y LUEGO PRESIONAR EL BOTON ACEPTAR", vbCritical, "Emisión"
    End If
    Dim Xelimptimbre As Double
    data_estud.RecordSource = "Select * from estudios where codest =" & 995
    data_estud.Refresh
    If data_estud.Recordset.RecordCount > 0 Then
       Xelimptimbre = data_estud.Recordset("cons")
    Else
       Xelimptimbre = 76
    End If
''' Ver acá donde dice 3 si tiene promoción
    Cuantaslineas = 0
    Do While Not data_emitiq.Recordset.EOF
       data_emision.RecordSource = "Select * from " & Cabemi & " where cliente2 =" & data_emitiq.Recordset("mat") & " and nro_linea not in (12) and mesc =" & Xmes & " and anioc =" & Xano
       data_emision.Refresh
       If data_emision.Recordset.RecordCount > 0 Then
          data_emision.Recordset.MoveLast
          Cuantaslineas = data_emision.Recordset.RecordCount
          data_emision.RecordSource = "Select * from " & Cabemi & " where cliente2 =" & data_emitiq.Recordset("mat") & " and cod_srv ='" & "882" & "'"
          data_emision.Refresh
          If data_emision.Recordset.RecordCount > 0 Then
             data_emision.Recordset.Edit
             data_emision.Recordset("cantidad") = data_emision.Recordset("cantidad") + 1
             data_emision.Recordset("imp_srv") = data_emitiq.Recordset("imp")
             data_emision.Recordset("monto") = data_emision.Recordset("monto") + data_emitiq.Recordset("imp")
             data_emision.Recordset.Update
          Else
             data_emision.Recordset.AddNew
             data_emision.Recordset("cantidad") = 1
             data_emision.Recordset("serie") = "SR"
             data_emision.Recordset("nro_doc") = 0
             data_emision.Recordset("cod_srv") = "882"
             data_emision.Recordset("descrip") = "DEUDAS POR SERVICIOS"
             data_emision.Recordset("imp_srv") = data_emitiq.Recordset("imp")
             data_emision.Recordset("nro_linea") = Cuantaslineas + 1 ' 3 sería si tiene promoción
             data_emision.Recordset("fecha2") = CDate(labfec.Caption)
             data_emision.Recordset("tipo_cod") = "INT1"
             data_emision.Recordset("indic_fact") = 1
             data_emision.Recordset("monto") = data_emitiq.Recordset("imp")
             data_emision.Recordset("cliente2") = data_emitiq.Recordset("mat")
             data_emision.Recordset("mesc") = Xmes
             data_emision.Recordset("anioc") = Xano
             data_emision.Recordset.Update
          End If
       End If
       Cuantaslineas = 0
       data_emitiq.Recordset.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
'timbres
    If data_rectiq.Recordset.RecordCount > 0 Then
       data_rectiq.Recordset.MoveLast
       ProgressBar1.Max = ProgressBar1.Max + data_rectiq.Recordset.RecordCount
       data_rectiq.Recordset.MoveFirst
    Else
       MsgBox "Atención!! no hay timbres para emisión, PUEDE CARGARLAS AHORA Y LUEGO PRESIONAR EL BOTON ACEPTAR", vbCritical, "Emisión"
    End If
    Dim Contarlineasemi As Integer
    Contarlineasemi = 0
    Do While Not data_rectiq.Recordset.EOF
       data_emision.RecordSource = "Select * from " & Cabemi & " where cliente2 =" & data_rectiq.Recordset("mat") & " and nro_linea not in (12) and mesc =" & Xmes & " and anioc =" & Xano
       data_emision.Refresh
       If data_emision.Recordset.RecordCount > 0 Then
          data_emision.Recordset.MoveLast
          Contarlineasemi = data_emision.Recordset.RecordCount
          data_emision.RecordSource = "Select * from " & Cabemi & " where cliente2 =" & data_rectiq.Recordset("mat") & " and cod_srv ='" & "995" & "'"
          data_emision.Refresh
          If data_emision.Recordset.RecordCount > 0 Then
             data_emision.Recordset.Edit
             data_emision.Recordset("cantidad") = data_emision.Recordset("cantidad") + 1
'             data_emision.Recordset("imp_srv") = Xelimptimbre
             data_emision.Recordset("monto") = data_emision.Recordset("monto") + Xelimptimbre
             data_emision.Recordset.Update
          Else
             data_emision.Recordset.AddNew
             data_emision.Recordset("cantidad") = 1
             data_emision.Recordset("serie") = "SR"
             data_emision.Recordset("nro_doc") = 0
             data_emision.Recordset("cod_srv") = "995"
             data_emision.Recordset("descrip") = "TIMBRE PROFESIONAL"
             data_emision.Recordset("imp_srv") = Xelimptimbre
             data_emision.Recordset("nro_linea") = Contarlineasemi + 1
             data_emision.Recordset("fecha2") = CDate(labfec.Caption)
             data_emision.Recordset("tipo_cod") = "INT1"
             data_emision.Recordset("indic_fact") = 1
             data_emision.Recordset("monto") = Xelimptimbre
             data_emision.Recordset("cliente2") = data_rectiq.Recordset("mat")
             data_emision.Recordset("mesc") = Xmes
             data_emision.Recordset("anioc") = Xano
             data_emision.Recordset.Update
          End If
          Contarlineasemi = 0
       End If
       data_rectiq.Recordset.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
       
'regularizar totales en cabezal
'    MsgBox "Terminado tickets y timbres"
    data_emicopi.RecordSource = Nomemi
    data_emicopi.Refresh
    Dim Cantlincab As Integer
    Cantlincab = 1
    If data_emicopi.Recordset.RecordCount > 0 Then
       data_emicopi.Recordset.MoveFirst
       Do While Not data_emicopi.Recordset.EOF
          data_emision.RecordSource = "Select * from " & Cabemi & " where cliente2 =" & data_emicopi.Recordset("cliente") & " and cod_srv not in ('881') and mesc =" & data_emicopi.Recordset("mes") & " and anioc =" & data_emicopi.Recordset("ano")
          data_emision.Refresh
          If data_emision.Recordset.RecordCount > 0 Then
             Do While Not data_emision.Recordset.EOF
                If data_emision.Recordset("cod_srv") = "995" Then
                   data_emicopi.Recordset.Edit
                   data_emicopi.Recordset("tiquet") = data_emision.Recordset("monto")
                   data_emicopi.Recordset("total") = data_emicopi.Recordset("total") + data_emision.Recordset("monto")
                   data_emicopi.Recordset.Update
                   Cantlincab = Cantlincab + 1
                End If
                If data_emision.Recordset("cod_srv") = "882" Then
                   data_emicopi.Recordset.Edit
                   data_emicopi.Recordset("deudas") = data_emision.Recordset("monto")
                   data_emicopi.Recordset("total") = data_emicopi.Recordset("total") + data_emision.Recordset("monto")
                   data_emicopi.Recordset.Update
                   Cantlincab = Cantlincab + 1
                End If
                If data_emision.Recordset("cod_srv") = "884" Then
                   data_emicopi.Recordset.Edit
                   data_emicopi.Recordset("deudaap") = data_emision.Recordset("monto")
                   ParaelIva = data_emicopi.Recordset("total") + data_emision.Recordset("monto")
                   data_emicopi.Recordset("total") = data_emicopi.Recordset("total") + data_emision.Recordset("monto")
                   Xivanuevo = ParaelIva / 1.1 * 0.1
                   data_emicopi.Recordset("iva") = Xivanuevo
                   data_emicopi.Recordset("servi") = ParaelIva - Xivanuevo
                   data_emicopi.Recordset.Update
                   Cantlincab = Cantlincab + 1
                End If
                data_emision.Recordset.MoveNext
             Loop
          End If
          data_emicopi.Recordset.Edit
          data_emicopi.Recordset("numero") = Cantlincab
          data_emicopi.Recordset.Update
          Cantlincab = 1
          data_emicopi.Recordset.MoveNext
       Loop
    End If
    labnomemi.Caption = Nomemi
    labcabemi.Caption = Cabemi
'numeracion
    MiBase.Close
    frm_emision.MousePointer = 0
    MsgBox "Proceso de emisión terminado, se comienza el envío de documentos a DGI", vbInformation
'    End
'    b_env_Click
    Command1.Enabled = False
    Command2.Enabled = False
    DoEvents
    b_etck_Click '''no borrar
'    b_fin_Click
    
End If

'Exit Sub
'Yaesta:
'       If Err.Number = 3010 Then
'          MsgBox "Ya existe emisión " + Nomemi & " CLI:" & data_cli.Recordset("cl_codigo"), vbInformation, "Emisión"
'          frm_emision.MousePointer = 0
'          frm_emision.Hide
'       Else
 '         MsgBox "Error al generar ERROR:" + Str(Err.Number) + Err.Description & " CLI:" & data_cli.Recordset("cl_codigo"), vbCritical, "Emisión"
'          frm_emision.Hide
'       End If

End Sub


Private Sub Command3_Click()

data_emicopi.DatabaseName = App.path & "\db1121.mdb"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deu.ConnectionString = "dsn=" & Xconexrmt

   

    data_emicopi.RecordSource = "emi1121"
    data_emicopi.Refresh

   data_emicopi.Recordset.MoveFirst

'   MsgBox "Terminado eliminación de no aceptados."
   Label1.Caption = "Cargando emisión definitiva a la base de datos..."
   Data1.RecordSource = "emi1121"
   Data1.Refresh
   DoEvents
'   MsgBox "Se Comienza a cargar la emisión definitiva."
   Dim Xcontarregsemi As Long
   Xcontarregsemi = 0
   Do While Not data_emicopi.Recordset.EOF
'      Data1.RecordSource = "Select * from emi1016 where cliente =" & data_emicopi.Recordset("cliente") & " and documento =" & data_emicopi.Recordset("documento")
'      Data1.Refresh
'      If Data1.Recordset.RecordCount > 0 Then
'      Else
      Data1.Recordset.AddNew
      Data1.Recordset("cliente") = data_emicopi.Recordset("cliente")
      Data1.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
      Data1.Recordset("nom_cnv") = data_emicopi.Recordset("nom_cnv")
      Data1.Recordset("ruc") = data_emicopi.Recordset("ruc")
      Data1.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
      Data1.Recordset("apellidos") = data_emicopi.Recordset("apellidos")
      Data1.Recordset("cedula") = data_emicopi.Recordset("cedula")
      Data1.Recordset("cod") = data_emicopi.Recordset("cod")
      Data1.Recordset("fecha") = data_emicopi.Recordset("fecha")
      Data1.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
      Data1.Recordset("documento") = data_emicopi.Recordset("documento")
      Data1.Recordset("tipo") = data_emicopi.Recordset("tipo")
      Data1.Recordset("importe") = data_emicopi.Recordset("importe")
      Data1.Recordset("debe_haber") = data_emicopi.Recordset("debe_haber")
      Data1.Recordset("moneda") = data_emicopi.Recordset("moneda")
      Data1.Recordset("origen") = data_emicopi.Recordset("origen")
      Data1.Recordset("operador") = data_emicopi.Recordset("operador")
      Data1.Recordset("hora") = data_emicopi.Recordset("hora")
      Data1.Recordset("dir_cli") = data_emicopi.Recordset("dir_cli")
      Data1.Recordset("loc_cli") = data_emicopi.Recordset("loc_cli")
      Data1.Recordset("tel_cli") = data_emicopi.Recordset("tel_cli")
      Data1.Recordset("nro_superv") = data_emicopi.Recordset("nro_superv")
      Data1.Recordset("nom_superv") = data_emicopi.Recordset("nom_superv")
      Data1.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
      Data1.Recordset("nom_vende") = data_emicopi.Recordset("nom_vende")
      Data1.Recordset("grupo") = data_emicopi.Recordset("grupo")
      Data1.Recordset("numero") = data_emicopi.Recordset("numero")
      Data1.Recordset("zona") = data_emicopi.Recordset("zona")
      Data1.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
      Data1.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
      Data1.Recordset("mes") = data_emicopi.Recordset("mes")
      Data1.Recordset("ano") = data_emicopi.Recordset("ano")
      Data1.Recordset("color_rec") = data_emicopi.Recordset("color_rec")
      Data1.Recordset("fecha_ing") = data_emicopi.Recordset("fecha_ing")
      Data1.Recordset("fecha_nac") = data_emicopi.Recordset("fecha_nac")
      Data1.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
      Data1.Recordset("promo") = data_emicopi.Recordset("promo")
      Data1.Recordset("descimp") = data_emicopi.Recordset("descimp")
      Data1.Recordset("descpor") = data_emicopi.Recordset("descpor")
      Data1.Recordset("deudas") = data_emicopi.Recordset("deudas")
      Data1.Recordset("servi") = data_emicopi.Recordset("servi")
      Data1.Recordset("iva") = data_emicopi.Recordset("iva")
      Data1.Recordset("total") = data_emicopi.Recordset("total")
      Data1.Recordset.Update
      data_emicopi.Recordset.MoveNext
   Loop
   
'   MsgBox "Se cargó la emisión a la tabla definitiva del sistema", vbInformation
   
   data_deu.RecordSource = "deudas"
   data_deu.Refresh
   
   Label1.Caption = "Cargando emisión a la deuda del socio...Aguarde!"
   DoEvents
   Xcontarregsemi = 0
   data_emicopi.Recordset.MoveFirst
   Do While Not data_emicopi.Recordset.EOF
      data_deu.Recordset.AddNew
      data_deu.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
      data_deu.Recordset("nom_cnv") = Mid(data_emicopi.Recordset("nom_cnv"), 1, 20)
      data_deu.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
      data_deu.Recordset("cliente") = data_emicopi.Recordset("cliente")
      data_deu.Recordset("nombre") = data_emicopi.Recordset("apellidos")
      data_deu.Recordset("fecha") = data_emicopi.Recordset("fecha")
      data_deu.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
      data_deu.Recordset("documento") = data_emicopi.Recordset("documento")
      data_deu.Recordset("importe") = data_emicopi.Recordset("importe")
      data_deu.Recordset("moneda") = data_emicopi.Recordset("moneda")
      data_deu.Recordset("origen") = "EMISION..." & Trim(str(data_emicopi.Recordset("mes"))) & "/" & Trim(str(data_emicopi.Recordset("ano")))
      data_deu.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
      data_deu.Recordset("grupo") = data_emicopi.Recordset("grupo")
      data_deu.Recordset("saldo_cc") = 0
      data_deu.Recordset("mes") = data_emicopi.Recordset("mes")
      data_deu.Recordset("ano") = data_emicopi.Recordset("ano")
      data_deu.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
      data_deu.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
      data_deu.Recordset("estado_cta") = 1
      data_deu.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
      data_deu.Recordset("deudas") = data_emicopi.Recordset("deudas")
      data_deu.Recordset("total") = data_emicopi.Recordset("total")
      data_deu.Recordset("servi") = data_emicopi.Recordset("servi")
      data_deu.Recordset("iva") = data_emicopi.Recordset("iva")
      data_deu.Recordset("promo") = data_emicopi.Recordset("promo")
      data_deu.Recordset("descimp") = data_emicopi.Recordset("descimp")
      data_deu.Recordset("descpor") = data_emicopi.Recordset("descpor")
      data_deu.Recordset("nro_superv") = 50
      data_deu.Recordset.Update
      data_emicopi.Recordset.MoveNext
   Loop
   MsgBox "Terminado"
'   MsgBox "Se cargó la emisión a la deuda de los socios", vbInformation



End Sub

Private Sub Command4_Click()
'labnomemi.Caption = "EMI1016"
'labcabemi.Caption = "CAB1016"
'b_etck_Click

End Sub

Private Sub Command5_Click()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'   MsgBox "Comienza a cargar la emi"
   DoEvents
data_emicopi.DatabaseName = App.path & "\emisnueva.mdb"
'data_emicopi.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\emisnueva.mdb"
data_emicopi.RecordSource = "EMI1216"
data_emicopi.Refresh
data_emicopi.Recordset.MoveFirst
   Do While Not data_emicopi.Recordset.EOF
      Data1.RecordSource = "Select * from emi1216 where cliente =" & data_emicopi.Recordset("cliente") & " and documento =" & data_emicopi.Recordset("documento")
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
      Else
        Data1.Recordset.AddNew
        Data1.Recordset("cliente") = data_emicopi.Recordset("cliente")
        Data1.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
        Data1.Recordset("nom_cnv") = data_emicopi.Recordset("nom_cnv")
        Data1.Recordset("ruc") = data_emicopi.Recordset("ruc")
        Data1.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
        Data1.Recordset("apellidos") = data_emicopi.Recordset("apellidos")
        Data1.Recordset("cedula") = data_emicopi.Recordset("cedula")
        Data1.Recordset("cod") = data_emicopi.Recordset("cod")
        Data1.Recordset("fecha") = data_emicopi.Recordset("fecha")
        Data1.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
        Data1.Recordset("documento") = data_emicopi.Recordset("documento")
        Data1.Recordset("tipo") = data_emicopi.Recordset("tipo")
        Data1.Recordset("importe") = data_emicopi.Recordset("importe")
        Data1.Recordset("debe_haber") = data_emicopi.Recordset("debe_haber")
        Data1.Recordset("moneda") = data_emicopi.Recordset("moneda")
        Data1.Recordset("origen") = data_emicopi.Recordset("origen")
        Data1.Recordset("operador") = data_emicopi.Recordset("operador")
        Data1.Recordset("hora") = data_emicopi.Recordset("hora")
        Data1.Recordset("dir_cli") = data_emicopi.Recordset("dir_cli")
        Data1.Recordset("loc_cli") = data_emicopi.Recordset("loc_cli")
        Data1.Recordset("tel_cli") = data_emicopi.Recordset("tel_cli")
        Data1.Recordset("nro_superv") = data_emicopi.Recordset("nro_superv")
        Data1.Recordset("nom_superv") = data_emicopi.Recordset("nom_superv")
        Data1.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
        Data1.Recordset("nom_vende") = data_emicopi.Recordset("nom_vende")
        Data1.Recordset("grupo") = data_emicopi.Recordset("grupo")
        Data1.Recordset("numero") = data_emicopi.Recordset("numero")
        Data1.Recordset("zona") = data_emicopi.Recordset("zona")
        Data1.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
        Data1.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
        Data1.Recordset("mes") = data_emicopi.Recordset("mes")
        Data1.Recordset("ano") = data_emicopi.Recordset("ano")
        Data1.Recordset("color_rec") = data_emicopi.Recordset("color_rec")
        Data1.Recordset("fecha_ing") = data_emicopi.Recordset("fecha_ing")
        Data1.Recordset("fecha_nac") = data_emicopi.Recordset("fecha_nac")
        Data1.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
        Data1.Recordset("deudas") = data_emicopi.Recordset("deudas")
        Data1.Recordset("servi") = data_emicopi.Recordset("servi")
        Data1.Recordset("iva") = data_emicopi.Recordset("iva")
        Data1.Recordset("total") = data_emicopi.Recordset("total")
        Data1.Recordset.Update
      End If
      data_emicopi.Recordset.MoveNext
   Loop
   
'   MsgBox "Se cargó la emisión a la tabla definitiva del sistema", vbInformation
   
   data_deu.RecordSource = "deudas"
   data_deu.Refresh
   
   Label1.Caption = "Cargando emisión a la deuda del socio..."
   DoEvents
   
   data_emicopi.Recordset.MoveFirst
   Do While Not data_emicopi.Recordset.EOF
      data_deu.Recordset.AddNew
      data_deu.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
      data_deu.Recordset("nom_cnv") = Mid(data_emicopi.Recordset("nom_cnv"), 1, 20)
      data_deu.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
      data_deu.Recordset("cliente") = data_emicopi.Recordset("cliente")
      data_deu.Recordset("nombre") = data_emicopi.Recordset("apellidos")
      data_deu.Recordset("fecha") = data_emicopi.Recordset("fecha")
      data_deu.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
      data_deu.Recordset("documento") = data_emicopi.Recordset("documento")
      data_deu.Recordset("importe") = data_emicopi.Recordset("importe")
      data_deu.Recordset("moneda") = data_emicopi.Recordset("moneda")
      data_deu.Recordset("origen") = "EMISION..." & Trim(str(data_emicopi.Recordset("mes"))) & "/" & Trim(str(data_emicopi.Recordset("ano")))
      data_deu.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
      data_deu.Recordset("grupo") = data_emicopi.Recordset("grupo")
      data_deu.Recordset("saldo_cc") = 0
      data_deu.Recordset("mes") = data_emicopi.Recordset("mes")
      data_deu.Recordset("ano") = data_emicopi.Recordset("ano")
      data_deu.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
      data_deu.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
      data_deu.Recordset("estado_cta") = 1
      data_deu.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
      data_deu.Recordset("deudas") = data_emicopi.Recordset("deudas")
      data_deu.Recordset("total") = data_emicopi.Recordset("total")
      data_deu.Recordset("servi") = data_emicopi.Recordset("servi")
      data_deu.Recordset("iva") = data_emicopi.Recordset("iva")
      data_deu.Recordset("nro_superv") = 50
      data_deu.Recordset.Update
      data_emicopi.Recordset.MoveNext
   Loop
'   MsgBox "Se cargó la emisión a la deuda de los socios", vbInformation
   
'   data_ctrolmes.Refresh
   data_ctrolmes.Recordset.Edit
   data_ctrolmes.Recordset("salidas") = 12
   data_ctrolmes.Recordset("entradas") = 2016
   data_ctrolmes.Recordset.Update
   MsgBox "Proceso de emisión terminado, se generan los informes"
   Label1.Caption = "Proceso de EMISION TERMINADO!"


End Sub

Private Sub Form_Load()
Dim Lafecven As Date

data_informe.DatabaseName = App.path & "\informes.mdb"
'data_informe.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_ultrec2.DatabaseName = App.path & "\controles.mdb"
data_ultrec2.RecordSource = "nrosrec"
data_ultrec2.Refresh

data_ultrec.DatabaseName = App.path & "\controles.mdb"
data_ultrec.RecordSource = "ultnro"
data_ultrec.Refresh
'data_cnv.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_cnv.RecordSource = "convenio"
'data_cnv.Refresh
data_ultemisim.DatabaseName = App.path & "\controles.mdb"
data_ultemisim.RecordSource = "ultsim"
data_ultemisim.Refresh

data_promos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_clipromo.Connect = "odbc;dsn=" & Xconexrmt & ";"

'data_par.Connect = "odbc;dsn=sapplocal;"
'data_par.RecordSource = "paramsapp"
'data_par.Refresh

data_eror.DatabaseName = App.path & "\erores.mdb"
data_eror.RecordSource = "erores"
data_eror.Refresh

'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_refinan.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_ultemiemi.DatabaseName = App.path & "\controles.mdb"
data_ultemiemi.RecordSource = "ultemi"
data_ultemiemi.Refresh

data_noemite.DatabaseName = App.path & "\noemite.mdb"
data_noemite.RecordSource = "noemite"
data_noemite.Refresh

'data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_cli.ConnectionString = "dsn=" & Xconexrmt

data_emitiq.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_emitiq.RecordSource = "emitiq"
data_emitiq.Refresh

data_rectiq.DatabaseName = App.path & "\env_tiq.mdb"
data_rectiq.RecordSource = "EMITIQ"
data_rectiq.Refresh
'data_ctrolmes.DatabaseName = App.Path & "\sapp.mdb"

data_ctrolmes.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ctrolmes.RecordSource = "saldos"
data_ctrolmes.Refresh

Dim Xdig, Xrut, Xtot, Xtot2, Xfactor, i As Integer

''' Acá colocar el proceso automático de las refinanciaciones

If data_emitiq.Recordset.RecordCount > 0 Then
   data_emitiq.Recordset.MoveFirst
   Do While Not data_emitiq.Recordset.EOF
      data_emitiq.Recordset.Delete
      data_emitiq.Recordset.MoveNext
   Loop
End If

data_cnv.RecordSource = "Select * from convenio where cnv_emite ='" & "SI" & "' and cnv_ruc is not null"
data_cnv.Refresh
If data_cnv.Recordset.RecordCount > 0 Then
   data_cnv.Recordset.MoveFirst
   Do While Not data_cnv.Recordset.EOF
      i = 0
      If data_cnv.Recordset("cnv_ruc") <> "" Then
        If Len(Trim(data_cnv.Recordset("cnv_ruc"))) = 12 Then
           If IsNumeric(data_cnv.Recordset("cnv_ruc")) Then
              Xdig = Val(Mid(data_cnv.Recordset("cnv_ruc"), 12, 1))
              Xrut = Val(Mid(data_cnv.Recordset("cnv_ruc"), 1, 12))
              Xtot = 0
              Xfactor = 2
              For i = 1 To 11
                  If i = 1 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 4
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 2 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 3
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 3 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 2
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 4 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 9
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 5 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 8
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 6 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 7
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 7 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 6
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 8 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 5
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 9 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 4
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 10 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 3
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 11 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 2
                     Xtot2 = Xtot2 + Xtot
                  End If
              Next
              Xtot = Xtot2 Mod 11
              If Xtot > 0 Then
                 Xtot = 11 - Xtot
              Else
                 Xdig = 0
              End If
              If Xtot = 11 Then
                 Xdig = 0
              Else
                 Xdig = Xtot
              End If
              If Xdig = Val(Mid(data_cnv.Recordset("cnv_ruc"), 12, 1)) Then
              Else
                 MsgBox "El convenio " & data_cnv.Recordset("cnv_codigo") & " tiene un error en el RUT, debe modificar para poder generar", vbCritical
                 If WElusuario = "JFERNAN" Then
                 Else
                    Command1.Enabled = False
                 End If
              End If
           Else
              MsgBox "El convenio " & data_cnv.Recordset("cnv_codigo") & " tiene un error en el RUT, debe modificar para poder generar", vbCritical
              If WElusuario = "JFERNAN" Then
              Else
                 Command1.Enabled = False
              End If
           End If
        Else
           MsgBox "El convenio " & data_cnv.Recordset("cnv_codigo") & " tiene un error en el RUT, debe modificar para poder generar", vbCritical
           If WElusuario = "JFERNAN" Then
           Else
              Command1.Enabled = False
           End If
        End If
      End If
      Xtot2 = 0
      data_cnv.Recordset.MoveNext
   Loop
End If

'data_cnv.RecordSource = "convenio"
'data_cnv.Refresh

data_deu.ConnectionString = "dsn=" & Xconexrmt

Dim XFecc As Date
Dim Xmess, Xanoo As Integer
Dim Xfecentexto2, Xfechasta2 As String

'XFecc = Date + 25
XFecc = Date + 15

Xmess = Month(XFecc)
Xanoo = Year(XFecc)

labmes.Caption = Xmess
labano.Caption = Xanoo

If Month(Date) = Xmess Then
   If Xmess < 10 Then
      If Day(Date) > 9 Then
         Xfecentexto2 = Trim(str(Day(Date))) & "/0" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
         Xfechasta2 = "30/0" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
      Else
         Xfecentexto2 = "0" & Trim(str(Day(Date))) & "/0" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
         Xfechasta2 = "30/0" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
      End If
   Else
      If Day(Date) > 9 Then
         Xfecentexto2 = Trim(str(Day(Date))) & "/" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
         Xfechasta2 = "30/" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
      Else
         Xfecentexto2 = "0" & Trim(str(Day(Date))) & "/" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
         Xfechasta2 = "30/" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
      End If
   End If
Else
   If Xmess < 10 Then
      Xfecentexto2 = "01/0" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
      Xfechasta2 = "30/0" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
   Else
      Xfecentexto2 = "01/" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
      Xfechasta2 = "30/" & Trim(str(Xmess)) & "/" & Trim(str(Xanoo))
   End If
End If

labfec.Caption = Xfecentexto2

If data_ctrolmes.Recordset("salidas") = labmes.Caption And data_ctrolmes.Recordset("entradas") = labano.Caption Then
   MsgBox "El mes que desea generar YA FUE GENERADO, VERIFIQUE!!", vbCritical
   Command1.Enabled = False
End If
'se cargan las refinanciaciones



MsgBox "Cargar refinanciaciones...", vbOKOnly

frm_emision.MousePointer = 11

data_refinan.RecordSource = "SELECT * FROM deudas where origen >='" & "Refinan" & "' and fecha_pago is null and mes_r =" & Xmess & " and anio_r =" & Xanoo
data_refinan.Refresh
If data_refinan.Recordset.RecordCount > 0 Then
   data_refinan.Recordset.MoveFirst
   Do While Not data_refinan.Recordset.EOF
      Lafecven = data_refinan.Recordset("fecha") + data_refinan.Recordset("nro_superv")
      If Val(Month(Lafecven)) = Val(labmes.Caption) Then
         data_emitiq.Recordset.AddNew
         data_emitiq.Recordset("mat") = data_refinan.Recordset("cliente")
         data_emitiq.Recordset("nombre") = data_refinan.Recordset("nombre")
         data_emitiq.Recordset("imp") = data_refinan.Recordset("total")
         data_emitiq.Recordset("fecha") = Date
         data_emitiq.Recordset.Update
      End If
      data_refinan.Recordset.MoveNext
   Loop
   data_emitiq.Refresh
   MsgBox "Se cargaron refinanciaciones correctamente!"
End If
frm_emision.MousePointer = 0

MsgBox "ATENCION!! RECUERDE REALIZAR CONTROL DE RUTAS ANTES DE GENERAR LA EMISION.", vbCritical, "EMISION"

If frmabm.Visible = True Then
   MsgBox "Atención! Está abierta la ficha de socio, cierre la ficha y vuelva a intentar", vbCritical
   Command1.Enabled = False
End If

End Sub

Private Function createTableEmision(tablename As String) As Boolean
    Dim conODBCDirect As DAO.Connection
    Dim rsODBCDirect As DAO.Recordset
    Dim WrkODBC As Workspace
    Dim strConn As String
    strConn = "odbc;dsn=" & Xconexrmt & ";"
    Set WrkODBC = CreateWorkspace("", "root", "sapp1987", dbUseODBC)
    Set conODBCDirect = WrkODBC.OpenConnection("", , , strConn)
    On Error Resume Next
    conODBCDirect.Execute ("call mktbemision('" & tablename & "')")
    If Err <> 0 Then
        createTableEmision = False
    Else
        createTableEmision = True
    End If
    

End Function


Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Function EstaInicializado() As Boolean

    EstaInicializado = False

    If objPosCfe Is Nothing Or Not objPosCfe.Inicializado Then
        MsgBox "Debe inicializar el POS, comunique a Administración"
        Set objPosCfe = Nothing
        Exit Function
    End If

    EstaInicializado = True
End Function

Private Sub DesplegarInfoEstadoCfe(Mensaje As String, ResultadoCfe As ResultadoCfe)

    If ResultadoCfe Is Nothing Then
        MsgBox Mensaje
        Exit Sub
    End If

    If Not ResultadoCfe.OperacionEjecutada Or ResultadoCfe.EstadoCfe Is Nothing Then
        If ResultadoCfe.Mensaje <> vbNullString Then Mensaje = Mensaje & ": " & ResultadoCfe.Mensaje
        MsgBox Mensaje
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.Error Then
        Mensaje = Mensaje & ", ocurrió un error"
        If ResultadoCfe.EstadoCfe.Mensaje <> vbNullString Then _
            Mensaje = Mensaje & ": " & ResultadoCfe.EstadoCfe.Mensaje
        MsgBox Mensaje
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.SerieNumeroCfe Is Nothing Then
        MsgBox "El CFE no trae número de folio, no se puede terminar la factura"
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.DatosCae Is Nothing Then
        MsgBox "El CFE no trae datos del CAE, no se puede terminar la factura"
        Exit Sub
    End If

    If (CInt(ResultadoCfe.EstadoCfe.SerieNumeroCfe.TipoCFE) < 200) Then
        Dim strFile As String
        strFile = App.path & "\qr.bmp"
        Dim objresultado As Resultado
        Set objresultado = objPosCfe.GenerarQr(ResultadoCfe.EstadoCfe.DatosQr, 100, strFile)

        Dim strMensaje As String
        strMensaje = "No se pudo generar el QR"

        If objresultado Is Nothing Then
            MsgBox strMensaje
            Exit Sub
        End If

        If Not objresultado.OperacionExitosa Then
            If objresultado.Mensaje <> vbNullString Then strMensaje = strMensaje & ": " & objresultado.Mensaje
            MsgBox strMensaje
            Exit Sub
        End If

'        imgQr.Picture = LoadPicture(strFile)
    End If
'    MsgBox "ES:" & ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
    If Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 0 Or _
       Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 11 Then
       labserie.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.serie
       labnrofact.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
'       MsgBox "ES:" & ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
'       MsgBox "ES:" & data_emicopi.Recordset("cliente")
       data_emicopi.Recordset.Edit
       labvence.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Vencimiento)
       labautoriza.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Autorizacion)
       labcae.Caption = labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroDesde) & " - " & labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroHasta)
       labcodseg.Caption = CStr(ResultadoCfe.EstadoCfe.CodigoSeguridad)
       If Len(labvence.Caption) = 8 Then
          labvenceok.Caption = Mid(labvence.Caption, 7, 2) & "/" & Mid(labvence.Caption, 5, 2) & "/" & Mid(labvence.Caption, 1, 4)
       Else
          labvenceok.Caption = "31/12/2016"
       End If
       data_emicopi.Recordset("fvence") = CDate(labvenceok.Caption)
       data_emicopi.Recordset("Autoriza") = Val(labautoriza.Caption)
       data_emicopi.Recordset("RangoCAE") = Trim(labcae.Caption)
       data_emicopi.Recordset("CodSeg") = Trim(labcodseg.Caption)
       Picture1.Picture = LoadPicture(App.path & "\qr.bmp")
       data_emicopi.Recordset.Update
    Else
       data_eror.Recordset.AddNew
       data_eror.Recordset("nro") = Val(ResultadoCfe.EstadoCfe.CodigoRespuesta)
       data_eror.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
       data_eror.Recordset("hora") = Format(Time, "HH:mm")
       data_eror.Recordset("obs") = "SOCIO: " & Trim(str(data_emicopi.Recordset("cliente"))) & " CAT:"
       data_eror.Recordset.Update
       MsgBox "Comprobante RECHAZADO, anote y luego verifique. Se continúa con la numeración! MAT: " & data_emicopi.Recordset("cliente"), vbInformation
'       End
    End If
    
'    MsgBox "SON:" & labserie.Caption & " " & labnrofac.Caption
'    MsgBox "Serie: " & ResultadoCfe.EstadoCfe.SerieNumeroCfe.serie & vbNewLine & _
'        "Numero: " & CStr(ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero) & vbNewLine & _
'        "CAE autorización: " & ResultadoCfe.EstadoCfe.DatosCae.Autorizacion & vbNewLine & _
'        "CAE vencimiento: " & ResultadoCfe.EstadoCfe.DatosCae.Vencimiento & vbNewLine & _
'        "CAE numero desde: " & ResultadoCfe.EstadoCfe.DatosCae.NumeroDesde & vbNewLine & _
'        "CAE numero desde: " & ResultadoCfe.EstadoCfe.DatosCae.NumeroHasta & vbNewLine & _
'        "Contenido QR: " & ResultadoCfe.EstadoCfe.DatosQr & vbNewLine & _
'        "Código de seguridad: " & ResultadoCfe.EstadoCfe.CodigoSeguridad & vbNewLine & _
'        "Código de respuesta: " & ResultadoCfe.EstadoCfe.CodigoRespuesta & vbNewLine & _
'        "Fecha de firma: " & ResultadoCfe.EstadoCfe.FechaFirma & vbNewLine & _
'        "GUID: " & ResultadoCfe.EstadoCfe.Guid & vbNewLine & _
'        "Mensaje: " & ResultadoCfe.EstadoCfe.Mensaje & vbNewLine & _
'        "Pendiente de envío: " & CStr(ResultadoCfe.EstadoCfe.PendienteDeEnvio) & vbNewLine

    strUltimoGuid = ResultadoCfe.EstadoCfe.Guid
    Set objUltimaSerieNumero = ResultadoCfe.EstadoCfe.SerieNumeroCfe

    'cmdConsultaXguid.Enabled = True
    'cmdConsultaXnumero.Enabled = True
End Sub


Public Sub Generar_anual(ByVal Xmatricula As Long)

Dim Xind As Integer
Xind = 0
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xmesanual, Xanoanual As Integer
Dim Xfecanual, Xfechastaanual As Date
Dim DescustrAnual As String
Dim TotdescuentoAnual, XivanuevoA, XparaelivaA As Double
Dim Nombaseemi, Nombasecab As String
Dim Nomtabemi, Nomtabcab As String
Dim Xfecaamm As Date

Dim XvenctextAnual As String

Nombaseemi = "db"
Nombasecab = "db"
Nomtabemi = "emi"
Nomtabcab = "cab"
XivanuevoA = 0
DescustrAnual = ""
TotdescuentoAnual = 0
XparaelivaA = 0
Xfecanual = Date + 15
'Xfecanual = Date + 25

Xmesanual = Month(Xfecanual)
Xanoanual = Year(Xfecanual)

Xfecaamm = CDate(labfec.Caption)

If Xmesanual < 10 Then
   Nombaseemi = Nombaseemi & "0" & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
   Nombasecab = Nombasecab & "0" & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
   Nomtabemi = Nomtabemi & "0" & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
   Nomtabcab = Nomtabcab & "0" & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
Else
   Nombaseemi = Nombaseemi & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
   Nombasecab = Nombasecab & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
   Nomtabemi = Nomtabemi & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
   Nomtabcab = Nomtabcab & Trim(str(Xmesanual)) & Mid(Trim(str(Xanoanual)), 3, 2)
End If

If Xmesanual = 12 Then
   Xmesanual = 1
   Xanoanual = Xanoanual + 1
Else
   Xmesanual = Xmesanual + 1
   Xanoanual = Xanoanual
End If

If Xmesanual > 9 Then
   XvenctextAnual = "20/" & Trim(str(Xmesanual)) & "/" & Trim(str(Xanoanual))
Else
   XvenctextAnual = "20/0" & Trim(str(Xmesanual)) & "/" & Trim(str(Xanoanual))
End If
Xfechastaanual = CDate(XvenctextAnual)

data_emianual.DatabaseName = App.path & "\" & Nombaseemi
data_emianual.RecordSource = Nomtabemi
data_emianual.Refresh

data_cabanual.DatabaseName = App.path & "\" & Nombasecab
data_cabanual.RecordSource = Nomtabcab
data_cabanual.Refresh

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from clientes where cl_codigo =" & Xmatricula
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   For Xind = 1 To 11

       data_emianual.Recordset.AddNew
       data_emianual.Recordset("cod_cnv") = data_cli.Recordset("cl_codconv")
       data_emianual.Recordset("nom_cnv") = Mid(data_cli.Recordset("cl_nomconv"), 1, 40)
       data_emianual.Recordset("debe_haber") = 101 'tipo de cfe e-ticket crédito
       If IsNull(data_cnv.Recordset("cnv_ruc")) = False Then
          If data_cnv.Recordset("cnv_ruc") <> "" Then
             data_emianual.Recordset("ruc") = Mid(data_cnv.Recordset("cnv_ruc"), 1, 20)
             data_emianual.Recordset("debe_haber") = 111 'tipo de cfe
             If IsNull(data_cnv.Recordset("cnv_entre")) = False Then
                If Trim(data_cnv.Recordset("cnv_entre")) <> "" Then
                   data_emianual.Recordset("apellidos") = Mid(data_cnv.Recordset("cnv_entre"), 1, 60)
                Else
                   data_emianual.Recordset("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
                End If
             Else
                data_emianual.Recordset("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
             End If
          Else
             data_emianual.Recordset("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
          End If
       Else
          data_emianual.Recordset("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 60)
       End If
       data_emianual.Recordset("tipocta") = "SR" 'Serie
       data_emianual.Recordset("cliente") = data_cli.Recordset("cl_codigo")
       data_emianual.Recordset("cedula") = Int(data_cli.Recordset("cl_cedula"))
       data_emianual.Recordset("cod") = data_cli.Recordset("cl_codced")
       data_emianual.Recordset("fecha") = Xfecaamm
       data_emianual.Recordset("tipodoc") = "UYU" 'moneda
       data_emianual.Recordset("documento") = 0
       data_emianual.Recordset("tipo") = "EMISION"
       If IsNull(data_cnv.Recordset("cnv_precio")) = False Then
          If data_cnv.Recordset("cnv_precio") > 0 Then
             XivanuevoA = data_cnv.Recordset("cnv_precio") / 1.1 * 0.1
             data_emianual.Recordset("servi") = data_cnv.Recordset("cnv_precio") - XivanuevoA
             data_emianual.Recordset("servi") = Format(data_emianual.Recordset("servi"), "0.00")
             data_emianual.Recordset("importe") = data_cnv.Recordset("cnv_precio")
          Else
             XivanuevoA = 0
             data_emianual.Recordset("importe") = 0
             data_emianual.Recordset("servi") = 0
          End If
       Else
          XivanuevoA = 0
          data_emianual.Recordset("importe") = 0
          data_emianual.Recordset("servi") = 0
       End If
       data_emianual.Recordset("moneda") = 2 'fpago crédito
       
       data_emianual.Recordset("origen") = "Cuota " + Trim(str(Xmesanual)) + "/" + Trim(str(Xanoanual))
       data_emianual.Recordset("operador") = WElusuario
       data_emianual.Recordset("hora") = Format(Time, "HH:mm")
       If IsNull(data_cli.Recordset("cl_dircobr")) = True Then
          If IsNull(data_cli.Recordset("cl_direcci")) = False Then
             If data_cli.Recordset("cl_direcci") <> "" Then
                data_emianual.Recordset("dir_cli") = Mid(data_cli.Recordset("cl_direcci"), 1, 70)
             Else
                data_emianual.Recordset("dir_cli") = "S/D"
             End If
          Else
             data_emianual.Recordset("dir_cli") = "S/D"
          End If
       Else
          If data_cli.Recordset("cl_dircobr") <> "" Then
             data_emianual.Recordset("dir_cli") = Mid(data_cli.Recordset("cl_dircobr"), 1, 50)
          Else
             If IsNull(data_cli.Recordset("cl_direcci")) = False Then
                If data_cli.Recordset("cl_direcci") <> "" Then
                   data_emianual.Recordset("dir_cli") = Mid(data_cli.Recordset("cl_direcci"), 1, 70)
                Else
                   data_emianual.Recordset("dir_cli") = "S/D"
                End If
             Else
                data_emianual.Recordset("dir_cli") = "S/D"
             End If
          End If
       End If
       If IsNull(data_cli.Recordset("cl_entre")) = False Then
          If Len(data_cli.Recordset("cl_entre")) > 0 Then
             data_emianual.Recordset("loc_cli") = Mid(data_cli.Recordset("cl_entre"), 1, 30)
          End If
       End If
       If IsNull(data_cli.Recordset("cl_dpto")) = False Then
          If IsNull(data_cli.Recordset("cl_telefon")) = False Then
             If Trim(data_cli.Recordset("cl_telefon")) = "NO APLICA" Then
                If Trim(data_cli.Recordset("cl_dpto")) = "NO APLICA" Then
                   data_emianual.Recordset("tel_cli") = "Sin Tel."
                Else
                   data_emianual.Recordset("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 10)
                End If
             Else
                If Trim(data_cli.Recordset("cl_dpto")) = "NO APLICA" Then
                   data_emianual.Recordset("tel_cli") = Mid(data_cli.Recordset("cl_telefon"), 1, 9)
                Else
                   data_emianual.Recordset("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 10) & "/" & Mid(data_cli.Recordset("cl_telefon"), 1, 9)
                End If
             End If
          Else
             data_emianual.Recordset("tel_cli") = Mid(data_cli.Recordset("cl_dpto"), 1, 18)
          End If
       Else
          If IsNull(data_cli.Recordset("cl_telefon")) = False Then
             If data_cli.Recordset("cl_telefon") <> "" Then
                data_emianual.Recordset("tel_cli") = Mid(data_cli.Recordset("cl_telefon"), 1, 20)
             Else
                data_emianual.Recordset("tel_cli") = "Sin Tel."
             End If
          Else
             data_emianual.Recordset("tel_cli") = "Sin Tel."
          End If
       End If
       data_emianual.Recordset("nro_superv") = 1
       data_emianual.Recordset("nom_superv") = "SUPERVISOR GENERAL"
       data_emianual.Recordset("fecha_cobr") = Xfechastaanual 'servicio hasta
       data_emianual.Recordset("nro_vende") = data_cli.Recordset("cl_nrovend")
       If IsNull(data_cli.Recordset("cl_nomvend")) = False Then
          data_emianual.Recordset("nom_vende") = data_cli.Recordset("cl_nomvend")
       End If
       data_emianual.Recordset("grupo") = data_cli.Recordset("cl_grupo")
       data_emianual.Recordset("numero") = 0
       If IsNull(data_cli.Recordset("cl_zona")) = False Then
          If data_cli.Recordset("cl_zona") = "" Then
             data_emianual.Recordset("zona") = "Sin Zona"
          Else
             data_emianual.Recordset("zona") = Mid(data_cli.Recordset("cl_zona"), 1, 30)
          End If
       Else
          data_emianual.Recordset("zona") = "Sin Zona"
       End If
       data_emianual.Recordset("nro_cobr") = data_cli.Recordset("cl_nrocobr")
       If IsNull(data_cli.Recordset("cl_nomcobr")) = False Then
          If Trim(data_cli.Recordset("cl_nomcobr")) = "" Then
             data_emianual.Recordset("nom_cobr") = "Sin Cob"
          Else
             data_emianual.Recordset("nom_cobr") = data_cli.Recordset("cl_nomcobr")
          End If
       Else
          data_emianual.Recordset("nom_cobr") = "Sin Cob"
       End If
       data_emianual.Recordset("mes") = Xmesanual
       data_emianual.Recordset("ano") = Xanoanual
       If IsNull(data_cnv.Recordset("cnv_colrec")) = False Then
          If Trim(data_cnv.Recordset("cnv_colrec")) = "" Then
             data_emianual.Recordset("color_rec") = "B"
          Else
             data_emianual.Recordset("color_rec") = data_cnv.Recordset("cnv_colrec")
          End If
       Else
          data_emianual.Recordset("color_rec") = "B"
       End If
       If data_cli.Recordset("cl_fecing") <> "" Then
          data_emianual.Recordset("fecha_ing") = Format(data_cli.Recordset("cl_fecing"), "dd/mm/yyyy")
       End If
       If data_cli.Recordset("cl_fnac") <> "" Then
          data_emianual.Recordset("fecha_nac") = Format(data_cli.Recordset("cl_fnac"), "dd/mm/yyyy")
       End If
        
       DescustrAnual = "0." & data_promos.Recordset("descu_por")
       TotdescuentoAnual = data_cnv.Recordset("cnv_precio") * CDbl(DescustrAnual)
       
''''VER cambio de proximo mes de emisión para pagos anual después de generar
'          If data_promos.Recordset("descrip") = "Pago anual" Then
       data_emianual.Recordset("descimp") = -TotdescuentoAnual
       data_emianual.Recordset("descpor") = data_promos.Recordset("descu_por")
       data_emianual.Recordset("promo") = data_promos.Recordset("descrip")
          
       data_emianual.Recordset("total") = data_cnv.Recordset("cnv_precio") - TotdescuentoAnual
       XparaelivaA = data_cnv.Recordset("cnv_precio") - TotdescuentoAnual
       
       XivanuevoA = XparaelivaA / 1.1 * 0.1
       data_emianual.Recordset("servi") = XparaelivaA - XivanuevoA
       data_emianual.Recordset("servi") = Format(data_emianual.Recordset("servi"), "0.00")
       data_emianual.Recordset("tiquet") = 0
       data_emianual.Recordset("deudas") = 0
       data_emianual.Recordset("iva") = Format(XivanuevoA, "0.00")
       data_emianual.Recordset.Update
''aquí
       data_cabanual.Recordset.AddNew
       data_cabanual.Recordset("serie") = "DS"
       data_cabanual.Recordset("nro_doc") = TotdescuentoAnual
       data_cabanual.Recordset("cod_srv") = "881"
       data_cabanual.Recordset("descrip") = "CUOTA MENSUAL"
       data_cabanual.Recordset("imp_srv") = Format(data_cnv.Recordset("cnv_precio"), "Standard")
       data_cabanual.Recordset("nro_linea") = 1
       data_cabanual.Recordset("fecha2") = CDate(labfec.Caption)
       data_cabanual.Recordset("mesc") = Xmesanual
       data_cabanual.Recordset("anioc") = Xanoanual
       data_cabanual.Recordset("tipo_cod") = "INT1"
       data_cabanual.Recordset("indic_fact") = 2
       data_cabanual.Recordset("cantidad") = 1
       data_cabanual.Recordset("monto") = Format(data_cnv.Recordset("cnv_precio"), "Standard")
       data_cabanual.Recordset("cliente2") = data_cli.Recordset("cl_codigo")
       data_cabanual.Recordset.Update
'          Xcountemi = Xcountemi + 1
       data_cabanual.Recordset.AddNew
       data_cabanual.Recordset("serie") = "SR"
       data_cabanual.Recordset("nro_doc") = 0
       data_cabanual.Recordset("cod_srv") = "883"
       data_cabanual.Recordset("descrip") = "PROMOCION " & data_promos.Recordset("descu_por") & "% " & data_promos.Recordset("descrip")
       data_cabanual.Recordset("imp_srv") = Format(-TotdescuentoAnual, "Standard")
       data_cabanual.Recordset("nro_linea") = 12
       data_cabanual.Recordset("fecha2") = CDate(labfec.Caption)
       data_cabanual.Recordset("mesc") = Xmesanual
       data_cabanual.Recordset("anioc") = Xanoanual
       data_cabanual.Recordset("tipo_cod") = "INT1"
       data_cabanual.Recordset("indic_fact") = 2
       data_cabanual.Recordset("cantidad") = 1
       data_cabanual.Recordset("monto") = Format(-TotdescuentoAnual, "Standard")
       data_cabanual.Recordset("cliente2") = data_cli.Recordset("cl_codigo")
       data_cabanual.Recordset.Update
       If Xmesanual = 12 Then
          Xmesanual = 1
          Xanoanual = Xanoanual + 1
       Else
          Xmesanual = Xmesanual + 1
          Xanoanual = Xanoanual
       End If
   
       If Xmesanual > 9 Then
          XvenctextAnual = "20/" & Trim(str(Xmesanual)) & "/" & Trim(str(Xanoanual))
       Else
          XvenctextAnual = "20/0" & Trim(str(Xmesanual)) & "/" & Trim(str(Xanoanual))
       End If
       Xfechastaanual = CDate(XvenctextAnual)
   Next Xind
      
   data_cli.Recordset("mesproxemi") = Xmesanual
   data_cli.Recordset("anoproxemi") = Xanoanual
   data_cli.Recordset.Update

End If

Xrecclii.Close
ConbdSapp.Close

End Sub

