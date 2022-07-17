VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infarq 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Arqueo"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9300
   Icon            =   "frm_infarq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_infdeu 
      Caption         =   "data_infdeu"
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
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_arqloc 
      Caption         =   "data_arqloc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_cobotra 
      Caption         =   "data_cobotra"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   3960
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2040
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7200
      TabIndex        =   20
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6600
      Top             =   5400
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
   Begin VB.TextBox txt_cob 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton b_cance 
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
      Left            =   8280
      Picture         =   "frm_infarq.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton b_acep 
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
      Left            =   240
      Picture         =   "frm_infarq.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Procesar"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Opciones de informes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9015
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FF0000&
         Caption         =   "Sólo Sauce"
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
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   4080
         Width           =   2535
      End
      Begin VB.OptionButton Option15 
         BackColor       =   &H00FF0000&
         Caption         =   "Atrasos actuales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   34
         Top             =   2280
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc data_inflin 
         Height          =   495
         Left            =   5880
         Top             =   2880
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "data_inflin"
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
      Begin VB.Data data_infcli 
         Caption         =   "data_infcli"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data data_infresum 
         Caption         =   "data_infresum"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_infarq 
         Caption         =   "data_infarq"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc data_abm 
         Height          =   330
         Left            =   6000
         Top             =   1440
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
         Caption         =   "data_abm"
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
      Begin MSAdodcLib.Adodc data_ent 
         Height          =   375
         Left            =   600
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "data_ent"
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
      Begin MSAdodcLib.Adodc data_emi 
         Height          =   330
         Left            =   0
         Top             =   3120
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
         Caption         =   "data_emi"
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
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   330
         Left            =   1800
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSAdodcLib.Adodc data_cob 
         Height          =   330
         Left            =   5640
         Top             =   2880
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
         Caption         =   "data_cob"
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
      Begin MSAdodcLib.Adodc data_arq 
         Height          =   375
         Left            =   6240
         Top             =   3000
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
         Caption         =   "data_arq"
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
         Height          =   330
         Left            =   6480
         Top             =   2880
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.OptionButton Option14 
         BackColor       =   &H00FF0000&
         Caption         =   "Notas de crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   495
         Left            =   7680
         TabIndex        =   32
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00FF0000&
         Caption         =   "A cobrar o de Baja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   31
         Top             =   1800
         Width           =   2535
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FF0000&
         Caption         =   "Pagos en Base"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   1200
         Width           =   2535
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe de N.Crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FF0000&
         Caption         =   "Cambios de cobrador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   2880
         Width           =   2535
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FF0000&
         Caption         =   "Talones no pasados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FF0000&
         Caption         =   "Todos los cobradores"
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
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF0000&
         Caption         =   "Solo Fact. Convenios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   3720
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FF0000&
         Caption         =   "Desde Historial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mff 
         Height          =   375
         Left            =   4320
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe detallado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Width           =   2535
      End
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.TextBox txt_ano 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txt_mes 
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
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe Contabilidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   2280
         Width           =   2535
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF0000&
         Caption         =   "Cobranza por Cob"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   1800
         Width           =   2535
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF0000&
         Caption         =   "Inf. Nuevas entregas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe de Cobrados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe Devoluciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe de Bajas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe de Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Emitir Arqueo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "MES/AÑO ARQUEO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label labcob 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3240
      TabIndex        =   15
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COBRADOR:"
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
      Height          =   405
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   120
      Picture         =   "frm_infarq.frx":0F56
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "frm_infarq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
Dim Xm, Xa As Integer
Dim XCat As String
Dim Xcant, XImp, Xtotmut, Xtotcomi, Xtotsapp, Xtotcomsapp, Xtotacomp As Double
Dim Xmiar As String
Dim Xcantem, Xpesosem, Xdifcan, Xdifpes, Xtotcobc, Xtotcobp As Double
Dim Xmesdear As String
Dim Xtotrecbase, Xtotpesosbase As Double
Xtotrecbase = 0
Xtotpesosbase = 0
On Error GoTo Xinfarqer

Xmesdear = "arq"
'If Check3.value = 1 Then
'   data_arq.DatabaseName = ""
'   data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Else
'   data_arq.DatabaseName = ""
'   data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
'End If
If txt_mes.Text > 9 Then
   Xmesdear = Xmesdear & Trim(str(txt_mes.Text))
Else
   Xmesdear = Xmesdear & "0" & Trim(str(txt_mes.Text))
End If
Xmesdear = Xmesdear & Mid(txt_ano.Text, 3, 2)
Xtotsapp = 0
Xtotacomp = 0
Xtotcomsapp = 0
Xcantem = 0
Xpesosem = 0
Xdifcan = 0
Xdifpes = 0
b_acep.Enabled = False
b_cance.Enabled = False
Dim Xeldia As Integer
Dim Xfd, Xfh As String

Xeldia = Day(DateSerial(txt_ano.Text, txt_mes.Text + 1, 0))
If Val(txt_mes.Text) > 9 Then
   Xfd = "01/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
   If Xeldia > 9 Then
      Xfh = Trim(str(Xeldia)) & "/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
   Else
      Xfh = "0" & Trim(str(Xeldia)) & "/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
   End If
Else
   Xfd = "01/" & "0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
   If Xeldia > 9 Then
      Xfh = Trim(str(Xeldia)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
   Else
      Xfh = "0" & Trim(str(Xeldia)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
   End If
End If

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infarq"
MiBaseact.Execute "Delete * from infarqc"
data_infarq.RecordSource = "infarq"
data_infarq.Refresh
data_infresum.RecordSource = "infarqc"
data_infresum.Refresh
If txt_cob.Text <> "" Then
   If Option1.Value = True Then
      frm_infarq.MousePointer = 11
      Text1.Text = "C"
      If Check3.Value = 1 Then
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select * from " & Xmesdear & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
''            data_arq.RecordSource = "Select * from " & Xmesdear & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
            data_arq.RecordSource = "Select * from " & Xmesdear & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "'"
            data_arq.Refresh
         End If
      Else
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
'''            data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
            data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "'"
            data_arq.Refresh
         End If
      End If
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
         Do While Not data_arq.Recordset.EOF
            MiBaseact.Execute "Insert into infarq (importe,arqueo,cat,color,mes,ano,cob,codzon,nomcob,total" & _
            ") values (" & data_arq.Recordset("importe") & ",'" & data_arq.Recordset("arqueo") & "','" & data_arq.Recordset("cat") & "'," & _
            "'" & data_arq.Recordset("color") & "'," & data_arq.Recordset("mes") & "," & _
            data_arq.Recordset("ano") & "," & data_arq.Recordset("cob") & "," & _
            data_arq.Recordset("codzon") & ",'" & data_arq.Recordset("nomcob") & "'," & data_arq.Recordset("total") & ")"

            data_arq.Recordset.MoveNext
         Loop
'         MiBaseact.Execute "Update infarq set color = 'C' where cat not in ('EMERC','EMERJ','EMERF','EMERP','ESPEM','EMERNE','EMERN') and color not in ('R','A','V','F','C')"
         data_infarq.Refresh
         data_infarq.Recordset.MoveFirst
         Do While Not data_infarq.Recordset.EOF
            If data_infarq.Recordset("color") = "R" Or _
               data_infarq.Recordset("color") = "A" Or _
               data_infarq.Recordset("color") = "V" Or _
               data_infarq.Recordset("color") = "F" Or _
               data_infarq.Recordset("cat") = "EMERC" Or _
               data_infarq.Recordset("cat") = "EMERJ" Or _
               data_infarq.Recordset("cat") = "EMERF" Or _
               data_infarq.Recordset("cat") = "EMERG" Or _
               data_infarq.Recordset("cat") = "EMERP" Or _
               data_infarq.Recordset("cat") = "ESPEM" Or _
               data_infarq.Recordset("cat") = "EMERNE" Or _
               data_infarq.Recordset("cat") = "EMERN" Or _
               data_infarq.Recordset("color") = "C" Then
               data_infarq.Recordset.MoveNext
           Else
               data_infarq.Recordset.Edit
               data_infarq.Recordset("color") = "C"
               data_infarq.Recordset.Update
               data_infarq.Recordset.MoveNext
            End If
         Loop
         data_infarq.RecordSource = "Select * from infarq order by ano,mes,color"
         data_infarq.Refresh
         data_infarq.Recordset.MoveFirst
         Xm = data_infarq.Recordset("mes")
         Xa = data_infarq.Recordset("ano")
         XCat = data_infarq.Recordset("color")
         Xcant = 0
         Do While Not data_infarq.Recordset.EOF
            If data_infarq.Recordset("mes") = Xm And data_infarq.Recordset("color") = XCat Then
'               If data_infarq.Recordset("color") = XCat Then
                  Xcant = Xcant + 1
                  XImp = XImp + data_infarq.Recordset("total")
                  Xm = data_infarq.Recordset("mes")
                  Xa = data_infarq.Recordset("ano")
                  XCat = data_infarq.Recordset("color")
                  data_infarq.Recordset.MoveNext
             Else
'                  data_infarq.Recordset.MovePrevious
'                  Xcant = Xcant + 1
'                  XImp = XImp + data_infarq.Recordset("total")
'                  data_infarq.Recordset.MoveNext
                  data_infresum.Recordset.AddNew
                  data_infresum.Recordset("cob") = data_infarq.Recordset("cob")
                  data_infresum.Recordset("nomcob") = data_infarq.Recordset("nomcob")
                  data_infresum.Recordset("mes") = Xm
                  data_infresum.Recordset("ano") = Xa
                  data_infresum.Recordset("color") = XCat
                  data_infresum.Recordset("totimp") = XImp
                  data_infresum.Recordset("totrec") = Xcant
                  data_infresum.Recordset.Update
                  Xtotsapp = Xtotsapp + XImp
                  XCat = data_infarq.Recordset("color")
                  Xcant = 0
                  XImp = 0
                  Xm = data_infarq.Recordset("mes")
                  Xa = data_infarq.Recordset("ano")
              End If
         Loop
         data_infarq.Recordset.MovePrevious
         data_infresum.Recordset.AddNew
         data_infresum.Recordset("cob") = data_infarq.Recordset("cob")
         data_infresum.Recordset("nomcob") = data_infarq.Recordset("nomcob")
         data_infresum.Recordset("mes") = Xm
         data_infresum.Recordset("ano") = Xa
         data_infresum.Recordset("color") = XCat
         data_infresum.Recordset("totimp") = XImp
         data_infresum.Recordset("totrec") = Xcant
         data_infresum.Recordset.Update
         Xtotsapp = Xtotsapp + XImp
         XCat = data_infarq.Recordset("color")
         Xcant = 0
         XImp = 0
         
'         data_ent.Recordset.FindFirst "cobrador =" & txt_cob.Text
         data_ent.RecordSource = "Select * from entregas where cobrador =" & txt_cob.Text
         data_ent.Refresh
         If data_ent.Recordset.RecordCount > 0 Then
            If data_infresum.Recordset.RecordCount > 0 Then
               data_infresum.Recordset.MoveFirst
               Do While Not data_infresum.Recordset.EOF
                  data_infresum.Recordset.Edit
                  data_infresum.Recordset("entrega") = data_ent.Recordset("pesos")
                  If txt_cob.Text <> "" Then
                     data_cobotra.RecordSource = "Select * from cobrador where cb_numero =" & txt_cob.Text
                     data_cobotra.Refresh
                     If data_cobotra.Recordset.RecordCount > 0 Then
                        If IsNull(data_cobotra.Recordset("cb_recpes")) = False Then
                           data_infresum.Recordset("comcima") = data_cobotra.Recordset("cb_recpes")
                        Else
                           data_infresum.Recordset("comcima") = 0
                        End If
                     End If
                  End If
                  data_infresum.Recordset.Update
                  data_infresum.Recordset.MoveNext
               Loop
            End If
         Else
            MsgBox "No se encontraron entregas", vbInformation, "Mensaje"
         End If
         
         If Dir("C:\mutuales\mutual.dbf") <> "" Then
            FileCopy "C:\mutuales\mutual.dbf", App.path & "\mutual.dbf"
            Data1.DatabaseName = App.path
            Data1.RecordSource = "mutual"
            Data1.RecordsetType = 1
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.FindFirst "cob =" & txt_cob.Text
               If Not Data1.Recordset.NoMatch Then
                  If data_infresum.Recordset.RecordCount > 0 Then
                     data_infresum.Recordset.MoveFirst
                     Do While Not data_infresum.Recordset.EOF
                        data_infresum.Recordset.Edit
                        data_infresum.Recordset("totuniv") = Data1.Recordset("pesuni")
                        data_infresum.Recordset("comuni") = Data1.Recordset("pesuni") / 1.1 * 0.0315
                        Xtotmut = Data1.Recordset("pesuni")
                        data_infresum.Recordset("totccou") = Data1.Recordset("pescco")
                        data_infresum.Recordset("comccou") = Data1.Recordset("pescco") / 1.1 * 0.0315
                        Xtotmut = Xtotmut + Data1.Recordset("pescco")
                        data_infresum.Recordset("totimpasa") = Data1.Recordset("pesimp")
                        data_infresum.Recordset("comimp") = Data1.Recordset("pesimp") / 1.1 * 0.0315
                        Xtotmut = Xtotmut + Data1.Recordset("pesimp")
                        data_infresum.Recordset("totcgali") = Data1.Recordset("pesgal")
                        data_infresum.Recordset("comgal") = Data1.Recordset("pesgal") / 1.1 * 0.0315
                        Xtotmut = Xtotmut + Data1.Recordset("pesgal")
                        data_infresum.Recordset("totevang") = Data1.Recordset("peseva")
                        data_infresum.Recordset("comeva") = Data1.Recordset("peseva") / 1.1 * 0.0315
                        Xtotmut = Xtotmut + Data1.Recordset("peseva")
                        data_infresum.Recordset("totsmi") = Data1.Recordset("pessmi")
                        data_infresum.Recordset("comsmi") = Data1.Recordset("pessmi") / 1.1 * 0.0315
                        Xtotmut = Xtotmut + Data1.Recordset("pessmi")
                        data_infresum.Recordset("totacom") = Data1.Recordset("pesaco")
                        data_infresum.Recordset("comacom") = Data1.Recordset("pesaco") / 1.1 * 0.07
                        Xtotmut = Xtotmut
                        data_infresum.Recordset("totmut") = Xtotmut
                        data_infresum.Recordset.Update
                        data_infresum.Recordset.MoveNext
                     Loop
                     data_infresum.Recordset.MoveFirst
                     Xtotsapp = Xtotsapp + data_infresum.Recordset("totacom")
                     Xtotacomp = Xtotacomp + data_infresum.Recordset("comacom")
                     Do While Not data_infresum.Recordset.EOF
                        data_infresum.Recordset.Edit
                        Xtotcomi = data_infresum.Recordset("comuni")
                        Xtotcomi = Xtotcomi + data_infresum.Recordset("comccou")
                        Xtotcomi = Xtotcomi + data_infresum.Recordset("comimp")
                        Xtotcomi = Xtotcomi + data_infresum.Recordset("comgal")
                        Xtotcomi = Xtotcomi + data_infresum.Recordset("comeva")
                        Xtotcomi = Xtotcomi + data_infresum.Recordset("comsmi")
                        data_infresum.Recordset("comodo") = Xtotcomi
                        If IsNull(data_infresum.Recordset("comcima")) = False Then
                           Xtotcomi = Xtotcomi + data_infresum.Recordset("comcima")
                        End If
                        data_infresum.Recordset.Update
                        data_infresum.Recordset.MoveNext
                     Loop
                     data_infresum.Recordset.MoveFirst
                     Do While Not data_infresum.Recordset.EOF
                        If txt_cob.Text <> 694 Then
                           data_infresum.Recordset.Edit
                           data_infresum.Recordset("coma") = data_infresum.Recordset("totimp") / 1.1 * 0.07
                           data_infresum.Recordset.Update
                           data_infresum.Recordset.MoveNext
                        Else
                           If data_infresum.Recordset("color") = "C" Then
                              data_infresum.Recordset.Edit
                              data_infresum.Recordset("coma") = data_infresum.Recordset("totimp") * 0.0315
                              data_infresum.Recordset.Update
                              data_infresum.Recordset.MoveNext
                           Else
                              If data_infresum.Recordset("color") = "A" Then
                                 data_infresum.Recordset.Edit
                                 data_infresum.Recordset("coma") = data_infresum.Recordset("totimp") * 0.07
                                 data_infresum.Recordset.Update
                                 data_infresum.Recordset.MoveNext
                              Else
                                 data_infresum.Recordset.Edit
                                 data_infresum.Recordset("coma") = data_infresum.Recordset("totimp") * 0.07
                                 data_infresum.Recordset.Update
                                 data_infresum.Recordset.MoveNext
                              End If
                           End If
                        End If
                     Loop
                     data_infresum.Recordset.MoveFirst
                     Xtotcomi = 0
                     Xcant = 0
                     Do While Not data_infresum.Recordset.EOF
                        Xtotcomi = Xtotcomi + data_infresum.Recordset("coma")
                        Xtotcomsapp = Xtotcomsapp + Xtotcomi
                        Xcant = Xcant + data_infresum.Recordset("totimp")
                        data_infresum.Recordset.MoveNext
                     Loop
                     data_infresum.Recordset.MoveFirst
                     Xtotcomi = data_infresum.Recordset("comodo") + data_infresum.Recordset("comcima")
                     Xtotacomp = Xtotsapp / 1.1 * 0.07
                     Xtotcomi = Xtotcomi + Xtotacomp 'con esto hago el total de comisión sapp
                     Xcant = Xcant + data_infresum.Recordset("totmut")
                     Xtotmut = 0
                     XImp = Xtotcomi / 1.1
                     Xtotmut = XImp * 0.1
                     XImp = Xtotcomi - Xtotmut
                     Xtotacomp = data_infresum.Recordset("totmut")
                     Do While Not data_infresum.Recordset.EOF
                        data_infresum.Recordset.Edit
                        '''data_infresum.Recordset("sd")
                        data_infresum.Recordset("ivatot") = Xtotmut
                        data_infresum.Recordset("totodon") = Xtotcomi
                        data_infresum.Recordset("totimpu") = XImp
                        data_infresum.Recordset("comm") = Xcant
                        data_infresum.Recordset("iva2") = data_infresum.Recordset("entrega") - Xtotsapp - Xtotacomp
                        data_infresum.Recordset("totcudam") = Xtotsapp / 1.1 * 0.07
'                        data_infresum.Recordset("totcudam") = data_infresum.Recordset("totcudam") + Xtotacomp
                        data_infresum.Recordset("totimpasa") = Xtotsapp
                        data_infresum.Recordset("totcima") = Xtotsapp + Xtotacomp 'es el total de mutuales
                        data_infresum.Recordset.Update
                        data_infresum.Recordset.MoveNext
                     Loop
                  End If
               End If
               Data1.DatabaseName = ""
               Data1.RecordSource = ""
               frm_infarq.MousePointer = 0
               b_acep.Enabled = True
               Data1.Refresh
               data_infresum.RecordSource = "Select * from infarqc order by ano,mes,color"
               data_infresum.Refresh
               CrystalReport1.ReportFileName = App.path & "\infarq.rpt"
               CrystalReport1.ReportTitle = "INFORME DE ARQUEO COBRADOR: " & txt_cob.Text & " " & labcob.Caption & " MES...:" & txt_mes.Text & "/" & txt_ano.Text
               CrystalReport1.Action = 1
            End If
         Else
            MsgBox "No se encontró el archivo MUTUAL", vbInformation, "Mensaje"
            
         End If
      End If
      frm_infarq.MousePointer = 0
   End If
         
   If Option2.Value = True Then
      Text1.Text = "P"
      frm_infarq.MousePointer = 11
      If txt_cob.Text = 0 Then
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         End If
      Else
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         End If
      End If
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
         Do While Not data_arq.Recordset.EOF
            data_infarq.Recordset.AddNew
            data_infarq.Recordset("matricula") = data_arq.Recordset("matricula")
            data_infarq.Recordset("nrorec") = data_arq.Recordset("nrorec")
            data_infarq.Recordset("color") = data_arq.Recordset("color")
            data_infarq.Recordset("mes") = data_arq.Recordset("mes")
            data_infarq.Recordset("ano") = data_arq.Recordset("ano")
            data_infarq.Recordset("cob") = data_arq.Recordset("cob")
            data_infarq.Recordset("nombre") = Mid(data_arq.Recordset("nombre"), 1, 30)
            If IsNull(data_arq.Recordset("cob")) = False Then
               If txt_cob.Text = 0 Then
'                  data_cob.Recordset.FindFirst "cb_numero =" & data_arq.Recordset("cob")
                  data_cob.RecordSource = "Select * from cobrador where cb_numero =" & data_arq.Recordset("cob")
                  data_cob.Refresh
                  If data_cob.Recordset.RecordCount > 0 Then
                     data_infarq.Recordset("nomcob") = Mid(data_cob.Recordset("cb_nombre"), 1, 35)
                  Else
                     data_infarq.Recordset("nomcob") = data_arq.Recordset("nomcob")
                  End If
               Else
                  data_infarq.Recordset("nomcob") = Mid(labcob.Caption, 1, 35)
               End If
            Else
               data_infarq.Recordset("nomcob") = ""
            End If
            data_infarq.Recordset("total") = data_arq.Recordset("total")
            data_infarq.Recordset.Update
            data_arq.Recordset.MoveNext
         Loop
      Else
         frm_infarq.MousePointer = 0
         MsgBox "No hay registros"
      End If
      frm_infarq.MousePointer = 0
      b_acep.Enabled = True
      data_infarq.RecordSource = "Select * from infarq order by ano,mes"
      data_infarq.Refresh
      
      If Check2.Value = 1 Then
         CrystalReport1.ReportFileName = App.path & "\infpendd.rpt"
      Else
         CrystalReport1.ReportFileName = App.path & "\infpend.rpt"
      End If
      CrystalReport1.ReportTitle = "INFORME DE PENDIENTES COBRADOR: " & txt_cob.Text & " " & labcob.Caption & " MES...:" & txt_mes.Text & "/" & txt_ano.Text
      CrystalReport1.Action = 1
   End If
   If Option3.Value = True Then
      Text1.Text = "B"
      frm_infarq.MousePointer = 11
      If txt_cob.Text = 0 Then
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         End If
      Else
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         End If
      End If
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
         Do While Not data_arq.Recordset.EOF
            data_infarq.Recordset.AddNew
            data_infarq.Recordset("matricula") = data_arq.Recordset("matricula")
            data_infarq.Recordset("nrorec") = data_arq.Recordset("nrorec")
            data_infarq.Recordset("color") = data_arq.Recordset("color")
            data_infarq.Recordset("mes") = data_arq.Recordset("mes")
            data_infarq.Recordset("ano") = data_arq.Recordset("ano")
            data_infarq.Recordset("nombre") = Mid(data_arq.Recordset("nombre"), 1, 30)
            data_infarq.Recordset("cob") = data_arq.Recordset("cob")
            If IsNull(data_arq.Recordset("cob")) = False Then
               If txt_cob.Text = 0 Then
'                  data_cob.Recordset.FindFirst "cb_numero =" & data_arq.Recordset("cob")
                  data_cob.RecordSource = "select * from cobrador where cb_numero =" & data_arq.Recordset("cob")
                  data_cob.Refresh
                  If data_cob.Recordset.RecordCount > 0 Then
                     data_infarq.Recordset("nomcob") = Mid(data_cob.Recordset("cb_nombre"), 1, 35)
                  Else
                     data_infarq.Recordset("nomcob") = data_arq.Recordset("nomcob")
                  End If
               Else
                  data_infarq.Recordset("nomcob") = Mid(labcob.Caption, 1, 35)
               End If
            Else
               data_infarq.Recordset("nomcob") = ""
            End If
            data_infarq.Recordset("total") = data_arq.Recordset("total")
            data_infarq.Recordset.Update
            data_arq.Recordset.MoveNext
         Loop
      Else
         frm_infarq.MousePointer = 0
         MsgBox "No hay registros"
      End If
      frm_infarq.MousePointer = 0
      b_acep.Enabled = True
      data_infarq.RecordSource = "Select * from infarq order by ano,mes"
      data_infarq.Refresh
      If Check2.Value = 1 Then
         CrystalReport1.ReportFileName = App.path & "\infpendd.rpt"
      Else
         CrystalReport1.ReportFileName = App.path & "\infpend.rpt"
      End If
      CrystalReport1.ReportTitle = "INFORME DE BAJAS COBRADOR: " & txt_cob.Text & " " & labcob.Caption & " MES...:" & txt_mes.Text & "/" & txt_ano.Text
      CrystalReport1.Action = 1
   End If
   If Option4.Value = True Then
      Text1.Text = "D"
      frm_infarq.MousePointer = 11
      If txt_cob.Text = 0 Then
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         End If
      Else
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         End If
      End If
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
         Do While Not data_arq.Recordset.EOF
            data_infarq.Recordset.AddNew
            data_infarq.Recordset("matricula") = data_arq.Recordset("matricula")
            data_infarq.Recordset("nrorec") = data_arq.Recordset("nrorec")
            data_infarq.Recordset("color") = data_arq.Recordset("color")
            data_infarq.Recordset("mes") = data_arq.Recordset("mes")
            data_infarq.Recordset("ano") = data_arq.Recordset("ano")
            data_infarq.Recordset("nombre") = Mid(data_arq.Recordset("nombre"), 1, 30)
            data_infarq.Recordset("cob") = data_arq.Recordset("cob")
            If IsNull(data_arq.Recordset("cob")) = False Then
               If txt_cob.Text = 0 Then
'                  data_cob.Recordset.FindFirst "cb_numero =" & data_arq.Recordset("cob")
                  data_cob.RecordSource = "Select * from cobrador where cb_numero =" & data_arq.Recordset("cob")
                  data_cob.Refresh
                  If data_cob.Recordset.RecordCount > 0 Then
                     data_infarq.Recordset("nomcob") = Mid(data_cob.Recordset("cb_nombre"), 1, 35)
                  Else
                     data_infarq.Recordset("nomcob") = data_arq.Recordset("nomcob")
                  End If
               Else
                  data_infarq.Recordset("nomcob") = Mid(labcob.Caption, 1, 35)
               End If
            Else
               data_infarq.Recordset("nomcob") = ""
            End If
            data_infarq.Recordset("total") = data_arq.Recordset("total")
            data_infarq.Recordset.Update
            data_arq.Recordset.MoveNext
         Loop
      Else
         frm_infarq.MousePointer = 0
         MsgBox "No hay registros"
      End If
      frm_infarq.MousePointer = 0
      b_acep.Enabled = True
      data_infarq.RecordSource = "Select * from infarq order by ano,mes"
      data_infarq.Refresh
      If Check2.Value = 1 Then
         CrystalReport1.ReportFileName = App.path & "\infpendd.rpt"
      Else
         CrystalReport1.ReportFileName = App.path & "\infpend.rpt"
      End If
      CrystalReport1.ReportTitle = "INFORME DEVOLUCIONES COBRADOR: " & txt_cob.Text & " " & labcob.Caption & " MES...:" & txt_mes.Text & "/" & txt_ano.Text
      CrystalReport1.Action = 1
   End If
   If Option5.Value = True Then
      Dim Xnomcli As String
      frm_infarq.MousePointer = 11
      Text1.Text = "C"
      If txt_cob.Text = 0 Then
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               If Check5.Value = 1 Then
                  data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon =" & 815
               Else
                  data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               End If
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               If Check5.Value = 1 Then
                  data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon =" & 815
               Else
                  data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon <>" & 100
               End If
               data_arq.Refresh
            End If
         End If
      Else
         If Check3.Value = 1 Then
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Trim(Xmesdear) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         Else
            If Check1.Value = 1 Then
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
               data_arq.Refresh
            End If
         End If
      End If
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
         Do While Not data_arq.Recordset.EOF
            If IsNull(data_arq.Recordset("nombre")) = False Then
               Xnomcli = Replace(data_arq.Recordset("nombre"), "'", chr(37))
            Else
               Xnomcli = "XX"
            End If
            MiBaseact.Execute "Insert into infarq (matricula,nrorec,color,mes,ano,nombre,cob,codzon,nomcob,total" & _
            ") values (" & data_arq.Recordset("matricula") & "," & data_arq.Recordset("nrorec") & _
            ",'" & data_arq.Recordset("color") & "'," & data_arq.Recordset("mes") & "," & _
            data_arq.Recordset("ano") & ",'" & Xnomcli & "'," & data_arq.Recordset("cob") & "," & _
            data_arq.Recordset("codzon") & ",'" & data_arq.Recordset("nomcob") & "'," & data_arq.Recordset("total") & ")"
            
            data_arq.Recordset.MoveNext
         Loop
      End If
      data_infarq.RecordSource = "Select * from infarq order by ano,mes"
      data_infarq.Refresh
      frm_infarq.MousePointer = 0
      b_acep.Enabled = True
      If Check2.Value = 1 Then
         CrystalReport1.ReportFileName = App.path & "\infpendd.rpt"
      Else
         CrystalReport1.ReportFileName = App.path & "\infpend.rpt"
      End If
      If Check5.Value = 1 Then
         CrystalReport1.ReportTitle = "INFORME DE RECIBOS COBRADOS COBRADOR: --SAUCE--  MES...:" & txt_mes.Text & "/" & txt_ano.Text
      Else
         CrystalReport1.ReportTitle = "INFORME DE RECIBOS COBRADOS COBRADOR: " & txt_cob.Text & " " & labcob.Caption & " MES...:" & txt_mes.Text & "/" & txt_ano.Text
      End If
      CrystalReport1.Action = 1
   End If
   If Option6.Value = True Then
      Dim Nomemi As String
      frm_infarq.MousePointer = 11
      If txt_mes.Text <> "" Then
         If txt_ano.Text <> "" Then
            If mf.Text <> "__/__/____" Then
               Nomemi = "emi"
               If Month(mf.Text) > 9 Then
                  Nomemi = Trim(Nomemi) + Trim(str(Month(mf.Text))) + Trim(Mid(mf.Text, 9, 2))
               Else
                  Nomemi = Trim(Nomemi) + Trim("0") + Trim(str(Month(mf.Text))) + Trim(Mid(mf.Text, 9, 2))
               End If
               If txt_cob.Text <> "" Then
                  If txt_cob.Text = 0 Then
                     data_emi.RecordSource = "Select * from " & Nomemi & " where fecha >='" & Format(mf.Text, "yyyy-mm-dd") & "'"
                     data_emi.Refresh
                  Else
                     data_emi.RecordSource = "Select * from " & Nomemi & " where fecha >='" & Format(mf.Text, "yyyy-mm-dd") & "' And nro_cobr =" & txt_cob.Text
                     data_emi.Refresh
                  End If
                  If data_emi.Recordset.RecordCount > 0 Then
                     data_emi.Recordset.MoveFirst
                     Do While Not data_emi.Recordset.EOF
                        data_infarq.Recordset.AddNew
                        data_infarq.Recordset("matricula") = data_emi.Recordset("cliente")
                        data_infarq.Recordset("nombre") = Mid(data_emi.Recordset("apellidos"), 1, 30)
                        data_infarq.Recordset("nrorec") = data_emi.Recordset("documento")
                        data_infarq.Recordset("color") = data_emi.Recordset("color_rec")
                        data_infarq.Recordset("mes") = data_emi.Recordset("mes")
                        data_infarq.Recordset("ano") = data_emi.Recordset("ano")
                        data_infarq.Recordset("cob") = data_emi.Recordset("nro_cobr")
                        data_infarq.Recordset("nomcob") = data_emi.Recordset("nom_cobr")
                        data_infarq.Recordset("total") = data_emi.Recordset("total")
                        data_infarq.Recordset.Update
                        data_emi.Recordset.MoveNext
                     Loop
                     frm_infarq.MousePointer = 0
                     b_acep.Enabled = True
                     data_infarq.RecordSource = "Select * from infarq order by ano,mes"
                     data_infarq.Refresh
                     CrystalReport1.ReportFileName = App.path & "\infnuev.rpt"
                     CrystalReport1.ReportTitle = "INFORME DE NUEVAS ENTREGAS COBRADOR: " & txt_cob.Text & " " & labcob.Caption & " MES...:" & txt_mes.Text & "/" & txt_ano.Text
                     CrystalReport1.Action = 1
                  End If
               Else
                  MsgBox "No ingresó cobrador", vbInformation, "Mensaje"
                  txt_cob.SetFocus
               End If
            End If
        End If
     End If
   End If
      
   If Option7.Value = True Then
      Dim Nomemdos As String
      Dim Xtotpes As Double
      Dim Xtotrec As Double
      Dim Xtotpor As Double
      Dim Xcob As Long
      Dim Xnomcob As String
      Dim Xtotemi As Double
      Dim Xnewentre As Date
      Dim Nomem3 As String

      MiBaseact.Execute "Delete * from infvtas"
      data_inflin.RecordSource = "infvtas"
      data_inflin.Refresh
        Nomem3 = "emi"
        If txt_mes.Text > 9 Then
           Nomem3 = Trim(Nomem3) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
        Else
           Nomem3 = Trim(Nomem3) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
        End If
        If txt_cob.Text = 0 Then
           data_emi.RecordSource = "Select * from " & Nomem3 & " where nro_cobr =" & 615
        Else
           data_emi.RecordSource = "Select * from " & Nomem3 & " where nro_cobr =" & txt_cob.Text
        End If
        data_emi.Refresh
        If data_emi.Recordset.RecordCount > 0 Then
           data_emi.Recordset.MoveFirst
           If IsNull(data_emi.Recordset("fecha")) = False Then
              Xnewentre = data_emi.Recordset("fecha") + 1
           Else
              Xnewentre = Date + 1
           End If
        Else
           Xnewentre = Date + 1
        End If
      Xcob = 0
      MsgBox "FECHA DE NUEVAS ENTREGAS: " & Xnewentre, vbExclamation, "INFORMES ARQUEO"
      frm_infarq.MousePointer = 11
      Text1.Text = "C"
      MiBaseact.Execute "Delete * from infarqc"
      data_infresum.RecordSource = "infarqc"
      data_infresum.Refresh
      If txt_mes.Text <> "" Then
         If txt_ano.Text <> "" Then
            Nomemdos = "emi"
            Xmiar = "arq"
            If txt_mes.Text > 9 Then
               Nomemdos = Trim(Nomemdos) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
               Xmiar = Trim(Xmiar) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
            Else
               Nomemdos = Trim(Nomemdos) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
               Xmiar = Trim(Xmiar) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
            End If
            If Check3.Value = 1 Then
               If Check1.Value = 1 Then
                  If txt_cob.Text = 0 Then
                     data_arq.RecordSource = "Select * from " & Xmiar & " where arqueo in ('C','E','P') and codzon =" & 100 & " and cob not in (0) order by cob"
                  Else
                     data_arq.RecordSource = "Select * from " & Xmiar & " where arqueo in ('C','E','P') and codzon =" & 100 & " and cob =" & txt_cob.Text & " order by cob"
                  End If
                  data_arq.Refresh
               Else
                  If txt_cob.Text = 0 Then
                     data_arq.RecordSource = "Select * from " & Xmiar & " where arqueo in ('C','E','P') and codzon <>" & 100 & " and cob not in (0) order by cob"
                  Else
                     data_arq.RecordSource = "Select * from " & Xmiar & " where arqueo in ('C','E','P') and codzon <>" & 100 & " and cob =" & txt_cob.Text & " order by cob"
                  End If
                  data_arq.Refresh
               End If
            Else
               If Check1.Value = 1 Then
                  If txt_cob.Text = 0 Then
                     data_arq.RecordSource = "Select * from arqueo where arqueo in ('C','E','P') and codzon =" & 100 & " and cob not in (0) order by cob"
                  Else
                     data_arq.RecordSource = "Select * from arqueo where arqueo in ('C','E','P') and codzon =" & 100 & " and cob =" & txt_cob.Text & " order by cob"
                  End If
                  data_arq.Refresh
               Else
                  If txt_cob.Text = 0 Then
                     data_arq.RecordSource = "Select * from arqueo where arqueo in ('C','E','P') and codzon <>" & 100 & " and cob not in (0) order by cob"
                  Else
                     data_arq.RecordSource = "Select * from arqueo where arqueo in ('C','E','P') and codzon <>" & 100 & " and cob =" & txt_cob.Text & " order by cob"
                  End If
                  data_arq.Refresh
               End If
            End If
            data_arq.Recordset.MoveFirst
            If IsNull(data_arq.Recordset("cob")) = False Then
               Xcob = data_arq.Recordset("cob")
            Else
               Xcob = 0
            End If
            DoEvents
            Do While Not data_arq.Recordset.EOF
               If Xcob = data_arq.Recordset("cob") Then
                  If data_arq.Recordset("arqueo") = "C" Then
                     Xtotpes = Xtotpes + data_arq.Recordset("total")
                     Xtotrec = Xtotrec + 1
                     Xcob = data_arq.Recordset("cob")
                     If IsNull(data_arq.Recordset("nomcob")) = False Then
                        Xnomcob = data_arq.Recordset("nomcob")
                     Else
                        Xnomcob = "SC"
                     End If
                     data_arq.Recordset.MoveNext
                  Else
                     Xcob = data_arq.Recordset("cob")
                     Xnomcob = data_arq.Recordset("nomcob")
                     data_arq.Recordset.MoveNext
                  End If
               Else
                  If Check1.Value = 1 Then
                     data_emi.RecordSource = "Select * from linmmdd where grupo =" & Xcob & " and tipo <>'" & "NOTA CR" & "'"
                     data_emi.Refresh
                  Else
                     If Xnewentre <> Date Then
    '                     data_emi.Connect = ""
    '                     data_emi.DatabaseName = App.Path & "\emisiones.mdb"
                         data_emi.RecordSource = "Select * from " & Nomemdos & " where nro_cobr =" & Xcob & " and fecha <'" & Format(Xnewentre, "yyyy-mm-dd") & "'"
                         data_emi.Refresh
                     Else
    '                     data_emi.Connect = ""
    '                     data_emi.DatabaseName = App.Path & "\emisiones.mdb"
                         data_emi.RecordSource = "Select * from " & Nomemdos & " where nro_cobr =" & Xcob
                         data_emi.Refresh
                     End If
                  End If
                  Xtotemi = 0
                  If data_emi.Recordset.RecordCount > 0 Then
'                     data_emi.Recordset.MoveFirst
'                     data_emi.Recordset.MoveLast
                     Xtotemi = data_emi.Recordset.RecordCount
                  End If
                  If Xtotemi > 0 Then
                     Xtotpor = Xtotrec / Xtotemi
                  Else
                     Xtotpor = 0
                  End If
                  Xtotpor = Xtotpor * 100
                  If Xcob <> 615 And _
                     Xcob <> 616 And _
                     Xcob <> 603 And _
                     Xcob <> 602 And _
                     Xcob <> 635 And _
                     Xcob <> 679 And _
                     Xcob <> 636 And Xcob <> 676 And _
                     Xcob <> 672 And Xcob <> 653 And _
                     Xcob <> 8 And Xcob <> 201 And _
                     Xcob <> 1 And Xcob <> 606 And _
                     Xcob <> 685 And Xcob <> 512 And _
                     Xcob <> 10 And Xcob <> 208 And _
                     Xcob <> 113 And Xcob <> 209 Then
                     data_infresum.Recordset.AddNew
                     data_infresum.Recordset("cob") = Xcob
                     data_infresum.Recordset("nomcob") = Xnomcob
                     data_infresum.Recordset("mes") = txt_mes.Text
                     data_infresum.Recordset("ano") = txt_ano.Text
                     data_infresum.Recordset("totimp") = Xtotpes
                     data_infresum.Recordset("totrec") = Xtotrec
                     data_infresum.Recordset("totrecu") = Xtotemi
                     data_infresum.Recordset("desc1") = Trim(str(Val(Int(Xtotpor)))) + " %"
                     
                     data_lin.RecordSource = "Select * from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh, "yyyy-mm-dd") & "' and grupo =" & Xcob
                     data_lin.Refresh
                     If data_lin.Recordset.RecordCount > 0 Then
                        data_lin.Recordset.MoveFirst
                        Do While Not data_lin.Recordset.EOF
                           Xtotpesosbase = Xtotpesosbase + data_lin.Recordset("tot_lin")
                           Xtotrecbase = Xtotrecbase + 1
                           data_lin.Recordset.MoveNext
                        Loop
                     End If
                     data_infresum.Recordset("promcel") = Xtotrecbase
                     data_infresum.Recordset("totcel") = Xtotpesosbase
                     data_infresum.Recordset.Update
                     Xcob = data_arq.Recordset("cob")
'                     If IsNull(data_arq.Recordset("cob")) = False Then
'                        If txt_cob.Text = 0 Then
'                           data_cob.RecordSource = "Select * from cobrador where cb_numero =" & data_arq.Recordset("cob")
'                           data_cob.Refresh
'                           If data_cob.Recordset.RecordCount > 0 Then
'                              Xnomcob = Mid(data_cob.Recordset("cb_nombre"), 1, 35)
'                           Else
                              Xnomcob = data_arq.Recordset("nomcob")
'                           End If
'                        Else
'                           Xnomcob = Mid(labcob.Caption, 1, 35)
'                        End If
'                     Else
'                        Xnomcob = "S/C"
'                     End If
                     Xtotpes = 0
                     Xtotrec = 0
                     Xtotpesosbase = 0
                     Xtotrecbase = 0
                  Else
                     Xtotrec = 0
                     Xtotpes = 0
                     Xtotpesosbase = 0
                     Xtotrecbase = 0
                     data_lin.RecordSource = "Select * from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh, "yyyy-mm-dd") & "' and grupo =" & Xcob
                     data_lin.Refresh
                     If data_lin.Recordset.RecordCount > 0 Then
                        data_lin.Recordset.MoveFirst
                        Do While Not data_lin.Recordset.EOF
                           Xtotpes = Xtotpes + data_lin.Recordset("tot_lin")
                           Xtotrec = Xtotrec + 1
                           data_lin.Recordset.MoveNext
                        Loop
                     End If
                     If Xtotemi > 0 Then
                        Xtotpor = Xtotrec / Xtotemi
                     Else
                        Xtotpor = 0
                     End If
                     Xtotpor = Xtotpor * 100
                     
                     data_infresum.Recordset.AddNew
                     data_infresum.Recordset("cob") = Xcob
                     data_infresum.Recordset("nomcob") = Xnomcob
                     data_infresum.Recordset("mes") = txt_mes.Text
                     data_infresum.Recordset("ano") = txt_ano.Text
                     data_infresum.Recordset("totimp") = Xtotpes
                     data_infresum.Recordset("totrec") = Xtotrec
                     data_infresum.Recordset("totrecu") = Xtotemi
                     data_infresum.Recordset("desc1") = Trim(str(Val(Int(Xtotpor)))) + " %"
                     data_infresum.Recordset.Update
                                          
                     Xcob = data_arq.Recordset("cob")
'                     If IsNull(data_arq.Recordset("cob")) = False Then
'                        If txt_cob.Text = 0 Then
'                           data_cob.RecordSource = "Select * from cobrador where cb_numero =" & data_arq.Recordset("cob")
'                           data_cob.Refresh
'                           If data_cob.Recordset.RecordCount > 0 Then
'                              Xnomcob = Mid(data_cob.Recordset("cb_nombre"), 1, 35)
'                           Else
                              Xnomcob = data_arq.Recordset("nomcob")
'                           End If
'                        Else
'                           Xnomcob = Mid(labcob.Caption, 1, 35)
'                        End If
'                     Else
'                        Xnomcob = "S/C"
'                     End If
                     Xtotpes = 0
                     Xtotrec = 0
                     Xtotpesosbase = 0
                     Xtotrecbase = 0
                  End If
               End If
            Loop
            
            data_arq.Recordset.MovePrevious
            If Check1.Value = 1 Then
               data_emi.RecordSource = "Select * from linmmdd where grupo =" & Xcob & " and tipo <>'" & "NOTA CR" & "'"
               data_emi.Refresh
            Else
               data_emi.RecordSource = "Select * from " & Nomemdos & " where nro_cobr =" & Xcob & " and fecha <='" & Format(Xnewentre, "yyyy-mm-dd") & "'"
               data_emi.Refresh
            End If
            If data_emi.Recordset.RecordCount > 0 Then
'               data_emi.Recordset.MoveFirst
               data_emi.Recordset.MoveLast
               Xtotemi = data_emi.Recordset.RecordCount
            End If
            If Xtotemi = 0 Then
               Xtotpor = 0
            Else
               Xtotpor = Xtotrec / Xtotemi
               Xtotpor = Xtotpor * 100
            End If
            If Xcob <> 615 And _
               Xcob <> 616 And _
               Xcob <> 603 And _
               Xcob <> 602 And _
               Xcob <> 635 And _
               Xcob <> 679 And _
               Xcob <> 636 And _
               Xcob <> 672 And Xcob <> 653 And _
               Xcob <> 8 And Xcob <> 201 And _
               Xcob <> 1 And Xcob <> 606 And _
               Xcob <> 685 And Xcob <> 512 And _
               Xcob <> 10 And Xcob <> 208 And _
               Xcob <> 113 And Xcob <> 209 Then
            Else
               Xtotrec = 0
               Xtotpes = 0
               data_lin.RecordSource = "Select * from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh, "yyyy-mm-dd") & "' and grupo =" & Xcob
               data_lin.Refresh
               If data_lin.Recordset.RecordCount > 0 Then
                  data_lin.Recordset.MoveFirst
                  Do While Not data_lin.Recordset.EOF
                     Xtotpes = Xtotpes + data_lin.Recordset("tot_lin")
                     Xtotrec = Xtotrec + 1
                     data_lin.Recordset.MoveNext
                  Loop
               End If
               If Xtotemi > 0 Then
                  Xtotpor = Xtotrec / Xtotemi
               Else
                  Xtotpor = 0
               End If
               Xtotpor = Xtotpor * 100
            End If
'            If Xtotpor <= 100 Then
               data_infresum.Recordset.AddNew
               data_infresum.Recordset("cob") = Xcob
               data_infresum.Recordset("nomcob") = Xnomcob
               data_infresum.Recordset("mes") = txt_mes.Text
               data_infresum.Recordset("ano") = txt_ano.Text
               data_infresum.Recordset("totimp") = Xtotpes
               data_infresum.Recordset("totrec") = Xtotrec
               data_infresum.Recordset("totrecu") = Xtotemi
               data_infresum.Recordset("desc1") = Trim(str(Format(Xtotpor, "Standard"))) + " %"
               data_infresum.Recordset.Update
'            End If
            Xcob = data_arq.Recordset("cob")
            Xtotpes = 0
            Xtotrec = 0
            Xtotpesosbase = 0
            Xtotrecbase = 0
            
            'cobranza de base cobrador cero
            data_lin.RecordSource = "Select * from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh, "yyyy-mm-dd") & "' and grupo =" & 0
            data_lin.Refresh
            If data_lin.Recordset.RecordCount > 0 Then
               data_lin.Recordset.MoveFirst
               Do While Not data_lin.Recordset.EOF
                  Xtotpes = Xtotpes + data_lin.Recordset("tot_lin")
                  Xtotrec = Xtotrec + 1
                  data_lin.Recordset.MoveNext
               Loop
            End If
            Xtotpor = 100
            
            data_infresum.Recordset.AddNew
            data_infresum.Recordset("cob") = 0
            data_infresum.Recordset("nomcob") = "Base sin Cobr."
            data_infresum.Recordset("mes") = txt_mes.Text
            data_infresum.Recordset("ano") = txt_ano.Text
            data_infresum.Recordset("totimp") = Xtotpes
            data_infresum.Recordset("totrec") = Xtotrec
            data_infresum.Recordset("totrecu") = 0
            data_infresum.Recordset("desc1") = Trim(str(Val(Int(Xtotpor)))) + " %"
            data_infresum.Recordset.Update
            frm_infarq.MousePointer = 0
            b_acep.Enabled = True
            data_infresum.RecordSource = "Select * from infarqc order by cob"
            data_infresum.Refresh
            CrystalReport1.ReportFileName = App.path & "\infcobarq.rpt"
            CrystalReport1.Action = 1
            
         End If
      End If
      Xnewentre = Date
      frm_infarq.MousePointer = 0
   End If
   
   If Option8.Value = True Or Option9.Value = True Or Option10.Value = True Or Option12.Value = True Or _
      Option14.Value = True Or Option13.Value = True Or Option15.Value = True Then
      If Option10.Value = True Or Option12.Value = True Or Option13.Value = True Or Option15.Value = True Then
         Command3_Click
      Else
         If Option14.Value = True Then
            frm_vtasxfac.Show vbModal
         Else
            Command2_Click
         End If
      End If
   End If
   Xcantem = 0
   Xpesosem = 0
   Xtotrec = 0
   Xtotpes = 0
   Xtotimpo = 0
   Xtotiva = 0
   Xtotdeu = 0
   Xtottiq = 0
End If
b_cance.Enabled = True
b_acep.Enabled = True

Exit Sub

Xinfarqer:
          If Err.Number = 53 Then
             MsgBox "No existe el archivo"
          Else
             MsgBox "Error en el informe, Verifique!", Err.Description
          End If
          
End Sub

Private Sub b_cance_Click()
Unload Me


End Sub

Private Sub Command1_Click()
   
       Dim MiBaseact As Database
       Dim Unasesact As Workspace
       Set Unasesact = Workspaces(0)
       Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
       
       MiBaseact.Execute "Delete * from infarqc"
       data_infresum.RecordSource = "infarqc"
       data_infresum.Refresh
   
   If Option7.Value = True Then
      Dim Nomemdos As String
      Dim Xtotpes As Double
      Dim Xtotrec As Double
      Dim Xtotpor As Long
      Dim Xcob As Long
      Dim Xnomcob As String
      Dim Xtotemi As Double
      Xcob = 0
      frm_infarq.MousePointer = 11
      Text1.Text = "C"
'      If data_infarq.Recordset.RecordCount > 0 Then
'         data_infarq.Recordset.MoveFirst
'         Do While Not data_infarq.Recordset.EOF
'            data_infarq.Recordset.Delete
'            data_infarq.Recordset.MoveNext
'         Loop
'      End If
      
      If txt_mes.Text <> "" Then
         If txt_ano.Text <> "" Then
            Nomemdos = "emi"
            If txt_mes.Text > 9 Then
               Nomemdos = Trim(Nomemdos) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
            Else
               Nomemdos = Trim(Nomemdos) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
            End If
            data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' order by cob"
            data_arq.Refresh
            data_arq.Recordset.MoveFirst
            Xcob = data_arq.Recordset("cob")
            Do While Not data_arq.Recordset.EOF
               If Xcob = data_arq.Recordset("cob") Then
                  If data_arq.Recordset("arqueo") = "C" Then
                     Xtotpes = Xtotpes + data_arq.Recordset("total")
                     Xtotrec = Xtotrec + 1
                     Xcob = data_arq.Recordset("cob")
                     If IsNull(data_arq.Recordset("nomcob")) = False Then
                        Xnomcob = data_arq.Recordset("nomcob")
                     Else
                        Xnomcob = Trim(str(data_arq.Recordset("cob")))
                     End If
                     data_arq.Recordset.MoveNext
                  Else
                     Xcob = data_arq.Recordset("cob")
                     data_arq.Recordset.MoveNext
                  End If
               Else
                  data_emi.RecordSource = "Select * from " & Nomemdos & " where nro_cobr =" & Xcob
                  data_emi.Refresh
                  If data_emi.Recordset.RecordCount > 0 Then
                     data_emi.Recordset.MoveFirst
                     data_emi.Recordset.MoveLast
                     Xtotemi = data_emi.Recordset.RecordCount
                  End If
                  Xtotpor = Xtotrec / Xtotemi
                  Xtotpor = Xtotpor * 100
                  If data_arq.Recordset("cob") <> 615 And _
                     data_arq.Recordset("cob") <> 616 And _
                     data_arq.Recordset("cob") <> 603 And _
                     data_arq.Recordset("cob") <> 602 And _
                     data_arq.Recordset("cob") <> 635 And _
                     data_arq.Recordset("cob") <> 632 And data_arq.Recordset("cob") <> 606 And _
                     data_arq.Recordset("cob") <> 636 And _
                     data_arq.Recordset("cob") <> 672 And _
                     data_arq.Recordset("cob") <> 8 And _
                     data_arq.Recordset("cob") <> 1 And _
                     data_arq.Recordset("cob") <> 10 Then
                     data_infresum.Recordset.AddNew
                     data_infresum.Recordset("cob") = Xcob
                     data_infresum.Recordset("nomcob") = Xnomcob
                     data_infresum.Recordset("mes") = txt_mes.Text
                     data_infresum.Recordset("ano") = txt_ano.Text
                     data_infresum.Recordset("totimp") = Xtotpes
                     data_infresum.Recordset("totrec") = Xtotrec
                     data_infresum.Recordset("totrecu") = Xtotemi
                     data_infresum.Recordset("desc1") = Trim(str(Xtotpor)) + " %"
                     data_infresum.Recordset.Update
                     Xcob = data_arq.Recordset("cob")
                     If IsNull(data_arq.Recordset("nomcob")) = False Then
                        Xnomcob = data_arq.Recordset("nomcob")
                     Else
                        Xnomcob = Trim(str(data_arq.Recordset("cob")))
                     End If
                     Xtotpes = 0
                     Xtotrec = 0
                 Else
                     Xcob = data_arq.Recordset("cob")
                     If IsNull(data_arq.Recordset("nomcob")) = False Then
                        Xnomcob = data_arq.Recordset("nomcob")
                     Else
                        Xnomcob = Trim(str(data_arq.Recordset("cob")))
                     End If
                     Xtotpes = 0
                     Xtotrec = 0
                 End If
               End If
            Loop
            data_arq.Recordset.MovePrevious
            data_emi.RecordSource = "Select * from " & Nomemdos & " where nro_cobr =" & Xcob
            data_emi.Refresh
            If data_emi.Recordset.RecordCount > 0 Then
               data_emi.Recordset.MoveFirst
               data_emi.Recordset.MoveLast
               Xtotemi = data_emi.Recordset.RecordCount
            End If
            Xtotpor = Xtotrec / Xtotemi
            Xtotpor = Xtotpor / 100
            data_infresum.Recordset.AddNew
            data_infresum.Recordset("cob") = Xcob
            data_infresum.Recordset("nomcob") = Xnomcob
            data_infresum.Recordset("mes") = txt_mes.Text
            data_infresum.Recordset("ano") = txt_ano.Text
            data_infresum.Recordset("totimp") = Xtotpes
            data_infresum.Recordset("totrec") = Xtotrec
            data_infresum.Recordset("totrecu") = Xtotemi
            data_infresum.Recordset("desc1") = Trim(str(Xtotpor)) + " %"
            data_infresum.Recordset.Update
            Xcob = data_arq.Recordset("cob")
            Xtotpes = 0
            Xtotrec = 0
            frm_infarq.MousePointer = 0
            b_acep.Enabled = True
            data_infresum.RecordSource = "Select * from infarqc order by cob"
            data_infresum.Refresh
            CrystalReport1.ReportFileName = App.path & "\infcobarq.rpt"
            CrystalReport1.Action = 1
            
         End If
      End If
      frm_infarq.MousePointer = 0
   End If
   
   If Option8.Value = True Then
'      Dim Xtotpes As Double
'      Dim Xtotrec As Double
      
      
      frm_infarq.MousePointer = 11
      Text1.Text = "C"
      data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "'"
      data_arq.Refresh
      data_arq.Recordset.MoveFirst
      Do While Not data_arq.Recordset.EOF
         If data_arq.Recordset("cob") <> 615 And _
            data_arq.Recordset("cob") <> 616 And _
            data_arq.Recordset("cob") <> 603 And _
            data_arq.Recordset("cob") <> 602 And _
            data_arq.Recordset("cob") <> 635 And _
            data_arq.Recordset("cob") <> 632 And _
            data_arq.Recordset("cob") <> 636 And _
            data_arq.Recordset("cob") <> 672 And data_arq.Recordset("cob") <> 606 And _
            data_arq.Recordset("cob") <> 8 And _
            data_arq.Recordset("cob") <> 1 And _
            data_arq.Recordset("cob") <> 685 And _
            data_arq.Recordset("cob") <> 10 And _
            data_arq.Recordset("cob") <> 113 Then
            Xtotrec = Xtotrec + 1
            Xtotpes = Xtotpes + data_arq.Recordset("total")
            If IsNull(data_arq.Recordset("tiquet")) = False Then
               Xtotpes = Xtotpes - data_arq.Recordset("tiquet")
            End If
         End If
         data_arq.Recordset.MoveNext
      Loop
      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "COBRADOS"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset.Update
      Xtotrec = 0
      Xtotpes = 0

      Text1.Text = "P"
      data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "'"
      data_arq.Refresh
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
      End If
      Do While Not data_arq.Recordset.EOF
         If data_arq.Recordset("cob") <> 615 And _
            data_arq.Recordset("cob") <> 616 And _
            data_arq.Recordset("cob") <> 603 And _
            data_arq.Recordset("cob") <> 602 And _
            data_arq.Recordset("cob") <> 635 And _
            data_arq.Recordset("cob") <> 632 And _
            data_arq.Recordset("cob") <> 636 And _
            data_arq.Recordset("cob") <> 672 And _
            data_arq.Recordset("cob") <> 8 And _
            data_arq.Recordset("cob") <> 1 And data_arq.Recordset("cob") <> 606 And _
            data_arq.Recordset("cob") <> 685 And _
            data_arq.Recordset("cob") <> 10 Then
            Xtotrec = Xtotrec + 1
            Xtotpes = Xtotpes + data_arq.Recordset("total")
         End If
         data_arq.Recordset.MoveNext
      Loop
      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "PENDIENTES"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset.Update
      Xtotrec = 0
      Xtotpes = 0

      Text1.Text = "B"
      data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "'"
      data_arq.Refresh
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
      End If
      Do While Not data_arq.Recordset.EOF
         If data_arq.Recordset("cob") <> 615 And _
            data_arq.Recordset("cob") <> 616 And _
            data_arq.Recordset("cob") <> 603 And _
            data_arq.Recordset("cob") <> 602 And _
            data_arq.Recordset("cob") <> 635 And _
            data_arq.Recordset("cob") <> 632 And _
            data_arq.Recordset("cob") <> 636 And _
            data_arq.Recordset("cob") <> 672 And data_arq.Recordset("cob") <> 606 And _
            data_arq.Recordset("cob") <> 8 And _
            data_arq.Recordset("cob") <> 1 And _
            data_arq.Recordset("cob") <> 685 And _
            data_arq.Recordset("cob") <> 10 Then
            Xtotrec = Xtotrec + 1
            Xtotpes = Xtotpes + data_arq.Recordset("total")
         End If
         data_arq.Recordset.MoveNext
      Loop
      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "BAJAS"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset.Update
      Xtotrec = 0
      Xtotpes = 0
   
      Text1.Text = "D"
      data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "'"
      data_arq.Refresh
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
      End If
      Do While Not data_arq.Recordset.EOF
         If data_arq.Recordset("cob") <> 615 And _
            data_arq.Recordset("cob") <> 616 And _
            data_arq.Recordset("cob") <> 603 And _
            data_arq.Recordset("cob") <> 602 And _
            data_arq.Recordset("cob") <> 635 And _
            data_arq.Recordset("cob") <> 632 And _
            data_arq.Recordset("cob") <> 636 And _
            data_arq.Recordset("cob") <> 672 And data_arq.Recordset("cob") <> 606 And _
            data_arq.Recordset("cob") <> 8 And _
            data_arq.Recordset("cob") <> 1 And _
            data_arq.Recordset("cob") <> 685 And _
            data_arq.Recordset("cob") <> 10 Then
            Xtotrec = Xtotrec + 1
            Xtotpes = Xtotpes + data_arq.Recordset("total")
         End If
         data_arq.Recordset.MoveNext
      Loop
      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "DEVOLUCION"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset.Update
   End If
   Xtotrec = 0
   Xtotpes = 0
   
   If txt_mes.Text <> "" Then
      If txt_ano.Text <> "" Then
         Nomemdos = "emi"
         If txt_mes.Text > 9 Then
            Nomemdos = Trim(Nomemdos) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         Else
            Nomemdos = Trim(Nomemdos) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         End If
         data_emi.RecordSource = "Select * from " & Nomemdos & " order by nro_cobr"
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               If Format(data_emi.Recordset("fecha"), "yyyy/mm/dd") >= Format(mff.Text, "yyyy/mm/dd") Then
                  Xtotrec = Xtotrec + 1
                  Xtotpes = Xtotpes + data_emi.Recordset("total")
               End If
               data_emi.Recordset.MoveNext
            Loop
            data_infresum.Recordset.AddNew
            data_infresum.Recordset("nomcob") = "NUEVAS ENTREGAS"
            data_infresum.Recordset("totrec") = Xtotrec
            data_infresum.Recordset("totimp") = Xtotpes
            data_infresum.Recordset("mes") = txt_mes.Text
            data_infresum.Recordset("ano") = txt_ano.Text
            data_infresum.Recordset.Update
         End If
         data_infresum.RecordSource = "select * from infarqc"
         data_infresum.Refresh
         frm_infarq.MousePointer = 0
         b_acep.Enabled = True
         CrystalReport1.ReportFileName = App.path & "\infcontab.rpt"
         CrystalReport1.Action = 1
      End If
   End If

End Sub

Private Sub Command2_Click()
Dim Xarqq, Xmiar As String
Dim Xmar, Xaar As Integer
Dim Xtottiq, Xtotimpo, Xtotdeu, Xtotiva, Xeliva As Double
Dim Xtotrecbor, Xtotpesbor As Double
Dim Xelmesact, Xelanoact As Integer
If Month(Date) = 12 Then
   Xelmesact = 1
Else
   Xelmesact = Month(Date) + 1
End If
Xelanoact = Year(Date)
Xarqq = "arq"
frm_infarq.MousePointer = 11
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"

data_inflin.RecordSource = "infvtas"
data_inflin.Refresh

MiBaseact.Execute "Delete * from infarqc"

data_infresum.RecordSource = "infarqc"
data_infresum.Refresh

MiBaseact.Execute "Delete * from infarq"

data_arqloc.RecordSource = "infarq"
data_arqloc.Refresh

If Option8.Value = True Then
   Dim Xeldia44 As Integer
   Dim Xfd44, Xfh44 As String
   
   Xeldia44 = Day(DateSerial(txt_ano.Text, txt_mes.Text + 1, 0))
   If Val(txt_mes.Text) > 9 Then
      Xfd44 = "01/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
      If Xeldia44 > 9 Then
         Xfh44 = Trim(str(Xeldia44)) & "/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
      Else
         Xfh44 = "0" & Trim(str(Xeldia44)) & "/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
      End If
   Else
      Xfd44 = "01/" & "0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
      If Xeldia44 > 9 Then
         Xfh44 = Trim(str(Xeldia44)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
      Else
         Xfh44 = "0" & Trim(str(Xeldia44)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
      End If
   End If
'   MiBaseact.Execute "Select sum(importe) from arq0317 where arqueo = 'C'"
   Dim xxx As Double
'   data_lin.RecordSource = "Select fecha,cod_prod,cod_cli,tot_lin,mes_paga,ano_paga,grupo from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd44, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh44, "yyyy-mm-dd") & "' and" & _
'   " grupo not in (601,699,680,649,605,650,678,656,202)"
'   data_lin.Refresh
'   data_lin.RecordSource = "Select sum(tot_lin) as totlin from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd44, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh44, "yyyy-mm-dd") & "'"
'   data_lin.Refresh
'   xxx = data_lin.Recordset("totlin")
   
'   If data_lin.Recordset.RecordCount > 0 Then
'      data_lin.Recordset.MoveFirst
'      Do While Not data_lin.Recordset.EOF
'         MiBaseact.Execute "Insert into infvtas (fecha,cod_prod,cod_cli,tot_lin,mes_paga,ano_paga,grupo) values " & _
'         "('" & data_lin.Recordset("fecha") & "'," & data_lin.Recordset("cod_prod") & "," & data_lin.Recordset("cod_cli") & "," & _
'         data_lin.Recordset("tot_lin") & "," & data_lin.Recordset("mes_paga") & "," & data_lin.Recordset("ano_paga") & "," & data_lin.Recordset("grupo") & ")"
'         data_lin.Recordset.MoveNext
'      Loop
'      data_inflin.Refresh
'   End If

      Text1.Text = "C"
      Xmiar = "arq"
      If Check3.Value = 1 Then
         If txt_mes.Text > 9 Then
            Xmiar = Trim(Xmiar) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         Else
            Xmiar = Trim(Xmiar) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         End If
      Else
         Xmiar = "arqueo"
      End If
      Xtotrec = 0
      Xtotpes = 0
      Xtotiva = 0
      Xtottiq = 0
      Xtotdeu = 0
      Xtotimpo = 0
      Xtotrecbor = 0
      Xtotpesbor = 0
      data_arq.RecordSource = "Select sum(total) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotpes = Xtotpes + data_arq.Recordset("totlin")
      data_arq.RecordSource = "Select count(*) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotrec = Xtotrec + data_arq.Recordset("totlin")
      Xtotrecbor = Xtotrec
      Xtotpesbor = Xtotpes
      data_arq.RecordSource = "Select sum(importe) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotimpo = Xtotimpo + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(iva) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotiva = Xtotiva + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(tiquet) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtottiq = Xtottiq + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(varia) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotdeu = Xtotdeu + data_arq.Recordset("totlin")
      
      If Check3.Value = 1 Then
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select * from " & Xmiar & " where arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
            data_arq.RecordSource = "Select importe,tiquet,varia,total,cob from " & Xmiar & " where arqueo ='" & Text1.Text & "' and codzon <>" & 100 & _
            " and cob not in (11,6,5,616,615,602,635,653,209,10,1,636)"
            data_arq.Refresh
         End If
      Else
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select importe,tiquet,varia,total,cob from arqueo where arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
            data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon <>" & 100 & _
            " and cob not in (11,6,5,616,615,602,635,653,209,10,1,636)"
            data_arq.Refresh
         End If
      End If
      
'      If data_arq.Recordset.RecordCount > 0 Then
'         data_arq.Recordset.MoveFirst
'         Do While Not data_arq.Recordset.EOF
'            MiBaseact.Execute "Insert into infarq (importe,tiquet,varia,total,cob" & _
'            ") values (" & data_arq.Recordset("importe") & "," & data_arq.Recordset("tiquet") & "," & data_arq.Recordset("varia") & "," & _
'            data_arq.Recordset("total") & "," & data_arq.Recordset("cob") & ")"
'            data_arq.Recordset.MoveNext
'         Loop
'         data_arqloc.Refresh
'      End If
'      data_arqloc.Recordset.MoveFirst
      
      data_lin.RecordSource = "Select sum(tot_lin) as totlin from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd44, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh44, "yyyy-mm-dd") & "' and" & _
      " grupo in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111) and mes_paga not in (0)"
      data_lin.Refresh
      Xtotpes = Xtotpes + data_lin.Recordset("totlin")
      data_lin.RecordSource = "Select count(*) as totlin from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xfd44, "yyyy-mm-dd") & "' and fecha <='" & Format(Xfh44, "yyyy-mm-dd") & "' and" & _
      " grupo in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111) and mes_paga not in (0)"
      data_lin.Refresh
      Xtotrec = Xtotrec + data_lin.Recordset("totlin")
      Xtotrecbor = Xtotrec
      Xtotpesbor = Xtotpes

      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "COBRADOS"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset("comr") = Xtottiq
      data_infresum.Recordset("coma") = Xtotdeu
      data_infresum.Recordset("comm") = Xtotiva
      data_infresum.Recordset("comc") = Xtotimpo
      data_infresum.Recordset.Update
      Xtotcobc = Xtotrec
      Xtotcobp = Xtotpes
      Xtotrec = 0
      Xtotpes = 0
      Xtotimpo = 0
      Xtottiq = 0
      Xtotdeu = 0
      Xtotiva = 0
      Xmiar = "arq"
      If txt_mes.Text = 1 Then
         Xmar = 12
         Xaar = txt_ano.Text - 1
      Else
         Xmar = txt_mes.Text - 1
         Xaar = txt_ano.Text
      End If
      If Xmar > 9 Then
         Xarqq = Trim(Xarqq) + Trim(str(Xmar)) + Trim(Mid(str(Xaar), 4, 2))
      Else
         Xarqq = Trim(Xarqq) + Trim("0") + Trim(str(Xmar)) + Trim(Mid(str(Xaar), 4, 2))
      End If
      Text1.Text = "P"
'      data_arq.DatabaseName = ""
      data_arq.ConnectionString = "dsn=" & Xconexrmt
      Xcantem = 0
      Xpesosem = 0
      Xdifcan = Xdifcan + Xtotrec
      Xdifpes = Xdifpes + Xtotpes
      Xtotrec = 0
      Xtotpes = 0
      Xtotimpo = 0
      Xtottiq = 0
      Xtotdeu = 0
      Xtotiva = 0
      Text1.Text = "P"
'      data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
'      data_arq.DatabaseName = App.Path & "\sapp.mdb"
      If Check3.Value = 1 Then
         If txt_mes.Text > 9 Then
            Xmiar = Trim(Xmiar) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         Else
            Xmiar = Trim(Xmiar) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         End If
      Else
         Xmiar = "arqueo"
      End If
      data_arq.RecordSource = "Select sum(total) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotpes = Xtotpes + data_arq.Recordset("totlin")
      data_arq.RecordSource = "Select count(*) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotrec = Xtotrec + data_arq.Recordset("totlin")
'      Xtotrecbor = Xtotrec
'      Xtotpesbor = Xtotpes
      data_arq.RecordSource = "Select sum(importe) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotimpo = Xtotimpo + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(iva) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotiva = Xtotiva + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(tiquet) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtottiq = Xtottiq + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(varia) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotdeu = Xtotdeu + data_arq.Recordset("totlin")
      
      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "PENDIENTES FINAL"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset("comr") = Xtottiq
      data_infresum.Recordset("coma") = Xtotdeu
      data_infresum.Recordset("comm") = Xtotiva
      data_infresum.Recordset("comc") = Xtotimpo
      data_infresum.Recordset.Update
      Xdifcan = Xdifcan - Xtotrec
      Xdifpes = Xdifpes - Xtotpes
      Xtotrec = 0
      Xtotpes = 0
      Xtotimpo = 0
      Xtottiq = 0
      Xtotdeu = 0
      Xtotiva = 0
      Text1.Text = "B"
      
      data_arq.RecordSource = "Select sum(total) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotpes = Xtotpes + data_arq.Recordset("totlin")
      data_arq.RecordSource = "Select count(*) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotrec = Xtotrec + data_arq.Recordset("totlin")
'      Xtotrecbor = Xtotrec
'      Xtotpesbor = Xtotpes
      data_arq.RecordSource = "Select sum(importe) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotimpo = Xtotimpo + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(iva) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotiva = Xtotiva + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(tiquet) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtottiq = Xtottiq + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(varia) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotdeu = Xtotdeu + data_arq.Recordset("totlin")
      
      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "BAJAS"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset("comr") = Xtottiq
      data_infresum.Recordset("coma") = Xtotdeu
      data_infresum.Recordset("comm") = Xtotiva
      data_infresum.Recordset("comc") = Xtotimpo
      data_infresum.Recordset.Update
      Xdifcan = Xdifcan - Xtotrec
      Xdifpes = Xdifpes - Xtotpes
      Xtotrec = 0
      Xtotpes = 0
      Xtotimpo = 0
      Xtottiq = 0
      Xtotdeu = 0
      Xtotiva = 0
      Text1.Text = "D"

      data_arq.RecordSource = "Select sum(total) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotpes = Xtotpes + data_arq.Recordset("totlin")
      data_arq.RecordSource = "Select count(*) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotrec = Xtotrec + data_arq.Recordset("totlin")
'      Xtotrecbor = Xtotrec
'      Xtotpesbor = Xtotpes
      data_arq.RecordSource = "Select sum(importe) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotimpo = Xtotimpo + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(iva) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotiva = Xtotiva + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(tiquet) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtottiq = Xtottiq + data_arq.Recordset("totlin")
      
      data_arq.RecordSource = "Select sum(varia) as totlin from " & Xmiar & " where arqueo ='" & Text1.Text & "' and" & _
      " cob not in (615,616,603,602,635,209,606,636,208,672,653,8,201,1,685,512,10,110,113,679,111)"
      data_arq.Refresh
      Xtotdeu = Xtotdeu + data_arq.Recordset("totlin")

      data_infresum.Recordset.AddNew
      data_infresum.Recordset("nomcob") = "DEVOLUCION"
      data_infresum.Recordset("totrec") = Xtotrec
      data_infresum.Recordset("totimp") = Xtotpes
      data_infresum.Recordset("mes") = txt_mes.Text
      data_infresum.Recordset("ano") = txt_ano.Text
      data_infresum.Recordset("comr") = Xtottiq
      data_infresum.Recordset("coma") = Xtotdeu
      data_infresum.Recordset("comm") = Xtotiva
      data_infresum.Recordset("comc") = Xtotimpo
      data_infresum.Recordset.Update
      Xdifcan = Xdifcan - Xtotrec
      Xdifpes = Xdifpes - Xtotpes
      Xtotrec = 0
      Xtotpes = 0
      Xtotimpo = 0
      Xtottiq = 0
      Xtotdeu = 0
      Xtotiva = 0
      Xcantem = 0
      Xpesosem = 0
      Dim Xtotivados As Double
      Xtotivados = 0
      If txt_mes.Text <> "" Then
         If txt_ano.Text <> "" Then
            Nomemdos = "emi"
            If txt_mes.Text > 9 Then
               Nomemdos = Trim(Nomemdos) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
            Else
               Nomemdos = Trim(Nomemdos) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
            End If
            If Check1.Value = 1 Then
               Xcantem = 0
               Xpesosem = 0
            Else
                data_emi.RecordSource = "Select count(*) as totlin from " & Nomemdos & " where fecha <'" & Format(mff.Text, "yyyy/mm/dd") & "'"
                data_emi.Refresh
                Xcantem = data_emi.Recordset("totlin")
                data_emi.RecordSource = "Select sum(total) as totlin from " & Nomemdos & " where fecha <='" & Format(mff.Text, "yyyy/mm/dd") & "'"
                data_emi.Refresh
                Xpesosem = data_emi.Recordset("totlin")
                
                data_emi.RecordSource = "Select * from " & Nomemdos & " where fecha >='" & Format(mff.Text, "yyyy/mm/dd") & "'"
                data_emi.Refresh
                If data_emi.Recordset.RecordCount > 0 Then
                   data_emi.Recordset.MoveFirst
                   Do While Not data_emi.Recordset.EOF
                      If Format(data_emi.Recordset("fecha"), "yyyy/mm/dd") >= Format(mff.Text, "yyyy/mm/dd") Then
                        If data_emi.Recordset("nro_cobr") <> 0 Then
'                           data_emi.Recordset("nro_cobr") <> 616 And _
'                           data_emi.Recordset("nro_cobr") <> 603 And _
'                           data_emi.Recordset("nro_cobr") <> 602 And data_emi.Recordset("nro_cobr") <> 209 And _
'                           data_emi.Recordset("nro_cobr") <> 635 And data_emi.Recordset("nro_cobr") <> 208 And _
'                           data_emi.Recordset("nro_cobr") <> 636 And data_emi.Recordset("nro_cobr") <> 606 And _
'                           data_emi.Recordset("nro_cobr") <> 672 And data_emi.Recordset("nro_cobr") <> 653 And _
'                           data_emi.Recordset("nro_cobr") <> 8 And data_emi.Recordset("nro_cobr") <> 201 And _
'                           data_emi.Recordset("nro_cobr") <> 1 And _
'                           data_emi.Recordset("nro_cobr") <> 685 And data_emi.Recordset("nro_cobr") <> 512 And _
'                           data_emi.Recordset("nro_cobr") <> 10 And _
'                           data_emi.Recordset("nro_cobr") <> 113 And data_emi.Recordset("nro_cobr") <> 679 Then
                           Xtotrec = Xtotrec + 1
                           Xtotpes = Xtotpes + data_emi.Recordset("total")
                           Xtotimpo = Xtotimpo + data_emi.Recordset("importe")
                           If IsNull(data_emi.Recordset("tiquet")) = False Then
                              Xtottiq = Xtottiq + data_emi.Recordset("tiquet")
                           End If
                           If IsNull(data_emi.Recordset("iva")) = False Then
                              Xtotiva = Xtotiva + data_emi.Recordset("iva")
                           Else
                              Xtotivados = data_emi.Recordset("total") / 1.1
                              Xtotivados = Xtotivados * 0.1
                              Xtotiva = Xtotiva + Xtotivados
                           End If
                           If IsNull(data_emi.Recordset("deudas")) = False Then
                              Xtotdeu = Xtotdeu + data_emi.Recordset("deudas")
                           End If
                        End If
                      Else
                        If data_emi.Recordset("nro_cobr") <> 615 And _
                           data_emi.Recordset("nro_cobr") <> 616 And _
                           data_emi.Recordset("nro_cobr") <> 603 And _
                           data_emi.Recordset("nro_cobr") <> 602 And data_emi.Recordset("nro_cobr") <> 209 And _
                           data_emi.Recordset("nro_cobr") <> 635 And data_emi.Recordset("nro_cobr") <> 208 And _
                           data_emi.Recordset("nro_cobr") <> 636 And data_emi.Recordset("nro_cobr") <> 606 And _
                           data_emi.Recordset("nro_cobr") <> 672 And data_emi.Recordset("nro_cobr") <> 653 And _
                           data_emi.Recordset("nro_cobr") <> 8 And data_emi.Recordset("nro_cobr") <> 201 And _
                           data_emi.Recordset("nro_cobr") <> 1 And _
                           data_emi.Recordset("nro_cobr") <> 685 And data_emi.Recordset("nro_cobr") <> 512 And _
                           data_emi.Recordset("nro_cobr") <> 10 And _
                           data_emi.Recordset("nro_cobr") <> 113 And data_emi.Recordset("nro_cobr") <> 679 Then
                           Xcantem = Xcantem + 1
                           Xpesosem = Xpesosem + data_emi.Recordset("total")
                        Else
                           Xcantem = Xcantem + 1
                           Xpesosem = Xpesosem + data_emi.Recordset("total")
                        End If
                      End If
                      data_emi.Recordset.MoveNext
                   Loop
                   data_infresum.Recordset.AddNew
                   data_infresum.Recordset("nomcob") = "NUEVAS ENTREGAS"
                   data_infresum.Recordset("totrec") = Xtotrec
                   data_infresum.Recordset("totimp") = Xtotpes
                   data_infresum.Recordset("mes") = txt_mes.Text
                   data_infresum.Recordset("ano") = txt_ano.Text
                   data_infresum.Recordset("comr") = Xtottiq
                   data_infresum.Recordset("coma") = Xtotdeu
                   data_infresum.Recordset("comm") = Xtotiva
                   data_infresum.Recordset("comc") = Xtotimpo
                   data_infresum.Recordset.Update
                   Xdifcan = Xdifcan + Xtotrec
                   Xdifpes = Xdifpes + Xtotpes
                   Xdifcan = Xdifcan + Xcantem
                   Xdifpes = Xdifpes + Xpesosem
                   data_infresum.Recordset.MoveFirst
                   Do While Not data_infresum.Recordset.EOF
                      data_infresum.Recordset.Edit
                      data_infresum.Recordset("totrecu") = Xcantem
                      data_infresum.Recordset("totimpu") = Xpesosem
'                      data_infresum.Recordset("iva1") = Xdifcan
'                      data_infresum.Recordset("iva2") = Xdifpes
                      data_infresum.Recordset("iva1") = Xtotrecbor
                      data_infresum.Recordset("iva2") = Xtotpesbor
                      
                      data_infresum.Recordset("entrega") = Xtotcobc - Xdifcan
                      data_infresum.Recordset("quesob") = Xtotcobp - Xdifpes
                      data_infresum.Recordset.Update
                      data_infresum.Recordset.MoveNext
                   Loop
                End If
            End If
            data_infresum.RecordSource = "select * from infarqc"
            data_infresum.Refresh
            frm_infarq.MousePointer = 0
            b_acep.Enabled = True
            CrystalReport1.ReportFileName = App.path & "\infcontabd.rpt"
            CrystalReport1.Action = 1
         End If
      End If
End If

If Option9.Value = True Then
    MiBaseact.Execute "Delete * from infarq"
    data_infarq.RecordSource = "infarq"
    data_infarq.Refresh
    Xmiar = "arq"
      If Check3.Value = 1 Then
         If txt_mes.Text > 9 Then
            Xmiar = Trim(Xmiar) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         Else
            Xmiar = Trim(Xmiar) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
         End If
      Else
         Xmiar = "arqueo"
      End If
   
   Text1.Text = "E"
   If txt_cob.Text = 0 Then
      If Check3.Value = 1 Then
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select * from " & Trim(Xmiar) & " where arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
            data_arq.RecordSource = "Select * from " & Trim(Xmiar) & " where arqueo ='" & Text1.Text & "' and codzon <>" & 100
            data_arq.Refresh
         End If
      Else
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
            data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' and codzon <>" & 100
            data_arq.Refresh
         End If
      End If
   Else
      If Check3.Value = 1 Then
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select * from " & Trim(Xmiar) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
            data_arq.RecordSource = "Select * from " & Trim(Xmiar) & " where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
            data_arq.Refresh
         End If
      Else
         If Check1.Value = 1 Then
            data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon =" & 100
            data_arq.Refresh
         Else
            data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & Text1.Text & "' and codzon <>" & 100
            data_arq.Refresh
         End If
      End If
   End If
   If data_arq.Recordset.RecordCount > 0 Then
      data_arq.Recordset.MoveFirst
      Do While Not data_arq.Recordset.EOF
         data_infarq.Recordset.AddNew
         data_infarq.Recordset("matricula") = data_arq.Recordset("matricula")
         data_infarq.Recordset("nrorec") = data_arq.Recordset("nrorec")
         data_infarq.Recordset("color") = data_arq.Recordset("color")
         data_infarq.Recordset("mes") = data_arq.Recordset("mes")
         data_infarq.Recordset("ano") = data_arq.Recordset("ano")
         data_infarq.Recordset("cob") = data_arq.Recordset("cob")
         data_infarq.Recordset("nombre") = Mid(data_arq.Recordset("nombre"), 1, 30)
         If IsNull(data_arq.Recordset("cob")) = False Then
            If txt_cob.Text = 0 Then
'               data_cob.Recordset.FindFirst "cb_numero =" & data_arq.Recordset("cob")
               data_cob.RecordSource = "Select * from cobrador where cb_numero =" & data_arq.Recordset("cob")
               data_cob.Refresh
               If data_cob.Recordset.RecordCount > 0 Then
                  data_infarq.Recordset("nomcob") = Mid(data_cob.Recordset("cb_nombre"), 1, 35)
               Else
                  data_infarq.Recordset("nomcob") = data_arq.Recordset("nomcob")
               End If
            Else
               data_infarq.Recordset("nomcob") = Mid(labcob.Caption, 1, 35)
            End If
         Else
            data_infarq.Recordset("nomcob") = ""
         End If
         data_infarq.Recordset("total") = data_arq.Recordset("total")
         data_infarq.Recordset.Update
         data_arq.Recordset.MoveNext
      Loop
   End If
   data_infarq.RecordSource = "Select * from infarq order by ano,mes"
   data_infarq.Refresh
   
   If Check2.Value = 1 Then
      CrystalReport1.ReportFileName = App.path & "\infpendd.rpt"
   Else
      CrystalReport1.ReportFileName = App.path & "\infpend.rpt"
   End If
   frm_infarq.MousePointer = 0
   b_acep.Enabled = True
   CrystalReport1.ReportTitle = "INFORME DE FACTURAS NO INGRESADAS EN ARQUEO " & txt_cob.Text & " " & labcob.Caption & " MES...:" & txt_mes.Text & "/" & txt_ano.Text
   CrystalReport1.Action = 1
End If
frm_infarq.MousePointer = 0

End Sub

Private Sub Command3_Click()
Dim Xdesde, Xhasta As String
Dim Xdia As Integer
Dim Nomlaemi As String
Dim SioNosj As String

Nomlaemi = "emi"
Dim Nombrearq As String
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
On Error GoTo Quepasa3

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\infdeuemi.mdb")

MiBaseact.Execute "Delete * from emision"
data_infdeu.RecordSource = "emision"
data_infdeu.Refresh

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
data_infcli.RecordSource = "infcli"
data_infcli.Refresh

If Option10.Value = True Then
   If txt_cob.Text <> 0 Then
      frm_infarq.MousePointer = 0
      MsgBox "El informe se emite para todos los cobradores", vbInformation
   End If
   frm_infarq.MousePointer = 11
    If txt_mes.Text <> "" And txt_ano.Text <> "" Then
       If txt_mes.Text > 9 Then
          Nomlaemi = Trim(Nomlaemi) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
       Else
          Nomlaemi = Trim(Nomlaemi) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
       End If
       Xdia = Day(DateSerial(txt_ano.Text, txt_mes.Text + 1, 0))
       If Val(txt_mes.Text) > 9 Then
          Xdesde = "01/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
          If Xdia > 9 Then
             Xhasta = Trim(str(Xdia)) & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
          Else
             Xhasta = "0" & Trim(str(Xdia)) & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
          End If
       Else
          Xdesde = "01/" & "0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
          If Xdia > 9 Then
             Xhasta = Trim(str(Xdia)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
          Else
             Xhasta = "0" & Trim(str(Xdia)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
          End If
       End If
       
       data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(Xdesde, "yyyy-mm-dd") & "' and fecha <='" & Format(Xhasta, "yyyy-mm-dd") & "' and cl_motivo ='" & "CAMBIO DE COBRADOR" & "'"
       data_abm.Refresh
       If data_abm.Recordset.RecordCount > 0 Then
          data_abm.Recordset.MoveFirst
          Do While Not data_abm.Recordset.EOF
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_abm.Recordset("cl_codigo")
             data_cli.Refresh
             If data_cli.Recordset.RecordCount > 0 Then
                If data_cli.Recordset("cl_nrocobr") > 0 Then
                   data_infcli.Recordset.AddNew
                   data_infcli.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                   data_infcli.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                   data_infcli.Recordset("cl_fecing") = data_abm.Recordset("fecha")
                   data_infcli.Recordset("cl_nombre") = data_abm.Recordset("usuario")
                   data_infcli.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                   data_infcli.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_infcli.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                   data_infcli.Recordset("cl_ultmesp") = txt_mes.Text
                   data_infcli.Recordset("cl_ultanop") = txt_ano.Text
                   data_infcli.Recordset("cl_atrasoa") = data_cli.Recordset("cl_atrasoa")
        '            data_arq.RecordSource = "Select * from arqueo where arqueo ='" & Text1.Text & "' order by cob"
        '            data_arq.Refresh
        '            data_arq.Recordset.MoveFirst
                   data_emi.RecordSource = "Select * from " & Nomlaemi & " where cliente =" & data_cli.Recordset("cl_codigo")
                   data_emi.Refresh
                   If data_emi.Recordset.RecordCount > 0 Then
                      data_emi.Recordset.MoveFirst
                      data_infcli.Recordset("cl_nrovend") = data_emi.Recordset("nro_cobr")
                   Else
                      data_infcli.Recordset("cl_nrovend") = 0
                   End If
                   data_infcli.Recordset.Update
                End If
             End If
             data_abm.Recordset.MoveNext
          Loop
          data_infcli.Refresh
          frm_infarq.MousePointer = 0
          b_acep.Enabled = True
          CrystalReport1.ReportFileName = App.path & "\infcambioscob.rpt"
          CrystalReport1.Action = 1
          
       End If
    End If
End If

If Option12.Value = True Then
   frm_infarq.MousePointer = 11
   If txt_mes.Text <> "" And txt_ano.Text <> "" Then
      If txt_mes.Text > 9 Then
         Nomlaemi = Trim(Nomlaemi) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
      Else
         Nomlaemi = Trim(Nomlaemi) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
      End If
      Xdia = Day(DateSerial(txt_ano.Text, txt_mes.Text + 1, 0))
      If Val(txt_mes.Text) > 9 Then
         Xdesde = "01/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
         If Xdia > 9 Then
            Xhasta = Trim(str(Xdia)) & "/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
         Else
            Xhasta = "0" & Trim(str(Xdia)) & "/" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
         End If
      Else
         Xdesde = "01/" & "0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
         If Xdia > 9 Then
            Xhasta = Trim(str(Xdia)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
         Else
            Xhasta = "0" & Trim(str(Xdia)) & "/0" & Trim(txt_mes.Text) & "/" & Trim(txt_ano.Text)
         End If
      End If
      data_abm.RecordSource = "Select * from linmmdd where cod_prod =" & 999 & " and fecha >='" & Format(Xdesde, "yyyy-mm-dd") & "' and fecha <='" & Format(Xhasta, "yyyy-mm-dd") & "' and " & _
      "grupo in (" & txt_cob.Text & ")"
      data_abm.Refresh
      If data_abm.Recordset.RecordCount > 0 Then
         data_abm.Recordset.MoveFirst
         Do While Not data_abm.Recordset.EOF
            If data_abm.Recordset("grupo") = 615 Or _
               data_abm.Recordset("grupo") = 616 Or _
               data_abm.Recordset("grupo") = 635 Or _
               data_abm.Recordset("grupo") = 602 Or _
               data_abm.Recordset("grupo") = 113 Or _
               data_abm.Recordset("grupo") = 653 Or _
               data_abm.Recordset("grupo") = 672 Or _
               data_abm.Recordset("grupo") = 1 Or _
               data_abm.Recordset("grupo") = 10 Or _
               data_abm.Recordset("grupo") = 201 Or _
               data_abm.Recordset("grupo") = 512 Or _
               data_abm.Recordset("grupo") = 636 Or _
               data_abm.Recordset("grupo") = 685 Or _
               data_abm.Recordset("grupo") = 208 Or _
               data_abm.Recordset("grupo") = 209 Or _
               data_abm.Recordset("grupo") = 8 Or _
               data_abm.Recordset("grupo") = 0 Or data_abm.Recordset("grupo") = 703 Then
            Else
               data_infcli.Recordset.AddNew
               data_infcli.Recordset("cl_fecing") = data_abm.Recordset("fecha")
               data_infcli.Recordset("cl_codigo") = data_abm.Recordset("cod_cli")
               data_infcli.Recordset("cl_apellid") = data_abm.Recordset("nom_cli")
               data_infcli.Recordset("saldo_cc") = data_abm.Recordset("tot_lin")
               data_infcli.Recordset("cl_nrocobr") = data_abm.Recordset("grupo")
               data_infcli.Recordset("cl_nrovend") = data_abm.Recordset("base")
               data_infcli.Recordset("cl_codconv") = data_abm.Recordset("convenio")
               data_infcli.Recordset("cl_ultmesp") = data_abm.Recordset("mes_paga")
               data_infcli.Recordset("cl_ultanop") = data_abm.Recordset("ano_paga")
               data_infcli.Recordset.Update
            End If
            data_abm.Recordset.MoveNext
         Loop
'         data_infcli.Refresh
'         MiBaseact.Execute "Delete * from infcli where cl_nrocobr <>" & txt_cob.Text
         frm_infarq.MousePointer = 0
         b_acep.Enabled = True
         data_infcli.RecordSource = "Select * from infcli order by cl_nrocobr"
         data_infcli.Refresh
         data_infcli.Recordset.MoveFirst
         cr2.ReportFileName = App.path & "\infcobbase.rpt"
         cr2.ReportTitle = "Informe de cobranza en Base desde:" & Xdesde & " hasta:" & Xhasta
         cr2.Action = 1
      Else
         frm_infarq.MousePointer = 0
         MsgBox "No hay registros"
      End If
   End If
End If
If Option13.Value = True Then
   Dim Xmatcob As Long
   Dim Xcantrec As Integer
   
   Nombrearq = "arq"
'   data_arq.DatabaseName = ""
'   data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
   If txt_mes.Text > 9 Then
      Nombrearq = Nombrearq & Trim(str(txt_mes.Text)) & Mid(Trim(str(txt_ano.Text)), 3, 2)
   Else
      Nombrearq = Nombrearq & "0" & Trim(str(txt_mes.Text)) & Mid(Trim(str(txt_ano.Text)), 3, 2)
   End If
   b_acep.Enabled = False
   If txt_cob.Text <> "" Then
      frm_infarq.MousePointer = 11
      MiBaseact.Execute "Delete * from infarq"
      data_infarq.Refresh
      If Check3.Value = 1 Then
         If Check1.Value = 1 Then
            If txt_cob.Text = 0 Then
               data_arq.RecordSource = "Select * from " & Nombrearq & " where arqueo ='" & "P" & "' and codzon =" & 100 & " order by matricula,ano,mes"
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Nombrearq & " where cob =" & txt_cob.Text & " and arqueo ='" & "P" & "' and codzon =" & 100 & " order by matricula,ano,mes"
               data_arq.Refresh
            End If
         Else
            If txt_cob.Text = 0 Then
               data_arq.RecordSource = "Select * from " & Nombrearq & " where arqueo ='" & "P" & "' order by matricula,ano,mes"
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from " & Nombrearq & " where cob =" & txt_cob.Text & " and arqueo ='" & "P" & "' order by matricula,ano,mes"
               data_arq.Refresh
            End If
         End If
      Else
         If Check1.Value = 1 Then
            If txt_cob.Text = 0 Then
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & "P" & "' and codzon =" & 100 & " order by matricula,ano,mes"
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & "P" & "' and codzon =" & 100 & " order by matricula,ano,mes"
               data_arq.Refresh
            End If
         Else
            If txt_cob.Text = 0 Then
               data_arq.RecordSource = "Select * from arqueo where arqueo ='" & "P" & "' order by matricula,ano,mes"
               data_arq.Refresh
            Else
               data_arq.RecordSource = "Select * from arqueo where cob =" & txt_cob.Text & " and arqueo ='" & "P" & "' order by matricula,ano,mes"
               data_arq.Refresh
            End If
         End If
      End If
      If data_arq.Recordset.RecordCount > 0 Then
         data_arq.Recordset.MoveFirst
         Xmatcob = data_arq.Recordset("matricula")
         Xcantrec = 0
         Do While Not data_arq.Recordset.EOF
            If Xmatcob = data_arq.Recordset("matricula") Then
               Xcantrec = Xcantrec + 1
               Xmatcob = data_arq.Recordset("matricula")
               data_arq.Recordset.MoveNext
            Else
               data_arq.Recordset.MovePrevious
               If Xcantrec >= 3 Then
                  If Check3.Value = 1 Then
                     data_abm.RecordSource = "Select * from " & Nombrearq & " where matricula =" & data_arq.Recordset("matricula")
                     data_abm.Refresh
                  Else
                     data_abm.RecordSource = "Select * from arqueo where matricula =" & data_arq.Recordset("matricula")
                     data_abm.Refresh
                  End If
                  If data_abm.Recordset.RecordCount > 0 Then
                     data_abm.Recordset.MoveFirst
                     Do While Not data_abm.Recordset.EOF
                        data_infarq.Recordset.AddNew
                        data_infarq.Recordset("matricula") = data_abm.Recordset("matricula")
                        data_infarq.Recordset("nrorec") = data_abm.Recordset("nrorec")
                        data_infarq.Recordset("color") = data_abm.Recordset("color")
                        data_infarq.Recordset("mes") = data_abm.Recordset("mes")
                        data_infarq.Recordset("ano") = data_abm.Recordset("ano")
                        data_infarq.Recordset("nombre") = Mid(data_abm.Recordset("nombre"), 1, 30)
                        data_infarq.Recordset("cob") = data_abm.Recordset("cob")
                        data_infarq.Recordset("nomcob") = data_abm.Recordset("nomcob")
                        data_infarq.Recordset("total") = data_abm.Recordset("total")
                        data_infarq.Recordset.Update
                        data_abm.Recordset.MoveNext
                     Loop
                  End If
               End If
               Xcantrec = 0
               data_arq.Recordset.MoveNext
               Xmatcob = data_arq.Recordset("matricula")
            End If
         Loop
         b_acep.Enabled = True
         frm_infarq.MousePointer = 0
         MsgBox "Proceso terminado"

         cr2.ReportFileName = App.path & "\infcobdevol.rpt"
         cr2.ReportTitle = "Informe de facturas a entregar por el cobrador MES:" & txt_mes.Text & "/" & txt_ano.Text
         cr2.Action = 1
         
      Else
         MsgBox "No hay registros"
      End If
   Else
      MsgBox "No ingresó cobrador"
   End If
End If

If Option15.Value = True Then
   Dim xmm, Xaa As Integer
   If txt_cob.Text = "" Then
      txt_cob.Text = 0
   End If
   If txt_mes.Text <> "" And txt_ano.Text <> "" Then
      SioNosj = MsgBox("Desea incluir la cobranza de San Jacinto?", vbInformation + vbYesNo)
      
       frm_infarq.MousePointer = 11
      
      If txt_mes.Text = 12 Then
         xmm = 1
         Xaa = txt_ano.Text + 1
      Else
         xmm = txt_mes.Text + 1
         Xaa = txt_ano.Text
      End If
      If xmm > 9 Then
         Nomlaemi = Trim(Nomlaemi) + Trim(str(xmm)) + Trim(Mid(Xaa, 3, 2))
      Else
         Nomlaemi = Trim(Nomlaemi) + Trim("0") + Trim(str(xmm)) + Trim(Mid(Xaa, 3, 2))
      End If
      If txt_cob.Text = 0 Then
         If SioNosj = vbYes Then
            data_emi.RecordSource = "Select * from " & Nomlaemi & " where total >" & 0
            data_emi.Refresh
         Else
            data_emi.RecordSource = "Select * from " & Nomlaemi & " where total >" & 0 & " and nro_cobr not in (6,5,11)"
            data_emi.Refresh
         End If
      Else
         data_emi.RecordSource = "Select * from " & Nomlaemi & " where nro_cobr =" & txt_cob.Text & " and total >" & 0
         data_emi.Refresh
      End If
      If data_emi.Recordset.RecordCount > 0 Then
         data_emi.Recordset.MoveFirst
         Do While Not data_emi.Recordset.EOF
            data_infdeu.Recordset.AddNew
            data_infdeu.Recordset("cliente") = data_emi.Recordset("cliente")
            data_infdeu.Recordset("apellidos") = data_emi.Recordset("apellidos")
            data_infdeu.Recordset("cod_cnv") = data_emi.Recordset("cod_cnv")
            data_infdeu.Recordset("tel_cli") = data_emi.Recordset("tel_cli")
            data_infdeu.Recordset("nro_vende") = data_emi.Recordset("nro_vende")
            data_infdeu.Recordset("grupo") = data_emi.Recordset("grupo")
            data_infdeu.Recordset("nro_cobr") = data_emi.Recordset("nro_cobr")
            data_infdeu.Recordset("fecha_ing") = data_emi.Recordset("fecha_ing")
            data_infdeu.Recordset("total") = data_emi.Recordset("total")
            data_abm.RecordSource = "select * from deudas where cliente =" & data_emi.Recordset("cliente") & " and mes>" & 0 & " and fecha_pago is null order by ano,mes"
            data_abm.Refresh
            If data_abm.Recordset.RecordCount > 0 Then
               If data_abm.Recordset("mes") = 1 Then
                  data_infdeu.Recordset("ult_mesp") = 12
                  data_infdeu.Recordset("ult_aniop") = data_abm.Recordset("ano") - 1
               Else
                  data_infdeu.Recordset("ult_mesp") = data_abm.Recordset("mes") - 1
                  data_infdeu.Recordset("ult_aniop") = data_abm.Recordset("ano")
               End If
               data_abm.Recordset.MoveLast
               If data_abm.Recordset.RecordCount = 1 Then
                  data_infdeu.Recordset("detalle") = "AL DÍA"
               Else
                  data_infdeu.Recordset("detalle") = "DEUDA " & data_abm.Recordset.RecordCount & " MESES"
               End If
            Else
               data_infdeu.Recordset("ult_mesp") = xmm
               data_infdeu.Recordset("ult_aniop") = Xaa
               data_infdeu.Recordset("detalle") = "AL DÍA"
            End If
            data_infdeu.Recordset.Update
            data_emi.Recordset.MoveNext
         Loop
      End If
      frm_infarq.MousePointer = 0
      data_infdeu.RecordSource = "select * from emision"
      data_infdeu.Refresh
      If Check2.Value = 1 Then
         cr2.ReportFileName = App.path & "\infdeuemi.rpt"
      Else
         cr2.ReportFileName = App.path & "\infdeuemin.rpt"
      End If
      cr2.ReportTitle = "Informe de atrasos emisión actual según arqueo de: " & txt_mes.Text & "/" & txt_ano.Text
      cr2.Action = 1
      MsgBox "Proceso terminado"
      
   Else
      MsgBox "No ingresó MES/AÑO"
   End If

End If

Exit Sub

Quepasa3:
         If Err.Number = 53 Then
            MsgBox "Archivo no encontrado"
         Else
            MsgBox "Verifique datos ingresados " & Err.Number & " " & Err.Description
         End If
         
End Sub

Private Sub Form_Load()
'data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.ConnectionString = "dsn=" & Xconexrmt
'data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_arq.ConnectionString = "dsn=" & Xconexrmt
data_infarq.DatabaseName = App.path & "\informes.mdb"
'data_infarq.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"

'data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_emi.ConnectionString = "dsn=" & Xconexrmt
'data_infresum.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_infresum.DatabaseName = App.path & "\informes.mdb"
'data_ent.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ent.ConnectionString = "dsn=" & Xconexrmt
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
'data_abm.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_abm.ConnectionString = "dsn=" & Xconexrmt
data_infcli.DatabaseName = App.path & "\informes.mdb"
'data_infcli.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_cobotra.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_inflin.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\informes.mdb"
data_arqloc.DatabaseName = App.path & "\informes.mdb"
data_infdeu.DatabaseName = App.path & "\infdeuemi.mdb"

'data_inflin.DatabaseName = App.Path & "\informes.mdb"
'data_inflin.RecordSource = "infvtas"
'data_inflin.Refresh

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub mf_GotFocus()
Dim Nomem3 As String

Nomem3 = "emi"
If txt_mes.Text > 9 Then
   Nomem3 = Trim(Nomem3) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
Else
   Nomem3 = Trim(Nomem3) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
End If
data_emi.RecordSource = "Select * from " & Nomem3 & " limit 50"
data_emi.Refresh
If data_emi.Recordset.RecordCount > 0 Then
   data_emi.Recordset.MoveFirst
   If IsNull(data_emi.Recordset("fecha")) = False Then
      mf.Text = data_emi.Recordset("fecha") + 1
   Else
      mf.Text = Date
   End If
Else
   mf.Text = Date
End If

End Sub

Private Sub mf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub

Private Sub mff_GotFocus()
Dim Nomem3 As String

Nomem3 = "emi"
If txt_mes.Text > 9 Then
   Nomem3 = Trim(Nomem3) + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
Else
   Nomem3 = Trim(Nomem3) + Trim("0") + Trim(str(txt_mes.Text)) + Trim(Mid(txt_ano.Text, 3, 2))
End If
data_emi.RecordSource = "Select * from " & Nomem3 & " limit 10"
data_emi.Refresh
If data_emi.Recordset.RecordCount > 0 Then
   data_emi.Recordset.MoveFirst
   If IsNull(data_emi.Recordset("fecha")) = False Then
      mff.Text = data_emi.Recordset("fecha") + 1
   Else
      mff.Text = Date
   End If
Else
   mff.Text = Date
End If

End Sub

Private Sub Option1_Click()
mf.Visible = False
Label2.Visible = False

End Sub

Private Sub Option2_Click()
mf.Visible = False
Label2.Visible = False

End Sub

Private Sub Option3_Click()
mf.Visible = False
Label2.Visible = False

End Sub

Private Sub Option4_Click()
mf.Visible = False
Label2.Visible = False

End Sub

Private Sub Option5_Click()
mf.Visible = False
Label2.Visible = False

End Sub

Private Sub Option6_Click()
mf.Visible = True
Label2.Visible = True
mf.SetFocus

End Sub

Private Sub Option7_Click()
mf.Visible = False
Label2.Visible = False

End Sub

Private Sub Option8_Click()

mf.Visible = False
mff.Visible = True
Label2.Visible = False


End Sub

Private Sub txt_ano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub

Private Sub txt_cob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mes.SetFocus
End If

End Sub

Private Sub txt_cob_LostFocus()
If txt_cob.Text <> "" Then
'   data_cob.Recordset.FindFirst "cb_numero =" & txt_cob.Text
   data_cob.RecordSource = "Select * from cobrador where cb_numero =" & txt_cob.Text
   data_cob.Refresh
   If data_cob.Recordset.RecordCount > 0 Then
      labcob.Caption = data_cob.Recordset("cb_nombre")
   Else
      MsgBox "No encontrado, verifique", vbCritical, "Mensaje"
      txt_cob.SetFocus
   End If
End If
   
End Sub

Private Sub txt_mes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ano.SetFocus
End If

End Sub
