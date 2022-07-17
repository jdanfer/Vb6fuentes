VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_convenios 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convenios"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12720
   Icon            =   "frm_convenios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12720
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_inftiquet 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_convenios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Informes de llamados con costo de AP"
      Top             =   7920
      Width           =   495
   End
   Begin VB.Data data_inftiq 
      Caption         =   "data_inftiq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin Crystal.CrystalReport crtiq 
      Left            =   10560
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox t_email 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   200
      TabIndex        =   63
      Top             =   2400
      Width           =   8655
   End
   Begin VB.CheckBox choculta 
      BackColor       =   &H00C00000&
      Caption         =   "Mantener oculto"
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
      Left            =   4320
      TabIndex        =   61
      Top             =   120
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Adodc1"
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
   Begin VB.Data data_abmconv 
      Caption         =   "data_abmconv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox t_dpto 
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
      Left            =   6360
      TabIndex        =   57
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4920
      TabIndex        =   55
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton b_hist 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Historial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      MouseIcon       =   "frm_convenios.frx":09CC
      MousePointer    =   99  'Custom
      Picture         =   "frm_convenios.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton b_fact 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      MouseIcon       =   "frm_convenios.frx":1118
      MousePointer    =   99  'Custom
      Picture         =   "frm_convenios.frx":1422
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox t_razon 
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
      Height          =   285
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   44
      Top             =   960
      Width           =   8655
   End
   Begin VB.TextBox txt_ruc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   42
      Top             =   2040
      Width           =   4695
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7800
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_cnvbusca 
      Caption         =   "data_cnvbusca"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton bbusca 
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
      Left            =   3480
      Picture         =   "frm_convenios.frx":1864
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Buscar datos"
      Top             =   7920
      Width           =   495
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton bimp 
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
      Left            =   4320
      Picture         =   "frm_convenios.frx":1DEE
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Informes"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   2640
      Picture         =   "frm_convenios.frx":2378
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Cancelar acción"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton bmodi 
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
      Left            =   1800
      Picture         =   "frm_convenios.frx":2902
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Editar"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   960
      Picture         =   "frm_convenios.frx":2E8C
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Grabar datos"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton balta 
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
      Left            =   120
      Picture         =   "frm_convenios.frx":3416
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Nuevo registro"
      Top             =   7920
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Otros datos administrativos"
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
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   12375
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF0000&
         Caption         =   "Lleva timbre al CONVENIO"
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
         Left            =   8400
         TabIndex        =   76
         Top             =   240
         Width           =   3255
      End
      Begin VB.CheckBox chafilia 
         BackColor       =   &H00FF0000&
         Caption         =   "Habilitado para afiliaciones"
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
         Left            =   8760
         TabIndex        =   67
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Data data_aranc 
         Caption         =   "data_aranc"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   65
         Text            =   "Combo1"
         Top             =   2040
         Width           =   4575
      End
      Begin MSAdodcLib.Adodc data_prec 
         Height          =   495
         Left            =   6960
         Top             =   600
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
         Caption         =   "data_prec"
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
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox t_nrocompra 
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
         Left            =   6240
         TabIndex        =   60
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox t_rub 
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
         Left            =   1920
         TabIndex        =   54
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox t_der 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   600
         Width           =   9735
      End
      Begin VB.TextBox txt_obs 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2520
         Width           =   9735
      End
      Begin MSMask.MaskEdBox fbaja 
         Height          =   375
         Left            =   10200
         TabIndex        =   30
         ToolTipText     =   "Al registrar fecha de baja el convenio no se mostrará en ninguna búsqueda"
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   8454143
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbomut 
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
         ItemData        =   "frm_convenios.frx":39A0
         Left            =   1920
         List            =   "frm_convenios.frx":39DA
         TabIndex        =   26
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FF0000&
         Caption         =   "Grupo Aranceles:"
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
         Left            =   120
         TabIndex        =   64
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FF0000&
         Caption         =   "Nro. de compra:"
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
         Left            =   4320
         TabIndex        =   59
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FF0000&
         Caption         =   "Rubro contable:"
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
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FF0000&
         Caption         =   "Convenio:"
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
         Height          =   855
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FFFF&
         Caption         =   "Fecha Baja:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   29
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF0000&
         Caption         =   "Grupo Mutual:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para emisión:"
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
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   12375
      Begin VB.TextBox cbogrupoap 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9240
         MaxLength       =   25
         TabIndex        =   74
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cbomesanio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_convenios.frx":3A70
         Left            =   10080
         List            =   "frm_convenios.frx":3A7A
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox t_implla 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   8640
         TabIndex        =   70
         ToolTipText     =   "Costo de los llamados cuando pasan el tope de contrato sin costo (IVA inc)"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox t_cantlla 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7800
         TabIndex        =   69
         ToolTipText     =   "Cantidad de llamados sin costo al mes"
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chdeuda 
         BackColor       =   &H00C00000&
         Caption         =   "Sin control de deuda"
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
         Left            =   7560
         TabIndex        =   66
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chtimbre 
         BackColor       =   &H00FF0000&
         Caption         =   "Lleva timbre al socio"
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
         Left            =   10080
         TabIndex        =   58
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cbovenc 
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
         ItemData        =   "frm_convenios.frx":3A8E
         Left            =   6720
         List            =   "frm_convenios.frx":3AA4
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cbofact 
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
         ItemData        =   "frm_convenios.frx":3AC1
         Left            =   5400
         List            =   "frm_convenios.frx":3ACB
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboaltasi 
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
         ItemData        =   "frm_convenios.frx":3AD7
         Left            =   6960
         List            =   "frm_convenios.frx":3AE1
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_cuenta 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cbosirec 
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
         ItemData        =   "frm_convenios.frx":3AED
         Left            =   4320
         List            =   "frm_convenios.frx":3AF7
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optodos 
         BackColor       =   &H00FF0000&
         Caption         =   "Todos los recibos"
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
         Height          =   375
         Left            =   2640
         TabIndex        =   22
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton opunosolo 
         BackColor       =   &H00FF0000&
         Caption         =   "Un solo recibo"
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
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txt_precio 
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
         Height          =   405
         Left            =   2760
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cbocolrec 
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
         ItemData        =   "frm_convenios.frx":3B03
         Left            =   1200
         List            =   "frm_convenios.frx":3B19
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cbomon 
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
         ItemData        =   "frm_convenios.frx":3B4A
         Left            =   120
         List            =   "frm_convenios.frx":3B54
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox vhasta 
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
      Begin MSMask.MaskEdBox vdesde 
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
      Begin VB.Label Label28 
         BackColor       =   &H00FF0000&
         Caption         =   "Grupo AP?"
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
         Left            =   8040
         TabIndex        =   73
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FF0000&
         Caption         =   "Mensual o Anual ?"
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
         Left            =   10080
         TabIndex        =   71
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FF0000&
         Caption         =   "Cant.Llamados y $$$"
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
         Left            =   7800
         TabIndex        =   68
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FF0000&
         Caption         =   "VENCE"
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
         Left            =   6720
         TabIndex        =   51
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FF0000&
         Caption         =   "FACTURA?"
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
         Left            =   5400
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF0000&
         Caption         =   "Permite ALTAS?"
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
         Height          =   375
         Left            =   5040
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF0000&
         Caption         =   "Nro.Cuenta"
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
         Left            =   4440
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF0000&
         Caption         =   "Emite?"
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
         Left            =   4320
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         Caption         =   "Importe:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         MouseIcon       =   "frm_convenios.frx":3B61
         MousePointer    =   99  'Custom
         TabIndex        =   18
         ToolTipText     =   "Haga click para ver los precios anteriores"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "COLOR:"
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
         Left            =   1200
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "Moneda"
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
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Vigencia:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox txt_tel 
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
      Height          =   285
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txt_localid 
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
      Height          =   285
      Left            =   1680
      MaxLength       =   45
      TabIndex        =   7
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txt_direc 
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
      Height          =   285
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1320
      Width           =   8655
   End
   Begin VB.TextBox txt_desc 
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
      Height          =   285
      Left            =   1680
      MaxLength       =   65
      TabIndex        =   3
      Top             =   600
      Width           =   8655
   End
   Begin VB.TextBox txt_cod 
      Enabled         =   0   'False
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FF0000&
      Caption         =   "E-mail:"
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
      Left            =   120
      TabIndex        =   62
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FF0000&
      Caption         =   "Dpto:"
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
      Left            =   4800
      TabIndex        =   56
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FF0000&
      Caption         =   "Razón social:"
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
      Left            =   120
      TabIndex        =   43
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FF0000&
      Caption         =   "RUT:"
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
      Left            =   4800
      TabIndex        =   41
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Localidad:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Dirección:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "CODIGO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   2910
      Left            =   6000
      Picture         =   "frm_convenios.frx":40EB
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   3885
   End
End
Attribute VB_Name = "frm_convenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_fact_Click()
Dim Xdig, Xrut, Xtot, Xtot2, Xfactor, i As Integer
Dim Xelrut As String
Xelrut = ""
Xelrut = frm_convenios.txt_ruc.Text

If WElusuario = "MCOSTA" Or WElusuario = "JFERNAN" Or WElusuario = "MPEREZ" Or XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "ADMINISTRADOR" Then
   If Xelrut <> "" Then
      If Len(Trim(Xelrut)) = 12 Then
         If IsNumeric(Xelrut) Then
            Xdig = Val(Mid(Xelrut, 12, 1))
            Xrut = Val(Mid(Xelrut, 1, 12))
            Xtot = 0
            Xfactor = 2
            For i = 1 To 11
                If i = 1 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 4
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 2 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 3
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 3 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 2
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 4 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 9
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 5 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 8
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 6 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 7
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 7 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 6
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 8 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 5
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 9 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 4
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 10 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 3
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 11 Then
                   Xtot = Val(Mid(Xelrut, i, 1)) * 2
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
            If Xdig = Val(Mid(Xelrut, 12, 1)) Then
               If cbofact.ListIndex = 1 Then
                  XcomoFactura = 1
                  frm_quefactcnv22.Show vbModal
               Else
                  MsgBox "Convenio no se realiza factura, VERIFIQUE!", vbCritical
               End If
            Else
               MsgBox "El RUT ingresado NO ES CORRECTO!"
            End If
         Else
            MsgBox "El RUT ingresado debe contener solo números"
         End If
      Else
         MsgBox "La cantidad de dígitos del RUT no es correcta"
      End If
   Else
      If cbofact.ListIndex = 1 Then
         If txt_cuenta.Text <> "" Then
            XcomoFactura = 2
            frm_quefactcnv22.Show vbModal
         Else
            MsgBox "Debe figurar número de cuenta para poder facturar"
         End If
      Else
         MsgBox "Convenio no se realiza factura, VERIFIQUE!", vbCritical
      End If
'      MsgBox "El campo del RUT está vacío"
   End If
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_hist_Click()
If WElusuario = "MCOSTA" Or WElusuario = "JFERNAN" Or WElusuario = "MPEREZ" Or XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Then
   If cbofact.ListIndex = 1 Then
      frm_estcnv.Show vbModal
   Else
      MsgBox "Convenio no se realiza factura, VERIFIQUE!", vbCritical
   End If
Else
   MsgBox "Usuario no autorizado"
End If
End Sub

Private Sub b_inftiquet_Click()

Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application
Dim Xfdesde, Xfhasta As String
Dim ListarGrupo As String

Xtotreg = 0
Xsub = 0

Dim Xlabrir3 As New Excel.Application

Xfdesde = InputBox("Ingrese fecha de inicio (Ej.01/01/2022)", "Informes")
Xfhasta = InputBox("Ingrese fecha de finalización (Ej.10/01/2022)", "Informes")

ListarGrupo = MsgBox("Desea listar el grupo seleccionado?", vbInformation + vbYesNo, "Convenios")
If listagrupo = vbYes Then
   If Trim(cbogrupoap.Text) <> "" Then
      If Trim(Xfdesde) <> "" And Trim(Xfhasta) <> "" Then
         data_inftiq.RecordSource = "Select * from convenio_tiquets where fecha >=#" & Format(Xfdesde, "yyyy-mm-dd") & "# And fecha <=#" & Format(Xfhasta, "yyyy-mm-dd") & "# and nom_grupo ='" & cbogrupoap.Text & "' order by fecha"
      Else
         data_inftiq.RecordSource = "Select * from convenio_tiquets where nom_grupo ='" & cbogrupoap.Text & "' order by fecha"
      End If
   Else
      If Trim(Xfdesde) <> "" And Trim(Xfhasta) <> "" Then
         data_inftiq.RecordSource = "Select * from convenio_tiquets where fecha >=#" & Format(Xfdesde, "yyyy-mm-dd") & "# And fecha <=#" & Format(Xfhasta, "yyyy-mm-dd") & "# and nom_grupo ='" & txt_cod.Text & "' order by fecha"
      Else
         data_inftiq.RecordSource = "Select * from convenio_tiquets where nom_grupo ='" & txt_cod.Text & "' order by fecha"
      End If
   End If
Else
   If Trim(Xfdesde) <> "" And Trim(Xfhasta) <> "" Then
      data_inftiq.RecordSource = "Select * from convenio_tiquets where fecha >=#" & Format(Xfdesde, "yyyy-mm-dd") & "# And fecha <=#" & Format(Xfhasta, "yyyy-mm-dd") & "# order by fecha"
   Else
      data_inftiq.RecordSource = "Select * from convenio_tiquets order by fecha"
   End If
End If
data_inftiq.Refresh
If data_inftiq.Recordset.RecordCount > 0 Then
   Xlin = 1
   XCol = 1
   Xtotreg = 0
   Xsub = 0
   Set Xobjexel22 = New Excel.Application
   Set Xlibexel22 = Xobjexel22.Workbooks.Add
   Set Xarchexel22 = Xlibexel22.Worksheets.Add
   Xarchexel22.Name = Trim("Controles")
   Xlibexel22.SaveAs ("C:\planillas\Llamados AP con costo.xls")
   Xarchtex = "C:\planillas\Llamados AP con costo.xls"

   Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
   Xlin = Xlin + 1
   XCol = XCol + 1
   Xarchexel22.Range("A1", "C3").Font.Size = 16
   Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
   Xarchexel22.Cells(Xlin, XCol) = "LLAMADOS CON COSTO DE A.P. DESDE: " & Xfdesde & " HASTA: " & Xfhasta
   XCol = 1
   Xlin = Xlin + 2
   Xnrocan = Xnrocan + Xlin
   Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
   Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
   Xarchexel22.Cells(Xlin, XCol) = "CATEG."
   XCol = XCol + 1
   Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 30
   Xarchexel22.Cells(Xlin, XCol) = "CONVENIO"
   XCol = XCol + 1
   Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "FECHA"
   XCol = XCol + 1
   Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
   Xarchexel22.Cells(Xlin, XCol) = "HORA"
   XCol = XCol + 1
   Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 30
   Xarchexel22.Cells(Xlin, XCol) = "NOMBRE PACIENTE"
   XCol = XCol + 1
   Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
   XCol = XCol + 1
   Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "FECHA_PAGO"
   XCol = XCol + 1
   Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "COMPROBANTE"
   XCol = XCol + 1
   Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "IMPORTE $."
   XCol = XCol + 1
   Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "GRUPO"
       
   Xlin = Xlin + 1
   XCol = 1
   data_inftiq.Recordset.MoveFirst
   Do While Not data_inftiq.Recordset.EOF
      Xarchexel22.Cells(Xlin, XCol) = data_inftiq.Recordset("id_convenio")
      XCol = XCol + 1
      data_cnvbusca.RecordSource = "select * from convenio where cnv_codigo ='" & data_inftiq.Recordset("id_convenio") & "'"
      data_cnvbusca.Refresh
      If data_cnvbusca.Recordset.RecordCount > 0 Then
         Xarchexel22.Cells(Xlin, XCol) = data_cnvbusca.Recordset("cnv_desc")
         XCol = XCol + 1
      Else
         Xarchexel22.Cells(Xlin, XCol) = "NO ENCONTRADO"
         XCol = XCol + 1
      End If
      Xarchexel22.Cells(Xlin, XCol) = "'" & data_inftiq.Recordset("fecha")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & data_inftiq.Recordset("hora")
      XCol = XCol + 1
      If IsNull(data_inftiq.Recordset("nombre")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_inftiq.Recordset("nombre")
      End If
      XCol = XCol + 1
      If IsNull(data_inftiq.Recordset("cedula")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_inftiq.Recordset("cedula")
      End If
      XCol = XCol + 1
      If IsNull(data_inftiq.Recordset("fecha_pago")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = "'" & data_inftiq.Recordset("fecha_pago")
      End If
      XCol = XCol + 1
      If IsNull(data_inftiq.Recordset("nro_doc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = "'" & data_inftiq.Recordset("nro_doc")
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inftiq.Recordset("importe")
      XCol = XCol + 1
      If IsNull(data_inftiq.Recordset("nom_grupo")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_inftiq.Recordset("nom_grupo")
      End If
      
      XCol = XCol + 1
      data_inftiq.Recordset.MoveNext
      XCol = 1
      Xlin = Xlin + 1
   Loop
   frm_prodmed.MousePointer = 0
   Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
   Xlibexel22.Save
   Xlibexel22.Close
   Xobjexel22.Quit
   Xlabrir.Workbooks.Open Xarchtex, , False
   Xlabrir.Visible = True
   Xlabrir.WindowState = xlMaximized
Else
   frm_prodmed.MousePointer = 0
   MsgBox "No hay registros para crear planilla."
End If

End Sub

Private Sub balta_Click()
If WElusuario = "JFERNAN" Or XWeltipoU = "ADMINISTRADOR" Or WElusuario = "MCOSTA" Or WElusuario = "MPEREZ" Or WElusuario = "GUSTAVO" Or XWeltipoU = "USUARIOS ADM" Then
    habilita
    borraconv
    txt_cod.SetFocus
    balta.Enabled = False
    bgraba.Enabled = True
    bmodi.Enabled = False
    bcance.Enabled = True
    bimp.Enabled = False
    bbusca.Enabled = False
    b_fact.Enabled = False
    b_hist.Enabled = False
    data_conv.Recordset.AddNew
    XAcnv = 1
Else
   MsgBox "Usuario no autorizado", vbCritical, "Convenios"
   Unload Me
End If

End Sub



Private Sub bbusca_Click()
frm_buscnvdos.Show vbModal

End Sub

Private Sub bcance_Click()
If XAcnv = 1 Then
   data_conv.Recordset.CancelUpdate
   borraconv
   deshab
   data_conv.Recordset.MoveLast
   igualar
   balta.Enabled = True
   bgraba.Enabled = False
   bmodi.Enabled = True
   bcance.Enabled = False
   bbusca.Enabled = True
   bimp.Enabled = True
   b_fact.Enabled = True
   b_hist.Enabled = True
   XAcnv = 0
Else
   borraconv
   deshab
   data_conv.Recordset.MoveLast
   igualar
   balta.Enabled = True
   bgraba.Enabled = False
   bmodi.Enabled = True
   bcance.Enabled = False
   bbusca.Enabled = True
   b_fact.Enabled = True
   b_hist.Enabled = True
   bimp.Enabled = True
   XAcnv = 0
End If

End Sub

Private Sub bgraba_Click()
Dim i, Xdig, Xrut, Xtot, Xfactor, Xtot2 As Integer
Dim Mifec As Date

If txt_ruc.Text <> "" Then
   If Len(Trim(txt_ruc.Text)) = 12 Then
         If IsNumeric(txt_ruc.Text) Then
            Xdig = Val(Mid(txt_ruc.Text, 12, 1))
            Xrut = Val(Mid(txt_ruc.Text, 1, 12))
            Xtot = 0
            Xfactor = 2
            For i = 1 To 11
                If i = 1 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 4
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 2 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 3
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 3 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 2
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 4 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 9
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 5 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 8
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 6 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 7
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 7 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 6
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 8 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 5
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 9 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 4
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 10 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 3
                   Xtot2 = Xtot2 + Xtot
                End If
                If i = 11 Then
                   Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 2
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
            If Xdig = Val(Mid(txt_ruc.Text, 12, 1)) Then
                If XAcnv = 1 Then
                   If txt_cod.Text <> "" And Combo1.Text <> "" Then
                      If txt_desc.Text <> "" Then
                         data_conv.Recordset("cnv_codigo") = txt_cod.Text
                         data_conv.Recordset("cnv_umpago") = choculta.Value
                         data_conv.Recordset("cnv_desc") = txt_desc.Text
                         data_conv.Recordset("cnv_entre") = t_razon.Text
                         data_conv.Recordset("cnv_direcc") = txt_direc.Text
                         data_conv.Recordset("cnv_local") = txt_localid.Text
                         data_conv.Recordset("cnv_tel") = txt_tel.Text
                         data_conv.Recordset("cnv_ruc") = txt_ruc.Text
                         data_conv.Recordset("cnv_sindeuda") = chdeuda.Value
                         data_conv.Recordset("cnv_gpoafilia") = chafilia.Value
                         If Trim(t_cantlla.Text) <> "" Then
                            data_conv.Recordset("cnv_cantcons") = t_cantlla.Text
                         End If
                         If Trim(t_implla.Text) <> "" Then
                            data_conv.Recordset("cnv_preccons") = t_implla.Text
                         End If
                         data_conv.Recordset("cnv_menanio") = cbomesanio.ListIndex
                         If Trim(cbogrupoap.Text) <> "" Then
                            data_conv.Recordset("cnv_grupoap") = cbogrupoap.Text
                         End If
                         If t_email.Text <> "" Then
                            data_conv.Recordset("cnv_correoe") = t_email.Text
                         End If
                         If t_rub.Text = "" Then
                            data_conv.Recordset("cnv_uapago") = 0
                         Else
                            data_conv.Recordset("cnv_uapago") = t_rub.Text
                         End If
                         If vdesde.Text <> "__/__/____" Then
                            data_conv.Recordset("cnv_desde") = Format(vdesde.Text, "dd/mm/yyyy")
                         End If
                         If vhasta.Text <> "__/__/____" Then
                            data_conv.Recordset("cnv_hasta") = Format(vhasta.Text, "dd/mm/yyyy")
                         End If
                         If cbomon.Text = "$U" Then
                            data_conv.Recordset("cnv_codmon") = 1
                            data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
                         Else
                            If cbomon.Text = "U$s" Then
                               data_conv.Recordset("cnv_codmon") = 2
                               data_conv.Recordset("cnv_nommon") = "DOLARES U.S.A."
                            Else
                               data_conv.Recordset("cnv_codmon") = 1
                               data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
                            End If
                         End If
                         If cbocolrec.Text = "ROJO" Then
                            data_conv.Recordset("cnv_colrec") = "R"
                         Else
                            If cbocolrec.Text = "AZUL" Then
                               data_conv.Recordset("cnv_colrec") = "A"
                            Else
                               If cbocolrec.Text = "MARRON" Then
                                  data_conv.Recordset("cnv_colrec") = "M"
                               Else
                                  If cbocolrec.Text = "VERDE" Then
                                     data_conv.Recordset("cnv_colrec") = "V"
                                  Else
                                     If cbocolrec.Text = "CELESTE" Then
                                        data_conv.Recordset("cnv_colrec") = "C"
                                     Else
                                        data_conv.Recordset("cnv_colrec") = ""
                                     End If
                                  End If
                               End If
                            End If
                         End If
                         If txt_precio.Text <> "" Then
                            data_conv.Recordset("cnv_precio") = txt_precio.Text
                         Else
                            data_conv.Recordset("cnv_precio") = 0
                         End If
                         If cbosirec.Text = "SI" Then
                            data_conv.Recordset("cnv_emite") = "SI"
                         Else
                            If cbosirec.Text = "NO" Then
                               data_conv.Recordset("cnv_emite") = "NO"
                            Else
                               data_conv.Recordset("cnv_emite") = "NO"
                            End If
                         End If
                         If txt_cuenta.Text <> "" Then
                            data_conv.Recordset("cnv_cuenta") = txt_cuenta.Text
                         Else
                            data_conv.Recordset("cnv_cuenta") = 0
                         End If
                         If cboaltasi.Text = "NO" Then
                            data_conv.Recordset("cnv_alta") = "NO"
                         Else
                            data_conv.Recordset("cnv_alta") = "SI"
                         End If
                         If opunosolo.Value = True Then
                            data_conv.Recordset("cnv_cant_r") = 1
                         Else
                            If optodos.Value = True Then
                               data_conv.Recordset("cnv_cant_r") = 2
                            Else
                               data_conv.Recordset("cnv_cant_r") = 1
                            End If
                         End If
                         If cbomut.Text <> "" Then
                            data_conv.Recordset("cnv_grupo") = cbomut.Text
                         Else
                            data_conv.Recordset("cnv_grupo") = ""
                         End If
                         If fbaja.Text <> "__/__/____" Then
                            data_conv.Recordset("cnv_fbaja") = Format(fbaja.Text, "dd/mm/yyyy")
                         Else
                
                         End If
                         If txt_obs.Text <> "" Then
                            data_conv.Recordset("cnv_ctrato") = txt_obs.Text
                         End If
                         data_conv.Recordset("cnv_pmserv") = cbofact.ListIndex
                         If t_der.Text <> "" Then
                            data_conv.Recordset("cnv_motbaj") = t_der.Text
                         Else
                            data_conv.Recordset("cnv_motbaj") = ""
                         End If
                         If cbovenc.ListIndex > 0 Then
                            data_conv.Recordset("cnv_paserv") = Val(cbovenc.Text)
                         Else
                            data_conv.Recordset("cnv_paserv") = 0
                         End If
                         data_conv.Recordset("cnv_sald") = chtimbre.Value
                         If t_nrocompra.Text <> "" Then
                            data_conv.Recordset("cnv_email") = t_nrocompra.Text
                         End If
                         If Combo1.Text <> "" Then
                            data_aranc.RecordSource = "Select * from Aran_grupos where desc_gpo ='" & Combo1.Text & "'"
                            data_aranc.Refresh
                            If data_aranc.Recordset.RecordCount > 0 Then
                               data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id")
                            Else
                               MsgBox "Atención: No se encuentra el grupo de arancel ingresado. Verifique!", vbCritical
                               data_conv.Recordset("cnv_aran") = 0
                            End If
                         Else
                            data_conv.Recordset("cnv_aran") = 0
                         End If
                         data_conv.Recordset.Update
                         
                         data_prec.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & txt_cod.Text & "' order by cnv_desde DESC"
                         data_prec.Refresh
                         If data_prec.Recordset.RecordCount > 0 Then
                            data_prec.Recordset("cnv_hasta") = Format(vhasta.Text, "yyyy-mm-dd")
'                            data_prec.Recordset("cnv_codigo") = data_conv.Recordset("cnv_codigo")
                            data_prec.Recordset("cnv_desde") = Format(vdesde.Text, "yyyy-mm-dd")
                            data_prec.Recordset("precio") = txt_precio.Text
                            data_prec.Recordset.Update
                         Else
                            data_prec.Recordset.AddNew
                            data_prec.Recordset("cnv_hasta") = Format(vhasta.Text, "yyyy-mm-dd")
                            data_prec.Recordset("cnv_codigo") = txt_cod.Text
                            data_prec.Recordset("cnv_desde") = Format(vdesde.Text, "yyyy-mm-dd")
                            data_prec.Recordset("precio") = txt_precio.Text
                            data_prec.Recordset("moneda") = 1
                            data_prec.Recordset("usuario") = WElusuario
                            data_prec.Recordset.Update
                            
                         End If
                         data_cnvbusca.Refresh
                         borraconv
                         deshab
                         igualar
                         balta.Enabled = True
                         bgraba.Enabled = False
                         bmodi.Enabled = True
                         bcance.Enabled = False
                         bbusca.Enabled = True
                         bimp.Enabled = True
                         b_fact.Enabled = True
                         b_hist.Enabled = True
                         XAcnv = 0
                      Else
                         MsgBox "No ingresó descripción", vbCritical, "Convenios"
                         txt_desc.SetFocus
                      End If
                   Else
                      MsgBox "Verifique si ingresó Grupo de Aranceles o verifique código del convenio.", vbCritical, "Convenios"
                      txt_desc.SetFocus
                   End If
                Else
                   If txt_cod.Text <> "" And Combo1.Text <> "" Then
                      If txt_desc.Text <> "" Then
                         data_conv.Recordset.Edit
                         data_conv.Recordset("cnv_codigo") = txt_cod.Text
                         data_conv.Recordset("cnv_umpago") = choculta.Value
                         data_conv.Recordset("cnv_desc") = txt_desc.Text
                         data_conv.Recordset("cnv_entre") = t_razon.Text
                         data_conv.Recordset("cnv_direcc") = txt_direc.Text
                         data_conv.Recordset("cnv_local") = txt_localid.Text
                         data_conv.Recordset("cnv_tel") = txt_tel.Text
                         data_conv.Recordset("cnv_ruc") = txt_ruc.Text
                         data_conv.Recordset("cnv_sindeuda") = chdeuda.Value
                         data_conv.Recordset("cnv_gpoafilia") = chafilia.Value
                         If Trim(t_cantlla.Text) <> "" Then
                            data_conv.Recordset("cnv_cantcons") = t_cantlla.Text
                         Else
                            If IsNull(data_conv.Recordset("cnv_cantcons")) = False Then
                               data_conv.Recordset("cnv_cantcons") = Null
                            End If
                         End If
                         If Trim(t_implla.Text) <> "" Then
                            data_conv.Recordset("cnv_preccons") = t_implla.Text
                         Else
                            If IsNull(data_conv.Recordset("cnv_preccons")) = False Then
                               data_conv.Recordset("cnv_preccons") = Null
                            End If
                         End If
                         data_conv.Recordset("cnv_menanio") = cbomesanio.ListIndex
                         If Trim(cbogrupoap.Text) <> "" Then
                            data_conv.Recordset("cnv_grupoap") = cbogrupoap.Text
                         Else
                            If IsNull(data_conv.Recordset("cnv_grupoap")) = False Then
                               data_conv.Recordset("cnv_grupoap") = Null
                            End If
                         End If
                         If t_email.Text <> "" Then
                            If IsNull(data_conv.Recordset("cnv_correoe")) = False Then
                               If t_email.Text <> data_conv.Recordset("cnv_correoe") Then
                                  data_conv.Recordset("cnv_correoe") = t_email.Text
                               End If
                            Else
                               data_conv.Recordset("cnv_correoe") = t_email.Text
                            End If
                         Else
                            If IsNull(data_conv.Recordset("cnv_correoe")) = False Then
                               data_conv.Recordset("cnv_correoe") = Null
                            End If
                         End If
                         If t_rub.Text = "" Then
                            data_conv.Recordset("cnv_uapago") = 0
                         Else
                            data_conv.Recordset("cnv_uapago") = t_rub.Text
                         End If
                         
                         If vdesde.Text <> "__/__/____" Then
                            data_conv.Recordset("cnv_desde") = Format(vdesde.Text, "dd/mm/yyyy")
                         End If
                         If vhasta.Text <> "__/__/____" Then
                            data_conv.Recordset("cnv_hasta") = Format(vhasta.Text, "dd/mm/yyyy")
                         End If
                         If cbomon.Text = "$U" Then
                            data_conv.Recordset("cnv_codmon") = 1
                            data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
                         Else
                            If cbomon.Text = "U$s" Then
                               data_conv.Recordset("cnv_codmon") = 2
                               data_conv.Recordset("cnv_nommon") = "DOLARES U.S.A."
                            Else
                               data_conv.Recordset("cnv_codmon") = 1
                               data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
                            End If
                         End If
                         If cbocolrec.Text = "ROJO" Then
                            data_conv.Recordset("cnv_colrec") = "R"
                         Else
                            If cbocolrec.Text = "AZUL" Then
                               data_conv.Recordset("cnv_colrec") = "A"
                            Else
                               If cbocolrec.Text = "MARRON" Then
                                  data_conv.Recordset("cnv_colrec") = "M"
                               Else
                                  If cbocolrec.Text = "VERDE" Then
                                     data_conv.Recordset("cnv_colrec") = "V"
                                  Else
                                     If cbocolrec.Text = "CELESTE" Then
                                        data_conv.Recordset("cnv_colrec") = "C"
                                     Else
                                        data_conv.Recordset("cnv_colrec") = Null
                                     End If
                                  End If
                               End If
                            End If
                         End If
                         
                         If txt_precio.Text <> "" Then
                            data_conv.Recordset("cnv_precio") = txt_precio.Text
                         Else
                            data_conv.Recordset("cnv_precio") = 0
                         End If
                         If cbosirec.Text = "SI" Then
                            data_conv.Recordset("cnv_emite") = "SI"
                         Else
                            If cbosirec.Text = "NO" Then
                               data_conv.Recordset("cnv_emite") = "NO"
                            Else
                               data_conv.Recordset("cnv_emite") = "NO"
                            End If
                         End If
                         If txt_cuenta.Text <> "" Then
                            data_conv.Recordset("cnv_cuenta") = txt_cuenta.Text
                         Else
                            data_conv.Recordset("cnv_cuenta") = 0
                         End If
                         If cboaltasi.Text = "NO" Then
                            data_conv.Recordset("cnv_alta") = "NO"
                         Else
                            data_conv.Recordset("cnv_alta") = "SI"
                         End If
                         If opunosolo.Value = True Then
                            data_conv.Recordset("cnv_cant_r") = 1
                         Else
                            If optodos.Value = True Then
                               data_conv.Recordset("cnv_cant_r") = 2
                            Else
                               data_conv.Recordset("cnv_cant_r") = 1
                            End If
                         End If
                         If cbomut.Text <> "" Then
                            data_conv.Recordset("cnv_grupo") = cbomut.Text
                         Else
                            data_conv.Recordset("cnv_grupo") = Null
                         End If
                         If fbaja.Text <> "__/__/____" Then
                            data_conv.Recordset("cnv_fbaja") = Format(fbaja.Text, "dd/mm/yyyy")
                         Else
                            If IsNull(data_conv.Recordset("cnv_fbaja")) = False Then
                               data_conv.Recordset("cnv_fbaja") = Null
                            End If
                         End If
                         If txt_obs.Text <> "" Then
                            data_conv.Recordset("cnv_ctrato") = txt_obs.Text
                         End If
                         data_conv.Recordset("cnv_pmserv") = cbofact.ListIndex
                         If t_der.Text <> "" Then
                            data_conv.Recordset("cnv_motbaj") = t_der.Text
                         Else
                            data_conv.Recordset("cnv_motbaj") = Null
                         End If
                         If cbovenc.ListIndex > 0 Then
                            data_conv.Recordset("cnv_paserv") = Val(cbovenc.Text)
                         Else
                            data_conv.Recordset("cnv_paserv") = 0
                         End If
                         data_conv.Recordset("cnv_sald") = chtimbre.Value
                         If t_nrocompra.Text <> "" Then
                            data_conv.Recordset("cnv_email") = t_nrocompra.Text
                         Else
                            If IsNull(data_conv.Recordset("cnv_email")) = False Then
                               data_conv.Recordset("cnv_email") = Null
                            End If
                         End If
                         If Combo1.Text <> "" Then
                            data_aranc.RecordSource = "Select * from Aran_grupos where desc_gpo ='" & Combo1.Text & "'"
                            data_aranc.Refresh
                            If data_aranc.Recordset.RecordCount > 0 Then
                               If IsNull(data_conv.Recordset("cnv_aran")) = False Then
                                  If data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id") Then
                                  Else
                                     data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id")
                                  End If
                               Else
                                  data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id")
                               End If
                            Else
                               MsgBox "Atención: No se encuentra el grupo de arancel ingresado. Verifique!", vbCritical
                               data_conv.Recordset("cnv_aran") = 0
                            End If
                         Else
                            data_conv.Recordset("cnv_aran") = 0
                         End If
                         data_conv.Recordset.Update
                         data_prec.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & txt_cod.Text & "' and cnv_hasta >='" & Format(Date, "yyyy-mm-dd") & "' order by cnv_desde DESC"
                         data_prec.Refresh
                         If data_prec.Recordset.RecordCount > 0 Then
                            data_prec.Recordset.MoveFirst
                            If data_prec.Recordset("precio") <> txt_precio.Text Then
                               data_prec.Recordset("cnv_hasta") = Date - 1
                               data_prec.Recordset.Update
                               data_prec.Recordset.AddNew
                               data_prec.Recordset("cnv_codigo") = txt_cod.Text
                               data_prec.Recordset("cnv_desde") = Format(Date, "yyyy-mm-dd")
                               Mifec = Date + 365
                               data_prec.Recordset("cnv_hasta") = Mifec
                               data_prec.Recordset("precio") = Format(txt_precio.Text, "Standard")
                               data_prec.Recordset("moneda") = 1
                               data_prec.Recordset("usuario") = WElusuario
                               data_prec.Recordset.Update
                            End If
                         Else
                            data_prec.Recordset.AddNew
                            data_prec.Recordset("cnv_hasta") = Format(vhasta.Text, "yyyy-mm-dd")
                            data_prec.Recordset("cnv_codigo") = txt_cod.Text
                            data_prec.Recordset("cnv_desde") = Format(vdesde.Text, "yyyy-mm-dd")
                            data_prec.Recordset("precio") = txt_precio.Text
                            data_prec.Recordset("moneda") = 1
                            data_prec.Recordset("usuario") = WElusuario
                            data_prec.Recordset.Update
                            
                         End If
                        
                        data_abmconv.Recordset.AddNew
                        data_abmconv.Recordset("cnv_codigo") = txt_cod.Text
                        data_abmconv.Recordset("cnv_desc") = txt_desc.Text
                        data_abmconv.Recordset("cnv_ruc") = txt_ruc.Text
                        If vdesde.Text <> "__/__/____" Then
                           data_abmconv.Recordset("cnv_desde") = Format(vdesde.Text, "dd/mm/yyyy")
                        End If
                        If vhasta.Text <> "__/__/____" Then
                           data_abmconv.Recordset("cnv_hasta") = Format(vhasta.Text, "dd/mm/yyyy")
                        End If
                        data_abmconv.Recordset("cnv_nommon") = Format(Time, "HH:MM:ss")
                        If cbomut.Text <> "" Then
                           data_abmconv.Recordset("cnv_grupo") = cbomut.Text
                        End If
                        data_abmconv.Recordset("cnv_direcc") = WElusuario
                        data_abmconv.Recordset("cnv_fbaja") = Format(Date, "dd/mm/yyyy")
                        data_abmconv.Recordset("cnv_precio") = frm_menu.data_parse.Recordset("base")
                        data_abmconv.Recordset("cnv_local") = "GRABA MOD"
                        data_abmconv.Recordset.Update
                         
                         data_cnvbusca.Refresh
                         borraconv
                         deshab
                         igualar
                         balta.Enabled = True
                         bgraba.Enabled = False
                         bmodi.Enabled = True
                         bcance.Enabled = False
                         bbusca.Enabled = True
                         bimp.Enabled = True
                         b_fact.Enabled = True
                         b_hist.Enabled = True
                         XAcnv = 0
                      Else
                         MsgBox "No ingresó descripción", vbCritical, "Convenios"
                         txt_desc.SetFocus
                      End If
                   Else
                      MsgBox "Verifique si ingresó Grupo de Aranceles o verifique código del convenio.", vbCritical, "Convenios"
                      txt_desc.SetFocus
                   End If
                End If
            Else
                MsgBox "Error en el RUT ingresado, modifique y vuelva a grabar"
            End If
         Else
            MsgBox "Hay un error en el RUT ingresado", vbCritical
         End If
   Else
       MsgBox "Error en el RUT ingresado"
   End If
Else
    If XAcnv = 1 Then
       If txt_cod.Text <> "" And Combo1.Text <> "" Then
          If txt_desc.Text <> "" Then
             data_conv.Recordset("cnv_codigo") = txt_cod.Text
             data_conv.Recordset("cnv_desc") = txt_desc.Text
             data_conv.Recordset("cnv_entre") = t_razon.Text
             data_conv.Recordset("cnv_direcc") = txt_direc.Text
             data_conv.Recordset("cnv_local") = txt_localid.Text
             data_conv.Recordset("cnv_tel") = txt_tel.Text
             data_conv.Recordset("cnv_ruc") = txt_ruc.Text
             data_conv.Recordset("cnv_umpago") = choculta.Value
             data_conv.Recordset("cnv_sindeuda") = chdeuda.Value
             data_conv.Recordset("cnv_gpoafilia") = chafilia.Value
             If Trim(t_cantlla.Text) <> "" Then
                data_conv.Recordset("cnv_cantcons") = t_cantlla.Text
             End If
             If Trim(t_implla.Text) <> "" Then
                data_conv.Recordset("cnv_preccons") = t_implla.Text
             End If
             data_conv.Recordset("cnv_menanio") = cbomesanio.ListIndex
             If Trim(cbogrupoap.Text) <> "" Then
                data_conv.Recordset("cnv_grupoap") = cbogrupoap.Text
             End If
             If t_email.Text <> "" Then
                data_conv.Recordset("cnv_correoe") = t_email.Text
             End If
             If t_rub.Text = "" Then
                data_conv.Recordset("cnv_uapago") = 0
             Else
                data_conv.Recordset("cnv_uapago") = t_rub.Text
             End If
             If vdesde.Text <> "__/__/____" Then
                data_conv.Recordset("cnv_desde") = Format(vdesde.Text, "dd/mm/yyyy")
             End If
             If vhasta.Text <> "__/__/____" Then
                data_conv.Recordset("cnv_hasta") = Format(vhasta.Text, "dd/mm/yyyy")
             End If
             If cbomon.Text = "$U" Then
                data_conv.Recordset("cnv_codmon") = 1
                data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
             Else
                If cbomon.Text = "U$s" Then
                   data_conv.Recordset("cnv_codmon") = 2
                   data_conv.Recordset("cnv_nommon") = "DOLARES U.S.A."
                Else
                   data_conv.Recordset("cnv_codmon") = 1
                   data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
                End If
             End If
             If cbocolrec.Text = "ROJO" Then
                data_conv.Recordset("cnv_colrec") = "R"
             Else
                If cbocolrec.Text = "AZUL" Then
                   data_conv.Recordset("cnv_colrec") = "A"
                Else
                   If cbocolrec.Text = "MARRON" Then
                      data_conv.Recordset("cnv_colrec") = "M"
                   Else
                      If cbocolrec.Text = "VERDE" Then
                         data_conv.Recordset("cnv_colrec") = "V"
                      Else
                         If cbocolrec.Text = "CELESTE" Then
                            data_conv.Recordset("cnv_colrec") = "C"
                         Else
                            data_conv.Recordset("cnv_colrec") = ""
                         End If
                      End If
                   End If
                End If
             End If
             If txt_precio.Text <> "" Then
                data_conv.Recordset("cnv_precio") = txt_precio.Text
             Else
                data_conv.Recordset("cnv_precio") = 0
             End If
             If cbosirec.Text = "SI" Then
                data_conv.Recordset("cnv_emite") = "SI"
             Else
                If cbosirec.Text = "NO" Then
                   data_conv.Recordset("cnv_emite") = "NO"
                Else
                   data_conv.Recordset("cnv_emite") = "NO"
                End If
             End If
             If txt_cuenta.Text <> "" Then
                data_conv.Recordset("cnv_cuenta") = txt_cuenta.Text
             Else
                data_conv.Recordset("cnv_cuenta") = 0
             End If
             If cboaltasi.Text = "NO" Then
                data_conv.Recordset("cnv_alta") = "NO"
             Else
                data_conv.Recordset("cnv_alta") = "SI"
             End If
             If opunosolo.Value = True Then
                data_conv.Recordset("cnv_cant_r") = 1
             Else
                If optodos.Value = True Then
                   data_conv.Recordset("cnv_cant_r") = 2
                Else
                   data_conv.Recordset("cnv_cant_r") = 1
                End If
             End If
             If cbomut.Text <> "" Then
                data_conv.Recordset("cnv_grupo") = cbomut.Text
             Else
                data_conv.Recordset("cnv_grupo") = ""
             End If
             If fbaja.Text <> "__/__/____" Then
                data_conv.Recordset("cnv_fbaja") = Format(fbaja.Text, "dd/mm/yyyy")
             Else
    
             End If
             If txt_obs.Text <> "" Then
                data_conv.Recordset("cnv_ctrato") = txt_obs.Text
             End If
             data_conv.Recordset("cnv_pmserv") = cbofact.ListIndex
             If t_der.Text <> "" Then
                data_conv.Recordset("cnv_motbaj") = t_der.Text
             Else
                data_conv.Recordset("cnv_motbaj") = ""
             End If
             If cbovenc.ListIndex > 0 Then
                data_conv.Recordset("cnv_paserv") = Val(cbovenc.Text)
             Else
                data_conv.Recordset("cnv_paserv") = 0
             End If
             data_conv.Recordset("cnv_sald") = chtimbre.Value
             If t_nrocompra.Text <> "" Then
                data_conv.Recordset("cnv_email") = t_nrocompra.Text
             End If
             If Combo1.Text <> "" Then
                data_aranc.RecordSource = "Select * from Aran_grupos where desc_gpo ='" & Combo1.Text & "'"
                data_aranc.Refresh
                If data_aranc.Recordset.RecordCount > 0 Then
                   data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id")
                Else
                   MsgBox "Atención: No se encuentra el grupo de arancel ingresado. Verifique!", vbCritical
                   data_conv.Recordset("cnv_aran") = 0
                End If
             Else
                data_conv.Recordset("cnv_aran") = 0
             End If
             data_conv.Recordset.Update
             data_prec.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & txt_cod.Text & "' order by cnv_desde DESC"
             data_prec.Refresh
             If data_prec.Recordset.RecordCount > 0 Then
                data_prec.Recordset("cnv_hasta") = Format(vhasta.Text, "yyyy-mm-dd")
                data_prec.Recordset("cnv_desde") = Format(vdesde.Text, "yyyy-mm-dd")
                data_prec.Recordset("precio") = txt_precio.Text
                data_prec.Recordset.Update
             Else
                data_prec.Recordset.AddNew
                data_prec.Recordset("cnv_hasta") = Format(vhasta.Text, "yyyy-mm-dd")
                data_prec.Recordset("cnv_codigo") = txt_cod.Text
                data_prec.Recordset("cnv_desde") = Format(vdesde.Text, "yyyy-mm-dd")
                If Trim(txt_precio.Text) <> "" Then
                   data_prec.Recordset("precio") = txt_precio.Text
                Else
                   data_prec.Recordset("precio") = 0
                End If
                data_prec.Recordset("moneda") = 1
                data_prec.Recordset("usuario") = WElusuario
                data_prec.Recordset.Update
               
             End If
             
             data_cnvbusca.Refresh
             borraconv
             deshab
             igualar
             balta.Enabled = True
             bgraba.Enabled = False
             bmodi.Enabled = True
             bcance.Enabled = False
             bbusca.Enabled = True
             bimp.Enabled = True
             b_fact.Enabled = True
             b_hist.Enabled = True
             XAcnv = 0
          Else
             MsgBox "No ingresó descripción", vbCritical, "Convenios"
             txt_desc.SetFocus
          End If
       Else
          MsgBox "Verifique si ingresó Grupo de Aranceles o verifique código del convenio.", vbCritical, "Convenios"
          txt_desc.SetFocus
       End If
    Else
       If txt_cod.Text <> "" And Combo1.Text <> "" Then
          If txt_desc.Text <> "" Then
             data_conv.Recordset.Edit
             data_conv.Recordset("cnv_codigo") = txt_cod.Text
             data_conv.Recordset("cnv_desc") = txt_desc.Text
             data_conv.Recordset("cnv_entre") = t_razon.Text
             data_conv.Recordset("cnv_direcc") = txt_direc.Text
             data_conv.Recordset("cnv_local") = txt_localid.Text
             data_conv.Recordset("cnv_tel") = txt_tel.Text
             data_conv.Recordset("cnv_ruc") = txt_ruc.Text
             data_conv.Recordset("cnv_umpago") = choculta.Value
             data_conv.Recordset("cnv_sindeuda") = chdeuda.Value
             data_conv.Recordset("cnv_gpoafilia") = chafilia.Value
             If Trim(t_cantlla.Text) <> "" Then
                data_conv.Recordset("cnv_cantcons") = t_cantlla.Text
             Else
                If IsNull(data_conv.Recordset("cnv_cantcons")) = False Then
                   data_conv.Recordset("cnv_cantcons") = Null
                End If
             End If
             If Trim(t_implla.Text) <> "" Then
                data_conv.Recordset("cnv_preccons") = t_implla.Text
             Else
                If IsNull(data_conv.Recordset("cnv_preccons")) = False Then
                   data_conv.Recordset("cnv_preccons") = Null
                End If
             End If
             data_conv.Recordset("cnv_menanio") = cbomesanio.ListIndex
             If Trim(cbogrupoap.Text) <> "" Then
                data_conv.Recordset("cnv_grupoap") = cbogrupoap.Text
             Else
                If IsNull(data_conv.Recordset("cnv_grupoap")) = False Then
                   data_conv.Recordset("cnv_grupoap") = Null
                End If
             End If
             If t_email.Text <> "" Then
                If IsNull(data_conv.Recordset("cnv_correoe")) = False Then
                   If t_email.Text <> data_conv.Recordset("cnv_correoe") Then
                      data_conv.Recordset("cnv_correoe") = t_email.Text
                   End If
                Else
                   data_conv.Recordset("cnv_correoe") = t_email.Text
                End If
             Else
                If IsNull(data_conv.Recordset("cnv_correoe")) = False Then
                   data_conv.Recordset("cnv_correoe") = Null
                End If
                
             End If
             If t_rub.Text = "" Then
                data_conv.Recordset("cnv_uapago") = 0
             Else
                data_conv.Recordset("cnv_uapago") = t_rub.Text
             End If
             
             If vdesde.Text <> "__/__/____" Then
                data_conv.Recordset("cnv_desde") = Format(vdesde.Text, "dd/mm/yyyy")
             End If
             If vhasta.Text <> "__/__/____" Then
                data_conv.Recordset("cnv_hasta") = Format(vhasta.Text, "dd/mm/yyyy")
             End If
             If cbomon.Text = "$U" Then
                data_conv.Recordset("cnv_codmon") = 1
                data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
             Else
                If cbomon.Text = "U$s" Then
                   data_conv.Recordset("cnv_codmon") = 2
                   data_conv.Recordset("cnv_nommon") = "DOLARES U.S.A."
                Else
                   data_conv.Recordset("cnv_codmon") = 1
                   data_conv.Recordset("cnv_nommon") = "PESOS URUGUAYOS"
                End If
             End If
             If cbocolrec.Text = "ROJO" Then
                data_conv.Recordset("cnv_colrec") = "R"
             Else
                If cbocolrec.Text = "AZUL" Then
                   data_conv.Recordset("cnv_colrec") = "A"
                Else
                   If cbocolrec.Text = "MARRON" Then
                      data_conv.Recordset("cnv_colrec") = "M"
                   Else
                      If cbocolrec.Text = "VERDE" Then
                         data_conv.Recordset("cnv_colrec") = "V"
                      Else
                         If cbocolrec.Text = "CELESTE" Then
                            data_conv.Recordset("cnv_colrec") = "C"
                         Else
                            data_conv.Recordset("cnv_colrec") = ""
                         End If
                      End If
                   End If
                End If
             End If
             If txt_precio.Text <> "" Then
                data_conv.Recordset("cnv_precio") = txt_precio.Text
             Else
                data_conv.Recordset("cnv_precio") = 0
             End If
             If cbosirec.Text = "SI" Then
                data_conv.Recordset("cnv_emite") = "SI"
             Else
                If cbosirec.Text = "NO" Then
                   data_conv.Recordset("cnv_emite") = "NO"
                Else
                   data_conv.Recordset("cnv_emite") = "NO"
                End If
             End If
             If txt_cuenta.Text <> "" Then
                data_conv.Recordset("cnv_cuenta") = txt_cuenta.Text
             Else
                data_conv.Recordset("cnv_cuenta") = 0
             End If
             If cboaltasi.Text = "NO" Then
                data_conv.Recordset("cnv_alta") = "NO"
             Else
                data_conv.Recordset("cnv_alta") = "SI"
             End If
             If opunosolo.Value = True Then
                data_conv.Recordset("cnv_cant_r") = 1
             Else
                If optodos.Value = True Then
                   data_conv.Recordset("cnv_cant_r") = 2
                Else
                   data_conv.Recordset("cnv_cant_r") = 1
                End If
             End If
             If cbomut.Text <> "" Then
                data_conv.Recordset("cnv_grupo") = cbomut.Text
             Else
                data_conv.Recordset("cnv_grupo") = ""
             End If
             If fbaja.Text <> "__/__/____" Then
                data_conv.Recordset("cnv_fbaja") = Format(fbaja.Text, "dd/mm/yyyy")
             Else
                If IsNull(data_conv.Recordset("cnv_fbaja")) = False Then
                   data_conv.Recordset("cnv_fbaja") = Null
                End If
             End If
             If txt_obs.Text <> "" Then
                data_conv.Recordset("cnv_ctrato") = txt_obs.Text
             End If
             data_conv.Recordset("cnv_pmserv") = cbofact.ListIndex
             If t_der.Text <> "" Then
                data_conv.Recordset("cnv_motbaj") = t_der.Text
             Else
                data_conv.Recordset("cnv_motbaj") = ""
             End If
             If cbovenc.ListIndex > 0 Then
                data_conv.Recordset("cnv_paserv") = Val(cbovenc.Text)
             Else
                data_conv.Recordset("cnv_paserv") = 0
             End If
             data_conv.Recordset("cnv_sald") = chtimbre.Value
             If t_nrocompra.Text <> "" Then
                data_conv.Recordset("cnv_email") = t_nrocompra.Text
             Else
                If IsNull(data_conv.Recordset("cnv_email")) = False Then
                   data_conv.Recordset("cnv_email") = Null
                End If
             End If
             If Combo1.Text <> "" Then
                data_aranc.RecordSource = "Select * from Aran_grupos where desc_gpo ='" & Combo1.Text & "'"
                data_aranc.Refresh
                If data_aranc.Recordset.RecordCount > 0 Then
                   If IsNull(data_conv.Recordset("cnv_aran")) = False Then
                      If data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id") Then
                      Else
                         data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id")
                      End If
                   Else
                      data_conv.Recordset("cnv_aran") = data_aranc.Recordset("id")
                   End If
                Else
                   MsgBox "Atención: No se encuentra el grupo de arancel ingresado. Verifique!", vbCritical
                   data_conv.Recordset("cnv_aran") = 0
                End If
             Else
                data_conv.Recordset("cnv_aran") = 0
             End If
             data_conv.Recordset.Update
             data_prec.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & txt_cod.Text & "' and cnv_hasta >='" & Format(Date, "yyyy-mm-dd") & "' order by cnv_desde DESC"
             data_prec.Refresh
             If data_prec.Recordset.RecordCount > 0 Then
                data_prec.Recordset.MoveFirst
                If data_prec.Recordset("precio") <> txt_precio.Text Then
                   data_prec.Recordset("cnv_hasta") = Date - 1
                   data_prec.Recordset.Update
                   data_prec.Recordset.AddNew
                   data_prec.Recordset("cnv_codigo") = txt_cod.Text
                   data_prec.Recordset("cnv_desde") = Format(Date, "yyyy-mm-dd")
                   Mifec = Date + 365
                   data_prec.Recordset("cnv_hasta") = Mifec
                   data_prec.Recordset("precio") = Format(txt_precio.Text, "Standard")
                   data_prec.Recordset("moneda") = 1
                   data_prec.Recordset("usuario") = WElusuario
                   data_prec.Recordset.Update
                End If
             Else
                data_prec.Recordset.AddNew
                data_prec.Recordset("cnv_hasta") = Format(vhasta.Text, "yyyy-mm-dd")
                data_prec.Recordset("cnv_codigo") = txt_cod.Text
                data_prec.Recordset("cnv_desde") = Format(vdesde.Text, "yyyy-mm-dd")
                data_prec.Recordset("precio") = txt_precio.Text
                data_prec.Recordset("moneda") = 1
                data_prec.Recordset("usuario") = WElusuario
                data_prec.Recordset.Update
             End If
             
             data_cnvbusca.Refresh
             borraconv
             deshab
             igualar
             balta.Enabled = True
             bgraba.Enabled = False
             bmodi.Enabled = True
             bcance.Enabled = False
             bbusca.Enabled = True
             bimp.Enabled = True
             b_fact.Enabled = True
             b_hist.Enabled = True
             XAcnv = 0
          Else
             MsgBox "No ingresó descripción", vbCritical, "Convenios"
             txt_desc.SetFocus
          End If
       Else
          MsgBox "Verifique si ingresó Grupo de Aranceles o verifique código del convenio.", vbCritical, "Convenios"
          txt_desc.SetFocus
       End If
    End If
End If
         
End Sub

Private Sub bimp_Click()
frm_infconves.Show vbModal

End Sub

Private Sub bmodi_Click()
If WElusuario = "JFERNAN" Or XWeltipoU = "ADMINISTRADOR" Or WElusuario = "MCOSTA" Or WElusuario = "MPEREZ" Or WElusuario = "GUSTAVO" Or XWeltipoU = "USUARIOS ADM" Then
    data_conv.Recordset.FindFirst "cnv_codigo = '" & txt_cod.Text & "'"
    If Not data_conv.Recordset.NoMatch Then
       habilita
       txt_cod.SetFocus
       balta.Enabled = False
       bgraba.Enabled = True
       bmodi.Enabled = False
       bcance.Enabled = True
       bimp.Enabled = False
       bbusca.Enabled = False
       b_fact.Enabled = False
       b_hist.Enabled = False
       XAcnv = 0
       data_abmconv.Recordset.AddNew
       data_abmconv.Recordset("cnv_codigo") = txt_cod.Text
       data_abmconv.Recordset("cnv_desc") = txt_desc.Text
       data_abmconv.Recordset("cnv_ruc") = txt_ruc.Text
       If vdesde.Text <> "__/__/____" Then
          data_abmconv.Recordset("cnv_desde") = Format(vdesde.Text, "dd/mm/yyyy")
       End If
       If vhasta.Text <> "__/__/____" Then
          data_abmconv.Recordset("cnv_hasta") = Format(vhasta.Text, "dd/mm/yyyy")
       End If
       data_abmconv.Recordset("cnv_nommon") = Format(Time, "HH:MM:ss")
       If cbomut.Text <> "" Then
          data_abmconv.Recordset("cnv_grupo") = cbomut.Text
       End If
       data_abmconv.Recordset("cnv_direcc") = WElusuario
       data_abmconv.Recordset("cnv_fbaja") = Format(Date, "dd/mm/yyyy")
       data_abmconv.Recordset("cnv_precio") = frm_menu.data_parse.Recordset("base")
       data_abmconv.Recordset("cnv_local") = "BOTON MOD"
       data_abmconv.Recordset.Update
    End If
Else
   MsgBox "Usuario no autorizado", vbCritical, "Convenios"
   Unload Me
End If

End Sub

Private Sub cboaltasi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomut.SetFocus
End If

End Sub

Private Sub cbocolrec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_precio.SetFocus
End If

End Sub

Private Sub cbogrupoap_DblClick()
If Trim(cbogrupoap.Text) <> "" Then
   cbogrupoap.Text = ""
Else
   cbogrupoap.Text = txt_cod.Text
End If

End Sub

Private Sub cbogrupoap_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

End Sub

Private Sub cbomon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocolrec.SetFocus
End If

End Sub

Private Sub cbomut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fbaja.SetFocus
End If

End Sub

Private Sub cbosirec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cuenta.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xdig, Xrut, Xtot, Xtot2, Xfactor, i As Integer
If Len(Trim(txt_ruc.Text)) = 12 Then
   If IsNumeric(txt_ruc.Text) Then
      Xdig = Val(Mid(txt_ruc.Text, 12, 1))
      Xrut = Val(Mid(txt_ruc.Text, 1, 12))
      Xtot = 0
      Xfactor = 2
      For i = 1 To 11
          If i = 1 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 4
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 2 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 3
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 3 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 2
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 4 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 9
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 5 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 8
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 6 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 7
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 7 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 6
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 8 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 5
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 9 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 4
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 10 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 3
             Xtot2 = Xtot2 + Xtot
          End If
          If i = 11 Then
             Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 2
             Xtot2 = Xtot2 + Xtot
          End If
      Next
      Xtot = Xtot2 / 11
      If Xtot > 0 Then
         Xtot = 11 - Xtot
      Else
         Xdig = 0
      End If
      MsgBox "RUT:" & Xdig
   Else
      MsgBox "Verifique el RUT si tiene ingresado solo números"
   End If
Else
   MsgBox "Verifique si están todos los números del RUT"
End If


End Sub

Private Sub fbaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_obs.SetFocus
End If

End Sub

Private Sub fbaja_LostFocus()
If fbaja.Text <> "__/__/____" Then
   If IsDate(fbaja.Text) = False Then
      MsgBox "Error en fecha", vbCritical, "Convenios"
      txt_obs.SetFocus
   End If
End If

End Sub

Private Sub Form_Initialize()
data_conv.Recordset.MoveLast
igualar

End Sub

Private Sub Form_Load()
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.RecordSource = "Select * from convenio"
data_conv.Refresh
data_cnvbusca.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_inftiq.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_aranc.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_aranc.RecordSource = "Select * from Aran_grupos"
data_aranc.Refresh
Combo1.Clear
If data_aranc.Recordset.RecordCount > 0 Then
   Do While Not data_aranc.Recordset.EOF
      Combo1.AddItem data_aranc.Recordset("desc_gpo")
      data_aranc.Recordset.MoveNext
   Loop
End If
   
Adodc1.ConnectionString = "dsn=" & Xconexrmt
'data_cnvbusca.RecordSource = "convenio"
'data_cnvbusca.Refresh
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh
data_prec.ConnectionString = "dsn=" & Xconexrmt

data_abmconv.DatabaseName = App.path & "\abmconv.mdb"
data_abmconv.RecordSource = "abmconv"
data_abmconv.Refresh

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With
End Sub

Private Sub Label9_DblClick()
frm_histconve.Show vbModal

End Sub

Private Sub opunosolo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboaltasi.SetFocus
End If

End Sub

Private Sub t_cantlla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_implla.SetFocus
End If

End Sub

Private Sub t_implla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomesanio.SetFocus
End If

End Sub

Private Sub t_implla_LostFocus()
If Trim(t_implla.Text) <> "" Then
   t_implla.Text = Format(t_implla.Text, "Standard")

End If

End Sub

Private Sub t_razon_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))
If KeyAscii = 13 Then
   txt_direc.SetFocus
End If

End Sub

Private Sub txt_cod_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))
If KeyAscii = 13 Then
   txt_desc.SetFocus
End If

End Sub

Private Sub txt_cod_LostFocus()
If XAcnv = 1 Then
   If txt_cod.Text <> "" Then
       data_cnvbusca.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_cod.Text & "'"
       data_cnvbusca.Refresh
    '   data_cnvbusca.Recordset.FindFirst "cnv_codigo = '" & txt_cod.Text & "'"
    '   If Not data_cnvbusca.Recordset.NoMatch Then
       If data_cnvbusca.Recordset.RecordCount > 0 Then
          MsgBox "Ya existe convenio, VERIFIQUE", vbCritical, "Convenios"
          txt_desc.SetFocus
       End If
   End If
End If

End Sub

Private Sub txt_cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   opunosolo.SetFocus
End If

End Sub

Private Sub txt_desc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))
If KeyAscii = 13 Then
   t_razon.SetFocus
End If

End Sub

Private Sub txt_direc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_localid.SetFocus
End If

End Sub

Private Sub txt_localid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_tel.SetFocus
End If

End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbosirec.SetFocus
End If

End Sub

Private Sub txt_ruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   vdesde.SetFocus
End If

End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ruc.SetFocus
End If
End Sub

Private Sub vdesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   vhasta.SetFocus
End If

End Sub

Private Sub vdesde_LostFocus()
If IsDate(vdesde.Text) = False Then
   MsgBox "Error en fecha", vbCritical, "Convenios"
   vhasta.SetFocus
End If

End Sub

Private Sub vhasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomon.SetFocus
End If

End Sub

Private Sub vhasta_LostFocus()
If IsDate(vhasta.Text) = False Then
   MsgBox "Error en fecha", vbCritical, "Convenios"
   cbomon.SetFocus
End If

End Sub

Public Function borraconv()

txt_cod.Text = ""
txt_desc.Text = ""
txt_direc.Text = ""
txt_localid.Text = ""
txt_tel.Text = ""
txt_ruc.Text = ""
vdesde.Text = "__/__/____"
vhasta.Text = "__/__/____"
cbomon.Text = ""
t_cantlla.Text = ""
cbocolrec.Text = ""
txt_precio.Text = ""
cbosirec.Text = ""
txt_cuenta.Text = ""
cboaltasi.Text = ""
opunosolo.Value = True
cbomut.Text = ""
fbaja.Text = "__/__/____"
txt_obs.Text = ""
t_razon.Text = ""
cbofact.ListIndex = -1
cbovenc.ListIndex = -1
t_der.Text = ""
t_rub.Text = ""
t_nrocompra.Text = ""
t_email.Text = ""
chdeuda.Value = 0
chafilia.Value = 0
t_cantlla.Text = ""
t_implla.Text = ""
cbomesanio.ListIndex = -1
cbogrupoap.Text = ""
chtimbre.Value = 0

End Function

Public Function igualar()
If IsNull(data_conv.Recordset("cnv_codigo")) = False Then
   txt_cod.Text = data_conv.Recordset("cnv_codigo")
Else
   txt_cod.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_desc")) = False Then
   txt_desc.Text = data_conv.Recordset("cnv_desc")
Else
   txt_desc.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_direcc")) = False Then
   If IsNull(data_conv.Recordset("cnv_entre")) = False Then
      txt_direc.Text = data_conv.Recordset("cnv_direcc")
   Else
      txt_direc.Text = data_conv.Recordset("cnv_direcc")
   End If
Else
   txt_direc.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_uapago")) = False Then
   t_rub.Text = data_conv.Recordset("cnv_uapago")
Else
   t_rub.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_sindeuda")) = False Then
   chdeuda.Value = data_conv.Recordset("cnv_sindeuda")
Else
   chdeuda.Value = 0
End If
If IsNull(data_conv.Recordset("cnv_gpoafilia")) = False Then
   chafilia.Value = data_conv.Recordset("cnv_gpoafilia")
Else
   chafilia.Value = 0
End If
If IsNull(data_conv.Recordset("cnv_local")) = False Then
   txt_localid.Text = data_conv.Recordset("cnv_local")
Else
   txt_localid.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_correoe")) = False Then
   t_email.Text = data_conv.Recordset("cnv_correoe")
Else
   t_email.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_cantcons")) = False Then
   t_cantlla.Text = data_conv.Recordset("cnv_cantcons")
Else
   t_cantlla.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_tel")) = False Then
   txt_tel.Text = data_conv.Recordset("cnv_tel")
Else
   txt_tel.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_ruc")) = False Then
   txt_ruc.Text = Trim(data_conv.Recordset("cnv_ruc"))
Else
   txt_ruc.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_desde")) = False Then
   vdesde.Text = Format(data_conv.Recordset("cnv_desde"), "dd/mm/yyyy")
Else
   vdesde.Text = "__/__/____"
End If
If IsNull(data_conv.Recordset("cnv_hasta")) = False Then
   vhasta.Text = Format(data_conv.Recordset("cnv_hasta"), "dd/mm/yyyy")
Else
   vhasta.Text = "__/__/____"
End If
If IsNull(data_conv.Recordset("cnv_codmon")) = False Then
   If data_conv.Recordset("cnv_codmon") = 1 Then
      cbomon.ListIndex = 0
   Else
      If data_conv.Recordset("cnv_codmon") = 2 Then
         cbomon.ListIndex = 1
      Else
         cbomon.Text = ""
      End If
   End If
Else
   cbomon.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_colrec")) = False Then
   If data_conv.Recordset("cnv_colrec") = "R" Then
      cbocolrec.ListIndex = 0
   Else
      If data_conv.Recordset("cnv_colrec") = "A" Then
         cbocolrec.ListIndex = 1
      Else
         If data_conv.Recordset("cnv_colrec") = "M" Then
            cbocolrec.ListIndex = 2
         Else
            If data_conv.Recordset("cnv_colrec") = "V" Then
               cbocolrec.ListIndex = 3
            Else
               If data_conv.Recordset("cnv_colrec") = "C" Then
                  cbocolrec.ListIndex = 4
               Else
                  cbocolrec.Text = ""
               End If
            End If
         End If
      End If
   End If
Else
   cbocolrec.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_precio")) = False Then
   txt_precio.Text = data_conv.Recordset("cnv_precio")
Else
   txt_precio.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_emite")) = False Then
   If data_conv.Recordset("cnv_emite") = "SI" Then
      cbosirec.ListIndex = 1
   Else
      If data_conv.Recordset("cnv_emite") = "NO" Then
         cbosirec.ListIndex = 0
      Else
         cbosirec.ListIndex = 0
      End If
   End If
Else
   cbosirec.ListIndex = 0
End If
If IsNull(data_conv.Recordset("cnv_cuenta")) = False Then
   txt_cuenta.Text = data_conv.Recordset("cnv_cuenta")
Else
   txt_cuenta.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_alta")) = False Then
   If data_conv.Recordset("cnv_alta") = "SI" Then
      cboaltasi.ListIndex = 0
   Else
      cboaltasi.ListIndex = 1
   End If
Else
   cboaltasi.ListIndex = 1
End If
If IsNull(data_conv.Recordset("cnv_cant_r")) = False Then
   If data_conv.Recordset("cnv_cant_r") = 1 Then
      opunosolo.Value = True
   Else
      If data_conv.Recordset("cnv_cant_r") = 2 Then
         optodos.Value = True
      Else
         opunosolo.Value = True
      End If
   End If
Else
   opunosolo.Value = True
End If
If IsNull(data_conv.Recordset("cnv_umpago")) = False Then
   If data_conv.Recordset("cnv_umpago") > 2 Then
      choculta.Value = 0
   Else
      choculta.Value = data_conv.Recordset("cnv_umpago")
   End If
Else
   choculta.Value = 0
End If
If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
   If data_conv.Recordset("cnv_grupo") <> "" Then
      cbomut.Text = data_conv.Recordset("cnv_grupo")
   Else
      cbomut.Text = ""
   End If
Else
   cbomut.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_fbaja")) = False Then
   fbaja.Text = Format(data_conv.Recordset("cnv_fbaja"), "dd/mm/yyyy")
Else
   fbaja.Text = "__/__/____"
End If
If IsNull(data_conv.Recordset("cnv_ctrato")) = False Then
   txt_obs.Text = data_conv.Recordset("cnv_ctrato")
Else
   txt_obs.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_entre")) = False Then
   t_razon.Text = data_conv.Recordset("cnv_entre")
Else
   t_razon.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_pmserv")) = False Then
   cbofact.ListIndex = data_conv.Recordset("cnv_pmserv")
Else
   cbofact.ListIndex = -1
End If
If IsNull(data_conv.Recordset("cnv_motbaj")) = False Then
   t_der.Text = data_conv.Recordset("cnv_motbaj")
Else
   t_der.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_email")) = False Then
   t_nrocompra.Text = data_conv.Recordset("cnv_email")
Else
   t_nrocompra.Text = ""
End If

If IsNull(data_conv.Recordset("cnv_paserv")) = False Then
   If data_conv.Recordset("cnv_paserv") = 0 Then
      cbovenc.ListIndex = 0
   Else
      If data_conv.Recordset("cnv_paserv") = 15 Then
         cbovenc.ListIndex = 1
      Else
         If data_conv.Recordset("cnv_paserv") = 30 Then
            cbovenc.ListIndex = 2
         Else
            If data_conv.Recordset("cnv_paserv") = 60 Then
               cbovenc.ListIndex = 3
            Else
               If data_conv.Recordset("cnv_paserv") = 90 Then
                  cbovenc.ListIndex = 4
               Else
                  If data_conv.Recordset("cnv_paserv") = 120 Then
                     cbovenc.ListIndex = 5
                  Else
                     cbovenc.ListIndex = -1
                  End If
               End If
            End If
         End If
      End If
   End If
Else
   cbovenc.ListIndex = -1
End If
If IsNull(data_conv.Recordset("cnv_aran")) = True Then
   Combo1.ListIndex = -1
   Combo1.Text = ""
Else
   If data_conv.Recordset("cnv_aran") = 0 Then
      Combo1.ListIndex = -1
      Combo1.Text = ""
   Else
      data_aranc.RecordSource = "Select * from Aran_grupos where id =" & data_conv.Recordset("cnv_aran")
      data_aranc.Refresh
      If data_aranc.Recordset.RecordCount > 0 Then
         Combo1.Text = data_aranc.Recordset("desc_gpo")
      Else
         Combo1.ListIndex = -1
         Combo1.Text = ""
      End If
   End If
End If
If IsNull(data_conv.Recordset("cnv_cantcons")) = False Then
   t_cantlla.Text = data_conv.Recordset("cnv_cantcons")
Else
   t_cantlla.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_preccons")) = False Then
   t_implla.Text = data_conv.Recordset("cnv_preccons")
Else
   t_implla.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_menanio")) = False Then
   cbomesanio.ListIndex = data_conv.Recordset("cnv_menanio")
Else
   cbomesanio.ListIndex = -1
End If
If IsNull(data_conv.Recordset("cnv_grupoap")) = False Then
   cbogrupoap.Text = data_conv.Recordset("cnv_grupoap")
Else
   cbogrupoap.Text = ""
End If
If IsNull(data_conv.Recordset("cnv_sald")) = False Then
   chtimbre.Value = data_conv.Recordset("cnv_sald")
Else
   chtimbre.Value = 0
End If


End Function

Public Function habilita()
txt_cod.Enabled = True
txt_desc.Enabled = True
txt_direc.Enabled = True
txt_localid.Enabled = True
txt_tel.Enabled = True
txt_ruc.Enabled = True
t_razon.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True

End Function

Public Function deshab()
txt_cod.Enabled = False
txt_desc.Enabled = False
txt_direc.Enabled = False
txt_localid.Enabled = False
t_razon.Enabled = False
txt_tel.Enabled = False
txt_ruc.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False

End Function
