VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_especcovid 
   BackColor       =   &H00404000&
   Caption         =   "Formulario de agenda para Solicitud Test Antígeno/HNF"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14445
   Icon            =   "frm_especcovid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   14445
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_users 
      Caption         =   "data_users"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dat_par 
      Caption         =   "dat_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FF0000&
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
      Height          =   315
      ItemData        =   "frm_especcovid.frx":058A
      Left            =   120
      List            =   "frm_especcovid.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   87
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton modifConsultorio 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      Picture         =   "frm_especcovid.frx":05C5
      Style           =   1  'Graphical
      TabIndex        =   86
      ToolTipText     =   "modificar consultorio ocupado"
      Top             =   3480
      Width           =   495
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5760
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc data_buscnv 
      Height          =   375
      Left            =   2400
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "data_buscnv"
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
   Begin MSAdodcLib.Adodc data_busca2 
      Height          =   375
      Left            =   3840
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "data_busca2"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar"
      Height          =   375
      Left            =   120
      Picture         =   "frm_especcovid.frx":35F5
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3375
      Left            =   120
      TabIndex        =   79
      Top             =   4680
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BASE"
         Object.Width           =   1481
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NOMBRE DEL MEDICO"
         Object.Width           =   4657
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ESPECIALIDAD"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "H.INICIO"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID"
         Object.Width           =   1129
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Solo consultas actuales"
      Height          =   495
      Left            =   5280
      TabIndex        =   76
      Top             =   4080
      Width           =   1335
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6480
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Anotación de pacientes"
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
      Height          =   5055
      Left            =   6720
      TabIndex        =   42
      Top             =   3360
      Width           =   7575
      Begin VB.CommandButton b_excel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   600
         TabIndex        =   90
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox t_direc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Data data_conscli 
         Caption         =   "data_conscli"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Anota metas"
         Height          =   495
         Left            =   2040
         TabIndex        =   78
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton b_impcons 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         Picture         =   "frm_especcovid.frx":3B7F
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Imprimir la consulta seleccionada"
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton b_elianota 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         Picture         =   "frm_especcovid.frx":4109
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox t_d 
         Height          =   405
         Left            =   4680
         TabIndex        =   67
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox t_m 
         Height          =   375
         Left            =   4320
         TabIndex        =   66
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox t_a 
         Height          =   405
         Left            =   3720
         TabIndex        =   65
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton b_buscapac 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         Picture         =   "frm_especcovid.frx":4693
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Buscar los datos del paciente"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton b_modpac 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         Picture         =   "frm_especcovid.frx":4C1D
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Modifica el dato seleccionado"
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton b_agrega 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Picture         =   "frm_especcovid.frx":51A7
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Agrega los datos a la lista de pacientes anotados"
         Top             =   2280
         Width           =   375
      End
      Begin VB.ComboBox cbotipcons 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frm_especcovid.frx":5731
         Left            =   1680
         List            =   "frm_especcovid.frx":5733
         TabIndex        =   61
         Text            =   "cbotipcons"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox t_celu 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1680
         TabIndex        =   60
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cbosino 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frm_especcovid.frx":5735
         Left            =   6720
         List            =   "frm_especcovid.frx":573F
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox t_tellinea 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5760
         TabIndex        =   57
         Top             =   1080
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mfnac 
         Height          =   375
         Left            =   6120
         TabIndex        =   54
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12582912
         ForeColor       =   14737632
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
      Begin VB.TextBox t_conv 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   52
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_nompac 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1680
         TabIndex        =   49
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox t_codced 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5640
         TabIndex        =   48
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4560
         TabIndex        =   47
         Top             =   360
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   0
         TabIndex        =   45
         ToolTipText     =   "HACIENDO DOBLE CLICK SOBRE EL REGISTRO PUEDE CANCELAR LA CONSULTA"
         Top             =   2640
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16744576
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "Nro."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "Hora"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "c"
            Text            =   "Cédula"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "d"
            Text            =   "Nombre"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "e"
            Text            =   "Convenio"
            Object.Width           =   1869
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "f"
            Text            =   "Celular"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "g"
            Text            =   "Tel.Línea"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "h"
            Text            =   "Zona"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "i"
            Text            =   "Sint.?"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Dirección"
            Object.Width           =   5010
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Anotado por"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox t_mat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1680
         TabIndex        =   44
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fec.Nac"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6360
         TabIndex        =   91
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label labselfec 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   75
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Celular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Telef. de línea:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3840
         TabIndex        =   56
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Síntomas?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5520
         TabIndex        =   55
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zona:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Convenio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5520
         TabIndex        =   51
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cédula:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3480
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Matrícula:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton b_elifec 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   5160
      Picture         =   "frm_especcovid.frx":574B
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Eliminar fecha del especialista seleccionado"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Data data_cabfec 
      Caption         =   "data_cabfec"
      Connect         =   "odbc;dsn=sappnew"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_buscar 
      Caption         =   "data_buscar"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      Caption         =   "Datos de los MEDICOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   6720
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton b_cierramed 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         Picture         =   "frm_especcovid.frx":5CD5
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3720
         Width           =   375
      End
      Begin VB.Data data_medicossapp 
         Caption         =   "data_medicossapp"
         Connect         =   "odbc;dsn=sappnew"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * from medicos_esp order by nom_med"
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frm_especcovid.frx":625F
         Height          =   1575
         Left            =   120
         OleObjectBlob   =   "frm_especcovid.frx":627E
         TabIndex        =   38
         Top             =   2520
         Width           =   6135
      End
      Begin VB.CommandButton b_canmed 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         Picture         =   "frm_especcovid.frx":6C69
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton b_grabmed 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         Picture         =   "frm_especcovid.frx":71F3
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton b_modmed 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         Picture         =   "frm_especcovid.frx":777D
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton b_altamed 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Picture         =   "frm_especcovid.frx":7D07
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox t_codsapp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
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
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cboespec 
         BackColor       =   &H00C00000&
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
         Height          =   315
         ItemData        =   "frm_especcovid.frx":8291
         Left            =   1800
         List            =   "frm_especcovid.frx":82E6
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox t_nom 
         BackColor       =   &H00C00000&
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
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Top             =   720
         Width           =   3975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7560
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF8080&
         Caption         =   "CODIGO SAPP:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF8080&
         Caption         =   "ESPECIALIDAD"
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
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "NOMBRE:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label labcod 
         BackColor       =   &H00C00000&
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
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "CODIGO:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton b_infos 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      Picture         =   "frm_especcovid.frx":8452
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Informes del sistema"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_edimed 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      MouseIcon       =   "frm_especcovid.frx":89DC
      MousePointer    =   99  'Custom
      Picture         =   "frm_especcovid.frx":8F66
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Editar tabla de médicos"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Data data_medicos 
      Caption         =   "data_medicos"
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_base 
      Caption         =   "data_base"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Fechas disponibles"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   6720
      TabIndex        =   20
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton b_cancecons 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar fecha"
         Enabled         =   0   'False
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
         Left            =   5520
         Picture         =   "frm_especcovid.frx":94F0
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Cancelar la consulta seleccionada"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox t_codcons 
         Height          =   285
         Left            =   4800
         TabIndex        =   69
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_feccab 
         Height          =   285
         Left            =   5160
         TabIndex        =   68
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data data_lista 
         Caption         =   "data_lista"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "frm_especcovid.frx":9A7A
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "frm_especcovid.frx":9A94
         TabIndex        =   40
         Top             =   240
         Width           =   7215
      End
      Begin VB.Data data_fechas 
         Caption         =   "data_fechas"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSMask.MaskEdBox mnuevaf 
         Height          =   495
         Left            =   1920
         TabIndex        =   22
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   16711680
         ForeColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   15.75
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
      Begin VB.CommandButton b_nuevafecha 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Crear nueva fecha"
         Enabled         =   0   'False
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
         Left            =   120
         MouseIcon       =   "frm_especcovid.frx":A61F
         MousePointer    =   99  'Custom
         Picture         =   "frm_especcovid.frx":ABA9
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2640
         Width           =   1815
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_especcovid.frx":B133
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "frm_especcovid.frx":B147
      TabIndex        =   19
      Top             =   4560
      Width           =   6615
   End
   Begin VB.TextBox t_busca 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton b_elim 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      Picture         =   "frm_especcovid.frx":BE7A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Eliminar el registro de especialista seleccionado"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      Picture         =   "frm_especcovid.frx":C404
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancelar acción anterior"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      Picture         =   "frm_especcovid.frx":C98E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Guardar datos."
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      Picture         =   "frm_especcovid.frx":CF18
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Modificar registro seleccionado"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Picture         =   "frm_especcovid.frx":D4A2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Crear un nuevo registro de especialista"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Parámetros para la agenda"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Data data_consultahnf 
         Caption         =   "data_consultahnf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data data_aut 
         Caption         =   "data_aut"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox t_basedescsel 
         Height          =   285
         Left            =   2880
         TabIndex        =   84
         Top             =   2040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox t_especsel 
         Height          =   285
         Left            =   2880
         TabIndex        =   83
         Top             =   1440
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox t_basesel 
         Height          =   285
         Left            =   4680
         TabIndex        =   82
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox t_codmedsel 
         Height          =   375
         Left            =   4680
         TabIndex        =   81
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Data data_veoesp 
         Caption         =   "data_veoesp"
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
         Top             =   2520
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data abmesp 
         Caption         =   "abmesp"
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
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox t_idant 
         Height          =   375
         Left            =   3360
         TabIndex        =   77
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data data_borrados 
         Caption         =   "data_borrados"
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
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox t_cantp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   74
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox t_mm 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox t_espera 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSMask.MaskEdBox mhfin 
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16711680
         ForeColor       =   16777215
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
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
      Begin MSMask.MaskEdBox mhini 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16711680
         ForeColor       =   16777215
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
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
      Begin VB.ComboBox cbomedico 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2040
         TabIndex        =   4
         Top             =   960
         Width           =   4215
      End
      Begin VB.ComboBox cbobase 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label labussi 
         Height          =   255
         Left            =   840
         TabIndex        =   92
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label21 
         BackColor       =   &H00808080&
         Caption         =   "CANT. PACIENTES:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         Caption         =   "MINUTOS POR PACIENTE:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "EN ESPERA:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "HORA FINALIZA:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "HORA INICIO:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "SERVICIO:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "BASE:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid4 
      Height          =   855
      Left            =   3960
      OleObjectBlob   =   "frm_especcovid.frx":DA2C
      TabIndex        =   85
      Top             =   5280
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   5760
      Picture         =   "frm_especcovid.frx":E59B
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_especcovid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'agrego propiedad para comunicar con formulario frm_espeligeconsultorio'
Private id_hora_reserva As Integer
Private urlServicio As String
'mapeo entre nro de base y idBase'
Private nroBase_idBase As Dictionary


Private Sub b_agrega_Click()
Dim Xind, Xcant, Xnro As Long
Dim Xfecdeuda, Xlafechacons As Date
Dim Xloslabos, Xlacedconsulta As String
Dim Xcantlibres, Xellugar As Integer
Dim Xlafv As Date
Dim Xelcodigoaut, Xlapersona As String

Xlafv = Date

Xcantlibres = 0

Xloslabos = ""
Xlafechacons = Date
Dim Xcountt As Long
Dim Xdeudasiono As Integer
Dim Xind22 As Integer
Dim ExisteenHNF As Integer
ExisteenHNF = 0

Xdeudasiono = 0
Xcountt = 1

For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       Xcant = Xcant + 1
    End If
Next Xind
Xind = 0
If mfnac.Text <> "__/__/____" Then
   CalculaEdad (mfnac.Text)
Else
   t_a.Text = ""
   t_d.Text = ""
   t_m.Text = ""
End If
If t_mat.Text = "" Then
   t_mat.Text = 0
End If
If t_ced.Text = "" Then
   t_ced.Text = 0
End If

If t_ced.Text <> "" Then
   If WElusuario = "MARIAJOSE" Or WElusuario = "KARINAROMERO" Or WElusuario = "GMONZON" Or WElusuario = "ANOYA" Or WElusuario = "FDELEON" Or _
      WElusuario = "LROMERO" Or WElusuario = "ANAPAULA" Or WElusuario = "MSANCHEZ" Or labussi.Caption = "S" Then
   Else
      data_consultahnf.RecordSource = "select * from sol_hisopos where cedula =" & Val(t_ced.Text) & " and fecha >=#" & Format(Date, "yyyy/mm/dd") & "#"
      data_consultahnf.Refresh
      If data_consultahnf.Recordset.RecordCount > 0 Then
      Else
         ExisteenHNF = 1
      End If
   End If
End If
   
If ExisteenHNF = 1 Then
   MsgBox "ATENCION!! El paciente no figura ingresado en el formulario de solicitud HNF", vbCritical
Else
    If t_mat.Text <> "" Then
       If Xcant = 1 Then
          For Xind = 1 To ListView1.ListItems.count
              ListView1.ListItems(Xind).Selected = True
              If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                 Xnro = ListView1.ListItems(Xind).Text
                 Xellugar = Xind
                 data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                 data_lista.Refresh
                 If data_lista.Recordset.RecordCount > 0 Then
                    If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                       MsgBox "Ya existe un paciente anotado con este número, verifique!! o seleccione otro número.", vbCritical
                    Else
                                  
                       data_lista.Recordset.Edit
                       If t_mat.Text <> "" Then
                          data_lista.Recordset("mat_pac") = t_mat.Text
                       End If
                       If t_nompac.Text <> "" Then
                          data_lista.Recordset("nom_pac") = t_nompac.Text
                       End If
                      If t_ced.Text <> "" Then
                         data_lista.Recordset("ced_pac") = t_ced.Text & t_codced.Text
                      End If
                      If t_conv.Text <> "" Then
                         data_lista.Recordset("convenio") = t_conv.Text
                      End If
                      If t_celu.Text <> "" Then
                         data_lista.Recordset("cel_pac") = t_celu.Text
                      End If
                      If t_tellinea.Text <> "" Then
                         data_lista.Recordset("tel_pac") = t_tellinea.Text
                      End If
                      If mfnac.Text <> "__/__/____" Then
                         data_lista.Recordset("fec_nac") = Format(mfnac.Text, "dd/mm/yyyy")
                      End If
                      If cbosino.ListIndex >= 0 Then
                         data_lista.Recordset("taxi") = cbosino.Text
                      End If
                      If cbotipcons.ListIndex >= 0 Then
                         data_lista.Recordset("zona") = cbotipcons.Text
                      Else
                         data_lista.Recordset("zona") = "Sin dato"
                      End If
                      If t_direc.Text <> "" Then
                         data_lista.Recordset("direcci") = t_direc.Text
                      Else
                         data_lista.Recordset("direcci") = "Sin Dato"
                      End If
                      data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                      data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                      data_lista.Recordset("usua_anota") = WElusuario
                      data_lista.Recordset("usua_web") = "SAPP"
                      data_lista.Recordset.Update
                      abmesp.Recordset.AddNew
                      abmesp.Recordset("fecha") = Date
                      abmesp.Recordset("hora") = Format(Time, "HH:mm")
                      abmesp.Recordset("usuario") = WElusuario
                      abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
                      abmesp.Recordset("accion") = "ANOTACION"
                      abmesp.Recordset.Update
                      t_mat.Text = ""
                      t_ced.Text = ""
                      t_codced.Text = ""
                      t_nompac.Text = ""
                      t_celu.Text = ""
                      t_tellinea.Text = ""
                      t_conv.Text = ""
                      cbotipcons.Text = ""
                      cbosino.ListIndex = -1
                      t_direc.Text = ""
                      
                      data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " order by nro"
                      data_lista.Refresh
                      If data_lista.Recordset.RecordCount > 0 Then
                         data_lista.Recordset.MoveFirst
                         ListView1.ListItems.Clear
                         Do While Not data_lista.Recordset.EOF
                            If IsNull(data_lista.Recordset("nro")) = False Then
                               ListView1.ListItems.Add Xcountt, , data_lista.Recordset("nro")
                            Else
                               ListView1.ListItems.Add Xcountt, , "0"
                            End If
                            If IsNull(data_lista.Recordset("hora")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("hora")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                            End If
                            If IsNull(data_lista.Recordset("ced_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("ced_pac")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("nom_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("nom_pac")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("convenio")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("convenio")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("cel_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("cel_pac")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("tel_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tel_pac")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("zona")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("zona")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("taxi")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("taxi")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("direcci")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("direcci")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
                            End If
                            If IsNull(data_lista.Recordset("usua_anota")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("usua_anota")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
                            End If
                            
                            data_lista.Recordset.MoveNext
                            Xcountt = Xcountt + 1
                         Loop
                      End If
                    End If
                 Else
                    MsgBox "No se encuentra registro para actualizar"
                 End If
              End If
          Next Xind
     '             ListView1.ListItems.Item(Xellugar).Selected = True
     '             ListView1.ListItems.Item(Xellugar).EnsureVisible
     '             ListView1.SetFocus
          MsgBox "Paciente anotado correctamente.", vbInformation
          MsgBox "RECUERDE! Antes de anotar otro paciente, realizar doble click nuevamente en las fechas para cargar la lista.", vbCritical
       
       Else
          MsgBox "Debe seleccionar un solo registro."
       End If
    Else
        MsgBox "Ingrese CEDULA para anotar"
    End If
End If



End Sub

Private Sub b_altamed_Click()
XAlta = 1
b_modmed.Enabled = False
b_grabmed.Enabled = True
b_canmed.Enabled = True
b_altamed.Enabled = False
t_nom.Text = ""
cboespec.ListIndex = -1
t_codsapp.Text = ""
t_nom.SetFocus
DBGrid2.Enabled = False

End Sub

Private Sub b_buscapac_Click()
Xdeb = 15
frm_buscasocesp.Show vbModal

End Sub

Private Sub b_cance_Click()
b_nuevo.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cance.Enabled = False
b_elim.Enabled = True
b_edimed.Enabled = True
b_infos.Enabled = True
DBGrid1.Enabled = True
Frame1.Enabled = False
DBGrid1.SetFocus
cbobase.ListIndex = -1
cbomedico.ListIndex = -1
mhini.Text = "__:__"
mhfin.Text = "__:__"
t_espera.Text = ""
t_mm.Text = ""
t_cantp.Text = ""

Frame2.Enabled = True
Frame4.Enabled = True


End Sub

Private Sub b_cancecons_Click()
Dim Xdeseacance, Xobscancela As String
Xobscancela = ""
If data_cabfec.Recordset("cancela") = "SI" Then
   MsgBox "La consulta ya figura CANCELADA", vbCritical
Else
    Xdeseacance = MsgBox("Desea CANCELAR la fecha de reserva seleccionada? " & data_cabfec.Recordset("fecha"), vbInformation + vbYesNo)
    If Xdeseacance = vbYes Then
       t_feccab.Text = data_cabfec.Recordset("fecha")
       t_codcons.Text = data_cabfec.Recordset("cod_cons")
       Xobscancela = InputBox("INGRESE MOTIVO DE LA CANCELACION:")
       data_cabfec.Recordset.Edit
       data_cabfec.Recordset("cancela") = "SI"
       If Xobscancela <> "" Then
          data_cabfec.Recordset("motivo") = Xobscancela
       Else
          data_cabfec.Recordset("motivo") = "SIN DATOS"
       End If
       data_cabfec.Recordset("usuario") = WElusuario
       data_cabfec.Recordset("fecha_can") = Format(Date, "dd/mm/yyyy")
       data_cabfec.Recordset("hora_can") = Format(Time, "HH:mm")
       data_cabfec.Recordset.Update
        'cancelo consultorio'
        On Error Resume Next
        reservarConsultorio = especialidadReserva(t_especsel.Text)
        If reservarConsultorio Then
            id_hora_consultorio = data_cabfec.Recordset("id_hora_consultorio")
            basee = data_cabfec.Recordset("base")
            Set obj = consumirServicio("DELETE", urlServicio & "/bases/" & basee & "/consultorios/0/disponibilidades/" & id_hora_consultorio, "")
            'MsgBox obj.responseText
        End If
        On Error GoTo 0
       
    
       abmesp.Recordset.AddNew
       abmesp.Recordset("fecha") = Date
       abmesp.Recordset("hora") = Format(Time, "HH:mm")
       abmesp.Recordset("usuario") = WElusuario
       abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
       abmesp.Recordset("accion") = "CANCELA CONS"
       abmesp.Recordset.Update
    
       data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " order by nro"
       data_lista.Refresh
       If data_lista.Recordset.RecordCount > 0 Then
          data_lista.Recordset.MoveFirst
          ListView1.ListItems.Clear
          Do While Not data_lista.Recordset.EOF
             data_lista.Recordset.Edit
             data_lista.Recordset("cancela") = "SI"
             data_lista.Recordset.Update
             data_lista.Recordset.MoveNext
          Loop
          MsgBox "La consulta ha sido CANCELADA!", vbExclamation
       Else
          MsgBox "No existe fecha"
       End If
    End If
    frm_especcovid.MousePointer = 0
End If

End Sub

Private Sub b_canmed_Click()
XAlta = 0
b_modmed.Enabled = True
b_grabmed.Enabled = False
b_canmed.Enabled = False
b_altamed.Enabled = True
DBGrid2.Enabled = True
t_nom.Text = ""
cboespec.ListIndex = -1
t_codsapp.Text = ""
DBGrid2.SetFocus


End Sub

Private Sub b_cierramed_Click()
Frame3.Visible = False

End Sub

Private Sub b_edimed_Click()
If Frame3.Visible = False Then
   Frame4.Visible = False
   Frame3.Visible = True
Else
   Frame3.Visible = False
   Frame4.Visible = True
End If

End Sub

Private Sub b_elianota_Click()
Dim Xind, Xcant, Xnro As Long
Dim Xborralaconsulta As String
Dim Xcantlibres As Integer
Xcantlibres = 0

Xborralaconsulta = MsgBox("Desea borrar los ddat anotados en la lista?", vbInformation + vbYesNo)
If Xborralaconsulta = vbYes Then
    Xind = 0
    Xnro = 0
    Xcant = 0
    Dim Xcountt As Long
    Dim Xdeudasiono As Integer
    Xdeudasiono = 0
    Xcountt = 1

    For Xind = 1 To ListView1.ListItems.count
        ListView1.ListItems(Xind).Selected = True
        If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
           Xcant = Xcant + 1
        End If
    Next Xind
    Xind = 0
    
    If Xcant = 1 Then
       For Xind = 1 To ListView1.ListItems.count
           ListView1.ListItems(Xind).Selected = True
           If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
              Xnro = ListView1.ListItems(Xind).Text
              data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
              data_lista.Refresh
              If data_lista.Recordset.RecordCount > 0 Then
                 data_lista.Recordset.Edit
                 data_lista.Recordset("mat_pac") = Null
                 data_lista.Recordset("nom_pac") = Null
                 data_lista.Recordset("ced_pac") = Null
                 data_lista.Recordset("convenio") = Null
                 data_lista.Recordset("cel_pac") = Null
                 data_lista.Recordset("tel_pac") = Null
                 data_lista.Recordset("fec_nac") = Null
                 data_lista.Recordset("hcsiono") = Null
                 data_lista.Recordset("tipo_cons") = Null
                 data_lista.Recordset("tipo_consd") = Null
                 data_lista.Recordset("fec_anota") = Null
                 data_lista.Recordset("hora_anota") = Null
                 data_lista.Recordset("usua_anota") = Null
                 data_lista.Recordset("edad") = Null
                 data_lista.Recordset("usua_web") = Null
                 data_lista.Recordset("direcci") = Null
                 data_lista.Recordset("zona") = Null
                 data_lista.Recordset("taxi") = Null
                 data_lista.Recordset.Update
                 abmesp.Recordset.AddNew
                 abmesp.Recordset("fecha") = Date
                 abmesp.Recordset("hora") = Format(Time, "HH:mm")
                 abmesp.Recordset("usuario") = WElusuario
                 abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
                 abmesp.Recordset("accion") = "ELIMINA ANOTACION"
                 abmesp.Recordset.Update
                 
                 MsgBox "Registro eliminado!", vbInformation
                 data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " order by nro"
                 data_lista.Refresh
                 If data_lista.Recordset.RecordCount > 0 Then
                    data_lista.Recordset.MoveFirst
                    ListView1.ListItems.Clear
                    Do While Not data_lista.Recordset.EOF
                       If Xcantlibres <= 0 Then
                            If IsNull(data_lista.Recordset("nro")) = False Then
                               ListView1.ListItems.Add Xcountt, , data_lista.Recordset("nro")
                            Else
                               ListView1.ListItems.Add Xcountt, , "0"
                            End If
                            If IsNull(data_lista.Recordset("hora")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("hora")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                            End If
                            If IsNull(data_lista.Recordset("ced_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("ced_pac")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("nom_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("nom_pac")
                            Else
'                               If data_lista.Recordset("especial") = "PEDIATRIA" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Then
'                                  ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
'                               Else
                                  ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
'                                  Xcantlibres = Xcantlibres + 1
'                               End If
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                            End If
                            If IsNull(data_lista.Recordset("convenio")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("convenio")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("cel_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("cel_pac")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("tel_pac")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tel_pac")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("tipo_consd")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipo_consd")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                            End If
                            If IsNull(data_lista.Recordset("usua_web")) = False Then
                               If data_lista.Recordset("usua_web") = "SI" Then
                                  ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "WEB"
                               Else
                                  ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "SAPP"
                               End If
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "SAPP"
                            End If
                            If IsNull(data_lista.Recordset("obs")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("obs")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
                            End If
                            If IsNull(data_lista.Recordset("usua_anota")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("usua_anota")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
                            End If
                       End If
                       data_lista.Recordset.MoveNext
                       Xcountt = Xcountt + 1
                    Loop
                 Else
                    MsgBox "No se encuentra registro para actualizar"
                 End If
              End If
           End If
       Next Xind
    Else
       MsgBox "Debe seleccionar un solo registro"
    End If
End If


End Sub

Private Sub b_elifec_Click()
Dim Xlafecaborrar, Xsionoborra As String
Dim Xhayanotados As Integer
Dim id_hora_consultorio As Integer
Dim base As Integer
Xhayanotados = 0

Xlafecaborrar = InputBox("Ingrese la fecha a borrar")
If Xlafecaborrar <> "" Then
   Xsionoborra = MsgBox("Desea borrar la fecha del especialista? " & Data1.Recordset("nombre") & " BASE:" & Data1.Recordset("base"), vbExclamation + vbYesNo)
   If Xsionoborra = vbYes Then
      frm_especialistas.MousePointer = 11
      data_buscar.RecordSource = "Select * from t_fechas where fecha ='" & Xlafecaborrar & "' and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base")
      data_buscar.Refresh
      If data_buscar.Recordset.RecordCount > 0 Then
         data_buscar.Recordset.MoveFirst
         Do While Not data_buscar.Recordset.EOF
            If IsNull(data_buscar.Recordset("nom_pac")) = False And IsNull(data_buscar.Recordset("ced_pac")) = False Then
               Xhayanotados = 1
            End If
            data_buscar.Recordset.MoveNext
         Loop
         If Xhayanotados = 0 Then
            data_buscar.Recordset.MoveFirst
            Do While Not data_buscar.Recordset.EOF
               data_buscar.Recordset.Delete
               data_buscar.Recordset.MoveNext
            Loop
         Else
            If WElusuario = "COMPUTOS" Or WElusuario = "AACUÑA" Then
               data_buscar.Recordset.MoveFirst
               Do While Not data_buscar.Recordset.EOF
                  data_buscar.Recordset.Delete
                  data_buscar.Recordset.MoveNext
               Loop
            Else
               MsgBox "No se puede eliminar porque hay pacientes anotados", vbExclamation
            End If
         End If
      End If
      If Xhayanotados = 0 Then
         data_buscar.RecordSource = "Select * from t_cabfechas where fecha ='" & Xlafecaborrar & "' and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base")
         data_buscar.Refresh
         If data_buscar.Recordset.RecordCount > 0 Then
            data_buscar.Recordset.MoveFirst
            Do While Not data_buscar.Recordset.EOF
               On Error Resume Next 'si no encuentra el campo, o error en ws sigue
               base = data_buscar.Recordset("base")
               id_hora_consultorio = data_buscar.Recordset("id_hora_consultorio")
               data_buscar.Recordset.Delete
               data_buscar.Recordset.MoveNext
               'consumo ws para eliminar hora de consultorio en el caso de que tenga'
               If Not IsNull(id_hora_consultorio) Then
                reservarConsultorio = especialidadReserva(t_especsel.Text)
                If reservarConsultorio Then
                    Set obj = consumirServicio("DELETE", urlServicio & "/bases/" & base & "/consultorios/0/disponibilidades/" & id_hora_consultorio, "")
                    response = obj.responseText
                    ' MsgBox response
                End If
               End If
               On Error GoTo 0 ' desactivo error handler para que siga todo igual
            Loop
         End If
         abmesp.Recordset.AddNew
         abmesp.Recordset("fecha") = Date
         abmesp.Recordset("hora") = Format(Time, "HH:mm")
         abmesp.Recordset("usuario") = WElusuario
         abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
         abmesp.Recordset("accion") = "BORRA " & Format(Xlafecaborrar, "dd/mm/yyyy")
         abmesp.Recordset.Update
         frm_especialistas.MousePointer = 0
         MsgBox "Proceso terminado"
      Else
         If WElusuario = "COMPUTOS" Or WElusuario = "AACUÑA" Then
            data_buscar.RecordSource = "Select * from t_cabfechas where fecha ='" & Xlafecaborrar & "' and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base")
            data_buscar.Refresh
            If data_buscar.Recordset.RecordCount > 0 Then
               data_buscar.Recordset.MoveFirst
               Do While Not data_buscar.Recordset.EOF
                  data_buscar.Recordset.Delete
                  data_buscar.Recordset.MoveNext
               Loop
            End If
            abmesp.Recordset.AddNew
            abmesp.Recordset("fecha") = Date
            abmesp.Recordset("hora") = Format(Time, "HH:mm")
            abmesp.Recordset("usuario") = WElusuario
            abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
            abmesp.Recordset("accion") = "BORRA " & Format(Xlafecaborrar, "dd/mm/yyyy")
            abmesp.Recordset.Update
            frm_especialistas.MousePointer = 0
            MsgBox "Proceso terminado"
         End If
      End If
      data_cabfec.RecordSource = "Select * from t_cabfechas where cod_med =" & Data1.Recordset("cod_med") & " order by fecha"
      data_cabfec.Refresh
   End If
End If
      
   
End Sub

Private Sub b_elim_Click()
If WElusuario = "JFERNAN" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "SMPEREZ" Or WElusuario = "COMPUTOS" Then
   Dim Xestaseguro As String
   Xestaseguro = MsgBox("Desea borrar el especialista " & Data1.Recordset("nombre") & " ??", vbInformation + vbYesNo)
   If Xestaseguro = vbYes Then
      Data1.Recordset.Delete
      Data1.Refresh
      abmesp.Recordset.AddNew
      abmesp.Recordset("fecha") = Date
      abmesp.Recordset("hora") = Format(Time, "HH:mm")
      abmesp.Recordset("usuario") = WElusuario
      abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
      abmesp.Recordset("accion") = "ELIMINA ESP"
      abmesp.Recordset.Update
      
      MsgBox "REGISTRO ELIMINADO!"
   End If
Else
   MsgBox "Usuario no habilitado"
End If


End Sub

Private Sub b_excel_Click()
Frame3.Visible = False
Frame4.Visible = True
Dim Xcountt As Long
Xcountt = 1
'On Error GoTo Quepasoalimp

frm_especialistas.MousePointer = 11
t_feccab.Text = data_cabfec.Recordset("fecha")
t_codcons.Text = data_cabfec.Recordset("cod_cons")

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\infespec.mdb")
Dim Listatodo As String

MiBaseact.Execute "Delete * from lista"

data_inf.RecordSource = "lista"
data_inf.Refresh

Dim desde, hasta, Promo As String
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook
Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir3 As New Excel.Application
Dim Lafecfact As String
Lafecfact = ""
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0

'Listatodo = MsgBox("Desea listar sólo los registrados ?", vbInformation + vbYesNo)
'If Listatodo = vbYes Then
''data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and especial ='" & "HNF" & "' order by base,nro"
'Else
   data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " order by nro"
'End If

data_lista.Refresh
If data_lista.Recordset.RecordCount > 0 Then
   frm_especcovid.MousePointer = 11
   data_lista.Recordset.MoveLast
   data_lista.Recordset.MoveFirst
   
   Set Xobjexel22 = New Excel.Application
   Set Xlibexel22 = Xobjexel22.Workbooks.Add
   Set Xarchexel22 = Xlibexel22.Worksheets.Add
   Xarchexel22.Name = Trim("AgendaHNF")
   Xlibexel22.SaveAs ("C:\planillas\AgendaHNF.xls")
   Xarchtex = "C:\planillas\AgendaHNF.xls"
   Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
   Xlin = Xlin + 1
   XCol = XCol + 1
   Xarchexel22.Range("A1", "C3").Font.Size = 16
   Xarchexel22.Range("A" & Trim(str(Xlin)), "L" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
   Xarchexel22.Cells(Xlin, XCol) = "AGENDA " & Trim(data_cabfec.Recordset("nom_med")) & " DE BASE: " & Trim(str(data_cabfec.Recordset("base"))) & "  FECHA:" & Format(data_cabfec.Recordset("fecha"), "dd/mm/yyyy")
   XCol = 1
   Xlin = Xlin + 2
   Xnrocan = Xnrocan + Xlin
   Xarchexel22.Range("A" & Trim(str(Xlin)), "L" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
   Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
   Xarchexel22.Cells(Xlin, XCol) = "NRO."
   XCol = XCol + 1
   Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
   Xarchexel22.Cells(Xlin, XCol) = "HORA"
   XCol = XCol + 1
   Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 30
   Xarchexel22.Cells(Xlin, XCol) = "NOMBRE PACIENTE"
   XCol = XCol + 1
   Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 13
   Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
   XCol = XCol + 1
   Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "MATRICULA"
   XCol = XCol + 1
   Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "CELULAR"
   XCol = XCol + 1
   Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "TELEFONO"
   XCol = XCol + 1
   Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 30
   Xarchexel22.Cells(Xlin, XCol) = "DIRECCION"
   XCol = XCol + 1
   Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "LOCALIDAD"
   XCol = XCol + 1
   Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 10
   Xarchexel22.Cells(Xlin, XCol) = "SINT.?"
   XCol = XCol + 1
   Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "FECHA FACT"
   XCol = XCol + 1
   Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 12
   Xarchexel22.Cells(Xlin, XCol) = "FECHA NAC."
      
   Xlin = Xlin + 1
   XCol = 1
   
   Do While Not data_lista.Recordset.EOF
      
      data_inf.Recordset.AddNew
      data_inf.Recordset("fecha") = Format(data_lista.Recordset("fecha"), "dd/mm/yyyy")

'      If Listatodo = vbYes Then
'         data_inf.Recordset("medico") = data_lista.Recordset("nom_med") & " B." & Trim(str(data_lista.Recordset("base")))
'      Else
         data_inf.Recordset("medico") = data_lista.Recordset("nom_med")
'      End If
      data_inf.Recordset("espec") = data_lista.Recordset("especial")
      data_inf.Recordset("base") = data_lista.Recordset("base")
      data_inf.Recordset("nro") = data_lista.Recordset("nro")
      data_inf.Recordset("hora") = data_lista.Recordset("hora")
      If IsNull(data_lista.Recordset("nom_pac")) = False Then
         data_inf.Recordset("nom_pac") = data_lista.Recordset("nom_pac")
      End If
      If IsNull(data_lista.Recordset("ced_pac")) = False Then
         data_inf.Recordset("cedula") = data_lista.Recordset("ced_pac")
      End If
      If IsNull(data_lista.Recordset("mat_pac")) = False Then
         data_inf.Recordset("mat") = data_lista.Recordset("mat_pac")
      End If
      If IsNull(data_lista.Recordset("convenio")) = False Then
         data_inf.Recordset("convenio") = data_lista.Recordset("convenio")
      End If
      If IsNull(data_lista.Recordset("cel_pac")) = False Then
         data_inf.Recordset("celular") = data_lista.Recordset("cel_pac")
      End If
      If IsNull(data_lista.Recordset("tel_pac")) = False Then
         data_inf.Recordset("telef") = data_lista.Recordset("tel_pac")
      End If
      If IsNull(data_lista.Recordset("zona")) = False Then
         data_inf.Recordset("tipocons") = Mid(data_lista.Recordset("zona"), 1, 50)
      End If
      If IsNull(data_lista.Recordset("taxi")) = False Then
         data_inf.Recordset("hc") = data_lista.Recordset("taxi")
      End If
      If IsNull(data_lista.Recordset("direcci")) = False Then
         data_inf.Recordset("edad") = data_lista.Recordset("direcci")
      End If
      If IsNull(data_lista.Recordset("fec_nac")) = False Then
         data_inf.Recordset("fnac") = data_lista.Recordset("fec_nac")
      End If
'      If IsNull(data_lista.Recordset("cod_cons")) = False Then
'         data_inf.Recordset("codcons") = data_lista.Recordset("cod_cons")
'      End If
      If IsNull(data_lista.Recordset("usua_anota")) = False Then
         data_inf.Recordset("via") = Mid(data_lista.Recordset("usua_anota"), 1, 15)
      End If
      data_inf.Recordset.Update
      
      Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("nro")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("hora")
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("nom_pac")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("nom_pac")
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("ced_pac")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("ced_pac")
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("mat_pac")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("mat_pac")
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("cel_pac")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("cel_pac")
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("tel_pac")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("tel_pac")
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("direcci")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("direcci")
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("zona")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("zona")
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("taxi")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_lista.Recordset("taxi")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "NO"
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("ced_pac")) = False Then
         If Len(data_lista.Recordset("ced_pac")) = 6 Then
            Lafecfact = Devuelve_fecha_fact(Mid(data_lista.Recordset("ced_pac"), 1, 6))
         Else
            Lafecfact = Devuelve_fecha_fact(Mid(data_lista.Recordset("ced_pac"), 1, 7))
         End If
         If Trim(Lafecfact) <> "" Then
            Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Lafecfact, "dd/mm/yyyy")
         End If
      End If
      XCol = XCol + 1
      If IsNull(data_lista.Recordset("fec_nac")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_lista.Recordset("fec_nac"), "dd/mm/yyyy")
      End If
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      
      data_lista.Recordset.MoveNext
   Loop
   
    Xlin = Xlin + 1
    XCol = 1
    Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
    Xlin = Xlin + 1
    XCol = 1
    Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
    Xlibexel22.Save
    Xlibexel22.Close
    Xobjexel22.Quit
    Xlabrir3.Workbooks.Open Xarchtex, , False
    Xlabrir3.Visible = True
    Xlabrir3.WindowState = xlMaximized
   
   frm_especcovid.MousePointer = 0
   data_inf.RecordSource = "Select * from lista"
   data_inf.Refresh
   MsgBox "Proceso terminado", vbInformation
'   cr1.ReportFileName = App.path & "\inflistahnf.rpt"
'   cr1.Action = 1
      

Else
   frm_especialistas.MousePointer = 0
   MsgBox "No existe fecha"
End If

End Sub

Private Sub b_graba_Click()
Dim Xhd, xhh, Xmh, Xpac, Xlahor As Integer
Dim Xtexhora, Xtexth As String
Xtexhora = ""
Xlahor = 0
Xpac = 0
If XAlta = 1 Then
   If cbobase.ListIndex >= 0 And cbomedico.ListIndex >= 0 Then
      data_buscar.RecordSource = "Select * from medicos_esp where nom_med ='" & cbomedico.Text & "'"
      data_buscar.Refresh
      If data_buscar.Recordset.RecordCount > 0 Then
         Data1.Recordset.AddNew
         Data1.Recordset("especialidad") = data_buscar.Recordset("esp_med")
         Data1.Recordset("nombre") = data_buscar.Recordset("nom_med")
         Data1.Recordset("cod_med") = data_buscar.Recordset("id")
         Data1.Recordset("base") = Val(cbobase.Text)
         data_buscar.RecordSource = "Select * from bases_sapp where nro_base =" & cbobase.Text
         data_buscar.Refresh
         If data_buscar.Recordset.RecordCount > 0 Then
            Data1.Recordset("basedesc") = data_buscar.Recordset("desc_base")
         Else
            Data1.Recordset("basedesc") = "NO ENCONTRADO"
         End If
         If t_cantp.Text <> "" And t_mm.Text <> "" Then
            If t_cantp.Text > 0 Then
               If t_mm.Text > 0 Then
                  Xhd = Val(Mid(mhini.Text, 1, 2))
                  Xmh = Val(Mid(mhini.Text, 4, 2))
                  Xpac = 1
                  Do While Xpac <= Val(t_cantp.Text)
                     Xmh = Xmh + t_mm.Text
                     If Xmh >= 60 Then
                        If Xmh > 60 Then
                           Xmh = Xmh - 60
                        Else
                           Xmh = 0
                        End If
                        Xhd = Xhd + 1
                     End If
                     Xpac = Xpac + 1
                  Loop
                  If Xhd <= 9 Then
                     If Xmh <= 9 Then
                        Xtexhora = "0" & Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                     Else
                        Xtexhora = "0" & Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                     End If
                  Else
                     If Xmh <= 9 Then
                        Xtexhora = Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                     Else
                        Xtexhora = Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                     End If
                  End If
                  mhfin.Text = Format(Xtexhora, "HH:mm")
               End If
            Else
               If t_mm.Text > 0 Then
                  If t_mm.Text > 60 Then
                     Xlahor = t_mm.Text / 60
                     Xhd = Val(Mid(mhini.Text, 1, 2))
                     Xmh = Val(Mid(mhini.Text, 4, 2))
                     Xtexth = mhini.Text
                     Xpac = 0
                     Do While Xtexth <> mhfin.Text
                        Xmh = 0
                        Xhd = Xhd + Xlahor
                        Xpac = Xpac + 1
                        If Xhd <= 9 Then
                           If Xmh <= 9 Then
                              Xtexhora = "0" & Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                           Else
                              Xtexhora = "0" & Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                           End If
                        Else
                           If Xmh <= 9 Then
                              Xtexhora = Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                           Else
                              Xtexhora = Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                           End If
                        End If
                        Xtexth = Format(Xtexhora, "HH:mm")
                     Loop
                     t_cantp.Text = Xpac
                  Else
                     Xhd = Val(Mid(mhini.Text, 1, 2))
                     Xmh = Val(Mid(mhini.Text, 4, 2))
                     Xtexth = mhini.Text
                     Xpac = 0
                     Do While Xtexth <> mhfin.Text
                        Xmh = Xmh + t_mm.Text
                        If Xmh >= 60 Then
                           If Xmh > 60 Then
                              Xmh = Xmh - 60
                           Else
                              Xmh = 0
                           End If
                           Xhd = Xhd + 1
                        End If
                        Xpac = Xpac + 1
                        If Xhd <= 9 Then
                           If Xmh <= 9 Then
                              Xtexhora = "0" & Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                           Else
                              Xtexhora = "0" & Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                           End If
                        Else
                           If Xmh <= 9 Then
                              Xtexhora = Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                           Else
                              Xtexhora = Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                           End If
                        End If
                        Xtexth = Format(Xtexhora, "HH:mm")
                     Loop
                     t_cantp.Text = Xpac
                  End If
               End If
            
            End If
         End If
         If t_cantp.Text <> "" Then
            Data1.Recordset("cantpac") = t_cantp.Text
         Else
            Data1.Recordset("cantpac") = 0
         End If
         If mhini.Text <> "__:__" Then
            Data1.Recordset("horaini") = mhini.Text
         Else
            Data1.Recordset("horaini") = "00:00"
         End If
         If mhfin.Text <> "__:__" Then
            Data1.Recordset("horafin") = mhfin.Text
         Else
            Data1.Recordset("horafin") = "00:00"
         End If
         If Trim(t_mm.Text) <> "" Then
            Data1.Recordset("minutos") = t_mm.Text
         Else
            Data1.Recordset("minutos") = 0
         End If
         If Trim(t_espera.Text) <> "" Then
            Data1.Recordset("espera") = t_espera.Text
         Else
            Data1.Recordset("espera") = 0
         End If
         Data1.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
         Data1.Recordset.Update
         Data1.Refresh
         abmesp.Recordset.AddNew
         abmesp.Recordset("fecha") = Date
         abmesp.Recordset("hora") = Format(Time, "HH:mm")
         abmesp.Recordset("usuario") = WElusuario
         abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
         abmesp.Recordset("accion") = "CREA ESPEC"
         abmesp.Recordset.Update
         
         b_nuevo.Enabled = True
         b_modif.Enabled = True
         b_graba.Enabled = False
         b_cance.Enabled = False
         b_elim.Enabled = True
         b_edimed.Enabled = True
         b_infos.Enabled = True
         DBGrid1.Enabled = True
         Frame1.Enabled = False
         DBGrid1.SetFocus
         Frame2.Enabled = True
         Frame4.Enabled = True
      Else
         MsgBox "No se encontró el medico seleccionado"
      End If
   Else
      MsgBox "No seleccionó BASE o MEDICO"
   End If
End If
If XAlta = 0 Then
      data_buscar.RecordSource = "Select * from medicos_esp where nom_med ='" & cbomedico.Text & "'"
      data_buscar.Refresh
      If data_buscar.Recordset.RecordCount > 0 Then
         Data1.Recordset.Edit
         Data1.Recordset("especialidad") = data_buscar.Recordset("esp_med")
         Data1.Recordset("nombre") = data_buscar.Recordset("nom_med")
         Data1.Recordset("cod_med") = data_buscar.Recordset("id")
         Data1.Recordset("base") = Val(cbobase.Text)
         data_buscar.RecordSource = "Select * from bases_sapp where nro_base =" & cbobase.Text
         data_buscar.Refresh
         If data_buscar.Recordset.RecordCount > 0 Then
            Data1.Recordset("basedesc") = data_buscar.Recordset("desc_base")
         Else
            Data1.Recordset("basedesc") = "NO ENCONTRADO"
         End If
         If t_cantp.Text <> "" And t_mm.Text <> "" Then
            If t_cantp.Text > 0 Then
               If t_mm.Text > 0 Then
                  Xhd = Val(Mid(mhini.Text, 1, 2))
                  Xmh = Val(Mid(mhini.Text, 4, 2))
                  Xpac = 1
                  Do While Xpac <= Val(t_cantp.Text)
                     Xmh = Xmh + t_mm.Text
                     If Xmh >= 60 Then
                        If Xmh > 60 Then
                           Xmh = Xmh - 60
                        Else
                           Xmh = 0
                        End If
                        Xhd = Xhd + 1
                     End If
                     Xpac = Xpac + 1
                  Loop
                  If Xhd <= 9 Then
                     If Xmh <= 9 Then
                        Xtexhora = "0" & Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                     Else
                        Xtexhora = "0" & Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                     End If
                  Else
                     If Xmh <= 9 Then
                        Xtexhora = Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                     Else
                        Xtexhora = Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                     End If
                  End If
                  mhfin.Text = Format(Xtexhora, "HH:mm")
               End If
            Else
               If t_mm.Text > 0 Then
                  Xhd = Val(Mid(mhini.Text, 1, 2))
                  Xmh = Val(Mid(mhini.Text, 4, 2))
                  Xtexth = mhini.Text
                  Xpac = 0
                  Do While Xtexth <> mhfin.Text
                     Xmh = Xmh + t_mm.Text
                     If Xmh >= 60 Then
                        If Xmh > 60 Then
                           Xmh = Xmh - 60
                        Else
                           Xmh = 0
                        End If
                        Xhd = Xhd + 1
                     End If
                     Xpac = Xpac + 1
                     If Xhd <= 9 Then
                        If Xmh <= 9 Then
                           Xtexhora = "0" & Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                        Else
                           Xtexhora = "0" & Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                        End If
                     Else
                        If Xmh <= 9 Then
                           Xtexhora = Trim(str(Xhd)) & ":" & "0" & Trim(str(Xmh))
                        Else
                           Xtexhora = Trim(str(Xhd)) & ":" & Trim(str(Xmh))
                        End If
                     End If
                     Xtexth = Format(Xtexhora, "HH:mm")
                  Loop
                  t_cantp.Text = Xpac
               End If
            End If
         End If
         If t_cantp.Text <> "" Then
            Data1.Recordset("cantpac") = t_cantp.Text
         Else
            Data1.Recordset("cantpac") = 0
         End If
         If mhini.Text <> "__:__" Then
            Data1.Recordset("horaini") = mhini.Text
         Else
            Data1.Recordset("horaini") = "00:00"
         End If
         If mhfin.Text <> "__:__" Then
            Data1.Recordset("horafin") = mhfin.Text
         Else
            Data1.Recordset("horafin") = "00:00"
         End If
         If Trim(t_mm.Text) <> "" Then
            Data1.Recordset("minutos") = t_mm.Text
         Else
            Data1.Recordset("minutos") = 0
         End If
         If Trim(t_espera.Text) <> "" Then
            Data1.Recordset("espera") = t_espera.Text
         Else
            Data1.Recordset("espera") = 0
         End If
         Data1.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
         Data1.Recordset.Update
         Data1.Refresh
         abmesp.Recordset.AddNew
         abmesp.Recordset("fecha") = Date
         abmesp.Recordset("hora") = Format(Time, "HH:mm")
         abmesp.Recordset("usuario") = WElusuario
         abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
         abmesp.Recordset("accion") = "MODIF ESPEC"
         abmesp.Recordset.Update
         
         Data1.Recordset.FindFirst " id=" & t_idant.Text
         
         b_nuevo.Enabled = True
         b_modif.Enabled = True
         b_graba.Enabled = False
         b_cance.Enabled = False
         b_elim.Enabled = True
         b_edimed.Enabled = True
         b_infos.Enabled = True
         DBGrid1.Enabled = True
         Frame1.Enabled = False
         DBGrid1.SetFocus
         Frame2.Enabled = True
         Frame4.Enabled = True
      Else
         MsgBox "No se encontró el medico seleccionado"
      End If
End If
XAlta = 0
         

End Sub

Private Sub b_grabmed_Click()
If XAlta = 1 Then
   If t_nom.Text <> "" Then
      data_medicossapp.Recordset.FindFirst "nom_med ='" & t_nom.Text & "'"
      If Not data_medicossapp.Recordset.NoMatch Then
         MsgBox "Ya existe nombre de médico, VERIFIQUE!!"
      Else
         data_medicossapp.Recordset.AddNew
         data_medicossapp.Recordset("nom_med") = t_nom.Text
         data_medicossapp.Recordset("esp_med") = cboespec.Text
         If t_codsapp.Text <> "" Then
            data_medicossapp.Recordset("cod_sapp") = t_codsapp.Text
         End If
         data_medicossapp.Recordset("esp_cod") = cboespec.ListIndex
         data_medicossapp.Recordset.Update
         data_medicossapp.Refresh
         XAlta = 0
         b_modmed.Enabled = True
         b_grabmed.Enabled = False
         b_canmed.Enabled = False
         b_altamed.Enabled = True
         DBGrid2.Enabled = True
         DBGrid2.SetFocus
      End If
   Else
      MsgBox "Ingrese nombre para poder grabar"
   End If
Else
   If t_nom.Text <> "" Then
      data_medicossapp.Recordset.FindFirst "id =" & labcod.Caption
      If Not data_medicossapp.Recordset.NoMatch Then
         data_medicossapp.Recordset.Edit
         data_medicossapp.Recordset("nom_med") = t_nom.Text
         data_medicossapp.Recordset("esp_med") = cboespec.Text
         data_medicossapp.Recordset("cod_sapp") = t_codsapp.Text
         data_medicossapp.Recordset("esp_cod") = cboespec.ListIndex
         data_medicossapp.Recordset.Update
         XAlta = 0
         b_modmed.Enabled = True
         b_grabmed.Enabled = False
         b_canmed.Enabled = False
         b_altamed.Enabled = True
         DBGrid2.Enabled = True
         DBGrid2.SetFocus
      Else
         XAlta = 0
         b_modmed.Enabled = True
         b_grabmed.Enabled = False
         b_canmed.Enabled = False
         b_altamed.Enabled = True
         DBGrid2.Enabled = True
         DBGrid2.SetFocus
      End If
   Else
      MsgBox "Ingrese nombre para poder grabar."
   End If
End If

End Sub

Private Sub b_impcons_Click()
Dim Xaexcel As String

Xaexcel = MsgBox("Desea generar a excel?(No en base)", vbInformation + vbYesNo, "Agenda")
If Xaexcel = vbYes Then
   b_excel_Click
Else
    Frame3.Visible = False
    Frame4.Visible = True
    Dim Xcountt As Long
    Xcountt = 1
    'On Error GoTo Quepasoalimp
    
    frm_especialistas.MousePointer = 11
    t_feccab.Text = data_cabfec.Recordset("fecha")
    t_codcons.Text = data_cabfec.Recordset("cod_cons")
    
    Dim MiBaseact As Database
    Dim Unasesact As Workspace
    Set Unasesact = Workspaces(0)
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\infespec.mdb")
    Dim Listatodo As String
    
    MiBaseact.Execute "Delete * from lista"
    
    data_inf.RecordSource = "lista"
    data_inf.Refresh
        
    'Listatodo = MsgBox("Desea listar sólo los registrados ?", vbInformation + vbYesNo)
    'If Listatodo = vbYes Then
    ''data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and especial ='" & "HNF" & "' order by base,nro"
    'Else
       data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " order by nro"
    'End If
    
    data_lista.Refresh
    If data_lista.Recordset.RecordCount > 0 Then
       frm_especcovid.MousePointer = 11
       data_lista.Recordset.MoveLast
       data_lista.Recordset.MoveFirst
              
       Do While Not data_lista.Recordset.EOF
          
          data_inf.Recordset.AddNew
          data_inf.Recordset("fecha") = Format(data_lista.Recordset("fecha"), "dd/mm/yyyy")
    
    '      If Listatodo = vbYes Then
    '         data_inf.Recordset("medico") = data_lista.Recordset("nom_med") & " B." & Trim(str(data_lista.Recordset("base")))
    '      Else
             data_inf.Recordset("medico") = data_lista.Recordset("nom_med")
    '      End If
          data_inf.Recordset("espec") = data_lista.Recordset("especial")
          data_inf.Recordset("base") = data_lista.Recordset("base")
          data_inf.Recordset("nro") = data_lista.Recordset("nro")
          data_inf.Recordset("hora") = data_lista.Recordset("hora")
          If IsNull(data_lista.Recordset("nom_pac")) = False Then
             data_inf.Recordset("nom_pac") = data_lista.Recordset("nom_pac")
          End If
          If IsNull(data_lista.Recordset("ced_pac")) = False Then
             data_inf.Recordset("cedula") = data_lista.Recordset("ced_pac")
          End If
          If IsNull(data_lista.Recordset("mat_pac")) = False Then
             data_inf.Recordset("mat") = data_lista.Recordset("mat_pac")
          End If
          If IsNull(data_lista.Recordset("convenio")) = False Then
             data_inf.Recordset("convenio") = data_lista.Recordset("convenio")
          End If
          If IsNull(data_lista.Recordset("cel_pac")) = False Then
             data_inf.Recordset("celular") = data_lista.Recordset("cel_pac")
          End If
          If IsNull(data_lista.Recordset("tel_pac")) = False Then
             data_inf.Recordset("telef") = data_lista.Recordset("tel_pac")
          End If
          If IsNull(data_lista.Recordset("zona")) = False Then
             data_inf.Recordset("tipocons") = Mid(data_lista.Recordset("zona"), 1, 50)
          End If
          If IsNull(data_lista.Recordset("taxi")) = False Then
             data_inf.Recordset("hc") = data_lista.Recordset("taxi")
          End If
          If IsNull(data_lista.Recordset("direcci")) = False Then
             data_inf.Recordset("edad") = data_lista.Recordset("direcci")
          End If
          If IsNull(data_lista.Recordset("fec_nac")) = False Then
             data_inf.Recordset("fnac") = data_lista.Recordset("fec_nac")
          End If
    '      If IsNull(data_lista.Recordset("cod_cons")) = False Then
    '         data_inf.Recordset("codcons") = data_lista.Recordset("cod_cons")
    '      End If
          If IsNull(data_lista.Recordset("usua_anota")) = False Then
             data_inf.Recordset("via") = Mid(data_lista.Recordset("usua_anota"), 1, 15)
          End If
          data_inf.Recordset.Update
                    
          data_lista.Recordset.MoveNext
       Loop
       
       
       frm_especcovid.MousePointer = 0
       data_inf.RecordSource = "Select * from lista"
       data_inf.Refresh
       MsgBox "Proceso terminado", vbInformation
       cr1.ReportFileName = App.path & "\inflistahnf.rpt"
       cr1.Action = 1
          
    
    Else
       frm_especialistas.MousePointer = 0
       MsgBox "No existe fecha"
    End If
End If

'Exit Sub

'Quepasoalimp:
'             If Err.Number = 91 Then
'                MsgBox "Verifique si seleccionó la consulta"
'             Else
'                MsgBox "Verifique si tiene datos seleccionados"
'             End If
             
End Sub

Private Sub b_infos_Click()
'frm_infespenew.Show vbModal
MsgBox "Sin opción para informes"

End Sub

Private Sub b_modif_Click()
XAlta = 0
b_nuevo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_elim.Enabled = False
b_edimed.Enabled = False
b_infos.Enabled = False
DBGrid1.Enabled = False
Frame1.Enabled = True
cbobase.SetFocus

End Sub

Private Sub b_modmed_Click()
If labcod.Caption <> "" Then
    XAlta = 2
    b_modmed.Enabled = False
    b_grabmed.Enabled = True
    b_canmed.Enabled = True
    b_altamed.Enabled = False
    t_nom.SetFocus
    DBGrid2.Enabled = False
Else
    MsgBox "Seleccione registro para modificar"
    DBGrid2.SetFocus
    
End If

End Sub


Private Sub b_nuevafecha_Click()
Dim Xmensacrear, Xfecstr, Xhoracomi As String
Dim Xhorah, Xhorad, Xminh As Integer
Dim XCantp, Xespera, Xesperah As Integer
Dim Xmin, Xhor As String
Dim Xnropac As Integer
Dim Xcodconscrea As Double
Dim Xlafecanota As Date
Dim Xlafechasta2, Xfec1, Xfec2 As Date
Dim ultimo, Xdiasmes1, Xdiasmes2, Xmes1, Xmes2, Xano1, Xano2 As Integer
Dim Xdiasdif As Integer
Dim sToSend As String
Dim obj As Object
Dim response As String
Dim reservar As Boolean

Me.Pass_id_hora_reserva = 0
Xnropac = 1
If mnuevaf.Text <> "__/__/____" Then
    ' me fijo si la especialidad seleccionada reserva consultorio (ws.mdb)'
    reservarConsultorio = especialidadReserva(t_especsel.Text)
    ' ---- datos para Consumo web service ----
    base = nroBase_idBase.Item(Val(t_basesel.Text))
    medico = Val(t_codmedsel.Text)
    XfecstrGuiones = Format(mnuevaf.Text, "yyyy-mm-dd")
    horaInicio = mhini.Text
    horaFin = mhfin.Text
    
    If Format(mnuevaf.Text, "yyyy/mm/dd") > Format(Date, "yyyy/mm/dd") Then
        Xfecstr = Format(mnuevaf.Text, "dd/mm/yyyy")
        Xmensacrear = MsgBox("Desea crear las fechas para " & t_especsel.Text & " Sol." & cbomedico.Text & "?", vbInformation + vbYesNo)
        If Xmensacrear = vbYes Then
            frm_especcovid.MousePointer = 11
            data_buscar.RecordSource = "Select * from t_cabfechas order by id"
            data_buscar.Refresh
            If data_buscar.Recordset.RecordCount > 0 Then
                data_buscar.Recordset.MoveLast
                Xcodconscrea = data_buscar.Recordset("id") + 1
            Else
                Xcodconscrea = 1
            End If
            data_fechas.RecordSource = "Select * from t_fechas where fecha ='" & Trim(Xfecstr) & "' and cod_med =" & Val(t_codmedsel.Text) & " and base =" & Val(t_basesel.Text) & " and hora ='" & mhini.Text & "'"
            data_fechas.Refresh
            If data_fechas.Recordset.RecordCount > 0 Then
                frm_especcovid.MousePointer = 0
                MsgBox "Ya existe consulta creada con éstos parámetros, VERIFIQUE!!", vbInformation
            Else
                Xhorah = Val(Mid(mhfin.Text, 1, 2))
                Xhorad = Val(Mid(mhini.Text, 1, 2))
                Xminh = Val(Mid(mhini.Text, 4, 2))
                If t_espera.Text = "" Then
                   t_espera.Text = 0
                End If
                Xespera = t_espera.Text
                Do While Xnropac <= Val(t_cantp.Text)
                   If Xminh = 0 Then
                      Xmin = "00"
                   Else
                      If Xminh <= 9 Then
                         Xmin = "0" & Trim(str(Xminh))
                      Else
                         Xmin = Trim(str(Xminh))
                      End If
                   End If
                   If Xhorad = 0 Then
                      Xhor = "00"
                   Else
                      If Xhorad <= 9 Then
                         Xhor = "0" & Trim(str(Xhorad))
                      Else
                         Xhor = Trim(str(Xhorad))
                      End If
                   End If
                   data_fechas.Recordset.AddNew
                   data_fechas.Recordset("fecha") = Xfecstr
                   data_fechas.Recordset("hora") = Trim(Xhor) & ":" & Trim(Xmin)
                   data_fechas.Recordset("nro") = Xnropac
                   data_fechas.Recordset("cod_med") = Val(t_codmedsel.Text)
                   data_fechas.Recordset("nom_med") = cbomedico.Text
                   data_fechas.Recordset("especial") = t_especsel.Text
                   data_fechas.Recordset("base") = Val(t_basesel.Text)
                   data_fechas.Recordset("cod_cons") = Xcodconscrea
                   data_fechas.Recordset("hora_com") = mhini.Text
                   data_fechas.Recordset("cancela") = "NO"
                   data_fechas.Recordset("desc_local") = t_basedescsel.Text
                   data_fechas.Recordset.Update
                   Xnropac = Xnropac + 1
                   Xminh = Xminh + t_mm.Text
                   If Xminh >= 60 Then
                      If Xminh >= 120 Then
                         Xhorad = Xhorad + 2
                         Xminh = 0
                      Else
                         If Xminh > 60 Then
                            Xminh = Xminh - 60
                         Else
                            Xminh = 0
                         End If
                         Xhorad = Xhorad + 1
                      End If
                   End If
                Loop
                Xesperah = 0
                Do While Xesperah < Xespera
                   Xesperah = Xesperah + 1
                   data_fechas.Recordset.AddNew
                   data_fechas.Recordset("fecha") = Xfecstr
                   data_fechas.Recordset("hora") = "00:00"
                   data_fechas.Recordset("nro") = 99
                   data_fechas.Recordset("cod_med") = Val(t_codmedsel.Text)
                   data_fechas.Recordset("nom_med") = cbomedico.Text
                   data_fechas.Recordset("especial") = t_especsel.Text
                   data_fechas.Recordset("base") = Val(t_basesel.Text)
                   data_fechas.Recordset("nom_pac") = "LISTA DE ESPERA"
                   data_fechas.Recordset("cod_cons") = Xcodconscrea
                   data_fechas.Recordset("hora_com") = mhini.Text
                   data_fechas.Recordset("cancela") = "NO"
                   data_fechas.Recordset("desc_local") = t_basedescsel.Text
                   data_fechas.Recordset.Update
                Loop
                data_cabfec.Recordset.AddNew
                data_cabfec.Recordset("id") = Xcodconscrea
                data_cabfec.Recordset("fecha") = Xfecstr
                data_cabfec.Recordset("hora") = mhini.Text
                data_cabfec.Recordset("cod_med") = Val(t_codmedsel.Text)
                data_cabfec.Recordset("cod_cons") = Xcodconscrea
                data_cabfec.Recordset("des_fecha") = FormatDateTime(mnuevaf.Text, vbLongDate)
                data_cabfec.Recordset("base") = Val(t_basesel.Text)
                data_cabfec.Recordset("especial") = t_especsel.Text
                data_cabfec.Recordset("nom_med") = cbomedico.Text
                data_cabfec.Recordset("base_desc") = t_basedescsel.Text
                data_cabfec.Recordset("cancela") = "NO"
                data_cabfec.Recordset("hora_fin") = mhfin.Text
                data_cabfec.Recordset("cant_pac") = Xnropac - 1
                
                'agrego id de reserva'
                ultimo = Day(DateSerial(Year(Xfecstr), Month(Xfecstr) + 1, 0))
                If Month(Xfecstr) = 12 Then
                   Xmes1 = 1
                   Xmes2 = 2
                   Xano1 = Year(Xfecstr) + 1
                   Xano2 = Year(Xfecstr) + 1
                Else
                   If Month(Xfecstr) = 11 Then
                      Xmes1 = Month(Xfecstr) + 1
                      Xano1 = Year(Xfecstr)
                      Xmes2 = 1
                      Xano2 = Year(Xfecstr) + 1
                   Else
                      Xmes1 = Month(Xfecstr) + 1
                      Xano1 = Year(Xfecstr)
                      Xmes2 = Month(Xfecstr) + 1
                      Xano2 = Year(Xfecstr)
                   End If
                End If
                Xdiasmes1 = Day(DateSerial(Xano1, Xmes1 + 1, 0))
                Xdiasmes2 = Day(DateSerial(Xano2, Xmes2 + 1, 0))
                 
                If ultimo > 9 Then
                   If Month(Xfecstr) > 9 Then
                      Xlafecanota = CDate(Trim(str(ultimo)) & "/" & Trim(str(Month(Xfecstr))) & "/" & Trim(str(Year(Xfecstr))))
                   Else
                      Xlafecanota = CDate(Trim(str(ultimo)) & "/0" & Trim(str(Month(Xfecstr))) & "/" & Trim(str(Year(Xfecstr))))
                   End If
                Else
                   If Month(Xfecstr) > 9 Then
                      Xlafecanota = CDate("0" & Trim(str(ultimo)) & "/" & Trim(str(Month(Xfecstr))) & "/" & Trim(str(Year(Xfecstr))))
                   Else
                      Xlafecanota = CDate("0" & Trim(str(ultimo)) & "/0" & Trim(str(Month(Xfecstr))) & "/" & Trim(str(Year(Xfecstr))))
                   End If
                End If
                Xfec1 = CDate(Xfecstr)
                Xfec2 = CDate(Xlafecanota)
                Xdiasdif = DateDiff("d", Xfec1, Xfec2)
                      Xdiasdif = Xdiasdif + Xdiasmes1 + Xdiasmes2
                      Xlafecanota = CDate(mnuevaf.Text) - 11
'                      Xlafecanota = CDate(Xfecstr) - Xdiasdif
                      
                      data_cabfec.Recordset("fecha_anota") = Format(Xlafecanota, "dd/mm/yyyy")
                      data_cabfec.Recordset.Update
                ''frm_especialistas.MousePointer = 0
                MsgBox "CONSULTA CREADA", vbInformation
                frm_especcovid.MousePointer = 0
            End If
        End If
    End If
End If
frm_especcovid.MousePointer = 0



End Sub




Private Sub b_nuevo_Click()
XAlta = 1
b_nuevo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_elim.Enabled = False
b_edimed.Enabled = False
b_infos.Enabled = False
DBGrid1.Enabled = False
Frame1.Enabled = True
cbobase.ListIndex = -1
cbomedico.Text = ""
mhini.Text = "__:__"
mhfin.Text = "__:__"
t_espera.Text = ""
t_mm.Text = ""
t_cantp.Text = ""

cbobase.SetFocus
Frame3.Visible = False
Frame2.Enabled = False
Frame4.Enabled = False

End Sub

Private Sub cboespec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codsapp.SetFocus
End If

End Sub

Private Sub cbosino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_direc.SetFocus
End If

End Sub

Private Sub cbotipcons_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbosino.SetFocus
End If

End Sub



Private Sub Command1_Click()
Dim Xind, Xcant, Xnro As Long
Dim Xfecdeuda, Xlafechacons As Date
Dim Xloslabos, Xlacedconsulta As String
Dim Xellugar As Integer

Xloslabos = ""
Xlafechacons = Date

Xfecdeuda = Date - 30
Xind = 0
Xnro = 0
Xcant = 0
Dim Xcountt As Long
Dim Xdeudasiono As Integer

End Sub

Private Sub Command2_Click()
Dim Xind As Integer
Dim Xnro As String

For Xind = 1 To ListView2.ListItems.count
    If ListView2.ListItems.Item(Xind).Selected = True Then
       Xcant = Xcant + 1
       Xnro = ListView2.ListItems(Xind).SubItems(4)
    End If
Next Xind
Xind = 0
MsgBox "ES: " & Xnro & "CANT:" & Xcant

Dim Xfecconsdesp As String
Xfecconsdesp = Format(Date, "dd/mm/yyyy")

frm_especialistas.MousePointer = 11

If Check1.Value = 1 Then
   If WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or WElusuario = "AACUÑA" Or WElusuario = "SMPEREZ" Then
      data_cabfec.RecordSource = "Select * from t_cabfechas where cod_med =" & t_codmedsel.Text & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(fecha)"
      data_cabfec.Refresh
   Else
'      data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(fecha)"
      data_cabfec.RecordSource = "Select * from t_cabfechas where cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(fecha)"
      data_cabfec.Refresh
   End If
Else
   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " order by cdate(fecha)"
   data_cabfec.Refresh
End If

frm_especialistas.MousePointer = 0

End Sub

Private Sub DBGrid1_DblClick()

frm_especcovid.MousePointer = 11
If Frame4.Visible = True Then
   ListView1.ListItems.Clear
End If
labselfec.Caption = ""

t_idant.Text = Data1.Recordset("id")
t_codmedsel.Text = Data1.Recordset("cod_med")
t_basesel.Text = Data1.Recordset("base")
t_especsel.Text = Data1.Recordset("especialidad")
t_basedescsel.Text = Data1.Recordset("basedesc")

If IsNull(Data1.Recordset("base")) = False Then
   cbobase.Text = Data1.Recordset("base")
Else
   cbobase.ListIndex = -1
End If
If IsNull(Data1.Recordset("nombre")) = False Then
   cbomedico.Text = Data1.Recordset("nombre")
Else
   cbomedico.ListIndex = -1
End If
If IsNull(Data1.Recordset("horaini")) = False Then
   mhini.Text = Data1.Recordset("horaini")
Else
   mhini.Text = "__:__"
End If
If IsNull(Data1.Recordset("horafin")) = False Then
   mhfin.Text = Data1.Recordset("horafin")
Else
   mhfin.Text = "__:__"
End If
If IsNull(Data1.Recordset("minutos")) = False Then
   t_mm.Text = Data1.Recordset("minutos")
Else
   t_mm.Text = ""
End If
If IsNull(Data1.Recordset("cantpac")) = False Then
   t_cantp.Text = Data1.Recordset("cantpac")
Else
   t_cantp.Text = ""
End If
If IsNull(Data1.Recordset("espera")) = False Then
   t_espera.Text = Data1.Recordset("espera")
Else
   t_espera.Text = ""
End If

Dim Xfecconsdesp As String
Xfecconsdesp = Format(Date, "dd/mm/yyyy")
Dim strSQL As String
strSQL = "SELECT t.id,t.fecha,t.hora,t.cod_med,t.cod_cons,t.des_fecha,t.base,t.especial,t.nom_med,t.base_desc,t.cancela,t.motivo,t.usuario,t.fecha_can,t.hora_can,t.cant_pac,t.hora_fin,t.cant_pacok,t.fecha_anota,t.enviado,t.consult,t.id_hora_consultorio, c.desc_consultorio FROM (t_cabfechas t " & _
          "LEFT JOIN horas_consultorios h on t.id_hora_consultorio = h.id) " & _
          "LEFT JOIN consultorios c ON c.id=h.id_consultorio " & _
          "WHERE t.cod_med = " & t_codmedsel.Text & _
              "AND t.base = " & t_basesel.Text & _
              " AND cdate(t.fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & _
              "# ORDER BY cdate(t.fecha)"


If Check1.Value = 1 Then
   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(fecha)"
'      data_cabfec.RecordSource = "Select * from t_cabfechas where cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(fecha)"
'      data_cabfec.RecordSource = "Select * from t_cabfechas t where t.cod_med =" & t_codmedsel.Text & " and t.base =" & t_basesel.Text & " and cdate(t.fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(t.fecha)"
      
'      data_cabfec.RecordSource = strSQL
   data_cabfec.Refresh
Else
'   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " order by cdate(fecha)"
'   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & t_codmedsel.Text & " and base =" & t_basesel.Text & " order by cdate(fecha)"
   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format("16/04/2021", "yyyy/mm/dd") & "# order by cdate(fecha)"
   data_cabfec.Refresh
   
End If
labussi.Caption = ""
If Data1.Recordset("nombre") = "HNF Domicilio" Then
   If ControlUsuario("Agenda hisopados domicilio") = 1 Then
      b_elianota.Enabled = True
      b_agrega.Enabled = True
      labussi.Caption = "SI"
   Else
      b_elianota.Enabled = False
      b_agrega.Enabled = False
   End If
Else
   If ControlUsuario("Agenda hisopados") = 1 Then
      labussi.Caption = "SI"
      b_elianota.Enabled = True
      b_agrega.Enabled = True
   Else
      b_elianota.Enabled = True
      b_agrega.Enabled = True
   End If
End If

frm_especcovid.MousePointer = 0

End Sub

Private Sub DBGrid2_DblClick()
labcod.Caption = data_medicossapp.Recordset("id")
t_nom.Text = data_medicossapp.Recordset("nom_med")
If IsNull(data_medicossapp.Recordset("esp_cod")) = False Then
   cboespec.ListIndex = data_medicossapp.Recordset("esp_cod")
Else
   cboespec.ListIndex = -1
End If
If IsNull(data_medicossapp.Recordset("cod_sapp")) = False Then
   t_codsapp.Text = data_medicossapp.Recordset("cod_sapp")
Else
   t_codsapp.Text = ""
End If


End Sub


Private Sub DBGrid3_DblClick()
Frame3.Visible = False
Frame4.Visible = True
Dim Xcountt As Long
Dim Xcantlibres, Xcansin As Integer

Xcountt = 1
Xcantlibres = 0
Xcansin = 0
frm_especcovid.MousePointer = 11
t_feccab.Text = data_cabfec.Recordset("fecha")
t_codcons.Text = data_cabfec.Recordset("cod_cons")
labselfec.Caption = data_cabfec.Recordset("des_fecha")

If CDate(data_cabfec.Recordset("fecha_anota")) <= Format(Date, "yyyy/mm/dd") Then
   ListView1.Enabled = True
   b_agrega.Enabled = True
   b_elianota.Enabled = True
Else
   MsgBox "Fecha habilitada para anotar a partir de: " & data_cabfec.Recordset("fecha_anota")
   ListView1.Enabled = False
   b_agrega.Enabled = False
   b_elianota.Enabled = False
End If

If data_cabfec.Recordset("cancela") = "SI" Then
   MsgBox "La consulta se encuentra CANCELADA, solo puede imprimir la lista", vbExclamation
   ListView1.Enabled = False
   b_agrega.Enabled = False
   b_elianota.Enabled = False
Else
   ListView1.Enabled = True
   b_agrega.Enabled = True
   b_elianota.Enabled = True
End If

data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " order by nro"
data_lista.Refresh
If data_lista.Recordset.RecordCount > 0 Then
   data_lista.Recordset.MoveFirst
   ListView1.ListItems.Clear
   Do While Not data_lista.Recordset.EOF
      If Xcantlibres <= 0 Then
        If IsNull(data_lista.Recordset("nro")) = False Then
           ListView1.ListItems.Add Xcountt, , data_lista.Recordset("nro")
        Else
           ListView1.ListItems.Add Xcountt, , "0"
        End If
        If IsNull(data_lista.Recordset("hora")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("hora")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
        End If
        If IsNull(data_lista.Recordset("ced_pac")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("ced_pac")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
        End If
        If IsNull(data_lista.Recordset("nom_pac")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("nom_pac")
        Else
        
'           If data_lista.Recordset("especial") = "PEDIATRIA" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or data_lista.Recordset("especial") = "SICOLOGIA" Then
'              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
'           Else
              Xcansin = Xcansin + 1
'              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
      '     End If
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
        
        End If
        If IsNull(data_lista.Recordset("convenio")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("convenio")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
        End If
        If IsNull(data_lista.Recordset("cel_pac")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("cel_pac")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
        End If
        If IsNull(data_lista.Recordset("tel_pac")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tel_pac")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
        End If
        If IsNull(data_lista.Recordset("zona")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("zona")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
        End If
      
        If IsNull(data_lista.Recordset("taxi")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("taxi")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
        End If
        If IsNull(data_lista.Recordset("direcci")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("direcci")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
        End If
        If IsNull(data_lista.Recordset("usua_anota")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("usua_anota")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
        End If
      
      
      
      End If
      data_lista.Recordset.MoveNext
      Xcountt = Xcountt + 1
   
   Loop
   If Xcansin = 0 Then
      MsgBox "La agenda de esta fecha ESTA COMPLETA!", vbInformation
   End If
Else
   frm_especcovid.MousePointer = 0
   MsgBox "No existe fecha"
   ListView1.ListItems.Clear
End If

If Data1.Recordset("nombre") = "HNF Domicilio" Then
   If ControlUsuario("Agenda hisopados domicilio") = 1 Then
      b_elianota.Enabled = True
      b_agrega.Enabled = True
   Else
      b_elianota.Enabled = False
      b_agrega.Enabled = False
   End If
Else
   If ControlUsuario("Agenda hisopados") = 1 Then
      b_elianota.Enabled = True
      b_agrega.Enabled = True
   Else
      b_elianota.Enabled = True
      b_agrega.Enabled = True
   End If
End If

frm_especcovid.MousePointer = 0

End Sub

Private Sub Form_Load()



'On Error GoTo Quepasaaliniciar

'inicializo url de servicio y mapeo nro de bases a id_base'
Set nroBase_idBase = New Dictionary
nroBase_idBase.Add 1, 1
nroBase_idBase.Add 2, 2
nroBase_idBase.Add 3, 3
nroBase_idBase.Add 4, 4
nroBase_idBase.Add 6, 5
nroBase_idBase.Add 8, 6
nroBase_idBase.Add 10, 7
nroBase_idBase.Add 13, 8
nroBase_idBase.Add 16, 9
nroBase_idBase.Add 17, 10
nroBase_idBase.Add 18, 11
nroBase_idBase.Add 11, 12
Combo1.ListIndex = 0
urlServicio = GetParameters.getValor(1)

abmesp.DatabaseName = App.path & "\abmespec.mdb"
abmesp.RecordSource = "abmespec"
abmesp.Refresh
Data1.RecordSource = "Select * from t_espec where especialidad in ('HNF') order by base,nombre"
Data1.Refresh

data_aut.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_consultahnf.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_users.Connect = "odbc;dsn=" & Xconexrmt & ";"

'data_buscar.Connect = "ODBC;DSN=sappespecial;"
data_buscar.Connect = "ODBC;DSN=" & Xconexrmt & ";"

dat_par.DatabaseName = App.path & "\mensa.mdb"

'data_busca2.DatabaseName = App.Path & "\sapp.mdb"
data_busca2.ConnectionString = "DSN=" & Xconexrmt
'data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & 25048
'data_busca2.Refresh

data_fechas.Connect = "ODBC;DSN=" & Xconexrmt & ";"

data_conscli.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_borrados.Connect = "ODBC;DSN=" & Xconexrmt & ";"

data_cabfec.Connect = "ODBC;DSN=" & Xconexrmt & ";"

data_veoesp.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_veoesp.RecordSource = "Select * from t_espec where especialidad in ('HNF') order by base,nombre"
data_veoesp.Refresh

data_lista.Connect = "ODBC;DSN=" & Xconexrmt & ";"

data_buscnv.ConnectionString = "DSN=" & Xconexrmt
data_buscnv.RecordSource = "select * from zonas order by zo_nombre"
data_buscnv.Refresh
If data_buscnv.Recordset.RecordCount > 0 Then
   data_buscnv.Recordset.MoveFirst
   Do While Not data_buscnv.Recordset.EOF
      cbotipcons.AddItem data_buscnv.Recordset("zo_nombre")
      data_buscnv.Recordset.MoveNext
   Loop
End If
'data_buscnv.RecordSource = ""
'data_buscnv.Refresh

Check1.Value = 1

data_inf.DatabaseName = App.path & "\infespec.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\infespec.mdb"

data_base.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_base.RecordSource = "Select * from bases_sapp order by nro_base"
data_base.Refresh
If data_base.Recordset.RecordCount > 0 Then
   data_base.Recordset.MoveFirst
   Do While Not data_base.Recordset.EOF
      cbobase.AddItem data_base.Recordset("nro_base")
      data_base.Recordset.MoveNext
   Loop
End If

data_medicos.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_medicos.RecordSource = "Select * from medicos_esp order by nom_med"
data_medicos.Refresh
If data_medicos.Recordset.RecordCount > 0 Then
   data_medicos.Recordset.MoveFirst
   Do While Not data_medicos.Recordset.EOF
      cbomedico.AddItem data_medicos.Recordset("nom_med")
      data_medicos.Recordset.MoveNext
   Loop
End If
If WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or WElusuario = "AACUÑA" Or WElusuario = "COMPUTOS" Or WElusuario = "KARINAROMERO" Or WElusuario = "GMONZON" Or WElusuario = "LROMERO" Then
   b_nuevafecha.Enabled = True
   b_cancecons.Enabled = True
   b_nuevo.Enabled = True
   b_modif.Enabled = True
   b_elim.Enabled = True
   b_edimed.Enabled = True
   b_infos.Enabled = True
   If WElusuario = "AACUÑA" Or WElusuario = "SMPEREZ" Then
      b_elifec.Enabled = True
   Else
      b_elifec.Enabled = True
   End If
   mnuevaf.Enabled = True
Else
   If WElusuario = "ROXANA" Or WElusuario = "SMPEREZ" Then
      b_edimed.Enabled = True
   Else
      If ControlUsuario("Especialistas") = 1 Then
         b_nuevafecha.Enabled = True
         b_cancecons.Enabled = True
         b_nuevo.Enabled = True
         b_modif.Enabled = True
         b_elim.Enabled = True
         b_edimed.Enabled = True
         b_infos.Enabled = True
         b_elifec.Enabled = True
      Else
         mnuevaf.Enabled = False
         b_infos.Enabled = True
      End If
   End If
End If
If WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or WElusuario = "MSANCHEZ" Or welusuaario = "JONATHAN" Or WElusuario = "GFERNANDEZ" Or WElusuario = "AACUÑA" Or WElusuario = "KARINAROMERO" Then
Else
'   Data1.RecordSource = "Select * from t_espec where base =" & frm_menu.data_parse.Recordset("base") & " order by nombre"
'   Data1.Refresh
End If


Exit Sub

Quepasaaliniciar:
                MsgBox "Hola.", vbCritical
                 If Err.Number = 3155 Then
                    MsgBox "Error al actualizar"
                 Else
                    MsgBox "Error " & Err.Number & " Verifique si cuenta con conexión a internet para abrir el sistema"
                    Unload Me
                 End If
                 
End Sub

Private Sub mhdescan_Change()

End Sub


Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub ListView1_DblClick()
Dim Xmensacance, Xporquecance As String
Dim Xind, Xcant, Xnro As Long
Xind = 0
Xnro = 0
Xcant = 0
Dim Xcountt As Long
Xcountt = 1

Xmensacance = MsgBox("Desea cancelar la consulta del PACIENTE SELECCIONADO?", vbInformation + vbYesNo)
If Xmensacance = vbYes Then
   frm_motivocance.Show vbModal
End If


End Sub

Private Sub mfnac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbosino.SetFocus
End If

End Sub

Private Sub mhfin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cantp.SetFocus
End If

End Sub

Private Sub mhini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhfin.SetFocus
End If

End Sub



Private Sub t_busca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo1.ListIndex = 1 Then
      If t_busca.Text <> "" Then
         Data1.RecordSource = "Select * from t_espec where especialidad ='" & t_busca.Text & "' and especialidad ='" & "AFILIACIONES" & "' order by base"
         Data1.Refresh
      Else
         Data1.RecordSource = "Select * from t_espec where especialidad ='" & "AFILIACIONES" & "' order by especialidad,base"
         Data1.Refresh
      End If
   Else
      If t_busca.Text <> "" Then
         Data1.RecordSource = "Select * from t_espec where nombre Like '*" + t_busca.Text + "*' and especialidad ='" & "AFILIACIONES" & "' order by nombre,base"
         Data1.Refresh
      Else
         Data1.RecordSource = "Select * from t_espec where especialidad ='" & "AFILIACIONES" & "' order by base,nombre"
         Data1.Refresh
      End If
   End If
End If


End Sub

Private Sub t_cantp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mm.SetFocus
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codced.SetFocus
End If

End Sub

Private Sub t_ced_LostFocus()
If t_ced.Text <> "" Then
   If t_ced.Text <> "0" Then
         data_busca2.RecordSource = "Select * from clientes where cl_cedula =" & t_ced.Text
         data_busca2.Refresh
         If data_busca2.Recordset.RecordCount > 0 Then
            If data_busca2.Recordset("estado") = 2 Then
               MsgBox "ATENCION!!! Socio de baja, COMUNIQUE A EQUIPO COVID.", vbCritical
            End If
            
            t_mat.Text = data_busca2.Recordset("cl_codigo")
            t_codced.Text = data_busca2.Recordset("cl_codced")
            t_nompac.Text = data_busca2.Recordset("cl_apellid")
            t_conv.Text = data_busca2.Recordset("cl_codconv")
            If IsNull(data_busca2.Recordset("cl_dpto")) = False Then
               t_celu.Text = data_busca2.Recordset("cl_dpto")
            Else
               t_celu.Text = ""
            End If
            If IsNull(data_busca2.Recordset("cl_telefon")) = False Then
               t_tellinea.Text = data_busca2.Recordset("cl_telefon")
            Else
               t_tellinea.Text = ""
            End If
            If IsNull(data_busca2.Recordset("cl_fnac")) = False Then
               mfnac.Text = data_busca2.Recordset("cl_fnac")
            Else
               mfnac.Text = "__/__/____"
            End If
            If IsNull(data_busca2.Recordset("cl_direcci")) = False Then
               t_direc.Text = data_busca2.Recordset("cl_direcci")
            Else
               t_direc.Text = ""
            End If
            If IsNull(data_busca2.Recordset("cl_zona")) = False Then
               cbotipcons.Text = data_busca2.Recordset("cl_zona")
            Else
               cbotipcons.Text = ""
            End If
         Else
            t_mat.Text = 0
            t_codced.Text = ""
            t_nompac.Text = "Ingrese nombre"
            t_celu.Text = ""
            t_tellinea.Text = ""
            t_direc.Text = ""
            mfnac.Text = "__/__/____"
            t_conv.Text = ""
            cbotipcons.Text = ""
            cbosino.ListIndex = -1
         End If
   End If
End If

         
End Sub

Private Sub t_celu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_tellinea.SetFocus
End If

End Sub

Private Sub t_codced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nompac.SetFocus
End If

End Sub

Private Sub t_codced_LostFocus()
If Trim(t_codced.Text) = "" Then
   MsgBox "Ingrese dígito de cédula"
End If

End Sub

Private Sub t_conv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_celu.SetFocus
End If

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_agrega.SetFocus
End If

End Sub

Private Sub t_mat_LostFocus()
If t_mat.Text <> "" Then
   If Trim(t_mat.Text) <> "0" Then
      data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
      data_busca2.Refresh
      If data_busca2.Recordset.RecordCount > 0 Then
         If data_busca2.Recordset("estado") = 2 Then
            MsgBox "ATENCION!!! Socio de baja, COMUNIQUE A EQUIPO COVID.", vbCritical
         End If
         t_ced.Text = data_busca2.Recordset("cl_cedula")
         t_codced.Text = data_busca2.Recordset("cl_codced")
         t_nompac.Text = data_busca2.Recordset("cl_apellid")
         t_conv.Text = data_busca2.Recordset("cl_codconv")
         If IsNull(data_busca2.Recordset("cl_dpto")) = False Then
            t_celu.Text = data_busca2.Recordset("cl_dpto")
         Else
            t_celu.Text = ""
         End If
         If IsNull(data_busca2.Recordset("cl_telefon")) = False Then
            t_tellinea.Text = data_busca2.Recordset("cl_telefon")
         Else
            t_tellinea.Text = ""
         End If
         If IsNull(data_busca2.Recordset("cl_fnac")) = False Then
            mfnac.Text = data_busca2.Recordset("cl_fnac")
         Else
            mfnac.Text = "__/__/____"
         End If
         If IsNull(data_busca2.Recordset("cl_direcci")) = False Then
            t_direc.Text = data_busca2.Recordset("cl_direcci")
         Else
            t_direc.Text = ""
         End If
         If IsNull(data_busca2.Recordset("cl_zona")) = False Then
            cbotipcons.Text = data_busca2.Recordset("cl_zona")
         Else
            cbotipcons.Text = ""
         End If
         
      Else
         t_mat.Text = ""
         t_ced.Text = ""
         t_codced.Text = ""
         t_nom.Text = ""
         t_celu.Text = ""
         t_tellinea.Text = ""
         mfnac.Text = "__/__/____"
         t_conv.Text = ""
         cbotipcons.Text = ""
         cbosino.ListIndex = -1
      End If
   End If
End If

         
End Sub

Private Sub t_mm_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   t_espera.SetFocus
'End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboespec.SetFocus
End If

End Sub

Private Sub Text2_Change()

End Sub

Private Sub CalculaEdad(ByVal FNaci As Date)
Dim FAct As String
Dim Anios As String
Dim Meses As String
Dim Dias As String
Dim newday As String
Dim newmonth As String
Dim newyear As String

FAct = Format(Now, "dd/MM/yyyy")
FNaci = Format(FNaci, "dd/MM/yyyy")

'Calcula los años
Anios = DateDiff("yyyy", CDate(Format(FNaci, "dd/MM/yyyy")), CDate(FAct))
'Si el mes actual es menor que el mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) < Month(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 newmonth = Month(CDate(FAct)) + 12
 Else
 'Deje el mes actual tal y como estan
 newmonth = Month(CDate(FAct))
 End If

 'Si el mes actual es igual al mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) = Month(CDate(FNaci)) Then
 'Si el día de la fecha actual es menor al día de la fecha de nacimiento
 If Day(CDate(FAct)) < Day(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 End If
End If

If Day(CDate(FAct)) < Day(CDate(FNaci)) Then
newday = Day(CDate(FAct)) + 30
newmonth = newmonth - 1
Else
newday = Day(CDate(FAct))
End If

If Month(CDate(FNaci)) = Month(Date) Then
   Meses = 0
Else
   Meses = newmonth - Month(CDate(FNaci))
End If

If Meses < 0 And Anios = 0 Then
   Meses = Meses + 12
End If

Dias = newday - Day(CDate(FNaci))

If FNaci <= FAct Then

'Me.TextBox3.Text = Anios & " Años, " & Meses & " Meses, " & Dias & " Dias."
   t_a.Text = Anios
   If Month(Date) = Month(FNaci) Then
      If Day(Date) > Day(FNaci) Then
         Meses = Meses + 1
      Else
         Meses = 11
      End If
   End If
   t_m.Text = Meses
   t_d.Text = Dias
Else
   t_a.Text = ""
   t_m.Text = ""
   t_d.Text = ""
End If

End Sub



Private Sub t_nompac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_conv.SetFocus
End If

End Sub

Private Sub t_tellinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbotipcons.SetFocus
End If

End Sub



'AGREGO BOTON modifCOnsultorio'
Private Sub DBGrid3_LostFocus()
End Sub
Private Sub DBGrid3_Click()
    If DBGrid3.Row >= 0 Then
            modifConsultorio.Enabled = True
    End If
End Sub

Private Sub modifConsultorio_Click()
    If DBGrid3.Row < 0 Then
        MsgBox "Seleccione una fecha para modificar"
        Exit Sub
    End If
    On Error GoTo ModifError
    With frm_espmodificaconsultorio
      .PassUrl = urlServicio
      .PassBase = nroBase_idBase.Item(Val(data_cabfec.Recordset("base")))
      .PassFecha = data_cabfec.Recordset("fecha")
      .PassHoraInicio = data_cabfec.Recordset("hora")
      .PassHoraFin = data_cabfec.Recordset("hora_fin")
      .PassIdMedico = data_cabfec.Recordset("cod_med")
     .Show vbModal
     End With
     If frm_espmodificaconsultorio.PassReservado Then
        id_hora_reserva_previo = data_cabfec.Recordset("id_hora_consultorio")
        'consumir ws para eliminar hora, funcion recibe base, idHora'
        Set obj = consumirServicio("DELETE", urlServicio & "/bases/" & nroBase_idBase.Item(Val(data_cabfec.Recordset("base"))) & "/consultorios/0/disponibilidades/" & id_hora_reserva_previo, "")
        'MsgBox "delete " & obj.responseText
        
        'si esta todo ok, update en tcabfec'
        id_hora_reserva_new = frm_espmodificaconsultorio.PassIdHoraReserva
        If (id_hora_reserva_new <> 0) Then
            data_cabfec.Recordset.Edit
            data_cabfec.Recordset("id_hora_consultorio") = id_hora_reserva_new
            data_cabfec.Recordset.Update
        Else
           'si id_hora_reserva es 0, no quiere reservar consultorio'
            
            If (Not IsNull(data_cabfec.Recordset("id_hora_consultorio"))) Then
                data_cabfec.Recordset.Edit
                data_cabfec.Recordset("id_hora_consultorio") = Nothing
                data_cabfec.Recordset.Update
            End If
        End If
        
        
     End If
     Exit Sub

ModifError:
     MsgBox "error al modificar, comuniquese con computos (" & Err.Description & Err.Source & ")"
     GetParameters.log (Err.Description & Err.Source)
     On Error GoTo 0
     
     
End Sub


'AGREGO PROPIEDAD para comunicar formularios'
Public Property Get Pass_id_hora_reserva() As Integer
    Pass_id_hora_reserva = id_hora_reserva
End Property

Public Property Let Pass_id_hora_reserva(ByVal vNewValue As Integer)
 id_hora_reserva = vNewValue
End Property
'FIN AGREGO PROPIEDAD para comunicar formularios'


Public Function especialidadReserva(especialidad As String) As Boolean
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection

Set conn = New ADODB.Connection
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\ws.mdb"
Set rs = New ADODB.Recordset
rs.Open "SELECT valor FROM parametros WHERE nombre = 'especNoReservaCons' ", conn
reserva = True
If Not rs.RecordCount = 0 Then
    While Not rs.EOF
        If rs("valor") = especialidad Then
            reserva = False
        End If
        rs.MoveNext
    Wend
End If

especialidadReserva = reserva
conn.Close
End Function
Public Function Devuelve_fecha_fact(cedula As String) As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim DesdeFecha As Date
DesdeFecha = CDate("15/04/2021")
ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from linmmdd where ced_socio =" & Val(cedula) & " and cod_prod in (30081,30084,30085) and fecha >='" & Format(DesdeFecha, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_fecha_fact = Format(Xrecclii("fecha"), "dd/mm/yyyy")
Else
   Devuelve_fecha_fact = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function

