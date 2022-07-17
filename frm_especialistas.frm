VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_especialistas 
   BackColor       =   &H00404000&
   Caption         =   "Formulario de fechas para especialistas"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14445
   Icon            =   "frm_especialistas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   14445
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Crystal.CrystalReport cr4 
      Left            =   6480
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
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
      ItemData        =   "frm_especialistas.frx":058A
      Left            =   120
      List            =   "frm_especialistas.frx":0597
      Style           =   2  'Dropdown List
      TabIndex        =   88
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton modifConsultorio 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      Picture         =   "frm_especialistas.frx":05CD
      Style           =   1  'Graphical
      TabIndex        =   87
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
      Left            =   3600
      Top             =   2160
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
      Picture         =   "frm_especialistas.frx":35FD
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3375
      Left            =   120
      TabIndex        =   80
      Top             =   4560
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
      BackColor       =   12582912
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
      TabIndex        =   77
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
      BackColor       =   &H00C00000&
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
      Begin VB.ComboBox cbotipoconsu 
         Height          =   315
         ItemData        =   "frm_especialistas.frx":3B87
         Left            =   2880
         List            =   "frm_especialistas.frx":3B91
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   240
         Width           =   3735
      End
      Begin VB.Data data_conscli 
         Caption         =   "data_conscli"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Anota metas"
         Height          =   495
         Left            =   1800
         TabIndex        =   79
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton b_impcons 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         Picture         =   "frm_especialistas.frx":3BAD
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Imprimir la consulta seleccionada"
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton b_elianota 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         Picture         =   "frm_especialistas.frx":4137
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox t_d 
         Height          =   405
         Left            =   4680
         TabIndex        =   68
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox t_m 
         Height          =   375
         Left            =   3960
         TabIndex        =   67
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox t_a 
         Height          =   405
         Left            =   3360
         TabIndex        =   66
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton b_buscapac 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         Picture         =   "frm_especialistas.frx":46C1
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Buscar los datos del paciente"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton b_modpac 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         Picture         =   "frm_especialistas.frx":4C4B
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Modifica el dato seleccionado"
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton b_agrega 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Picture         =   "frm_especialistas.frx":51D5
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Agrega los datos a la lista de pacientes anotados"
         Top             =   2160
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
         ItemData        =   "frm_especialistas.frx":575F
         Left            =   6000
         List            =   "frm_especialistas.frx":576F
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   1680
         Width           =   1455
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
         Top             =   1320
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
         ItemData        =   "frm_especialistas.frx":57A9
         Left            =   3960
         List            =   "frm_especialistas.frx":57B3
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1680
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
         Left            =   5640
         TabIndex        =   57
         Top             =   1320
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mfnac 
         Height          =   375
         Left            =   1680
         TabIndex        =   54
         Top             =   1680
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
         TabIndex        =   52
         Top             =   960
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
         Top             =   960
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
         Left            =   6600
         TabIndex        =   48
         Top             =   600
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
         Left            =   5520
         TabIndex        =   47
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "HACIENDO DOBLE CLICK SOBRE EL REGISTRO PUEDE CANCELAR LA CONSULTA"
         Top             =   2520
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
         NumItems        =   12
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
            Text            =   "Tipo de consulta"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cédula"
            Object.Width           =   1939
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "d"
            Text            =   "Nombre"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "e"
            Text            =   "Convenio"
            Object.Width           =   1869
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "f"
            Text            =   "Celular"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "g"
            Text            =   "Tel.Línea"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "h"
            Text            =   "Tipo Cons."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Key             =   "i"
            Text            =   "VIA"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Observaciones"
            Object.Width           =   5010
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
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
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "Seleccione tipo de consulta:"
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
         TabIndex        =   90
         Top             =   240
         Width           =   2775
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
         TabIndex        =   76
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FF8080&
         Caption         =   "Tipo Cons."
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
         Height          =   255
         Left            =   4800
         TabIndex        =   61
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   3720
         TabIndex        =   56
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FF8080&
         Caption         =   "HC?"
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
         Height          =   255
         Left            =   3120
         TabIndex        =   55
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF8080&
         Caption         =   "F.Nacimiento:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   5520
         TabIndex        =   51
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4440
         TabIndex        =   46
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.CommandButton b_elifec 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   5160
      Picture         =   "frm_especialistas.frx":57BF
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Eliminar fecha del especialista seleccionado"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Data data_cabfec 
      Caption         =   "data_cabfec"
      Connect         =   "Access"
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
      Begin VB.Data data_buscaus 
         Caption         =   "data_buscaus"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton b_cierramed 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         Picture         =   "frm_especialistas.frx":5D49
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
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * from medicos_esp order by nom_med"
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frm_especialistas.frx":62D3
         Height          =   1575
         Left            =   120
         OleObjectBlob   =   "frm_especialistas.frx":62F2
         TabIndex        =   38
         Top             =   2520
         Width           =   6135
      End
      Begin VB.CommandButton b_canmed 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         Picture         =   "frm_especialistas.frx":6CDD
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
         Picture         =   "frm_especialistas.frx":7267
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton b_modmed 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         Picture         =   "frm_especialistas.frx":77F1
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton b_altamed 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Picture         =   "frm_especialistas.frx":7D7B
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
         ItemData        =   "frm_especialistas.frx":8305
         Left            =   1800
         List            =   "frm_especialistas.frx":835D
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
      Picture         =   "frm_especialistas.frx":84D4
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
      MouseIcon       =   "frm_especialistas.frx":8A5E
      MousePointer    =   99  'Custom
      Picture         =   "frm_especialistas.frx":8FE8
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
      Caption         =   "Fechas disponibles del especialista seleccionado"
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
         Caption         =   "Cancelar consulta"
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
         Picture         =   "frm_especialistas.frx":9572
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Cancelar la consulta seleccionada"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox t_codcons 
         Height          =   285
         Left            =   4560
         TabIndex        =   70
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_feccab 
         Height          =   285
         Left            =   5160
         TabIndex        =   69
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
         Bindings        =   "frm_especialistas.frx":9AFC
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "frm_especialistas.frx":9B16
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
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSMask.MaskEdBox mnuevaf 
         Height          =   495
         Left            =   2040
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
         MouseIcon       =   "frm_especialistas.frx":A6A1
         MousePointer    =   99  'Custom
         Picture         =   "frm_especialistas.frx":AC2B
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2640
         Width           =   1815
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "odbc;dsn=sappnew"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_especialistas.frx":B1B5
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "frm_especialistas.frx":B1C9
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
      Picture         =   "frm_especialistas.frx":BF04
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
      Picture         =   "frm_especialistas.frx":C48E
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
      Picture         =   "frm_especialistas.frx":CA18
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
      Picture         =   "frm_especialistas.frx":CFA2
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
      Picture         =   "frm_especialistas.frx":D52C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Crear un nuevo registro de especialista"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de los especialistas"
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
      Begin VB.CheckBox chconstel 
         BackColor       =   &H00404040&
         Caption         =   "Solo consulta telefónica"
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
         Height          =   375
         Left            =   3120
         TabIndex        =   89
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Data data_aut 
         Caption         =   "data_aut"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox t_basedescsel 
         Height          =   285
         Left            =   2880
         TabIndex        =   85
         Top             =   2040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox t_especsel 
         Height          =   285
         Left            =   2880
         TabIndex        =   84
         Top             =   1440
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox t_basesel 
         Height          =   285
         Left            =   4680
         TabIndex        =   83
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox t_codmedsel 
         Height          =   375
         Left            =   4680
         TabIndex        =   82
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
         TabIndex        =   78
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
         TabIndex        =   75
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
         ItemData        =   "frm_especialistas.frx":DAB6
         Left            =   2040
         List            =   "frm_especialistas.frx":DAB8
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
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
         TabIndex        =   74
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
         Caption         =   "MEDICO:"
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
      OleObjectBlob   =   "frm_especialistas.frx":DABA
      TabIndex        =   86
      Top             =   5280
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   5760
      Picture         =   "frm_especialistas.frx":E629
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_especialistas"
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
Dim Xrecconve As New ADODB.Recordset
Dim Xsqlstr As String
Dim Xfechacartas As Date
Dim XmatCtrol As Long
Dim VerificarCgal As Integer
VerificarCgal = 0
Xfechacartas = Xfechacartas - 150
dat_par.RecordSource = "mensaje"
dat_par.Refresh

If Trim(t_mat.Text) <> "" Then
   If Val(t_mat.Text) > 0 Then
       Verifica_datosJ
   Else
       DatosVerificadosOk = 0
   End If
Else
   DatosVerificadosOk = 0
End If

If Data1.Recordset("nombre") = "SILVA MONICA" And t_ced.Text = 6373958 Then
   MsgBox "ATENCION!! No es posible agendar este paciente para el médico seleccionado. Consulte con el médico!!!", vbCritical
Else
    If cbotipoconsu.ListIndex >= 0 And DatosVerificadosOk = 0 Then
    
       If t_conv.Text = "SMIN" Or t_conv.Text = "SMINA" Or t_conv.Text = "UNIVS" Or _
          t_conv.Text = "UNNSAM" Or t_conv.Text = "HEVANO" Or t_conv.Text = "EVNSAM" Or _
          t_conv.Text = "CCNOS" Or t_conv.Text = "CCNSAM" Or t_conv.Text = "GANOS" Or _
          t_conv.Text = "CASANO" Or t_conv.Text = "CASNSA" Then
          If t_mat.Text <> "" Then
             ConectarBD
             ConbdSapp.Open
             Xsqlstr = "Select * from linmmdd where cod_cli =" & t_mat.Text & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
             With Xrecconve
                 .CursorLocation = adUseClient
                 .CursorType = adOpenKeyset
                 .LockType = adLockOptimistic
                 .Open Xsqlstr, ConbdSapp, , , adCmdText
             End With
             If Xrecconve.RecordCount > 0 Then
                Xestaok = 0
                ConbdSapp.Close
                dat_par.RecordSource = "mensaje"
                dat_par.Refresh
                dat_par.Recordset.Edit
                dat_par.Recordset("text") = "RECUERDE! Confirmar socio con la mutualista."
                dat_par.Recordset.Update
                frm_mensajesvar.Show vbModal
             Else
                ConbdSapp.Close
                data_conscli.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
                data_conscli.Refresh
                If data_conscli.Recordset.RecordCount > 0 Then
                   If IsNull(data_conscli.Recordset("cl_decuota")) = False Then
                      If data_conscli.Recordset("cl_decuota") = 1 Or data_conscli.Recordset("cl_decuota") = 3 Or data_conscli.Recordset("cl_decuota") = 4 Then
                         dat_par.RecordSource = "mensaje"
                         dat_par.Refresh
                         dat_par.Recordset.Edit
                         If t_conv.Text = "SMIN" Or t_conv.Text = "SMINA" Then
                            dat_par.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                            & "Fotocopia CI vigente. Comprobante de domicilio:" _
                            & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumos," _
                            & " a nombre del cliente y del mes corriente o anterior." _
                            & " RECUERDE! Confirmar socio con la mutualista."
                         Else
                            dat_par.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                            & "Fotocopia de Cédula de identidad vigente." _
                            & " RECUERDE! Confirmar socio con la mutualista."
                         End If
                         dat_par.Recordset.Update
                         dat_par.Refresh
                         Xestaok = 25
                         frm_mensajesvar.Show vbModal
                      Else
                         If data_conscli.Recordset("cl_decuota") = 2 Then
                         Else
                            dat_par.RecordSource = "mensaje"
                            dat_par.Refresh
                            dat_par.Recordset.Edit
                            If t_conv.Text = "SMIN" Or t_conv.Text = "SMINA" Then
                               dat_par.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                               & "Fotocopia CI vigente. Comprobante de domicilio: " _
                                & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumos," _
                               & " a nombre del cliente y que sea del mes corriente o anterior." _
                               & " RECUERDE! Confirmar socio con la mutualista."
                            Else
                               dat_par.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                               & "Fotocopia de Cédula de identidad vigente." _
                               & " RECUERDE! Confirmar socio con la mutualista."
                            End If
                            dat_par.Recordset.Update
                            dat_par.Refresh
                            Xestaok = 25
                            frm_mensajesvar.Show vbModal
                         End If
                      End If
                   Else
                      ConectarBD
                      ConbdSapp.Open
                      Xsqlstr = "Select * from linmmdd where cod_cli =" & t_mat.Text & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                      With Xrecconve
                          .CursorLocation = adUseClient
                          .CursorType = adOpenKeyset
                          .LockType = adLockOptimistic
                          .Open Xsqlstr, ConbdSapp, , , adCmdText
                      End With
                      If Xrecconve.RecordCount > 0 Then
                         Xestaok = 0
                      Else
                         dat_par.Recordset.Edit
                         dat_par.Recordset("text") = "ATENCION!!! Si no realiza carta mutual:" & Chr(13) & " No tendrá derecho a los servicios NO URGENTES."
                         dat_par.Recordset.Update
                         frm_mensajesvar.Show vbModal
                         Xestaok = 25
                      End If
                      ConbdSapp.Close
                      
                      If data_conscli.Recordset("cl_decuota") = 1 Or data_conscli.Recordset("cl_decuota") = 3 Or data_conscli.Recordset("cl_decuota") = 4 Then
                         dat_par.RecordSource = "mensaje"
                         dat_par.Refresh
                         dat_par.Recordset.Edit
                         If t_conv.Text = "SMIN" Or t_conv.Text = "SMINA" Then
                            dat_par.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                            & "Fotocopia CI vigente. Comprobante de domicilio: " _
                            & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumos," _
                            & " a nombre del cliente y que sea del mes corriente o anterior." _
                            & " RECUERDE! Confirmar socio con la mutualista."
                         Else
                            dat_par.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                            & "Fotocopia de Cédula de identidad vigente." _
                            & " RECUERDE! Confirmar socio con la mutualista."
                         End If
                         dat_par.Recordset.Update
                         dat_par.Refresh
                         Xestaok = 25
                         frm_mensajesvar.Show vbModal
                      End If
                   End If
                Else
                   dat_par.RecordSource = "mensaje"
                   dat_par.Refresh
                   dat_par.Recordset.Edit
                   dat_par.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Verificar Documentación a presentar. Confirmar socio con la mutualista."
                   
                   dat_par.Recordset.Update
                   dat_par.Refresh
                   Xestaok = 25
                   frm_mensajesvar.Show vbModal
                End If
              End If
           End If
       Else
           Xestaok = 0
       End If
        
       If Xestaok = 25 Then
          If data_cabfec.Recordset("especial") = "VACUNACION" Or _
             data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
             data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
             data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
             data_cabfec.Recordset("especial") = "FISIOTERAPIA" Or data_cabfec.Recordset("especial") = "LABORATORIO" Or _
             cbotipcons.Text = "META" Or cbotipcons.Text = "RN (Recien Nacido)" Then
             Xestaok = 0
          End If
       End If
       
       If WElusuario = "JFERNAN" Or data_cabfec.Recordset("especial") = "PEDIATRIA" Or data_cabfec.Recordset("especial") = "SICOLOGIA" Or data_cabfec.Recordset("especial") = "ODONTOLOGIA" Then
          If data_cabfec.Recordset("especial") = "ODONTOLOGIA" Then
             If frm_menu.data_parse.Recordset("base") = 11 Or frm_menu.data_parse.Recordset("base") = 3 Then
                Command1_Click
             Else
                Xfecdeuda = Date - 30
                Xind = 0
                Xnro = 0
                Xcant = 0
                Xdeudasiono = 0
                Xcountt = 1
                If t_mat.Text <> "" Then
                   If t_celu.Text <> "" Then
                      data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
                      data_busca2.Refresh
                      If IsNull(data_busca2.Recordset("cl_dpto")) = True Then
                          data_busca2.Recordset("cl_dpto") = t_celu.Text
                          data_busca2.Recordset.Update
                      End If
                   End If
                End If
                    
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
                  
                If t_mat.Text <> "" Then
                   If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                      data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                      data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                      data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                      data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
                      Xdeudasiono = 0
                      Xestaok = 0
                   Else
                      Xdeb = 3
                      Wopszond = ""
                      data_busca2.RecordSource = "Select * from deudas where cliente =" & Val(t_mat.Text) & " and tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanciacion" & "' and mes =" & 0
                      data_busca2.Refresh
                      If data_busca2.Recordset.RecordCount > 0 Then
                         data_busca2.Recordset.MoveFirst
                         Do While Not data_busca2.Recordset.EOF
                            If IsNull(data_busca2.Recordset("nro_superv")) = False Then
                               Xlafv = data_busca2.Recordset("fecha") + data_busca2.Recordset("nro_superv")
                            Else
                               Xlafv = data_busca2.Recordset("fecha") + 30
                            End If
                            If Format(Xlafv, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                               Xdeudasiono = 9
                               Wxquepreg = 1 'es deuda por servicio
                            End If
                            data_busca2.Recordset.MoveNext
                         Loop
                      Else
                         Xdeudasiono = 0
                      End If
                             
                      data_busca2.RecordSource = "Select * from deudas where cliente =" & Val(t_mat.Text) & " and mes >" & 0 & " and fecha_pago is null and origen <='" & "Refinan" & "' order by ano,mes"
                      data_busca2.Refresh
                      If data_busca2.Recordset.RecordCount > 0 Then
                         data_busca2.Recordset.MoveLast
                         If data_busca2.Recordset.RecordCount > 2 Then
                            Xop4 = data_busca2.Recordset("mes")
                            Xop5 = data_busca2.Recordset("ano")
                            Xdeudasiono = 11
                            If Wxquepreg = 0 Then
                               Wxquepreg = 2 'es por cuota
                            End If
                         End If
                      End If
                               
                      data_busca2.RecordSource = "Select * from deudas where cliente =" & Val(t_mat.Text) & " and fecha_pago is null and origen >='" & "Refinan" & "'"
                      data_busca2.Refresh
                      If data_busca2.Recordset.RecordCount > 0 Then
                         data_busca2.Recordset.MoveFirst
                         Do While Not data_busca2.Recordset.EOF
                            If IsNull(data_busca2.Recordset("nro_superv")) = False Then
                               Xlafv = data_busca2.Recordset("fecha") + data_busca2.Recordset("nro_superv")
                            Else
                               Xlafv = data_busca2.Recordset("fecha") + 30
                            End If
                            If Format(Xlafv, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                               Xdeudasiono = 9
                               Wxquepreg = 3 'es por refinanc
                            End If
                            data_busca2.Recordset.MoveNext
                         Loop
                      End If
                       
                      If Xdeudasiono = 9 Or Xdeudasiono = 11 Then
                         MsgBox "Socio moroso, no se puede realizar agenda. Consulte con Administración.", vbCritical
                         Xhab = Val(t_mat.Text)
                      End If
                       
                      data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
                      data_busca2.Refresh
                      If data_busca2.Recordset.RecordCount > 0 Then
                         If IsNull(data_busca2.Recordset("saldo_chc2")) = False Then
                            If data_busca2.Recordset("saldo_chc2") = 1 Then
                               If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                                  data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                                  data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                                  data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                                  data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
                               Else
                                  Xdeudasiono = 17
                               End If
                            End If
                            If data_busca2.Recordset("cl_codconv") = "GANOS" Or _
                               data_busca2.Recordset("cl_codconv") = "SMIN" Or _
                               data_busca2.Recordset("cl_codconv") = "UNIVS" Or _
                               data_busca2.Recordset("cl_codconv") = "CPS" Or _
                               data_busca2.Recordset("cl_codconv") = "CASH" Or _
                               data_busca2.Recordset("cl_codconv") = "PART" Or _
                               data_busca2.Recordset("cl_codconv") = "UCM" Then
                               If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                                  data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                                  data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                                  data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                                  data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
                               Else
                                  Xdeudasiono = 13
                               End If
                            End If
                         End If
                         data_buscnv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_busca2.Recordset("cl_codconv") & "'"
                         data_buscnv.Refresh
                         If data_buscnv.Recordset.RecordCount > 0 Then
                            If IsNull(data_buscnv.Recordset("cnv_colrec")) = False Then
                               If data_buscnv.Recordset("cnv_colrec") = "M" Then
                                  If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                                     data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                                     data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                                     data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                                     data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
                                  Else
                                     Xdeudasiono = 12
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                    
                   If t_ced.Text <> "" And t_codced.Text <> "" Then
                      If cbotipcons.ListIndex <> 0 And (Xdeudasiono = 9 Or Xdeudasiono = 17 Or Xdeudasiono = 11) Then
                         If Xdeudasiono = 9 Then
                            MsgBox "ATENCION!! Socio con servicios crédito pendientes.No se puede agendar.", vbCritical, "Deudas"
                         Else
                            If Xdeudasiono = 11 Then
                               MsgBox "ATENCION!! Socio con cuotas pendientes de pago.No se puede agendar.", vbCritical, "Deudas"
                            Else
                               If Xdeudasiono = 17 Then
                                  MsgBox "ATENCION!! SOCIO CON SERVICIOS RESTRINGIDOS. NO SE PUEDE ANOTAR.", vbCritical, "AGENDA"
                               Else
                                  MsgBox "Error al anotar, verifique datos!", vbCritical, "AGENDA"
                               End If
                            End If
                         End If
                      Else
                         If Xdeudasiono = 12 Or Xdeudasiono = 13 Or Xestaok = 25 Then
                            If Xestaok = 25 Then
                               MsgBox "Debe realizar carta para poder anotarse", vbCritical
                               Xestaok = 0
                            Else
                               MsgBox "ATENCION!! Socio con categoría no habilitada para especialistas. Llame al 097215419", vbCritical
                            End If
                         Else
                            If Xdeudasiono = 11 Then
                               MsgBox "ATENCION!! Socio con deudas pendientes, no se puede anotar. Llame al 097215419", vbCritical
                            Else
                               If Xcant = 1 Then
                                  For Xind = 1 To ListView1.ListItems.count
                                      ListView1.ListItems(Xind).Selected = True
                                      If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                                         Xnro = ListView1.ListItems(Xind).Text
                                         Xellugar = Xind
                                         If cbomedico.Text = "FERTILAB" Then
                                            Xloslabos = InputBox("Ingrese los ANALISIS a realizar")
                                         End If
                                         Xlacedconsulta = t_ced.Text & t_codced.Text
                                         data_lista.RecordSource = "select * from t_fechas where cdate(fecha) >#" & Format(Xlafechacons, "yyyy/mm/dd") & "# and ced_pac ='" & Trim(Xlacedconsulta) & "' and cancela not in ('SI') and especial ='" & data_cabfec.Recordset("especial") & "'"
                                         data_lista.Refresh
                                         If data_lista.Recordset.RecordCount > 0 Then
                                            MsgBox "Ya se encuentra anotado para una consulta con ésta especialidad", vbExclamation
                                         Else
                                            data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                                            data_lista.Refresh
                                            If data_lista.Recordset.RecordCount > 0 Then
                                               If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                                                  MsgBox "Ya existe un paciente anotado, verifique!!", vbCritical
                                               Else
                                                  data_lista.Recordset.Edit
                                                  If t_mat.Text <> "" Then
                                                     data_lista.Recordset("mat_pac") = t_mat.Text
                                                  Else
                                                     data_lista.Recordset("mat_pac") = 0
                                                  End If
                                                  If t_nompac.Text <> "" Then
                                                     data_lista.Recordset("nom_pac") = t_nompac.Text
                                                  End If
                                                  data_lista.Recordset("tipoconsulta") = cbotipoconsu.Text
                                                  data_lista.Recordset("tipoconsultan") = cbotipoconsu.ListIndex
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
                                                  data_lista.Recordset("hcsiono") = cbosino.ListIndex
                                                  data_lista.Recordset("tipo_cons") = cbotipcons.ListIndex
                                                  If cbotipcons.ListIndex >= 0 Then
                                                     data_lista.Recordset("tipo_consd") = cbotipcons.Text
                                                  End If
                                                  data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                                                  data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                                                  data_lista.Recordset("usua_anota") = WElusuario
                                                  If t_a.Text <> "" Then
                                                     If t_m.Text <> "" Then
                                                        If t_d.Text <> "" Then
                                                           data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES " & t_d.Text & " DIAS"
                                                        Else
                                                           data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES "
                                                        End If
                                                     Else
                                                        data_lista.Recordset("edad") = t_a.Text & "AÑOS"
                                                     End If
                                                  Else
                                                     If t_m.Text <> "" Then
                                                        If t_d.Text <> "" Then
                                                           data_lista.Recordset("edad") = t_m.Text & "MESES " & t_d.Text & " DIAS"
                                                        Else
                                                           data_lista.Recordset("edad") = t_m.Text & "MESES "
                                                        End If
                                                     Else
                                                        If t_d.Text <> "" Then
                                                           data_lista.Recordset("edad") = t_d.Text & "DIAS"
                                                        End If
                                                     End If
                                                  End If
                                                  data_lista.Recordset("usua_web") = "SAPP"
                                                  If Xloslabos <> "" Then
                                                     data_lista.Recordset("obs") = Mid(Xloslabos, 1, 200)
                                                  End If
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
                                                  mfnac.Text = "__/__/____"
                                                  t_conv.Text = ""
                                                  cbotipcons.ListIndex = -1
                                                  cbosino.ListIndex = -1
                                                  cbotipoconsu.ListIndex = -1
                                                  data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " and nom_pac is null"
                                                  data_lista.Refresh
                                                  If data_lista.Recordset.RecordCount > 0 Then
                                                     If IsNull(data_cabfec.Recordset("completa")) = False Then
                                                        data_cabfec.Recordset.Edit
                                                        data_cabfec.Recordset("completa") = Null
                                                        data_cabfec.Recordset.Update
                                                     End If
                                                  Else
                                                     If IsNull(data_cabfec.Recordset("completa")) = False Then
                                                        If Int(data_cabfec.Recordset("completa")) <> 1 Then
                                                           data_cabfec.Recordset.Edit
                                                           data_cabfec.Recordset("completa") = 1
                                                           data_cabfec.Recordset.Update
                                                        End If
                                                     Else
                                                        data_cabfec.Recordset.Edit
                                                        data_cabfec.Recordset("completa") = 1
                                                        data_cabfec.Recordset.Update
                                                     End If
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
                                                            If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
                                                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
                                                            Else
                                                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
                                                  End If
                                               End If
                                            Else
                                               MsgBox "No se encuentra registro para actualizar"
                                            End If
                                         End If
                                      
                                      End If
                                  Next Xind
                                  ListView1.ListItems.Item(Xellugar).Selected = True
                                  ListView1.ListItems.Item(Xellugar).EnsureVisible
                                  ListView1.SetFocus
                               Else
                                  MsgBox "Debe seleccionar un solo registro"
                               End If
                            End If
                         End If
                      End If
                   Else
                      MsgBox "Ingrese CEDULA para anotar"
                   End If
                Else
                   MsgBox "No ingresó matrícula"
                End If
             End If
          Else
             Command1_Click
          End If
       Else
          Xfecdeuda = Date - 30
          Xind = 0
          Xnro = 0
          Xcant = 0
          Xdeudasiono = 0
          Xcountt = 1
          If t_mat.Text <> "" Then
             If t_celu.Text <> "" Then
                data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
                data_busca2.Refresh
                If IsNull(data_busca2.Recordset("cl_dpto")) = True Then
                   data_busca2.Recordset("cl_dpto") = t_celu.Text
                   data_busca2.Recordset.Update
                End If
             End If
          End If
            
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
            
          If t_mat.Text <> "" Then
             If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                data_cabfec.Recordset("especial") = "FISIOTERAPIA" Or _
                cbotipcons.Text = "META" Or cbotipcons.Text = "RN (Recien Nacido)" Then
                Xestaok = 0
                Xdeudasiono = 0
             Else
                Xdeb = 3
                Wopszond = ""
                data_busca2.RecordSource = "Select * from deudas where cliente =" & Val(t_mat.Text) & " and tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanciacion" & "' and mes =" & 0
                data_busca2.Refresh
                If data_busca2.Recordset.RecordCount > 0 Then
                   data_busca2.Recordset.MoveFirst
                   Do While Not data_busca2.Recordset.EOF
                      If IsNull(data_busca2.Recordset("nro_superv")) = False Then
                         Xlafv = data_busca2.Recordset("fecha") + data_busca2.Recordset("nro_superv")
                      Else
                         Xlafv = data_busca2.Recordset("fecha") + 30
                      End If
                      If Format(Xlafv, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                         Xdeudasiono = 9
                         Wxquepreg = 1 'es deuda por servicio
                      End If
                      data_busca2.Recordset.MoveNext
                   Loop
                Else
                   Xdeudasiono = 0
                End If
                         
                data_busca2.RecordSource = "Select * from deudas where cliente =" & Val(t_mat.Text) & " and mes >" & 0 & " and fecha_pago is null and origen <='" & "Refinan" & "' order by ano,mes"
                data_busca2.Refresh
                If data_busca2.Recordset.RecordCount > 0 Then
                   data_busca2.Recordset.MoveLast
                   If data_busca2.Recordset.RecordCount > 2 Then
                      Xop4 = data_busca2.Recordset("mes")
                      Xop5 = data_busca2.Recordset("ano")
                      Xdeudasiono = 11
                      If Wxquepreg = 0 Then
                         Wxquepreg = 2 'es por cuota
                      End If
                   End If
                End If
                         
                data_busca2.RecordSource = "Select * from deudas where cliente =" & Val(t_mat.Text) & " and fecha_pago is null and origen >='" & "Refinan" & "'"
                data_busca2.Refresh
                If data_busca2.Recordset.RecordCount > 0 Then
                   data_busca2.Recordset.MoveFirst
                   Do While Not data_busca2.Recordset.EOF
                      If IsNull(data_busca2.Recordset("nro_superv")) = False Then
                         Xlafv = data_busca2.Recordset("fecha") + data_busca2.Recordset("nro_superv")
                      Else
                         Xlafv = data_busca2.Recordset("fecha") + 30
                      End If
                      If Format(Xlafv, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                         Xdeudasiono = 9
                         Wxquepreg = 3 'es por refinanc
                      End If
                      data_busca2.Recordset.MoveNext
                   Loop
                End If
             End If
             If Xdeudasiono = 9 Or Xdeudasiono = 11 Then
                MsgBox "Socio moroso, no se puede realizar agenda. Consulte con Administración.", vbCritical
                Xhab = Val(t_mat.Text)
    '                  frm_autoriza.Show vbModal
    '                 Xelcodigoaut = InputBox("INGRESE CÓDIGO DE AUTORIZACIÓN:", "AUTORIZACIÓN", Wopszond)
             End If
                          
             data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
             data_busca2.Refresh
             If data_busca2.Recordset.RecordCount > 0 Then
                If IsNull(data_busca2.Recordset("saldo_chc2")) = False Then
                   If data_busca2.Recordset("saldo_chc2") = 1 Then
                      If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                         data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                         data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                         data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                         data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
                      Else
                         Xdeudasiono = 9
                      End If
                   End If
    ''''''''aquí particular no
                   
                   If data_busca2.Recordset("cl_codconv") = "GANOS" Or data_busca2.Recordset("cl_codconv") = "CAUTE" Or _
                      data_busca2.Recordset("cl_codconv") = "SMIN" Or data_busca2.Recordset("cl_codconv") = "CCASMU" Or _
                      data_busca2.Recordset("cl_codconv") = "UNIVS" Or data_busca2.Recordset("cl_codconv") = "ESEMM" Or _
                      data_busca2.Recordset("cl_codconv") = "CPS" Or data_busca2.Recordset("cl_codconv") = "EUCM" Or _
                      data_busca2.Recordset("cl_codconv") = "CASH" Or data_busca2.Recordset("cl_codconv") = "ESUAT" Or _
                      data_busca2.Recordset("cl_codconv") = "SEMM" Or data_busca2.Recordset("cl_codconv") = "CONVE" Or _
                      data_busca2.Recordset("cl_codconv") = "UCM" Then
                      If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                         data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                         data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                         data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                         data_cabfec.Recordset("especial") = "FISIOTERAPIA" Or _
                         data_cabfec.Recordset("especial") = "RADIOLOGIA" Then
                         Xestaok = 0
                      Else
                         Xestaok = 25
                         Xdeudasiono = 13
                      End If
                   Else
                      If data_busca2.Recordset("cl_codconv") = "PART" Then
                         Xestaok = 0
                         MsgBox "ATENCION! Categoría PARTICULARES, informe del costo correspondiente al paciente.", vbCritical
                      End If
                   End If
                End If
                data_buscnv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_busca2.Recordset("cl_codconv") & "'"
                data_buscnv.Refresh
                If data_buscnv.Recordset.RecordCount > 0 Then
                   If IsNull(data_buscnv.Recordset("cnv_colrec")) = False Then
                      If data_buscnv.Recordset("cnv_colrec") = "M" Then
                         If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                            data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                            data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                            data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                            data_cabfec.Recordset("especial") = "FISIOTERAPIA" Or _
                            data_cabfec.Recordset("especial") = "RADIOLOGIA" Then
                         Else
                            Xdeudasiono = 12
                         End If
                      End If
                   End If
                End If
             Else
                If Trim(t_conv.Text) = "GANOS" Or Trim(t_conv.Text) = "CAUTE" Or _
                   Trim(t_conv.Text) = "SMIN" Or Trim(t_conv.Text) = "CCASMU" Or _
                   Trim(t_conv.Text) = "UNIVS" Or Trim(t_conv.Text) = "EUCM" Or _
                   Trim(t_conv.Text) = "CPS" Or Trim(t_conv.Text) = "ESUAT" Or _
                   Trim(t_conv.Text) = "CASH" Or Trim(t_conv.Text) = "ESEMM" Or _
                   Trim(t_conv.Text) = "SEMM" Or Trim(t_conv.Text) = "CONVE" Or _
                   Trim(t_conv.Text) = "UCM" Then
                   If data_cabfec.Recordset("especial") = "VACUNACION" Or _
                      data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
                      data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
                      data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
                      data_cabfec.Recordset("especial") = "FISIOTERAPIA" Or _
                      data_cabfec.Recordset("especial") = "RADIOLOGIA" Then
                      Xestaok = 0
                   Else
                      Xestaok = 25
                      Xdeudasiono = 28
                   End If
                Else
                   If Trim(t_conv.Text) = "PART" Then
                      Xestaok = 0
                      MsgBox "ATENCION! Categoría PARTICULARES, informe del costo correspondiente al paciente.", vbCritical
                   End If
                End If
             End If
    '      End If
            
              If cbotipcons.ListIndex <> 0 And (Xdeudasiono = 9 Or Xdeudasiono = 17 Or Xdeudasiono = 11) Then
                 If Xdeudasiono = 9 Then
                    MsgBox "ATENCION!! Socio con servicios crédito pendientes.No se puede agendar.", vbCritical, "Deudas"
                 Else
                    If Xdeudasiono = 11 Then
                       MsgBox "ATENCION!! Socio con cuotas pendientes de pago.No se puede agendar.", vbCritical, "Deudas"
                    Else
                       If Xdeudasiono = 17 Then
                          MsgBox "ATENCION!! SOCIO CON SERVICIOS RESTRINGIDOS. NO SE PUEDE ANOTAR.", vbCritical, "AGENDA"
                       Else
                          MsgBox "Error al anotar, verifique datos!", vbCritical, "AGENDA"
                       End If
                    End If
                 End If
              Else
                 If Xdeudasiono = 12 Or Xdeudasiono = 13 Or Xestaok = 25 Or Xdeudasiono = 28 Then
                    If Xestaok = 25 Then
                       If Xdeudasiono = 28 Then
                          MsgBox "ATENCION!! Socio con categoría no habilitada para especialistas. Llame al 097215419", vbCritical
                       Else
                          MsgBox "Debe realizar carta para poder anotarse", vbCritical
                          Xestaok = 0
                       End If
                    Else
                       MsgBox "ATENCION!! Socio con categoría no habilitada para especialistas. Llame al 097215419", vbCritical
                    End If
                 Else
                    If Xdeudasiono = 11 Then
                       MsgBox "ATENCION!! Socio con deudas pendientes, no se puede anotar. Llame al 097215419", vbCritical
                    Else
                       If Xcant = 1 Then
                          For Xind = 1 To ListView1.ListItems.count
                                ListView1.ListItems(Xind).Selected = True
                                If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                                   Xnro = ListView1.ListItems(Xind).Text
                                   Xellugar = Xind
                                   If cbomedico.Text = "FERTILAB" Then
                                      Xloslabos = InputBox("Ingrese los ANALISIS a realizar")
                                   End If
                                   Xlacedconsulta = t_ced.Text & t_codced.Text
                                   If data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
                                      data_lista.RecordSource = "select * from t_fechas where cdate(fecha) <#" & Format("01/01/2010", "yyyy/mm/dd") & "# and ced_pac ='" & Trim(Xlacedconsulta) & "' and cancela not in ('SI') and especial ='" & data_cabfec.Recordset("especial") & "'"
                                      data_lista.Refresh
                                   Else
                                      data_lista.RecordSource = "select * from t_fechas where cdate(fecha) >#" & Format(Xlafechacons, "yyyy/mm/dd") & "# and ced_pac ='" & Trim(Xlacedconsulta) & "' and cancela not in ('SI') and especial ='" & data_cabfec.Recordset("especial") & "'"
                                      data_lista.Refresh
                                   End If
                                   If data_lista.Recordset.RecordCount > 0 Then
                                      MsgBox "Ya se encuentra anotado para una consulta con ésta especialidad", vbExclamation
                                   Else
                                      data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                                      data_lista.Refresh
                                      If data_lista.Recordset.RecordCount > 0 Then
                                         If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                                            MsgBox "Ya existe un paciente anotado, verifique!!", vbCritical
                                         Else
                                            data_lista.Recordset.Edit
                                            If t_mat.Text <> "" Then
                                               data_lista.Recordset("mat_pac") = t_mat.Text
                                            Else
                                               data_lista.Recordset("mat_pac") = 0
                                            End If
                                            If t_nompac.Text <> "" Then
                                               data_lista.Recordset("nom_pac") = t_nompac.Text
                                            End If
                                            If t_ced.Text <> "" Then
                                               data_lista.Recordset("ced_pac") = t_ced.Text & t_codced.Text
                                            End If
                                            data_lista.Recordset("tipoconsulta") = cbotipoconsu.Text
                                            data_lista.Recordset("tipoconsultan") = cbotipoconsu.ListIndex
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
                                            data_lista.Recordset("hcsiono") = cbosino.ListIndex
                                            data_lista.Recordset("tipo_cons") = cbotipcons.ListIndex
                                            If cbotipcons.ListIndex >= 0 Then
                                               data_lista.Recordset("tipo_consd") = cbotipcons.Text
                                            End If
                                            data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                                            data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                                            data_lista.Recordset("usua_anota") = WElusuario
                                            If t_a.Text <> "" Then
                                               If t_m.Text <> "" Then
                                                  If t_d.Text <> "" Then
                                                     data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES " & t_d.Text & " DIAS"
                                                  Else
                                                     data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES "
                                                  End If
                                               Else
                                                  data_lista.Recordset("edad") = t_a.Text & "AÑOS"
                                               End If
                                            Else
                                               If t_m.Text <> "" Then
                                                  If t_d.Text <> "" Then
                                                     data_lista.Recordset("edad") = t_m.Text & "MESES " & t_d.Text & " DIAS"
                                                  Else
                                                     data_lista.Recordset("edad") = t_m.Text & "MESES "
                                                  End If
                                               Else
                                                  If t_d.Text <> "" Then
                                                     data_lista.Recordset("edad") = t_d.Text & "DIAS"
                                                  End If
                                               End If
                                            End If
                                            data_lista.Recordset("usua_web") = "SAPP"
                                            If Xloslabos <> "" Then
                                               data_lista.Recordset("obs") = Mid(Xloslabos, 1, 200)
                                            End If
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
                                            mfnac.Text = "__/__/____"
                                            t_conv.Text = ""
                                            cbotipcons.ListIndex = -1
                                            cbosino.ListIndex = -1
                                            cbotipoconsu.ListIndex = -1
        
                                            data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " and nom_pac is null"
                                            data_lista.Refresh
                                            If data_lista.Recordset.RecordCount > 0 Then
                                               If IsNull(data_cabfec.Recordset("completa")) = False Then
                                                  data_cabfec.Recordset.Edit
                                                  data_cabfec.Recordset("completa") = Null
                                                  data_cabfec.Recordset.Update
                                               End If
                                            Else
                                               If IsNull(data_cabfec.Recordset("completa")) = False Then
                                                  If Int(data_cabfec.Recordset("completa")) <> 1 Then
                                                     data_cabfec.Recordset.Edit
                                                     data_cabfec.Recordset("completa") = 1
                                                     data_cabfec.Recordset.Update
                                                  End If
                                               Else
                                                  data_cabfec.Recordset.Edit
                                                  data_cabfec.Recordset("completa") = 1
                                                  data_cabfec.Recordset.Update
                                               End If
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
                                                      If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
                                                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
                                                      Else
                                                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
                                            
                                            End If
                                         End If
                                      Else
                                         MsgBox "No se encuentra registro para actualizar"
                                      End If
                                   End If
                                End If
                          Next Xind
                          ListView1.ListItems.Item(Xellugar).Selected = True
                          ListView1.ListItems.Item(Xellugar).EnsureVisible
                          ListView1.SetFocus
                       Else
                          MsgBox "Debe seleccionar un solo registro"
                       End If
                    End If
                 End If
              End If
          Else
             MsgBox "Ingrese matrícula para anotar"
          End If
       End If
    Else
       If DatosVerificadosOk <> 0 Then
          MsgBox "No ha confirmado datos, debe confirmar datos para agendar.", vbCritical
       Else
          MsgBox "Debe ingresar tipo de consulta Presencial o Telefónica", vbInformation
          cbotipoconsu.SetFocus
       End If
    End If
End If

Xdeudasiono = 0
Xestaok = 0


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

data_medicossapp.RecordSource = "select * from medicos_esp order by id DESC"
data_medicossapp.Refresh
If data_medicossapp.Recordset.RecordCount > 0 Then
   labcod.Caption = data_medicossapp.Recordset("id") + 1
Else
   labcod.Caption = 0
End If
data_medicossapp.RecordSource = "Select * from medicos_esp order by nom_med"
data_medicossapp.Refresh

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
    frm_especialistas.MousePointer = 0
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
Frame4.Visible = True

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

Xborralaconsulta = MsgBox("Desea borrar los datos anotados en la lista?", vbInformation + vbYesNo)
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
                 data_lista.Recordset("tipoconsulta") = Null
                 data_lista.Recordset("tipoconsultan") = Null
                 data_lista.Recordset.Update
                 abmesp.Recordset.AddNew
                 abmesp.Recordset("fecha") = Date
                 abmesp.Recordset("hora") = Format(Time, "HH:mm")
                 abmesp.Recordset("usuario") = WElusuario
                 abmesp.Recordset("base") = frm_menu.data_parse.Recordset("base")
                 abmesp.Recordset("accion") = "ELIMINA ANOTACION"
                 abmesp.Recordset.Update
                 
                 MsgBox "Registro eliminado!", vbInformation
                 data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " and nom_pac is null"
                 data_lista.Refresh
                 If data_lista.Recordset.RecordCount > 0 Then
                    If IsNull(data_cabfec.Recordset("completa")) = False Then
                       data_cabfec.Recordset.Edit
                       data_cabfec.Recordset("completa") = Null
                       data_cabfec.Recordset.Update
                    End If
                 Else
                    If IsNull(data_cabfec.Recordset("completa")) = False Then
                       If Int(data_cabfec.Recordset("completa")) <> 1 Then
                          data_cabfec.Recordset.Edit
                          data_cabfec.Recordset("completa") = 1
                          data_cabfec.Recordset.Update
                       End If
                    Else
                       data_cabfec.Recordset.Edit
                       data_cabfec.Recordset("completa") = 1
                       data_cabfec.Recordset.Update
                    End If
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
                            If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
                            Else
                               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
'''''               On Error Resume Next 'si no encuentra el campo, o error en ws sigue
               base = data_buscar.Recordset("base")
'''''               id_hora_consultorio = data_buscar.Recordset("id_hora_consultorio")
               data_buscar.Recordset.Delete
               data_buscar.Recordset.MoveNext
               'consumo ws para eliminar hora de consultorio en el caso de que tenga'
     '''''          If Not IsNull(id_hora_consultorio) Then
     '''''           reservarConsultorio = especialidadReserva(t_especsel.Text)
     '''''           If reservarConsultorio Then
     '''''               Set obj = consumirServicio("DELETE", urlServicio & "/bases/" & base & "/consultorios/0/disponibilidades/" & id_hora_consultorio, "")
     '''''               response = obj.responseText
                    ' MsgBox response
     '''''           End If
     '''''          End If
     '''''          On Error GoTo 0 ' desactivo error handler para que siga todo igual
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

If WElusuario = "COMPUTOS" Then
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
         Data1.Recordset("constel") = chconstel.Value
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
         If IsNull(Data1.Recordset("constel")) = False Then
            If chconstel.Value <> Data1.Recordset("constel") Then
               Data1.Recordset("constel") = chconstel.Value
            End If
         Else
            Data1.Recordset("constel") = chconstel.Value
         End If
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
Dim Lacedus As String
Dim Xnomus As String
Dim Xelnromed As Integer

Xnomus = ""
Lacedus = ""

If XAlta = 1 Then
   If t_nom.Text <> "" Then
      If cboespec.Text = "MED.GRAL." And Trim(t_codsapp.Text) = "" Then
         MsgBox "Debe ingresar código de médico.", vbCritical
      Else
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
            
            If cboespec.Text = "MED.GRAL." Then
               data_buscaus.RecordSource = "select * from meta_tres where m_mat =" & t_codsapp.Text
               data_buscaus.Refresh
               If data_buscaus.Recordset.RecordCount > 0 Then
                  Lacedus = data_buscaus.Recordset("m_nrofrm")
               Else
                  Lacedus = ""
               End If
               If Lacedus <> "" Then
                  data_buscaus.RecordSource = "select * from cap_ciap where cod_cap ='" & Lacedus & "'"
                  data_buscaus.Refresh
                  If data_buscaus.Recordset.RecordCount > 0 Then
                     Xnomus = data_buscaus.Recordset("des_cap")
                  Else
                     Xnomus = ""
                  End If
               End If
               If Xnomus <> "" Then
                  data_buscaus.RecordSource = "select * from usuarios where usuario ='" & Xnomus & "'"
                  data_buscaus.Refresh
                  If data_buscaus.Recordset.RecordCount > 0 Then
                     If IsNull(data_buscaus.Recordset("codmed")) = False Then
                        If data_buscaus.Recordset("codmed") <> Val(labcod.Caption) Then
                           data_buscaus.Recordset.Edit
                           data_buscaus.Recordset("codmed") = Val(labcod.Caption)
                           data_buscaus.Recordset.Update
                        End If
                     Else
                        data_buscaus.Recordset.Edit
                        data_buscaus.Recordset("codmed") = Val(labcod.Caption)
                        data_buscaus.Recordset.Update
                     End If
                     data_medicossapp.Recordset.Update
                  Else
                     MsgBox "No se encuentra nombre de usuario, consulte a informática.", vbCritical
                     data_medicossapp.Recordset.CancelUpdate
                  End If
               Else
                  MsgBox "No se encuentra médico con esta cédula, consulte a informática.", vbCritical
                  data_medicossapp.Recordset.CancelUpdate
               End If
            Else
               data_medicossapp.Recordset.Update
            End If
            data_medicossapp.Refresh
            XAlta = 0
            b_modmed.Enabled = True
            b_grabmed.Enabled = False
            b_canmed.Enabled = False
            b_altamed.Enabled = True
            DBGrid2.Enabled = True
            DBGrid2.SetFocus
         End If
      
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
Frame3.Visible = False
Frame4.Visible = True
Dim Xcountt As Long
Xcountt = 1
On Error GoTo Quepasoalimp

frm_especialistas.MousePointer = 11
t_feccab.Text = data_cabfec.Recordset("fecha")
t_codcons.Text = data_cabfec.Recordset("cod_cons")

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\infespec.mdb")

MiBaseact.Execute "Delete * from lista"

data_inf.RecordSource = "lista"
data_inf.Refresh

Data2.DatabaseName = App.path & "\infespec.mdb"

MiBaseact.Execute "Delete * from lista2"

Data2.RecordSource = "lista2"
Data2.Refresh

data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " order by nro"
data_lista.Refresh
If data_lista.Recordset.RecordCount > 0 Then
   data_lista.Recordset.MoveLast
   data_lista.Recordset.MoveFirst
   Do While Not data_lista.Recordset.EOF
      If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
         If data_lista.Recordset("tipoconsulta") = "Presencial" Then
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = Format(data_lista.Recordset("fecha"), "dd/mm/yyyy")
            data_inf.Recordset("medico") = data_lista.Recordset("nom_med")
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
            If IsNull(data_lista.Recordset("tipo_consd")) = False Then
               data_inf.Recordset("tipocons") = data_lista.Recordset("tipo_consd")
            End If
            If IsNull(data_lista.Recordset("hcsiono")) = False Then
               If data_lista.Recordset("hcsiono") = 0 Then
                  data_inf.Recordset("hc") = "SI"
               Else
                  data_inf.Recordset("hc") = "NO"
               End If
            End If
            If IsNull(data_lista.Recordset("edad")) = False Then
               data_inf.Recordset("edad") = data_lista.Recordset("edad")
            End If
            If IsNull(data_lista.Recordset("fec_nac")) = False Then
               data_inf.Recordset("fnac") = data_lista.Recordset("fec_nac")
            End If
            If IsNull(data_lista.Recordset("cod_cons")) = False Then
               data_inf.Recordset("codcons") = data_lista.Recordset("cod_cons")
            End If
            If IsNull(data_lista.Recordset("usua_anota")) = False Then
               data_inf.Recordset("via") = Mid(data_lista.Recordset("usua_anota"), 1, 15)
            End If
            data_inf.Recordset.Update
         End If
      Else
        data_inf.Recordset.AddNew
        data_inf.Recordset("fecha") = Format(data_lista.Recordset("fecha"), "dd/mm/yyyy")
        data_inf.Recordset("medico") = data_lista.Recordset("nom_med")
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
        If IsNull(data_lista.Recordset("tipo_consd")) = False Then
           data_inf.Recordset("tipocons") = data_lista.Recordset("tipo_consd")
        End If
        If IsNull(data_lista.Recordset("hcsiono")) = False Then
           If data_lista.Recordset("hcsiono") = 0 Then
              data_inf.Recordset("hc") = "SI"
           Else
              data_inf.Recordset("hc") = "NO"
           End If
        End If
        If IsNull(data_lista.Recordset("edad")) = False Then
           data_inf.Recordset("edad") = data_lista.Recordset("edad")
        End If
        If IsNull(data_lista.Recordset("fec_nac")) = False Then
           data_inf.Recordset("fnac") = data_lista.Recordset("fec_nac")
        End If
        If IsNull(data_lista.Recordset("cod_cons")) = False Then
           data_inf.Recordset("codcons") = data_lista.Recordset("cod_cons")
        End If
        If IsNull(data_lista.Recordset("usua_anota")) = False Then
           data_inf.Recordset("via") = Mid(data_lista.Recordset("usua_anota"), 1, 15)
        End If
        data_inf.Recordset.Update
            
      End If
      data_lista.Recordset.MoveNext
   Loop
   
   data_lista.Recordset.MoveFirst
   Do While Not data_lista.Recordset.EOF
      If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
         If data_lista.Recordset("tipoconsulta") = "Telefónica" Then
            Data2.Recordset.AddNew
            Data2.Recordset("fecha") = Format(data_lista.Recordset("fecha"), "dd/mm/yyyy")
            Data2.Recordset("medico") = data_lista.Recordset("nom_med")
            Data2.Recordset("espec") = data_lista.Recordset("especial")
            Data2.Recordset("base") = data_lista.Recordset("base")
            Data2.Recordset("nro") = data_lista.Recordset("nro")
            Data2.Recordset("hora") = data_lista.Recordset("hora")
            If IsNull(data_lista.Recordset("nom_pac")) = False Then
               Data2.Recordset("nom_pac") = data_lista.Recordset("nom_pac")
            End If
            If IsNull(data_lista.Recordset("ced_pac")) = False Then
               Data2.Recordset("cedula") = data_lista.Recordset("ced_pac")
            End If
            If IsNull(data_lista.Recordset("mat_pac")) = False Then
               Data2.Recordset("mat") = data_lista.Recordset("mat_pac")
            End If
            If IsNull(data_lista.Recordset("convenio")) = False Then
               Data2.Recordset("convenio") = data_lista.Recordset("convenio")
            End If
            If IsNull(data_lista.Recordset("cel_pac")) = False Then
               Data2.Recordset("celular") = data_lista.Recordset("cel_pac")
            End If
            If IsNull(data_lista.Recordset("tel_pac")) = False Then
               Data2.Recordset("telef") = data_lista.Recordset("tel_pac")
            End If
            If IsNull(data_lista.Recordset("tipo_consd")) = False Then
               Data2.Recordset("tipocons") = data_lista.Recordset("tipo_consd")
            End If
            If IsNull(data_lista.Recordset("hcsiono")) = False Then
               If data_lista.Recordset("hcsiono") = 0 Then
                  Data2.Recordset("hc") = "SI"
               Else
                  Data2.Recordset("hc") = "NO"
               End If
            End If
            If IsNull(data_lista.Recordset("edad")) = False Then
               Data2.Recordset("edad") = data_lista.Recordset("edad")
            End If
            If IsNull(data_lista.Recordset("fec_nac")) = False Then
               Data2.Recordset("fnac") = data_lista.Recordset("fec_nac")
            End If
            If IsNull(data_lista.Recordset("cod_cons")) = False Then
               Data2.Recordset("codcons") = data_lista.Recordset("cod_cons")
            End If
            If IsNull(data_lista.Recordset("usua_anota")) = False Then
               Data2.Recordset("via") = Mid(data_lista.Recordset("usua_anota"), 1, 15)
            End If
            Data2.Recordset.Update
         End If
      End If
      data_lista.Recordset.MoveNext
   Loop
   
   
   data_inf.RecordSource = "Select * from lista"
   data_inf.Refresh
   frm_especialistas.MousePointer = 0
   cr1.ReportFileName = App.path & "\inflistanew.rpt"
   cr1.Action = 1
   
   
   Data2.RecordSource = "Select * from lista2"
   Data2.Refresh
   frm_especialistas.MousePointer = 0
   MsgBox "Proceso terminado", vbInformation
   cr4.ReportFileName = App.path & "\inflistanewt.rpt"
   cr4.Action = 1

Else
   frm_especialistas.MousePointer = 0
   MsgBox "No existe fecha"
End If

Exit Sub

Quepasoalimp:
             If Err.Number = 91 Then
                MsgBox "Verifique si seleccionó la consulta"
             Else
                MsgBox "Verifique si tiene datos seleccionados"
             End If
             
End Sub

Private Sub b_infos_Click()
frm_infespenew.Show vbModal

End Sub

Private Sub b_modif_Click()

If WElusuario = "COMPUTOS" Then
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
Else
    If ControlUsuario("Especialistas") = 1 Then
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
    End If
End If

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
Dim BanderaPed As Integer
BanderaPed = 0

Me.Pass_id_hora_reserva = 0
Xnropac = 1
'''On Error GoTo Erroralcrear

If mnuevaf.Text <> "__/__/____" Then
    '''' me fijo si la especialidad seleccionada reserva consultorio (ws.mdb)'
    'reservarConsultorio = especialidadReserva(t_especsel.Text)
    '''' ---- datos para Consumo web service ----
    'On Error GoTo WSError
    base = nroBase_idBase.Item(Val(t_basesel.Text))
    medico = Val(t_codmedsel.Text)
    XfecstrGuiones = Format(mnuevaf.Text, "yyyy-mm-dd")
    horaInicio = mhini.Text
    horaFin = mhfin.Text
    'urlWS = urlServicio & "/bases/" & base & "/consultorios"
    'body = "medico=" & medico & "&inicio=" & XfecstrGuiones & "T" & horaInicio & ":00&fin=" & XfecstrGuiones & "T" & horaFin & ":00"
    
    If Format(mnuevaf.Text, "yyyy/mm/dd") >= Format(Date, "yyyy/mm/dd") Then
        Xfecstr = Format(mnuevaf.Text, "dd/mm/yyyy")
        Xmensacrear = MsgBox("Desea crear las fechas para " & t_especsel.Text & " DR." & cbomedico.Text & "?", vbInformation + vbYesNo)
        If Xmensacrear = vbYes Then
            frm_especialistas.MousePointer = 11
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
                frm_especialistas.MousePointer = 0
                MsgBox "Ya existe consulta creada con éstos parámetros, VERIFIQUE!!", vbInformation
     '           On Error GoTo 0
     '           GoTo Fin
            Else
                ''''' si la especialidad reserva consultorio llamo a web service, sino no
     '           If (reservarConsultorio) Then
                    ''''''verifico si ya reservaron consultorios para MG si el parametro de control esta activo'
     '               If (GetParameters.getActivo(2)) Then
     '                   Set obj1 = consumirServicio("GET", urlWS & "/0/disponibilidades/MG/reservado/" & XfecstrGuiones, "")
     '                   reservado = obj1.Status
     '                   If reservado = 400 Then
     '                       XnoreservadoMG = MsgBox("No fueron reservados los consultorios para Medicina General este mes, desea continuar igualmente?", vbInformation + vbYesNo)
     '                       If Not XnoreservadoMG = vbYes Then
                                '''''On Error GoTo 0
     '                           GoTo Fin
     '                       End If
     '                   End If
     '               End If
                    '''''FIN verifico si ya reservaron consultorios para MG'
                    
                    ''''' -- instancio conexion ws --
     '               Set obj = consumirServicio("PUT", urlWS & "/0/disponibilidades", body & "&forzar=false")
                    ''''' -- envio datos y me fijo en el status HTTP que devuelve --
     '               estado = obj.Status
    
     '               response = obj.responseText
     '               Set p = JSON.parse(response)
                
                    ''''' si el estado es 200 sigue flujo normal,
                    ''''' si el estado es 205, el consultorio es de una especialidad pero no exclusivo ==> cartel de confirmacion
                    ''''' si el estado es 404 (seguramente porque no hay lugar disponible) dejo elegir consultorio y calculo la superposicion
     '               Select Case estado
     '                   Case 200
     '                       id_hora_reserva = p.Item("horaConsultorio").Item("id")
     '                   Case 404
                            '''''' MsgBox JSON.toString(p.Item("superposiciones").Item(1).Item("id_consultorio"))
     '                       With frm_espeligeconsultorio
     '                           .PassVar = JSON.toString(p)
     '                           .PassUrlWS = urlWS
     '                           .PassBodyWS = body
     '                           .Show vbModal
     '                           Debug.Print .PassVar
                                ''''''si no reservo consultorio porque le dio cancelar, me voy'
     '                           If Not .PassReservoConsultorio Then
     '                               On Error GoTo 0
     '                               GoTo Fin
     '                           End If
     '                       End With
     '                   Case 205
                            ''''''pendiente, pueden haber 2 consultorios (base 6) y solo estoy ofreciendo 1, esta mal(de ultima puede modificar)'
     '                       XmensaNoExclusivo = MsgBox("El consultorio que se podría reservar es de especialista, pero no es exclusivo, desea continuar igualmente?", vbInformation + vbYesNo)
     '                       If XmensaNoExclusivo = vbYes Then
     '                           Set obj = consumirServicio("PUT", urlWS & "/0/disponibilidades", body & "&forzar=true")
     '                           Set pp = JSON.parse(obj.responseText)
     '                           id_hora_reserva = pp.Item("horaConsultorio").Item("id")
                                ''''''MsgBox id_hora_reserva
     '                           estado = obj.Status
     '                           On Error GoTo 0
                                '''''''GoTo ConfirmarCreacion
     '                       Else
     '                           On Error GoTo 0
     '                           GoTo Fin
     '                       End If
     '                   Case Else
                            ''''''sigo con la creacion de la fecha sin la reserva de consultorio'
     '               End Select
     '           End If
            End If
     '   Else
     '       On Error GoTo 0
     '       GoTo Fin
     '   End If
     '   On Error GoTo 0 '''''' desactivo error handler para que el codigo siga como estaba (sin error handler)
    'Else
    '    frm_especialistas.MousePointer = 0
    '    MsgBox "La fecha debe ser MAYOR a la fecha actual.", vbInformation
    '    On Error GoTo 0
    '    Exit Sub
    'End If
'Else
'    frm_especialistas.MousePointer = 0
'    MsgBox "Verifique si es correcta la fecha de consulta.", vbInformation
'    On Error GoTo 0
'    Exit Sub
'End If
'ConfirmarCreacion: ''''''''''cuando confirma creacion de fecha
            Xhorah = Val(Mid(mhfin.Text, 1, 2))
            Xhorad = Val(Mid(mhini.Text, 1, 2))
            Xminh = Val(Mid(mhini.Text, 4, 2))
            If t_espera.Text = "" Then
               t_espera.Text = 0
            End If
            Xespera = t_espera.Text
            BanderaPed = 0
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
               If cbobase.Text = "98" Or cbobase.Text = "99" Then
                  If Val(t_cantp.Text) >= 60 Then
                     If Xnropac = 31 Or Xnropac = 32 Or Xnropac = 33 Or Xnropac = 34 Or Xnropac = 35 Then
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
                        data_fechas.Recordset("ced_pac") = "9"
                        data_fechas.Recordset("mat_pac") = "9"
                        data_fechas.Recordset("nom_pac") = "RESERVADO-DESCANSO"
                        data_fechas.Recordset("tel_pac") = "9"
                        data_fechas.Recordset("cel_pac") = "9"
                        data_fechas.Recordset("convenio") = "NO"
                        data_fechas.Recordset("tipoconsulta") = "Reservado"
                        data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                        If cbobase.Text = "98" Or cbobase.Text = "99" Then
                           data_fechas.Recordset("sepuede") = 0
                        Else
                           If Trim(Xmin) = "00" Or Trim(Xmin) = "30" Or Trim(Xmin) = "36" Then
                              data_fechas.Recordset("sepuede") = 1
                           Else
                              data_fechas.Recordset("sepuede") = 0
                           End If
                        End If
                        data_fechas.Recordset.Update
                     Else
                        If Xnropac = 6 Or Xnropac = 12 Or Xnropac = 18 Or Xnropac = 24 Or _
                           Xnropac = 30 Or Xnropac = 36 Or Xnropac = 42 Or Xnropac = 48 Or _
                           Xnropac = 54 Or Xnropac = 60 Or Xnropac = 66 Or Xnropac = 72 Then
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
                           data_fechas.Recordset("ced_pac") = "9"
                           data_fechas.Recordset("mat_pac") = "9"
                           data_fechas.Recordset("nom_pac") = "RESERVADO-NO ANOTAR"
                           data_fechas.Recordset("tel_pac") = "9"
                           data_fechas.Recordset("cel_pac") = "9"
                           data_fechas.Recordset("convenio") = "NO"
                           data_fechas.Recordset("tipoconsulta") = "Reservado"
                           data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                           If cbobase.Text = "98" Or cbobase.Text = "99" Then
                              data_fechas.Recordset("sepuede") = 0
                           Else
                              If Trim(Xmin) = "00" Or Trim(Xmin) = "30" Or Trim(Xmin) = "36" Then
                                 data_fechas.Recordset("sepuede") = 1
                              Else
                                 data_fechas.Recordset("sepuede") = 0
                              End If
                           End If
                           data_fechas.Recordset.Update
                        Else
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
                           data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                           If cbobase.Text = "98" Or cbobase.Text = "99" Then
                              data_fechas.Recordset("sepuede") = 0
                           Else
                              If Trim(Xmin) = "00" Or Trim(Xmin) = "30" Or Trim(Xmin) = "36" Then
                                 data_fechas.Recordset("sepuede") = 1
                              Else
                                 data_fechas.Recordset("sepuede") = 0
                              End If
                           End If
                           data_fechas.Recordset.Update
                        End If
                     End If
                  Else
                     If Xnropac = 6 Or Xnropac = 12 Or Xnropac = 18 Or Xnropac = 24 Or _
                        Xnropac = 30 Or Xnropac = 36 Or Xnropac = 42 Or Xnropac = 48 Or _
                        Xnropac = 54 Or Xnropac = 60 Or Xnropac = 66 Or Xnropac = 72 Then
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
                        data_fechas.Recordset("ced_pac") = "9"
                        data_fechas.Recordset("mat_pac") = "9"
                        data_fechas.Recordset("nom_pac") = "RESERVADO-NO ANOTAR"
                        data_fechas.Recordset("tel_pac") = "9"
                        data_fechas.Recordset("cel_pac") = "9"
                        data_fechas.Recordset("convenio") = "NO"
                        data_fechas.Recordset("tipoconsulta") = "Reservado"
                        data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                        If cbobase.Text = "98" Or cbobase.Text = "99" Then
                           data_fechas.Recordset("sepuede") = 0
                        Else
                           If Trim(Xmin) = "00" Or Trim(Xmin) = "30" Or Trim(Xmin) = "36" Then
                              data_fechas.Recordset("sepuede") = 1
                           Else
                              data_fechas.Recordset("sepuede") = 0
                           End If
                        End If
                        data_fechas.Recordset.Update
                     Else
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
                        data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                        If cbobase.Text = "98" Or cbobase.Text = "99" Then
                           data_fechas.Recordset("sepuede") = 0
                        Else
                           If Trim(Xmin) = "00" Or Trim(Xmin) = "30" Or Trim(Xmin) = "36" Then
                              data_fechas.Recordset("sepuede") = 1
                           Else
                              data_fechas.Recordset("sepuede") = 0
                           End If
                        End If
                        data_fechas.Recordset.Update
                     End If
                  End If
               Else
                  If t_especsel.Text = "MED.GRAL." Then
                     If Val(Xhor) = 13 Then
                        If Trim(Xmin) = "00" Or Trim(Xmin) = "12" Or Trim(Xmin) = "24" Or Trim(Xmin) = "10" Or Trim(Xmin) = "20" Then
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
                           data_fechas.Recordset("ced_pac") = "9"
                           data_fechas.Recordset("mat_pac") = "9"
                           data_fechas.Recordset("nom_pac") = "RESERVADO Almuerzo"
                           data_fechas.Recordset("tel_pac") = "9"
                           data_fechas.Recordset("cel_pac") = "9"
                           data_fechas.Recordset("convenio") = "NO"
                           data_fechas.Recordset("tipoconsulta") = "Reservado"
                           data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                           data_fechas.Recordset("sepuede") = 0
                        Else
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
                           data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                           If Trim(Xmin) = "00" Or Trim(Xmin) = "24" Or Trim(Xmin) = "20" Then
                              data_fechas.Recordset("sepuede") = 1
                           Else
                              If Trim(Xmin) = "48" Or Trim(Xmin) = "40" Then
                                 data_fechas.Recordset("sepuede") = 2
                              Else
                                 data_fechas.Recordset("sepuede") = 0
                              End If
                           End If
                        End If
                     Else
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
                        data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                        If Trim(Xmin) = "00" Or Trim(Xmin) = "24" Or Trim(Xmin) = "20" Then
                           data_fechas.Recordset("sepuede") = 1
                        Else
                           If Trim(Xmin) = "48" Or Trim(Xmin) = "40" Then
                              data_fechas.Recordset("sepuede") = 2
                           Else
                              data_fechas.Recordset("sepuede") = 0
                           End If
                        End If
                     End If
                  Else
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
                        data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
                        If t_especsel.Text = "PEDIATRIA" Then
                           If Xnropac = 1 Or Xnropac = 2 Or Xnropac = 3 Or Xnropac = 4 Then
                              If Xnropac = 1 Or Xnropac = 2 Then
                                 data_fechas.Recordset("ced_pac") = "9999999"
                                 data_fechas.Recordset("mat_pac") = 0
                                 data_fechas.Recordset("nom_pac") = "Reservado RN/Etc"
                                 data_fechas.Recordset("tel_pac") = "999"
                                 data_fechas.Recordset("cel_pac") = "999"
                                 data_fechas.Recordset("usua_anota") = "RESERVA"
                                 data_fechas.Recordset("convenio") = "RN"
                                 data_fechas.Recordset("sepuede") = 0
                              Else
                                 data_fechas.Recordset("ced_pac") = "9999999"
                                 data_fechas.Recordset("mat_pac") = 0
                                 data_fechas.Recordset("nom_pac") = "Reservado"
                                 data_fechas.Recordset("tel_pac") = "999"
                                 data_fechas.Recordset("cel_pac") = "999"
                                 data_fechas.Recordset("usua_anota") = "RESERVA"
                                 data_fechas.Recordset("convenio") = "RN"
                                 data_fechas.Recordset("sepuede") = 0
                              End If
                           Else
                              If Trim(Xhor) = "13" Then
                                 If Val(Xmin) <= 15 Then
                                    data_fechas.Recordset("ced_pac") = "9999999"
                                    data_fechas.Recordset("mat_pac") = 0
                                    data_fechas.Recordset("nom_pac") = "ALMUERZO"
                                    data_fechas.Recordset("tel_pac") = "999"
                                    data_fechas.Recordset("cel_pac") = "999"
                                    data_fechas.Recordset("usua_anota") = "RESERVA"
                                    data_fechas.Recordset("convenio") = "AL"
                                    data_fechas.Recordset("sepuede") = 0
                                 Else
                                    data_fechas.Recordset("sepuede") = 0
                                 End If
                              Else
'                                 If Trim(Xmin) = "00" Or Trim(Xmin) = "30" Then
                                 If Trim(Xmin) = "30" Then
                                    If BanderaPed = 0 Then
''                                    If Trim(Xmin) = "30" Then
                                       data_fechas.Recordset("sepuede") = 1
                                       BanderaPed = 1
                                    Else
                                       data_fechas.Recordset("sepuede") = 2
                                       BanderaPed = 0
                                    End If
                                 Else
                                    data_fechas.Recordset("sepuede") = 0
                                 End If
                              End If
                           End If
                        Else
                           If t_especsel.Text = "VACUNACION" Then
                              data_fechas.Recordset("sepuede") = 1
                           Else
                              If Xnropac = 5 Then
                                 data_fechas.Recordset("sepuede") = 2
                              Else
                                 data_fechas.Recordset("sepuede") = 1
                              End If
                           End If
                        End If
                  End If
                  data_fechas.Recordset.Update
               End If
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
               data_fechas.Recordset("fecha_cons") = CDate(Xfecstr)
               data_fechas.Recordset("sepuede") = 0
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
            '''''''agrego id de reserva'
            If id_hora_reserva <> 0 Then
        '      On Error Resume Next ''''''si no esta el campo en la bd no se cae el programa
        '      data_cabfec.Recordset("id_hora_consultorio") = id_hora_reserva
        '      On Error GoTo 0
            End If
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
            If t_especsel.Text = "CIRUGIA" Or _
               cbomedico.Text = "GARCIA NORA" Or _
               t_especsel.Text = "GASTROENTEROLOGIA" Then
               Xdiasdif = Xdiasdif + Xdiasmes1
               Xlafecanota = CDate(Xfecstr) - Xdiasdif
               data_cabfec.Recordset("fecha_anota") = Format(Xlafecanota, "dd/mm/yyyy")
               data_cabfec.Recordset("fecha_cons") = CDate(Xfecstr)
               data_cabfec.Recordset("fecha_anotad") = CDate(Xlafecanota)
               data_cabfec.Recordset.Update
            Else
               If t_especsel.Text = "MED.GRAL." Then
                  Xlafecanota = Format(Date, "dd/mm/yyyy")
                  data_cabfec.Recordset("fecha_anota") = Format(Xlafecanota, "dd/mm/yyyy")
                  data_cabfec.Recordset("fecha_cons") = CDate(Xfecstr)
                  data_cabfec.Recordset("fecha_anotad") = CDate(Xlafecanota)
                  data_cabfec.Recordset.Update
               Else
                  Xdiasdif = Xdiasdif + Xdiasmes1 + Xdiasmes2
                  Xlafecanota = CDate(Xfecstr) - Xdiasdif
                  data_cabfec.Recordset("fecha_anota") = Format(Xlafecanota, "dd/mm/yyyy")
                  data_cabfec.Recordset("fecha_cons") = CDate(Xfecstr)
                  data_cabfec.Recordset("fecha_anotad") = CDate(Xlafecanota)
                  data_cabfec.Recordset.Update
               End If
            End If
            frm_especialistas.MousePointer = 0
            MsgBox "CONSULTA CREADA", vbInformation
        End If
    End If
Else
    frm_especialistas.MousePointer = 0
    MsgBox "Verifique si es correcta la fecha de consulta.", vbInformation
End If

''''Exit Sub


'Fin:
'    frm_especialistas.MousePointer = 0
'    Exit Sub
    
'WSError:
'  MsgBox Err.Description & " (servicio de reservas caído o sin conexion, favor, contactarse con computos), podrá seguir pero el consultorio de la base no sera reservado. "
'  On Error GoTo 0 ''''''' desactivo error handler para que el resto quede como estaba
'  Resume ConfirmarCreacion

''Erroralcrear:
'             If Err.Number = 3150 Then
'                MsgBox "Error: " & Err.Description
'             Else
'                MsgBox "ERROR: " & Err.Description
'             End If
             
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
   cbotipcons.SetFocus
End If

End Sub

Private Sub cbotipcons_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_agrega.SetFocus
End If

End Sub



Private Sub cbotipoconsu_Click()
If cbotipoconsu.Text <> "Telefónica" Then
   If chconstel.Value = 1 Then
      MsgBox "Solo acepta consulta telefónica", vbInformation
      cbotipoconsu.ListIndex = 1
   End If
End If

End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Solo CMT" Then
   Data1.RecordSource = "Select * from t_espec where base =" & 98 & " and especialidad not in ('AFILIACIONES','HNF') order by base,nombre"
Else
   Data1.RecordSource = "Select * from t_espec where especialidad not in ('AFILIACIONES','HNF') order by base,nombre"
End If
Data1.Refresh


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
Xdeudasiono = 0
Xcountt = 1
If t_mat.Text <> "" Then
   If t_celu.Text <> "" Then
      data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
      data_busca2.Refresh
      If IsNull(data_busca2.Recordset("cl_dpto")) = True Then
'         data_busca2.Recordset.Edit
         data_busca2.Recordset("cl_dpto") = t_celu.Text
         data_busca2.Recordset.Update
      End If
   End If
End If

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

If Xestaok = 25 Then
   If data_cabfec.Recordset("especial") = "VACUNACION" Or _
      data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
      data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
      data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
      data_cabfec.Recordset("especial") = "FISIOTERAPIA" Or data_cabfec.Recordset("especial") = "LABORATORIO" Or _
      cbotipcons.Text = "META" Or cbotipcons.Text = "RN (Recien Nacido)" Then
      Xestaok = 0
   End If
End If

If t_mat.Text <> "" Then
   data_busca2.RecordSource = "Select * from deudas where cliente =" & t_mat.Text & " and fecha_pago is null and mes >" & 0
   data_busca2.Refresh
   If data_busca2.Recordset.RecordCount > 2 Then
      If data_cabfec.Recordset("especial") = "VACUNACION" Or _
         data_cabfec.Recordset("especial") = "SICOLOGIA" Or _
         data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
         data_cabfec.Recordset("especial") = "ODONTOLOGIA" Or _
         data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
         Xdeudasiono = 0
      Else
         Xdeudasiono = 11
      End If
   Else
      Xdeudasiono = 0
   End If
   data_busca2.RecordSource = "Select * from deudas where cliente =" & t_mat.Text & " and tipodoc ='" & "CRE" & "' and fecha_pago is null order by fecha"
   data_busca2.Refresh
   If data_busca2.Recordset.RecordCount > 0 Then
      Do While Not data_busca2.Recordset.EOF
         Xlafv = data_busca2.Recordset("fecha") + 30
         If Xlafv >= Date Then
         Else
            Xdeudasiono = 9
         End If
         data_busca2.Recordset.MoveNext
      Loop
   End If
End If

If t_ced.Text <> "" And t_codced.Text <> "" Then
    If Xdeudasiono >= 9 And cbotipcons.ListIndex <> 0 Then
       MsgBox "ATENCION!! Socio con servicios crédito pendientes o SERVICIOS RESTRINGIDOS!Llame al 097215419", vbCritical, "No se puede agendar"
       Xelcodigoaut = InputBox("SOCIO CON CRÉDITOS PENDIENTES O CUOTAS, INGRESE CODIGO DE AUTORIZACIÓN SI ES CLAVE 3", "SOCIO CON CRÉDITOS PENDIENTES")
       Xlapersona = InputBox("INGRESE NOMBRE DE RESPONSABLE QUE AUTORIZA", "RESPONSABLE QUE AUTORIZA")
       If Trim(Xelcodigoaut) <> "" Then
          If Trim(Xlapersona) <> "" Then
             data_aut.RecordSource = "select * from Codigos_aut"
             data_aut.Refresh
             data_aut.Recordset.AddNew
             data_aut.Recordset("fecha") = Date
             data_aut.Recordset("usuario") = Mid(Xlapersona, 1, 50)
             data_aut.Recordset("codaut") = Mid(Xelcodigoaut, 1, 45)
             If t_mat.Text <> "" Then
                data_aut.Recordset("socio") = t_mat.Text
             Else
                data_aut.Recordset("socio") = t_mat.Text
             End If
             data_aut.Recordset("modulo") = "RESERVA"
             data_aut.Recordset("usuario_caja") = WElusuario
             data_aut.Recordset.Update
             Xdeudasiono = 0
          Else
             MsgBox "Socio con créditos o cuotas(>3) pendientes, no se puede reservar. Comunique a Administración al 097215419.", vbCritical
          End If
       Else
          MsgBox "Socio con créditos o cuotas(>3) pendientes, no se puede reservar. Comunique a Administración al 097215419.", vbCritical
       End If
       If Xdeudasiono = 0 Then
       
            If Xcant = 1 Then
              For Xind = 1 To ListView1.ListItems.count
                  ListView1.ListItems(Xind).Selected = True
                  If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                     Xnro = ListView1.ListItems(Xind).Text
                     Xellugar = Xind
                     If cbomedico.Text = "FERTILAB" Then
                        Xloslabos = InputBox("Ingrese los ANALISIS a realizar")
                     End If
                     If data_cabfec.Recordset("especial") = "GINECOLOGIA" Or data_cabfec.Recordset("especial") = "PEDIATRIA" Then
                        data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                        data_lista.Refresh
    'JMOLINS,GUSTAVO
                        If data_lista.Recordset.RecordCount > 0 Then
                           If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                              MsgBox "Ya existe un paciente anotado, verifique!!", vbCritical
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
                              data_lista.Recordset("tipoconsulta") = cbotipoconsu.Text
                              data_lista.Recordset("tipoconsultan") = cbotipoconsu.ListIndex
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
                              data_lista.Recordset("hcsiono") = cbosino.ListIndex
                              data_lista.Recordset("tipo_cons") = cbotipcons.ListIndex
                              If cbotipcons.ListIndex >= 0 Then
                                 data_lista.Recordset("tipo_consd") = cbotipcons.Text
                              End If
                              data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                              data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                              data_lista.Recordset("usua_anota") = WElusuario
                              If t_a.Text <> "" Then
                                 If t_m.Text <> "" Then
                                    If t_d.Text <> "" Then
                                       data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES " & t_d.Text & " DIAS"
                                    Else
                                       data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES "
                                    End If
                                 Else
                                    data_lista.Recordset("edad") = t_a.Text & "AÑOS"
                                 End If
                              Else
                                 If t_m.Text <> "" Then
                                    If t_d.Text <> "" Then
                                       data_lista.Recordset("edad") = t_m.Text & "MESES " & t_d.Text & " DIAS"
                                    Else
                                       data_lista.Recordset("edad") = t_m.Text & "MESES "
                                    End If
                                 Else
                                    If t_d.Text <> "" Then
                                       data_lista.Recordset("edad") = t_d.Text & "DIAS"
                                    End If
                                 End If
                              End If
                              data_lista.Recordset("usua_web") = "SAPP"
                              If Xloslabos <> "" Then
                                 data_lista.Recordset("obs") = Mid(Xloslabos, 1, 200)
                              End If
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
                              mfnac.Text = "__/__/____"
                              t_conv.Text = ""
                              cbotipcons.ListIndex = -1
                              cbosino.ListIndex = -1
                              cbotipoconsu.ListIndex = -1
                              data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " and nom_pac is null"
                              data_lista.Refresh
                              If data_lista.Recordset.RecordCount > 0 Then
                                 If IsNull(data_cabfec.Recordset("completa")) = False Then
                                    data_cabfec.Recordset.Edit
                                    data_cabfec.Recordset("completa") = Null
                                    data_cabfec.Recordset.Update
                                 End If
                              Else
                                 If IsNull(data_cabfec.Recordset("completa")) = False Then
                                    If Int(data_cabfec.Recordset("completa")) <> 1 Then
                                       data_cabfec.Recordset.Edit
                                       data_cabfec.Recordset("completa") = 1
                                       data_cabfec.Recordset.Update
                                    End If
                                 Else
                                    data_cabfec.Recordset.Edit
                                    data_cabfec.Recordset("completa") = 1
                                    data_cabfec.Recordset.Update
                                 End If
                              End If
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
                                    If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
                                       ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
                                    Else
                                       ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
                                    
                                    data_lista.Recordset.MoveNext
                                    Xcountt = Xcountt + 1
                                 Loop
                              End If
                           End If
                        Else
                           MsgBox "No se encuentra registro para actualizar"
                        End If
                     Else
                        If data_cabfec.Recordset("especial") = "FISIOTERAPIA" Then
                           data_lista.RecordSource = "select * from t_fechas where cdate(fecha) <=#" & Format("01/01/2010", "dd/mm/yyyy") & "# and ced_pac ='" & Trim(Xlacedconsulta) & "' and especial ='" & data_cabfec.Recordset("especial") & "'"
                           data_lista.Refresh
                        Else
                           data_lista.RecordSource = "select * from t_fechas where cdate(fecha) >=#" & Format(Xlafechacons, "dd/mm/yyyy") & "# and ced_pac ='" & Trim(Xlacedconsulta) & "' and especial ='" & data_cabfec.Recordset("especial") & "'"
                           data_lista.Refresh
                        End If
                         If data_lista.Recordset.RecordCount > 0 Then
                            MsgBox "Ya se encuentra anotado para una consulta con ésta especialidad", vbExclamation
                         Else
                            data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                            data_lista.Refresh
        'JMOLINS,GUSTAVO
                            If data_lista.Recordset.RecordCount > 0 Then
                               If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                                  MsgBox "Ya existe un paciente anotado, verifique!!", vbCritical
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
                                  data_lista.Recordset("tipoconsulta") = cbotipoconsu.Text
                                  data_lista.Recordset("tipoconsultan") = cbotipoconsu.ListIndex
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
                                  data_lista.Recordset("hcsiono") = cbosino.ListIndex
                                  data_lista.Recordset("tipo_cons") = cbotipcons.ListIndex
                                  If cbotipcons.ListIndex >= 0 Then
                                     data_lista.Recordset("tipo_consd") = cbotipcons.Text
                                  End If
                                  data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                                  data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                                  data_lista.Recordset("usua_anota") = WElusuario
                                  If t_a.Text <> "" Then
                                     If t_m.Text <> "" Then
                                        If t_d.Text <> "" Then
                                           data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES " & t_d.Text & " DIAS"
                                        Else
                                           data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES "
                                        End If
                                     Else
                                        data_lista.Recordset("edad") = t_a.Text & "AÑOS"
                                     End If
                                  Else
                                     If t_m.Text <> "" Then
                                        If t_d.Text <> "" Then
                                           data_lista.Recordset("edad") = t_m.Text & "MESES " & t_d.Text & " DIAS"
                                        Else
                                           data_lista.Recordset("edad") = t_m.Text & "MESES "
                                        End If
                                     Else
                                        If t_d.Text <> "" Then
                                           data_lista.Recordset("edad") = t_d.Text & "DIAS"
                                        End If
                                     End If
                                  End If
                                  data_lista.Recordset("usua_web") = "SAPP"
                                  If Xloslabos <> "" Then
                                     data_lista.Recordset("obs") = Mid(Xloslabos, 1, 200)
                                  End If
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
                                  mfnac.Text = "__/__/____"
                                  t_conv.Text = ""
                                  cbotipcons.ListIndex = -1
                                  cbosino.ListIndex = -1
                                  cbotipoconsu.ListIndex = -1
                                  
                                  data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " and nom_pac is null"
                                  data_lista.Refresh
                                  If data_lista.Recordset.RecordCount > 0 Then
                                     If IsNull(data_cabfec.Recordset("completa")) = False Then
                                        data_cabfec.Recordset.Edit
                                        data_cabfec.Recordset("completa") = Null
                                        data_cabfec.Recordset.Update
                                     End If
                                  Else
                                     If IsNull(data_cabfec.Recordset("completa")) = False Then
                                        If Int(data_cabfec.Recordset("completa")) <> 1 Then
                                           data_cabfec.Recordset.Edit
                                           data_cabfec.Recordset("completa") = 1
                                           data_cabfec.Recordset.Update
                                        End If
                                     Else
                                        data_cabfec.Recordset.Edit
                                        data_cabfec.Recordset("completa") = 1
                                        data_cabfec.Recordset.Update
                                     End If
                                  End If
                                  
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
                                        If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
                                           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
                                        Else
                                           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
                                        
                                        data_lista.Recordset.MoveNext
                                        Xcountt = Xcountt + 1
                                     Loop
                                  End If
                               End If
                            Else
                               MsgBox "No se encuentra registro para actualizar"
                            End If
                         End If
                     End If
                  End If
              Next Xind
              ListView1.ListItems.Item(Xellugar).Selected = True
              ListView1.ListItems.Item(Xellugar).EnsureVisible
              ListView1.SetFocus
    
           Else
              MsgBox "Debe seleccionar un solo registro"
           End If
       
       
       End If
    Else
       If Xestaok = 25 Then
          MsgBox "No se puede agendar sin realizar carta", vbCritical
          Xestaok = 0
          Unload Me
       Else
           If Xcant = 1 Then
              For Xind = 1 To ListView1.ListItems.count
                  ListView1.ListItems(Xind).Selected = True
                  If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                     Xnro = ListView1.ListItems(Xind).Text
                     Xellugar = Xind
                     If cbomedico.Text = "FERTILAB" Then
                        Xloslabos = InputBox("Ingrese los ANALISIS a realizar")
                     End If
                     If data_cabfec.Recordset("especial") = "GINECOLOGIA" Or data_cabfec.Recordset("especial") = "PEDIATRIA" Then
                        data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                        data_lista.Refresh
    'JMOLINS,GUSTAVO
                        If data_lista.Recordset.RecordCount > 0 Then
                           If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                              MsgBox "Ya existe un paciente anotado, verifique!!", vbCritical
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
                              data_lista.Recordset("tipoconsulta") = cbotipoconsu.Text
                              data_lista.Recordset("tipoconsultan") = cbotipoconsu.ListIndex
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
                              data_lista.Recordset("hcsiono") = cbosino.ListIndex
                              data_lista.Recordset("tipo_cons") = cbotipcons.ListIndex
                              If cbotipcons.ListIndex >= 0 Then
                                 data_lista.Recordset("tipo_consd") = cbotipcons.Text
                              End If
                              data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                              data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                              data_lista.Recordset("usua_anota") = WElusuario
                              If t_a.Text <> "" Then
                                 If t_m.Text <> "" Then
                                    If t_d.Text <> "" Then
                                       data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES " & t_d.Text & " DIAS"
                                    Else
                                       data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES "
                                    End If
                                 Else
                                    data_lista.Recordset("edad") = t_a.Text & "AÑOS"
                                 End If
                              Else
                                 If t_m.Text <> "" Then
                                    If t_d.Text <> "" Then
                                       data_lista.Recordset("edad") = t_m.Text & "MESES " & t_d.Text & " DIAS"
                                    Else
                                       data_lista.Recordset("edad") = t_m.Text & "MESES "
                                    End If
                                 Else
                                    If t_d.Text <> "" Then
                                       data_lista.Recordset("edad") = t_d.Text & "DIAS"
                                    End If
                                 End If
                              End If
                              data_lista.Recordset("usua_web") = "SAPP"
                              If Xloslabos <> "" Then
                                 data_lista.Recordset("obs") = Mid(Xloslabos, 1, 200)
                              End If
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
                              mfnac.Text = "__/__/____"
                              t_conv.Text = ""
                              cbotipcons.ListIndex = -1
                              cbosino.ListIndex = -1
                              cbotipoconsu.ListIndex = -1
                              data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " and nom_pac is null"
                              data_lista.Refresh
                              If data_lista.Recordset.RecordCount > 0 Then
                                 If IsNull(data_cabfec.Recordset("completa")) = False Then
                                    data_cabfec.Recordset.Edit
                                    data_cabfec.Recordset("completa") = Null
                                    data_cabfec.Recordset.Update
                                 End If
                              Else
                                 If IsNull(data_cabfec.Recordset("completa")) = False Then
                                    If Int(data_cabfec.Recordset("completa")) <> 1 Then
                                       data_cabfec.Recordset.Edit
                                       data_cabfec.Recordset("completa") = 1
                                       data_cabfec.Recordset.Update
                                    End If
                                 Else
                                    data_cabfec.Recordset.Edit
                                    data_cabfec.Recordset("completa") = 1
                                    data_cabfec.Recordset.Update
                                 End If
                              End If
                              
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
                                    If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
                                       ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
                                    Else
                                       ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
                                    
                                    data_lista.Recordset.MoveNext
                                    Xcountt = Xcountt + 1
                                 Loop
                              End If
                           End If
                        Else
                           MsgBox "No se encuentra registro para actualizar"
                        End If
                     Else
                         data_lista.RecordSource = "select * from t_fechas where cdate(fecha) >=#" & Format(Xlafechacons, "dd/mm/yyyy") & "# and ced_pac ='" & Trim(Xlacedconsulta) & "' and especial ='" & data_cabfec.Recordset("especial") & "'"
                         data_lista.Refresh
                         If data_lista.Recordset.RecordCount > 0 Then
                            MsgBox "Ya se encuentra anotado para una consulta con ésta especialidad", vbExclamation
                         Else
                            data_lista.RecordSource = "select * from t_fechas where fecha ='" & t_feccab.Text & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                            data_lista.Refresh
                            If data_lista.Recordset.RecordCount > 0 Then
                               If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                                  MsgBox "Ya existe un paciente anotado, verifique!!", vbCritical
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
                                  data_lista.Recordset("tipoconsulta") = cbotipoconsu.Text
                                  data_lista.Recordset("tipoconsultan") = cbotipoconsu.ListIndex
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
                                  data_lista.Recordset("hcsiono") = cbosino.ListIndex
                                  data_lista.Recordset("tipo_cons") = cbotipcons.ListIndex
                                  If cbotipcons.ListIndex >= 0 Then
                                     data_lista.Recordset("tipo_consd") = cbotipcons.Text
                                  End If
                                  data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                                  data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                                  data_lista.Recordset("usua_anota") = WElusuario
                                  If t_a.Text <> "" Then
                                     If t_m.Text <> "" Then
                                        If t_d.Text <> "" Then
                                           data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES " & t_d.Text & " DIAS"
                                        Else
                                           data_lista.Recordset("edad") = t_a.Text & " AÑOS " & t_m.Text & " MESES "
                                        End If
                                     Else
                                        data_lista.Recordset("edad") = t_a.Text & "AÑOS"
                                     End If
                                  Else
                                     If t_m.Text <> "" Then
                                        If t_d.Text <> "" Then
                                           data_lista.Recordset("edad") = t_m.Text & "MESES " & t_d.Text & " DIAS"
                                        Else
                                           data_lista.Recordset("edad") = t_m.Text & "MESES "
                                        End If
                                     Else
                                        If t_d.Text <> "" Then
                                           data_lista.Recordset("edad") = t_d.Text & "DIAS"
                                        End If
                                     End If
                                  End If
                                  data_lista.Recordset("usua_web") = "SAPP"
                                  If Xloslabos <> "" Then
                                     data_lista.Recordset("obs") = Mid(Xloslabos, 1, 200)
                                  End If
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
                                  mfnac.Text = "__/__/____"
                                  t_conv.Text = ""
                                  cbotipcons.ListIndex = -1
                                  cbosino.ListIndex = -1
                                  cbotipoconsu.ListIndex = -1
                                  
                                  data_lista.RecordSource = "select * from t_fechas where fecha ='" & data_cabfec.Recordset("fecha") & "' and cod_cons =" & data_cabfec.Recordset("cod_cons") & " and nom_pac is null"
                                  data_lista.Refresh
                                  If data_lista.Recordset.RecordCount > 0 Then
                                     If IsNull(data_cabfec.Recordset("completa")) = False Then
                                        data_cabfec.Recordset.Edit
                                        data_cabfec.Recordset("completa") = Null
                                        data_cabfec.Recordset.Update
                                     End If
                                  Else
                                     If IsNull(data_cabfec.Recordset("completa")) = False Then
                                        If Int(data_cabfec.Recordset("completa")) <> 1 Then
                                           data_cabfec.Recordset.Edit
                                           data_cabfec.Recordset("completa") = 1
                                           data_cabfec.Recordset.Update
                                        End If
                                     Else
                                        data_cabfec.Recordset.Edit
                                        data_cabfec.Recordset("completa") = 1
                                        data_cabfec.Recordset.Update
                                     End If
                                  End If
                                  
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
                                        If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
                                           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
                                        Else
                                           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
                                        
                                        data_lista.Recordset.MoveNext
                                        Xcountt = Xcountt + 1
                                     Loop
                                  End If
                               End If
                            Else
                               MsgBox "No se encuentra registro para actualizar"
                            End If
                         End If
                     End If
                  End If
              Next Xind
              ListView1.ListItems.Item(Xellugar).Selected = True
              ListView1.ListItems.Item(Xellugar).EnsureVisible
              ListView1.SetFocus
    
           Else
              MsgBox "Debe seleccionar un solo registro"
           End If
       End If
    End If
Else
    MsgBox "Ingrese CEDULA para anotar"
End If

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
   If WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or WElusuario = "AACUÑA" Or WElusuario = "SMPEREZ" Or WElusuario = "KROMERO" Then
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

frm_especialistas.MousePointer = 11
If Frame4.Visible = True Then
   ListView1.ListItems.Clear
End If
labselfec.Caption = ""

t_idant.Text = Data1.Recordset("id")
t_codmedsel.Text = Data1.Recordset("cod_med")
t_basesel.Text = Data1.Recordset("base")
t_especsel.Text = Data1.Recordset("especialidad")
t_basedescsel.Text = Data1.Recordset("basedesc")
If IsNull(Data1.Recordset("constel")) = False Then
   chconstel.Value = Data1.Recordset("constel")
Else
   chconstel.Value = 0
End If

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
strSQL = "SELECT t.id,t.fecha,t.hora,t.cod_med,t.cod_cons,t.des_fecha,t.base,t.especial,t.nom_med,t.base_desc,t.cancela,t.motivo,t.usuario,t.fecha_can,t.hora_can,t.cant_pac,t.hora_fin,t.cant_pacok,t.fecha_anota,t.enviado,t.consult,t.id_hora_consultorio,t.fecha_cons,t.fecha_anotad,t.completa, c.desc_consultorio FROM (t_cabfechas t " & _
          "LEFT JOIN horas_consultorios h on t.id_hora_consultorio = h.id) " & _
          "LEFT JOIN consultorios c ON c.id=h.id_consultorio " & _
          "WHERE t.cod_med = " & t_codmedsel.Text & _
              "AND t.base = " & t_basesel.Text & _
              " AND cdate(t.fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & _
              "# ORDER BY cdate(t.fecha)"


If Check1.Value = 1 Then
   If WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or WElusuario = "AACUÑA" Or WElusuario = "SMPEREZ" Or WElusuario = "KROMERO" Then
      'antes data_cabfec.RecordSource = "Select * from t_cabfechas t where t.cod_med =" & t_codmedsel.Text & " and t.base =" & t_basesel.Text & " and cdate(t.fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(t.fecha)"
      data_cabfec.RecordSource = strSQL
      data_cabfec.Refresh
   Else
'      data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(fecha)"
'      data_cabfec.RecordSource = "Select * from t_cabfechas where cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " and cdate(fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(fecha)"
      'antes data_cabfec.RecordSource = "Select * from t_cabfechas t where t.cod_med =" & t_codmedsel.Text & " and t.base =" & t_basesel.Text & " and cdate(t.fecha) >=#" & Format(Xfecconsdesp, "yyyy/mm/dd") & "# order by cdate(t.fecha)"
      
      data_cabfec.RecordSource = strSQL
      data_cabfec.Refresh
   End If
Else
'   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " order by cdate(fecha)"
strSQL = "SELECT t.id,t.fecha,t.hora,t.cod_med,t.cod_cons,t.des_fecha,t.base,t.especial,t.nom_med,t.base_desc,t.cancela,t.motivo,t.usuario,t.fecha_can,t.hora_can,t.cant_pac,t.hora_fin,t.cant_pacok,t.fecha_anota,t.enviado,t.consult,t.id_hora_consultorio,t.fecha_cons,t.fecha_anotad,t.completa, c.desc_consultorio FROM (t_cabfechas t " & _
          "LEFT JOIN horas_consultorios h on t.id_hora_consultorio = h.id) " & _
          "LEFT JOIN consultorios c ON c.id=h.id_consultorio " & _
          "WHERE t.cod_med = " & t_codmedsel.Text & _
              "AND t.base = " & t_basesel.Text & _
              " AND cdate(t.fecha) >=#" & Format("16/04/2021", "yyyy/mm/dd") & _
              "# ORDER BY cdate(t.fecha)"
'   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & t_codmedsel.Text & " and base =" & t_basesel.Text & " order by cdate(fecha)"
   data_cabfec.RecordSource = strSQL
   data_cabfec.Refresh
   
End If

frm_especialistas.MousePointer = 0

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
Dim Xcantlibres As Integer
Xcountt = 1
Xcantlibres = 0
frm_especialistas.MousePointer = 11
t_feccab.Text = data_cabfec.Recordset("fecha")
t_codcons.Text = data_cabfec.Recordset("cod_cons")
labselfec.Caption = data_cabfec.Recordset("des_fecha")


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

If Format(CDate(data_cabfec.Recordset("fecha_anota")), "yyyy/mm/dd") <= Format(Date, "yyyy/mm/dd") Then
   ListView1.Enabled = True
   b_agrega.Enabled = True
   b_elianota.Enabled = True
Else
   MsgBox "Fecha habilitada para anotar a partir de: " & data_cabfec.Recordset("fecha_anota")
   ListView1.Enabled = False
   b_agrega.Enabled = False
   b_elianota.Enabled = False
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
        If IsNull(data_lista.Recordset("tipoconsulta")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipoconsulta")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin datos"
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
'              Xcantlibres = Xcantlibres + 1
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
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
        If IsNull(data_lista.Recordset("tipo_consd")) = False Then
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lista.Recordset("tipo_consd")
        Else
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
        End If
        If IsNull(data_lista.Recordset("usua_web")) = False Then
           If data_lista.Recordset("usua_web") = "WEB" Then
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
           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "Sin Datos"
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
   frm_especialistas.MousePointer = 0
   MsgBox "No existe fecha"
   ListView1.ListItems.Clear
End If
                      
If data_cabfec.Recordset("especial") = "VACUNACION" Or _
   data_cabfec.Recordset("especial") = "RADIOLOGIA" Or _
   data_cabfec.Recordset("especial") = "CARNE DE SALUD" Or _
   data_cabfec.Recordset("especial") = "LABORATORIO" Or _
   data_cabfec.Recordset("especial") = "NUTRICIONISTA" Or _
   data_cabfec.Recordset("especial") = "FISIOTERAPIA" Or _
   data_cabfec.Recordset("especial") = "OFTALMOLOGIA" Then
   cbotipoconsu.ListIndex = 0

Else
   cbotipoconsu.ListIndex = -1
End If
frm_especialistas.MousePointer = 0

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

data_aut.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_buscar.Connect = "ODBC;DSN=" & Xconexrmt & ";"

data_buscaus.Connect = "odbc;dsn=" & Xconexrmt & ";"


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
data_veoesp.RecordSource = "Select * from t_espec where especialidad not in ('AFILIACIONES','HNF') order by base,nombre"
data_veoesp.Refresh

data_lista.Connect = "ODBC;DSN=" & Xconexrmt & ";"

data_buscnv.ConnectionString = "DSN=" & Xconexrmt

Check1.Value = 1

Data1.Connect = "ODBC;DSN=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from t_espec where especialidad not in ('AFILIACIONES','HNF') order by base,nombre"
Data1.Refresh

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
If ControlUsuario("Especialistas") = 1 Then
   b_nuevafecha.Enabled = True
   b_cancecons.Enabled = True
   b_nuevo.Enabled = True
   b_modif.Enabled = True
   b_elim.Enabled = True
   b_edimed.Enabled = True
   b_infos.Enabled = True
   b_elifec.Enabled = True
   mnuevaf.Enabled = True
Else
   If WElusuario = "SDOMINGUEZ" Or WElusuario = "JFERNAN" Or WElusuario = "AACUÑA" Or WElusuario = "COMPUTOS" Or WElusuario = "KARINAROMERO" Or WElusuario = "CLOVRECICH" Or WElusuario = "MICAELA" Or WElusuario = "MARTINC" Or WElusuario = "KROMERO" Then
      b_nuevafecha.Enabled = True
      b_cancecons.Enabled = True
      b_nuevo.Enabled = True
      b_modif.Enabled = True
      b_elim.Enabled = True
      b_edimed.Enabled = True
      b_infos.Enabled = True
      mnuevaf.Enabled = True
   End If
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
         Data1.RecordSource = "Select * from t_espec where especialidad ='" & t_busca.Text & "' and especialidad not in ('AFILIACIONES','HNF') order by base"
         Data1.Refresh
      Else
         Data1.RecordSource = "Select * from t_espec where especialidad not in ('AFILIACIONES','HNF') order by especialidad,base"
         Data1.Refresh
      End If
   Else
      If t_busca.Text <> "" Then
         Data1.RecordSource = "Select * from t_espec where nombre Like '*" + t_busca.Text + "*' and especialidad not in ('AFILIACIONES','HNF') order by nombre,base"
         Data1.Refresh
      Else
         Data1.RecordSource = "Select * from t_espec where especialidad not in ('AFILIACIONES','HNF') order by base,nombre"
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
   t_nompac.SetFocus
End If

End Sub

Private Sub t_ced_LostFocus()
If t_ced.Text <> "" Then
   If t_ced.Text <> "0" Then
         data_busca2.RecordSource = "Select * from clientes where cl_cedula =" & t_ced.Text
         data_busca2.Refresh
         If data_busca2.Recordset.RecordCount > 0 Then
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
            If t_conv.Text = "APS" Then
               MsgBox "Categoría no habilitada.", vbCritical
               t_ced.Text = ""
               t_mat.Text = ""
               t_nompac.Text = ""
               t_celu.Text = ""
               t_tellinea.Text = ""
               mfnac.Text = "__/__/____"
            End If
         Else
            t_mat.Text = ""
            t_codced.Text = 0
            t_nompac.Text = ""
            t_celu.Text = ""
            t_tellinea.Text = ""
            mfnac.Text = "__/__/____"
            t_conv.Text = "PART"
            cbotipcons.ListIndex = -1
            cbosino.ListIndex = -1
         End If
   Else
        t_conv.Text = "PART"
   End If
Else
   t_conv.Text = "PART"
End If

         
End Sub

Private Sub t_celu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_tellinea.SetFocus
End If

End Sub

Private Sub t_conv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_celu.SetFocus
End If

End Sub

Private Sub t_conv_LostFocus()
If t_conv.Text <> "" Then
   If t_conv.Text = "APS" Then
      MsgBox "Categoría no habilitada.", vbCritical
      t_ced.Text = ""
      t_mat.Text = ""
      t_nompac.Text = ""
      t_celu.Text = ""
      t_tellinea.Text = ""
      mfnac.Text = "__/__/____"
   End If
End If
      
End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_ced.SetFocus
End If

End Sub

Private Sub t_mat_LostFocus()
If t_mat.Text <> "" Then
   If t_mat.Text <> "0" Then
      data_busca2.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
      data_busca2.Refresh
      If data_busca2.Recordset.RecordCount > 0 Then
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
         If t_conv.Text = "APS" Then
            MsgBox "Categoría no habilitada.", vbCritical
            t_ced.Text = ""
            t_mat.Text = ""
            t_nompac.Text = ""
            t_celu.Text = ""
            t_tellinea.Text = ""
            mfnac.Text = "__/__/____"
         End If
      Else
         t_mat.Text = ""
         t_ced.Text = ""
         t_codced.Text = ""
         t_nom.Text = ""
         t_celu.Text = ""
         t_tellinea.Text = ""
         mfnac.Text = "__/__/____"
         t_conv.Text = "PART"
         cbotipcons.ListIndex = -1
         cbosino.ListIndex = -1
      End If
   Else
      t_conv.Text = "PART"
   End If
Else
   t_conv.Text = "PART"
End If

         
End Sub

Private Sub t_mm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_espera.SetFocus
End If

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
   mfnac.SetFocus
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

Public Sub Verifica_datosJ()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Fecha_Datos As Date
Fecha_Datos = Date - 365

DatosVerificadosOk = 1

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from clientes_history where cl_codigo =" & Val(t_mat.Text) & " and fecha_modif >='" & Format(Fecha_Datos, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount <= 0 Then
   XorigenDatos = 3 'Agenda
Else
   DatosVerificadosOk = 0
End If

Xrecclii.Close
ConbdSapp.Close

If DatosVerificadosOk = 1 Then
   frm_valida_datos_socio.Show vbModal
End If

End Sub

Public Function Cgalicia() As Integer

Dim XsqlpromoF As String
Dim XreccliiAvisoF As New ADODB.Recordset

ConectarAvisoF
ConbdSappAvisoF.Open

If t_conv.Text <> "" Then
   XsqlpromoF = "Select * from convenio where cnv_codigo ='" & t_conv.Text & "' and cnv_grupo in ('CASA DE GALICIA')"
Else
   XsqlpromoF = "Select * from convenio where cnv_codigo ='" & "PART" & "'"
End If
With XreccliiAvisoF
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
End With
If XreccliiAvisoF.RecordCount > 0 Then
   Cgalicia = 1
Else
   If t_conv.Text = "CASANR" Then
      Cgalicia = 1
   Else
      Cgalicia = 0
   End If
End If
XreccliiAvisoF.Close
ConbdSappAvisoF.Close


End Function

