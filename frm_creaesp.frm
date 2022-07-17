VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_creaesp 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear fechas de especialistas / Anotación de pacientes"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "frm_creaesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_d 
      Height          =   375
      Left            =   8640
      TabIndex        =   44
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox t_m 
      Height          =   375
      Left            =   7920
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox t_a 
      Height          =   405
      Left            =   7320
      TabIndex        =   42
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Data data_deuda 
      Caption         =   "data_deuda"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data data_faltas 
      Caption         =   "data_faltas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Falta sin Aviso"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Falta con aviso"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_lin2 
      Caption         =   "data_lin2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Nro."
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Text            =   "Hora"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Text            =   "Matrícula"
         Object.Width           =   1764
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
         Text            =   "Cédula"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "f"
         Text            =   "Convenio"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "g"
         Text            =   "Fecha Nac."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "h"
         Text            =   "Teléfonos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "i"
         Text            =   "Tipo Consulta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "j"
         Text            =   "Tiene HC?"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "EDAD"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Datos del paciente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   10815
      Begin VB.TextBox t_codcedp 
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
         Left            =   8640
         TabIndex        =   41
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox t_telp 
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
         TabIndex        =   40
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox t_nomp 
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
         TabIndex        =   39
         Top             =   840
         Width           =   6975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modifica"
         Height          =   615
         Left            =   9240
         Picture         =   "frm_creaesp.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton b_agreg 
         BackColor       =   &H0080C0FF&
         Caption         =   "Agregar"
         Height          =   615
         Left            =   9240
         Picture         =   "frm_creaesp.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cbosn 
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
         ItemData        =   "frm_creaesp.frx":0F56
         Left            =   8040
         List            =   "frm_creaesp.frx":0F60
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox t_cedp 
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
         Left            =   7200
         TabIndex        =   30
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cbotip 
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
         ItemData        =   "frm_creaesp.frx":0F6C
         Left            =   6720
         List            =   "frm_creaesp.frx":0F79
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1800
         Width           =   2295
      End
      Begin MSMask.MaskEdBox mfn 
         Height          =   375
         Left            =   4800
         TabIndex        =   25
         Top             =   1320
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
      Begin VB.TextBox t_convp 
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
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton b_bus2 
         Caption         =   "Buscar..."
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
         Left            =   9240
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox t_matp 
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
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "H.C.?"
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
         Left            =   7440
         TabIndex        =   31
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CEDULA:"
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
         Left            =   5640
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tipo consulta"
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
         Left            =   5280
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "TELEFONOS:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "FEC.NAC:"
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
         Left            =   3720
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CONVENIO:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NOMBRE:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "MATRICULA:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FF80&
      Caption         =   "Consulta cancelada por:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frm_creaesp.frx":0FA1
      Left            =   3000
      List            =   "frm_creaesp.frx":0FAE
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton b_bussoc 
      BackColor       =   &H00FFFF80&
      Caption         =   "Buscar..."
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Buscar socio por..."
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton b_abrir 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ABRIR LISTA"
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
      Left            =   9120
      MouseIcon       =   "frm_creaesp.frx":0FE9
      MousePointer    =   99  'Custom
      Picture         =   "frm_creaesp.frx":12F3
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   1455
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENTES"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton b_sele 
      BackColor       =   &H00FFFF80&
      Caption         =   "Seleccionar"
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
      Left            =   4200
      MouseIcon       =   "frm_creaesp.frx":1735
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Anotar a especialista la matrícula ingresada"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txt_mat 
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
      Left            =   2040
      TabIndex        =   10
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton b_cierra 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CERRAR"
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   1095
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
      Top             =   6240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_fechas 
      Caption         =   "data_fechas"
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
      RecordSource    =   "fechasesp"
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_creaesp.frx":1A3F
      Height          =   3975
      Left            =   240
      OleObjectBlob   =   "frm_creaesp.frx":1A58
      TabIndex        =   7
      Top             =   2160
      Width           =   10695
   End
   Begin VB.CommandButton b_crea 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CREAR"
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
      Left            =   6960
      MouseIcon       =   "frm_creaesp.frx":2FBB
      MousePointer    =   99  'Custom
      Picture         =   "frm_creaesp.frx":32C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSMask.MaskEdBox mfec 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin VB.Label Label16 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frm_creaesp.frx":3707
      Height          =   615
      Left            =   6000
      TabIndex        =   45
      Top             =   6840
      Width           =   5055
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080FFFF&
      Caption         =   "Faltas a la consulta:"
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
      Left            =   120
      TabIndex        =   36
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11160
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "MATRICULA:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   3
      X1              =   0
      X2              =   11160
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Especialista:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_creaesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_abrir_Click()
Dim xcont, xhh, xmm As Integer
Dim xcomi, xhortex As String


If WNomesp >= "PEDIATRIA " And WNomesp <= "PG" Then
   Frame1.Visible = True
   DBGrid1.Visible = False
   ListView1.Visible = True
   Dim Xcountt As Long
   Xcountt = 1
   If mfec.Text <> "__/__/____" Then
       Wopsed = 0
       data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
       data_lin2.Refresh
       If data_lin2.Recordset.RecordCount > 0 Then
          data_lin2.Recordset.MoveFirst
          ListView1.ListItems.Clear
          Do While Not data_lin2.Recordset.EOF
             If IsNull(data_lin2.Recordset("cl_nrovend")) = False Then
                ListView1.ListItems.Add Xcountt, , data_lin2.Recordset("cl_nrovend")
             Else
                ListView1.ListItems.Add Xcountt, , "0"
             End If
             If IsNull(data_lin2.Recordset("cl_ruc")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_ruc")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
             End If
             If IsNull(data_lin2.Recordset("cl_atrasoa")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_atrasoa")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add Xcountt, , "0"
             End If
             If IsNull(data_lin2.Recordset("cl_nom_sup")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_nom_sup")
                If Len(data_lin2.Recordset("cl_nom_sup")) > 4 Then
                   Wopsed = Wopsed + 1
                End If
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
             End If
             If IsNull(data_lin2.Recordset("cl_zona")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_zona")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lin2.Recordset("cl_descpag")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_descpag")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
             End If
             If IsNull(data_lin2.Recordset("cl_fultmov")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_fultmov")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
             End If
             If IsNull(data_lin2.Recordset("cl_desc2")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc2")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
             End If
             If IsNull(data_lin2.Recordset("cl_atrasop")) = False Then
                If data_lin2.Recordset("cl_atrasop") = 0 Then
                   ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "META"
                Else
                   If data_lin2.Recordset("cl_atrasop") = 1 Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "CONSULTA"
                   Else
                      If data_lin2.Recordset("cl_atrasop") = 2 Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "RN (Recién Nacido)"
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                      End If
                   End If
                End If
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
             End If
             If IsNull(data_lin2.Recordset("cl_codconv")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_codconv")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
             End If
             If IsNull(data_lin2.Recordset("cl_numero")) = False Then
                If data_lin2.Recordset("cl_numero") = 2 Then
                   ListView1.ListItems.Item(Xcountt).ListSubItems(1).ForeColor = vbRed
                   ListView1.ListItems.Item(Xcountt).ListSubItems(2).ForeColor = vbRed
                   ListView1.ListItems.Item(Xcountt).ListSubItems(3).ForeColor = vbRed

                End If
             Else
'                ListView1.ListItems.Item(Xcountt).ListSubItems(1).ForeColor = vbRed
             End If
             If IsNull(data_lin2.Recordset("cl_desc1")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc1")
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
             End If
             If IsNull(data_lin2.Recordset("cl_val1")) = False Then
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_val1") & " AÑOS " & data_lin2.Recordset("cl_val2") & " MESES " & data_lin2.Recordset("cl_val3") & " DIAS"
             Else
                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
             End If
             Xcountt = Xcountt + 1
             data_lin2.Recordset.MoveNext
          Loop
       Else
          MsgBox "No existe fecha"
          b_cierra.SetFocus
       End If
    Else
       ListView1.ListItems.Clear
       MsgBox "No ingresó fecha", vbCritical, "Mensaje"
       b_cierra.SetFocus
    End If
Else
   Frame1.Visible = False
   If mfec.Text <> "__/__/____" Then
       xcont = 1
       xhh = Val(frm_espec.txt_hh.Text)
       xmm = Val(frm_espec.txt_mm.Text)
       data_lista.RecordSource = "select * from lista where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and base =" & frm_espec.txt_base.Text & " and cod ='" & frm_espec.txt_cod.Text & "'"
       data_lista.Refresh
       If data_lista.Recordset.RecordCount > 0 Then
       Else
          MsgBox "NO existe fecha", vbInformation, "Mensaje"
          b_cierra.SetFocus
       End If
   Else
       MsgBox "No ingresó fecha", vbCritical, "Mensaje"
       b_cierra.SetFocus
   End If
End If

data_fechas.RecordSource = "Select * from fechasesp where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cod ='" & Label2.Caption & "' and base =" & frm_espec.txt_base.Text
data_fechas.Refresh
If data_fechas.Recordset.RecordCount > 0 Then
   If IsNull(data_fechas.Recordset("codmed")) = False Then
      If data_fechas.Recordset("codmed") >= 0 Then
         MsgBox "La consulta seleccionada está CANCELADA"
         Combo1.ListIndex = data_fechas.Recordset("codmed")
         Check1.value = 1
         If ListView1.Visible = True Then
            ListView1.Enabled = False
         Else
            DBGrid1.Enabled = False
         End If
      Else
         Check1.value = 0
         Combo1.ListIndex = -1
         If ListView1.Visible = True Then
            ListView1.Enabled = True
         Else
            DBGrid1.Enabled = True
         End If
      End If
   Else
     Check1.value = 0
     Combo1.ListIndex = -1
      If ListView1.Visible = True Then
         ListView1.Enabled = True
      Else
         DBGrid1.Enabled = True
      End If
   End If
End If

End Sub

Private Sub b_agreg_Click()
Dim Xind, Xcant, Xnro As Long
Xind = 0
Xnro = 0
Xcant = 0
Dim Xcountt As Long
If t_matp.Text <> "" Then
   If t_telp.Text <> "" Then
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & t_matp.Text
      data_cli.Refresh
      If IsNull(data_cli.Recordset("cl_telefon")) = True Then
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_telefon") = t_telp.Text
         data_cli.Recordset.Update
      End If
   End If
End If
For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
       Xcant = Xcant + 1
    End If
Next Xind
Xind = 0
If mfn.Text <> "__/__/____" Then
   CalculaEdad (mfn.Text)
Else
   t_a.Text = ""
   t_d.Text = ""
   t_m.Text = ""
End If

If cbotip.ListIndex = 1 Then
    data_deuda.DatabaseName = App.Path & "\sapp.mdb"
    data_deuda.RecordSource = "Select * from deudas where cliente =" & t_matp.Text & " and tipodoc ='" & "CRE" & "' and fecha_pago is null and fecha >=#" & Format("01/12/2012", "yyyy/mm/dd") & "#"
    data_deuda.Refresh
    If data_deuda.Recordset.RecordCount > 0 Then
       Dim Xladat, Xhoy As Date
       Dim Xq As Integer
       Xhoy = Date
       Xq = 0
       data_deuda.Recordset.MoveFirst
       Do While Not data_deuda.Recordset.EOF
          If IsNull(data_deuda.Recordset("nro_superv")) = False Then
             Xladat = data_deuda.Recordset("fecha") + data_deuda.Recordset("nro_superv")
          Else
             Xladat = data_deuda.Recordset("fecha") + 15
          End If
          If Xladat < Xhoy Then
             Xq = 9
          End If
          data_deuda.Recordset.MoveNext
       Loop
       If Xq = 9 And cbotip.ListIndex = 1 Then
          MsgBox "ATENCION!! Socio con servicios crédito pendientes de PAGO", vbCritical, "DEUDA SOCIO"
          Dim Xcodaut As String
          Xcodaut = InputBox("Ingrese CODIGO DE AUTORIZACION PARA CONTINUAR", "Código autorización")
          If Xcodaut <> "" Then
            If Xcant = 1 Then
               For Xind = 1 To ListView1.ListItems.count
                   ListView1.ListItems(Xind).Selected = True
                   If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
                      Xnro = ListView1.ListItems(Xind).Text
                      data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Xnro
                      data_lin2.Refresh
                      If data_lin2.Recordset.RecordCount > 0 Then
                         If t_matp.Text <> "" Then
                            If t_nomp.Text <> "" And t_telp.Text <> "" Then
                               data_lin2.Recordset.Edit
                               data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                               data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                               If t_cedp.Text <> "" Then
                                  data_lin2.Recordset("cl_zona") = t_cedp.Text
                                  data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                               Else
                                  data_lin2.Recordset("cl_zona") = 0
                                  data_lin2.Recordset("cl_nomcobr") = 0
                               End If
                               If t_convp.Text <> "" Then
                                  data_lin2.Recordset("cl_descpag") = t_convp.Text
                               Else
                                  data_lin2.Recordset("cl_descpag") = "NO REG"
                               End If
                               If mfn.Text <> "__/__/____" Then
                                  data_lin2.Recordset("cl_fultmov") = mfn.Text
                               End If
                               If t_telp.Text <> "" Then
                                  data_lin2.Recordset("cl_desc2") = t_telp.Text
                               Else
                                  data_lin2.Recordset("cl_desc2") = "00"
                               End If
                               data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                               If cbosn.ListIndex >= 0 Then
                                  data_lin2.Recordset("cl_codconv") = cbosn.Text
                               Else
                                  data_lin2.Recordset("cl_codconv") = "NR"
                               End If
                               If t_a.Text <> "" Then
                                  data_lin2.Recordset("cl_val1") = t_a.Text
                               End If
                               If t_m.Text <> "" Then
                                  data_lin2.Recordset("cl_val2") = t_m.Text
                               End If
                               If t_d.Text <> "" Then
                                  data_lin2.Recordset("cl_val3") = t_d.Text
                               End If
                               data_lin2.Recordset("cl_desc1") = WElusuario
                               data_lin2.Recordset.Update
                               t_matp.Text = ""
                               t_nomp.Text = ""
                               t_cedp.Text = ""
                               t_codcedp.Text = ""
                               t_convp.Text = ""
                               mfn.Text = "__/__/____"
                               t_telp.Text = ""
                               cbotip.ListIndex = -1
                               cbosn.ListIndex = -1
                            Else
                               MsgBox "No ingresó datos"
                            End If
                         Else
                            MsgBox "Ingrese matrícula"
                         End If
                      Else
                         MsgBox "No hay datos a modificar"
                      End If
                   End If
               Next Xind
            Else
                If t_matp.Text <> "" Then
                   If t_nomp.Text <> "" And t_telp.Text <> "" Then
                      data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
                      data_lin2.Refresh
                      If data_lin2.Recordset.RecordCount > 0 Then
                         If Wopsed < ListView1.ListItems.count Then
                            Wopsed = Wopsed + 1
                            data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Wopsed
                            data_lin2.Refresh
                            If data_lin2.Recordset.RecordCount > 0 Then
                               data_lin2.Recordset.Edit
                               data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                               data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                               If t_cedp.Text <> "" Then
                                  data_lin2.Recordset("cl_zona") = t_cedp.Text
                                  data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                               Else
                                  data_lin2.Recordset("cl_zona") = 0
                                  data_lin2.Recordset("cl_nomcobr") = 0
                               End If
                               If t_convp.Text <> "" Then
                                  data_lin2.Recordset("cl_descpag") = t_convp.Text
                               Else
                                  data_lin2.Recordset("cl_descpag") = "NO REG"
                               End If
                               If mfn.Text <> "__/__/____" Then
                                  data_lin2.Recordset("cl_fultmov") = mfn.Text
                               End If
                               If t_telp.Text <> "" Then
                                  data_lin2.Recordset("cl_desc2") = t_telp.Text
                               Else
                                  data_lin2.Recordset("cl_desc2") = "00"
                               End If
                               data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                               If cbosn.ListIndex >= 0 Then
                                  data_lin2.Recordset("cl_codconv") = cbosn.Text
                               Else
                                  data_lin2.Recordset("cl_codconv") = "NR"
                               End If
                               If t_a.Text <> "" Then
                                  data_lin2.Recordset("cl_val1") = t_a.Text
                               End If
                               If t_m.Text <> "" Then
                                  data_lin2.Recordset("cl_val2") = t_m.Text
                               End If
                               If t_d.Text <> "" Then
                                  data_lin2.Recordset("cl_val3") = t_d.Text
                               End If
                               data_lin2.Recordset("cl_desc1") = WElusuario
                               data_lin2.Recordset.Update
                               t_matp.Text = ""
                               t_nomp.Text = ""
                               t_cedp.Text = ""
                               t_codcedp.Text = ""
                               t_convp.Text = ""
                               mfn.Text = "__/__/____"
                               t_telp.Text = ""
                               cbotip.ListIndex = -1
                               cbosn.ListIndex = -1
                            Else
                               MsgBox "No hay datos a modificar"
                            End If
                         Else
                            MsgBox "No hay números disponibles"
                         End If
                      Else
                         MsgBox "No hay datos a modificar"
                      End If
                   Else
                      MsgBox "Ingrese datos para la lista"
                   End If
                Else
                   MsgBox "Ingrese matrícula"
                End If
            End If
            data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
            data_lin2.Refresh
            Wopsed = 0
            Xcountt = 1
            If data_lin2.Recordset.RecordCount > 0 Then
                ListView1.ListItems.Clear
                data_lin2.Recordset.MoveFirst
                Do While Not data_lin2.Recordset.EOF
                   If IsNull(data_lin2.Recordset("cl_nrovend")) = False Then
                      ListView1.ListItems.Add Xcountt, , data_lin2.Recordset("cl_nrovend")
                   Else
                      ListView1.ListItems.Add Xcountt, , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_ruc")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_ruc")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                   End If
                   If IsNull(data_lin2.Recordset("cl_atrasoa")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_atrasoa")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_nom_sup")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_nom_sup")
                      If Len(data_lin2.Recordset("cl_nom_sup")) > 4 Then
                         Wopsed = Wopsed + 1
                      End If
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_zona")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_zona")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_descpag")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_descpag")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                   If IsNull(data_lin2.Recordset("cl_fultmov")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_fultmov")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                   If IsNull(data_lin2.Recordset("cl_desc2")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc2")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                    If IsNull(data_lin2.Recordset("cl_atrasop")) = False Then
                       If data_lin2.Recordset("cl_atrasop") = 0 Then
                          ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "META"
                       Else
                          If data_lin2.Recordset("cl_atrasop") = 1 Then
                             ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "CONSULTA"
                          Else
                             If data_lin2.Recordset("cl_atrasop") = 2 Then
                                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "RN (Recién Nacido)"
                             Else
                                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                             End If
                          End If
                       End If
                    Else
                       ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                    End If
                   If IsNull(data_lin2.Recordset("cl_codconv")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_codconv")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_desc1")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc1")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_val1")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_val1") & " AÑOS " & data_lin2.Recordset("cl_val2") & " MESES " & data_lin2.Recordset("cl_val3") & " DIAS"
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   Xcountt = Xcountt + 1
                   data_lin2.Recordset.MoveNext
                Loop
            Else
               MsgBox "No hay lista"
            End If
          Else
            MsgBox "Ingrese código de autorización", vbCritical
            Unload Me
          End If
       Else
            If Xcant = 1 Then
               For Xind = 1 To ListView1.ListItems.count
                   ListView1.ListItems(Xind).Selected = True
                   If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
                      Xnro = ListView1.ListItems(Xind).Text
                      data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Xnro
                      data_lin2.Refresh
                      If data_lin2.Recordset.RecordCount > 0 Then
                         If t_matp.Text <> "" Then
                            If t_nomp.Text <> "" And t_telp.Text <> "" Then
                               data_lin2.Recordset.Edit
                               data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                               data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                               If t_cedp.Text <> "" Then
                                  data_lin2.Recordset("cl_zona") = t_cedp.Text
                                  data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                               Else
                                  data_lin2.Recordset("cl_zona") = 0
                                  data_lin2.Recordset("cl_nomcobr") = 0
                               End If
                               If t_convp.Text <> "" Then
                                  data_lin2.Recordset("cl_descpag") = t_convp.Text
                               Else
                                  data_lin2.Recordset("cl_descpag") = "NO REG"
                               End If
                               If mfn.Text <> "__/__/____" Then
                                  data_lin2.Recordset("cl_fultmov") = mfn.Text
                               End If
                               If t_telp.Text <> "" Then
                                  data_lin2.Recordset("cl_desc2") = t_telp.Text
                               Else
                                  data_lin2.Recordset("cl_desc2") = "00"
                               End If
                               data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                               If cbosn.ListIndex >= 0 Then
                                  data_lin2.Recordset("cl_codconv") = cbosn.Text
                               Else
                                  data_lin2.Recordset("cl_codconv") = "NR"
                               End If
                               If t_a.Text <> "" Then
                                  data_lin2.Recordset("cl_val1") = t_a.Text
                               End If
                               If t_m.Text <> "" Then
                                  data_lin2.Recordset("cl_val2") = t_m.Text
                               End If
                               If t_d.Text <> "" Then
                                  data_lin2.Recordset("cl_val3") = t_d.Text
                               End If
                               data_lin2.Recordset("cl_desc1") = WElusuario
                               data_lin2.Recordset.Update
                               t_matp.Text = ""
                               t_nomp.Text = ""
                               t_cedp.Text = ""
                               t_codcedp.Text = ""
                               t_convp.Text = ""
                               mfn.Text = "__/__/____"
                               t_telp.Text = ""
                               cbotip.ListIndex = -1
                               cbosn.ListIndex = -1
                            Else
                               MsgBox "No ingresó datos"
                            End If
                         Else
                            MsgBox "Ingrese matrícula"
                         End If
                      Else
                         MsgBox "No hay datos a modificar"
                      End If
                   End If
               Next Xind
            Else
                If t_matp.Text <> "" Then
                   If t_nomp.Text <> "" And t_telp.Text <> "" Then
                      data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
                      data_lin2.Refresh
                      If data_lin2.Recordset.RecordCount > 0 Then
                         If Wopsed < ListView1.ListItems.count Then
                            Wopsed = Wopsed + 1
                            data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Wopsed
                            data_lin2.Refresh
                            If data_lin2.Recordset.RecordCount > 0 Then
                               data_lin2.Recordset.Edit
                               data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                               data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                               If t_cedp.Text <> "" Then
                                  data_lin2.Recordset("cl_zona") = t_cedp.Text
                                  data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                               Else
                                  data_lin2.Recordset("cl_zona") = 0
                                  data_lin2.Recordset("cl_nomcobr") = 0
                               End If
                               If t_convp.Text <> "" Then
                                  data_lin2.Recordset("cl_descpag") = t_convp.Text
                               Else
                                  data_lin2.Recordset("cl_descpag") = "NO REG"
                               End If
                               If mfn.Text <> "__/__/____" Then
                                  data_lin2.Recordset("cl_fultmov") = mfn.Text
                               End If
                               If t_telp.Text <> "" Then
                                  data_lin2.Recordset("cl_desc2") = t_telp.Text
                               Else
                                  data_lin2.Recordset("cl_desc2") = "00"
                               End If
                               data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                               If cbosn.ListIndex >= 0 Then
                                  data_lin2.Recordset("cl_codconv") = cbosn.Text
                               Else
                                  data_lin2.Recordset("cl_codconv") = "NR"
                               End If
                               If t_a.Text <> "" Then
                                  data_lin2.Recordset("cl_val1") = t_a.Text
                               End If
                               If t_m.Text <> "" Then
                                  data_lin2.Recordset("cl_val2") = t_m.Text
                               End If
                               If t_d.Text <> "" Then
                                  data_lin2.Recordset("cl_val3") = t_d.Text
                               End If
                               data_lin2.Recordset("cl_desc1") = WElusuario
                               data_lin2.Recordset.Update
                               t_matp.Text = ""
                               t_nomp.Text = ""
                               t_cedp.Text = ""
                               t_codcedp.Text = ""
                               t_convp.Text = ""
                               mfn.Text = "__/__/____"
                               t_telp.Text = ""
                               cbotip.ListIndex = -1
                               cbosn.ListIndex = -1
                            Else
                               MsgBox "No hay datos a modificar"
                            End If
                         Else
                            MsgBox "No hay números disponibles"
                         End If
                      Else
                         MsgBox "No hay datos a modificar"
                      End If
                   Else
                      MsgBox "Ingrese datos para la lista"
                   End If
                Else
                   MsgBox "Ingrese matrícula"
                End If
            End If
            data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
            data_lin2.Refresh
            Wopsed = 0
            Xcountt = 1
            If data_lin2.Recordset.RecordCount > 0 Then
                ListView1.ListItems.Clear
                data_lin2.Recordset.MoveFirst
                Do While Not data_lin2.Recordset.EOF
                   If IsNull(data_lin2.Recordset("cl_nrovend")) = False Then
                      ListView1.ListItems.Add Xcountt, , data_lin2.Recordset("cl_nrovend")
                   Else
                      ListView1.ListItems.Add Xcountt, , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_ruc")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_ruc")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                   End If
                   If IsNull(data_lin2.Recordset("cl_atrasoa")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_atrasoa")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_nom_sup")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_nom_sup")
                      If Len(data_lin2.Recordset("cl_nom_sup")) > 4 Then
                         Wopsed = Wopsed + 1
                      End If
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_zona")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_zona")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_descpag")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_descpag")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                   If IsNull(data_lin2.Recordset("cl_fultmov")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_fultmov")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                   If IsNull(data_lin2.Recordset("cl_desc2")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc2")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                    If IsNull(data_lin2.Recordset("cl_atrasop")) = False Then
                       If data_lin2.Recordset("cl_atrasop") = 0 Then
                          ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "META"
                       Else
                          If data_lin2.Recordset("cl_atrasop") = 1 Then
                             ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "CONSULTA"
                          Else
                             If data_lin2.Recordset("cl_atrasop") = 2 Then
                                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "RN (Recién Nacido)"
                             Else
                                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                             End If
                          End If
                       End If
                    Else
                       ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                    End If
                   If IsNull(data_lin2.Recordset("cl_codconv")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_codconv")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_desc1")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc1")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_val1")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_val1") & " AÑOS " & data_lin2.Recordset("cl_val2") & " MESES " & data_lin2.Recordset("cl_val3") & " DIAS"
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   Xcountt = Xcountt + 1
                   data_lin2.Recordset.MoveNext
                Loop
            Else
               MsgBox "No hay lista"
            End If
       End If
    Else
            If Xcant = 1 Then
               For Xind = 1 To ListView1.ListItems.count
                   ListView1.ListItems(Xind).Selected = True
                   If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
                      Xnro = ListView1.ListItems(Xind).Text
                      data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Xnro
                      data_lin2.Refresh
                      If data_lin2.Recordset.RecordCount > 0 Then
                         If t_matp.Text <> "" Then
                            If t_nomp.Text <> "" And t_telp.Text <> "" Then
                               data_lin2.Recordset.Edit
                               data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                               data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                               If t_cedp.Text <> "" Then
                                  data_lin2.Recordset("cl_zona") = t_cedp.Text
                                  data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                               Else
                                  data_lin2.Recordset("cl_zona") = 0
                                  data_lin2.Recordset("cl_nomcobr") = 0
                               End If
                               If t_convp.Text <> "" Then
                                  data_lin2.Recordset("cl_descpag") = t_convp.Text
                               Else
                                  data_lin2.Recordset("cl_descpag") = "NO REG"
                               End If
                               If mfn.Text <> "__/__/____" Then
                                  data_lin2.Recordset("cl_fultmov") = mfn.Text
                               End If
                               If t_telp.Text <> "" Then
                                  data_lin2.Recordset("cl_desc2") = t_telp.Text
                               Else
                                  data_lin2.Recordset("cl_desc2") = "00"
                               End If
                               data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                               If cbosn.ListIndex >= 0 Then
                                  data_lin2.Recordset("cl_codconv") = cbosn.Text
                               Else
                                  data_lin2.Recordset("cl_codconv") = "NR"
                               End If
                               If t_a.Text <> "" Then
                                  data_lin2.Recordset("cl_val1") = t_a.Text
                               End If
                               If t_m.Text <> "" Then
                                  data_lin2.Recordset("cl_val2") = t_m.Text
                               End If
                               If t_d.Text <> "" Then
                                  data_lin2.Recordset("cl_val3") = t_d.Text
                               End If
                               data_lin2.Recordset("cl_desc1") = WElusuario
                               data_lin2.Recordset.Update
                               t_matp.Text = ""
                               t_nomp.Text = ""
                               t_cedp.Text = ""
                               t_codcedp.Text = ""
                               t_convp.Text = ""
                               mfn.Text = "__/__/____"
                               t_telp.Text = ""
                               cbotip.ListIndex = -1
                               cbosn.ListIndex = -1
                            Else
                               MsgBox "No ingresó datos"
                            End If
                         Else
                            MsgBox "Ingrese matrícula"
                         End If
                      Else
                         MsgBox "No hay datos a modificar"
                      End If
                   End If
               Next Xind
            Else
                If t_matp.Text <> "" Then
                   If t_nomp.Text <> "" And t_telp.Text <> "" Then
                      data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
                      data_lin2.Refresh
                      If data_lin2.Recordset.RecordCount > 0 Then
                         If Wopsed < ListView1.ListItems.count Then
                            Wopsed = Wopsed + 1
                            data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Wopsed
                            data_lin2.Refresh
                            If data_lin2.Recordset.RecordCount > 0 Then
                               data_lin2.Recordset.Edit
                               data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                               data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                               If t_cedp.Text <> "" Then
                                  data_lin2.Recordset("cl_zona") = t_cedp.Text
                                  data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                               Else
                                  data_lin2.Recordset("cl_zona") = 0
                                  data_lin2.Recordset("cl_nomcobr") = 0
                               End If
                               If t_convp.Text <> "" Then
                                  data_lin2.Recordset("cl_descpag") = t_convp.Text
                               Else
                                  data_lin2.Recordset("cl_descpag") = "NO REG"
                               End If
                               If mfn.Text <> "__/__/____" Then
                                  data_lin2.Recordset("cl_fultmov") = mfn.Text
                               End If
                               If t_telp.Text <> "" Then
                                  data_lin2.Recordset("cl_desc2") = t_telp.Text
                               Else
                                  data_lin2.Recordset("cl_desc2") = "00"
                               End If
                               data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                               If cbosn.ListIndex >= 0 Then
                                  data_lin2.Recordset("cl_codconv") = cbosn.Text
                               Else
                                  data_lin2.Recordset("cl_codconv") = "NR"
                               End If
                               If t_a.Text <> "" Then
                                  data_lin2.Recordset("cl_val1") = t_a.Text
                               End If
                               If t_m.Text <> "" Then
                                  data_lin2.Recordset("cl_val2") = t_m.Text
                               End If
                               If t_d.Text <> "" Then
                                  data_lin2.Recordset("cl_val3") = t_d.Text
                               End If
                               data_lin2.Recordset("cl_desc1") = WElusuario
                               data_lin2.Recordset.Update
                               t_matp.Text = ""
                               t_nomp.Text = ""
                               t_cedp.Text = ""
                               t_codcedp.Text = ""
                               t_convp.Text = ""
                               mfn.Text = "__/__/____"
                               t_telp.Text = ""
                               cbotip.ListIndex = -1
                               cbosn.ListIndex = -1
                            Else
                               MsgBox "No hay datos a modificar"
                            End If
                         Else
                            MsgBox "No hay números disponibles"
                         End If
                      Else
                         MsgBox "No hay datos a modificar"
                      End If
                   Else
                      MsgBox "Ingrese datos para la lista"
                   End If
                Else
                   MsgBox "Ingrese matrícula"
                End If
            End If
            data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
            data_lin2.Refresh
            Wopsed = 0
            Xcountt = 1
            If data_lin2.Recordset.RecordCount > 0 Then
                ListView1.ListItems.Clear
                data_lin2.Recordset.MoveFirst
                Do While Not data_lin2.Recordset.EOF
                   If IsNull(data_lin2.Recordset("cl_nrovend")) = False Then
                      ListView1.ListItems.Add Xcountt, , data_lin2.Recordset("cl_nrovend")
                   Else
                      ListView1.ListItems.Add Xcountt, , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_ruc")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_ruc")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                   End If
                   If IsNull(data_lin2.Recordset("cl_atrasoa")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_atrasoa")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_nom_sup")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_nom_sup")
                      If Len(data_lin2.Recordset("cl_nom_sup")) > 4 Then
                         Wopsed = Wopsed + 1
                      End If
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_zona")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_zona")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                   End If
                   If IsNull(data_lin2.Recordset("cl_descpag")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_descpag")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                   If IsNull(data_lin2.Recordset("cl_fultmov")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_fultmov")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                   If IsNull(data_lin2.Recordset("cl_desc2")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc2")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                   End If
                    If IsNull(data_lin2.Recordset("cl_atrasop")) = False Then
                       If data_lin2.Recordset("cl_atrasop") = 0 Then
                          ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "META"
                       Else
                          If data_lin2.Recordset("cl_atrasop") = 1 Then
                             ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "CONSULTA"
                          Else
                             If data_lin2.Recordset("cl_atrasop") = 2 Then
                                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "RN (Recién Nacido)"
                             Else
                                ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                             End If
                          End If
                       End If
                    Else
                       ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                    End If
                   If IsNull(data_lin2.Recordset("cl_codconv")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_codconv")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_desc1")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc1")
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   If IsNull(data_lin2.Recordset("cl_val1")) = False Then
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_val1") & " AÑOS " & data_lin2.Recordset("cl_val2") & " MESES " & data_lin2.Recordset("cl_val3") & " DIAS"
                   Else
                      ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                   End If
                   
                   Xcountt = Xcountt + 1
                   data_lin2.Recordset.MoveNext
                Loop
            Else
               MsgBox "No hay lista"
            End If
    End If
Else
    If Xcant = 1 Then
       For Xind = 1 To ListView1.ListItems.count
           ListView1.ListItems(Xind).Selected = True
           If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
              Xnro = ListView1.ListItems(Xind).Text
              data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Xnro
              data_lin2.Refresh
              If data_lin2.Recordset.RecordCount > 0 Then
                 If t_matp.Text <> "" Then
                    If t_nomp.Text <> "" And t_telp.Text <> "" Then
                       data_lin2.Recordset.Edit
                       data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                       data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                       If t_cedp.Text <> "" Then
                          data_lin2.Recordset("cl_zona") = t_cedp.Text
                          data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                       Else
                          data_lin2.Recordset("cl_zona") = 0
                          data_lin2.Recordset("cl_nomcobr") = 0
                       End If
                       If t_convp.Text <> "" Then
                          data_lin2.Recordset("cl_descpag") = t_convp.Text
                       Else
                          data_lin2.Recordset("cl_descpag") = "NO REG"
                       End If
                       If mfn.Text <> "__/__/____" Then
                          data_lin2.Recordset("cl_fultmov") = mfn.Text
                       End If
                       If t_telp.Text <> "" Then
                          data_lin2.Recordset("cl_desc2") = t_telp.Text
                       Else
                          data_lin2.Recordset("cl_desc2") = "00"
                       End If
                       data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                       If cbosn.ListIndex >= 0 Then
                          data_lin2.Recordset("cl_codconv") = cbosn.Text
                       Else
                          data_lin2.Recordset("cl_codconv") = "NR"
                       End If
                       If t_a.Text <> "" Then
                          data_lin2.Recordset("cl_val1") = t_a.Text
                       End If
                       If t_m.Text <> "" Then
                          data_lin2.Recordset("cl_val2") = t_m.Text
                       End If
                       If t_d.Text <> "" Then
                          data_lin2.Recordset("cl_val3") = t_d.Text
                       End If
                       data_lin2.Recordset("cl_desc1") = WElusuario
                       data_lin2.Recordset.Update
                       t_matp.Text = ""
                       t_nomp.Text = ""
                       t_cedp.Text = ""
                       t_codcedp.Text = ""
                       t_convp.Text = ""
                       mfn.Text = "__/__/____"
                       t_telp.Text = ""
                       cbotip.ListIndex = -1
                       cbosn.ListIndex = -1
                    Else
                       MsgBox "No ingresó datos"
                    End If
                 Else
                    MsgBox "Ingrese matrícula"
                 End If
              Else
                 MsgBox "No hay datos a modificar"
              End If
           End If
       Next Xind
    Else
        If t_matp.Text <> "" Then
           If t_nomp.Text <> "" And t_telp.Text <> "" Then
              data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
              data_lin2.Refresh
              If data_lin2.Recordset.RecordCount > 0 Then
                 If Wopsed < ListView1.ListItems.count Then
                    Wopsed = Wopsed + 1
                    data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Wopsed
                    data_lin2.Refresh
                    If data_lin2.Recordset.RecordCount > 0 Then
                       data_lin2.Recordset.Edit
                       data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                       data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                       If t_cedp.Text <> "" Then
                          data_lin2.Recordset("cl_zona") = t_cedp.Text
                          data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                       Else
                          data_lin2.Recordset("cl_zona") = 0
                          data_lin2.Recordset("cl_nomcobr") = 0
                       End If
                       If t_convp.Text <> "" Then
                          data_lin2.Recordset("cl_descpag") = t_convp.Text
                       Else
                          data_lin2.Recordset("cl_descpag") = "NO REG"
                       End If
                       If mfn.Text <> "__/__/____" Then
                          data_lin2.Recordset("cl_fultmov") = mfn.Text
                       End If
                       If t_telp.Text <> "" Then
                          data_lin2.Recordset("cl_desc2") = t_telp.Text
                       Else
                          data_lin2.Recordset("cl_desc2") = "00"
                       End If
                       data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                       If cbosn.ListIndex >= 0 Then
                          data_lin2.Recordset("cl_codconv") = cbosn.Text
                       Else
                          data_lin2.Recordset("cl_codconv") = "NR"
                       End If
                       If t_a.Text <> "" Then
                          data_lin2.Recordset("cl_val1") = t_a.Text
                       End If
                       If t_m.Text <> "" Then
                          data_lin2.Recordset("cl_val2") = t_m.Text
                       End If
                       If t_d.Text <> "" Then
                          data_lin2.Recordset("cl_val3") = t_d.Text
                       End If
                       data_lin2.Recordset("cl_desc1") = WElusuario
                       data_lin2.Recordset.Update
                       t_matp.Text = ""
                       t_nomp.Text = ""
                       t_cedp.Text = ""
                       t_codcedp.Text = ""
                       t_convp.Text = ""
                       mfn.Text = "__/__/____"
                       t_telp.Text = ""
                       cbotip.ListIndex = -1
                       cbosn.ListIndex = -1
                    Else
                       MsgBox "No hay datos a modificar"
                    End If
                 Else
                    MsgBox "No hay números disponibles"
                 End If
              Else
                 MsgBox "No hay datos a modificar"
              End If
           Else
              MsgBox "Ingrese datos para la lista"
           End If
        Else
           MsgBox "Ingrese matrícula"
        End If
    End If
    data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
    data_lin2.Refresh
    Wopsed = 0
    Xcountt = 1
    If data_lin2.Recordset.RecordCount > 0 Then
        ListView1.ListItems.Clear
        data_lin2.Recordset.MoveFirst
        Do While Not data_lin2.Recordset.EOF
           If IsNull(data_lin2.Recordset("cl_nrovend")) = False Then
              ListView1.ListItems.Add Xcountt, , data_lin2.Recordset("cl_nrovend")
           Else
              ListView1.ListItems.Add Xcountt, , "0"
           End If
           If IsNull(data_lin2.Recordset("cl_ruc")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_ruc")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
           End If
           If IsNull(data_lin2.Recordset("cl_atrasoa")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_atrasoa")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
           End If
           If IsNull(data_lin2.Recordset("cl_nom_sup")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_nom_sup")
              If Len(data_lin2.Recordset("cl_nom_sup")) > 4 Then
                 Wopsed = Wopsed + 1
              End If
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
           End If
           If IsNull(data_lin2.Recordset("cl_zona")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_zona")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
           End If
           If IsNull(data_lin2.Recordset("cl_descpag")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_descpag")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
           End If
           If IsNull(data_lin2.Recordset("cl_fultmov")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_fultmov")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
           End If
           If IsNull(data_lin2.Recordset("cl_desc2")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc2")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
           End If
            If IsNull(data_lin2.Recordset("cl_atrasop")) = False Then
               If data_lin2.Recordset("cl_atrasop") = 0 Then
                  ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "META"
               Else
                  If data_lin2.Recordset("cl_atrasop") = 1 Then
                     ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "CONSULTA"
                  Else
                     If data_lin2.Recordset("cl_atrasop") = 2 Then
                        ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "RN (Recién Nacido)"
                     Else
                        ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                     End If
                  End If
               End If
            Else
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
            End If
           If IsNull(data_lin2.Recordset("cl_codconv")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_codconv")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
           End If
           If IsNull(data_lin2.Recordset("cl_desc1")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc1")
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
           End If
           If IsNull(data_lin2.Recordset("cl_val1")) = False Then
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_val1") & " AÑOS " & data_lin2.Recordset("cl_val2") & " MESES " & data_lin2.Recordset("cl_val3") & " DIAS"
           Else
              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
           End If
           Xcountt = Xcountt + 1
           data_lin2.Recordset.MoveNext
        Loop
    Else
       MsgBox "No hay lista"
    End If
            
End If
End Sub

Private Sub b_bus2_Click()
Xdeb = 9
frm_buscasocesp.Show vbModal

End Sub

Private Sub b_bussoc_Click()
frm_buscasocesp.Show vbModal

End Sub

Private Sub b_cierra_Click()
Unload Me

End Sub

Private Sub b_crea_Click()
Dim xcont, xhh, xmm, Xhaspac, Xlavuelta As Integer
Dim xcomi, xhortex As String
Dim Xquecantpac As Integer
Dim Xre As Long
Xre = data_parsec.Recordset("limite_dia") + 1

Xlavuelta = 0
If WNomesp >= "PEDIATRIA " And WNomesp <= "PG" Then
   If mfec.Text <> "__/__/____" Then
      xcont = 1
      xhh = Val(frm_espec.txt_hh.Text)
      xmm = Val(frm_espec.txt_mm.Text)
      data_lin2.RecordSource = "Select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
      data_lin2.Refresh
      If data_lin2.Recordset.RecordCount > 0 Then
         MsgBox "Ya existe fecha", vbInformation, "Mensaje"
         b_abrir_Click
      Else
         data_fechas.Recordset.AddNew
         data_fechas.Recordset("base") = frm_espec.txt_base.Text
         data_fechas.Recordset("cod") = frm_espec.txt_cod.Text
         data_fechas.Recordset("desc") = frm_espec.txt_desc.Text
         data_fechas.Recordset("fecha") = Format(mfec.Text, "dd/mm/yyyy")
         If xhh > 9 Then
            If xmm > 9 Then
               data_fechas.Recordset("descfec") = " H." + Trim(Str(xhh)) + ":" + Trim(Str(xmm)) + " " + Label5.Caption
            Else
               data_fechas.Recordset("descfec") = " H." + Trim(Str(xhh)) + ":0" + Trim(Str(xmm)) + " " + Label5.Caption
            End If
         Else
            If xmm > 9 Then
               data_fechas.Recordset("descfec") = " H.0" + Trim(Str(xhh)) + ":" + Trim(Str(xmm)) + " " + Label5.Caption
            Else
               data_fechas.Recordset("descfec") = " H.0" + Trim(Str(xhh)) + ":0" + Trim(Str(xmm)) + " " + Label5.Caption
            End If
         End If
         data_fechas.Recordset.Update
         If frm_espec.Check1.value = 1 Then
            Xhaspac = Val(frm_espec.txt_cantp.Text) * 2
         Else
            Xhaspac = Val(frm_espec.txt_cantp.Text)
         End If
         Xquecantpac = Val(frm_espec.txt_mmpp.Text)
         Do While xcont <= Xhaspac
            If frm_espec.Check1.value = 1 Then
               If Xlavuelta > 1 Then
                  Xlavuelta = 0
                  xmm = xmm + Xquecantpac
                  If xmm >= 60 Then
                     xmm = 0
                     xhh = xhh + 1
                  Else
                     xhh = xhh
                     xmm = xmm
                  End If
               Else
                  data_lin2.Recordset.AddNew
                  data_lin2.Recordset("cl_codigo") = Xre
                  data_lin2.Recordset("cl_nrovend") = xcont
                  data_lin2.Recordset("cl_fnac") = Format(mfec.Text, "dd/mm/yyyy")
                  data_lin2.Recordset("cl_grupo") = frm_espec.txt_base.Text
                  data_lin2.Recordset("cl_fax") = Mid(frm_espec.txt_cod.Text, 1, 5)
                  data_lin2.Recordset("cl_atrasoa") = 0
                  data_lin2.Recordset("cl_zona") = 0
                  data_lin2.Recordset("cl_nomcobr") = 0
'                  data_lin2.Recordset("cl_nom_sup") = " "
                  data_lin2.Recordset("cl_desc2") = " "
                  If xhh > 9 Then
                     If xmm > 9 Then
                        data_lin2.Recordset("cl_ruc") = Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                     Else
                        data_lin2.Recordset("cl_ruc") = Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                     End If
                  Else
                     If xmm > 9 Then
                        data_lin2.Recordset("cl_ruc") = "0" + Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                     Else
                        data_lin2.Recordset("cl_ruc") = "0" + Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                     End If
                  End If
                  data_lin2.Recordset.Update
                  xcont = xcont + 1
                  Xre = Xre + 1
                  Xlavuelta = Xlavuelta + 1
               End If
            Else
               data_lin2.Recordset.AddNew
               data_lin2.Recordset("cl_codigo") = Xre
               data_lin2.Recordset("cl_nrovend") = xcont
               data_lin2.Recordset("cl_fnac") = Format(mfec.Text, "dd/mm/yyyy")
               data_lin2.Recordset("cl_grupo") = frm_espec.txt_base.Text
               data_lin2.Recordset("cl_fax") = Mid(frm_espec.txt_cod.Text, 1, 5)
               data_lin2.Recordset("cl_atrasoa") = 0
               data_lin2.Recordset("cl_zona") = 0
               data_lin2.Recordset("cl_nomcobr") = 0
               'data_lista.Recordset("nompac") = " "
               'data_lista.Recordset("tel") = " "
               If xhh > 9 Then
                  If xmm > 9 Then
                     data_lin2.Recordset("cl_ruc") = Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                  Else
                     data_lin2.Recordset("cl_ruc") = Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                  End If
               Else
                  If xmm > 9 Then
                     data_lin2.Recordset("cl_ruc") = "0" + Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                  Else
                     data_lin2.Recordset("cl_ruc") = "0" + Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                  End If
               End If
               data_lin2.Recordset.Update
               xcont = xcont + 1
               Xre = Xre + 1
               xmm = xmm + Xquecantpac
               If xmm >= 60 Then
                  xmm = 0
                  xhh = xhh + 1
               Else
                  xhh = xhh
                  xmm = xmm
               End If
               Xlavuelta = 0
            End If
         Loop
         xcont = 1
         Xre = Xre + 1
         Do While xcont <= Val(frm_espec.txt_espera.Text)
            data_lin2.Recordset.AddNew
            data_lin2.Recordset("cl_codigo") = Xre
            data_lin2.Recordset("cl_nrovend") = xcont + 100
            data_lin2.Recordset("cl_fnac") = Format(mfec.Text, "dd/mm/yyyy")
            data_lin2.Recordset("cl_grupo") = frm_espec.txt_base.Text
            data_lin2.Recordset("cl_fax") = Mid(frm_espec.txt_cod.Text, 1, 5)
            data_lin2.Recordset("cl_atrasoa") = 0
            data_lin2.Recordset("cl_zona") = 0
            data_lin2.Recordset("cl_nomcobr") = 0
'            data_lista.Recordset("nompac") = " "
    '          data_lista.Recordset("obs") = " "
'              data_lista.Recordset("tel") = " "
            data_lin2.Recordset("cl_ruc") = "00:00"
            data_lin2.Recordset.Update
            xcont = xcont + 1
            Xre = Xre + 1
         Loop
         data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
         data_lin2.Refresh
         data_parsec.Recordset.Edit
         data_parsec.Recordset("limite_dia") = Xre
         data_parsec.Recordset.Update
         MsgBox "Fecha de consulta creada"
         b_abrir_Click
      End If
   Else
      MsgBox "No ingresó fecha", vbCritical, "Mensaje"
      mfec.SetFocus
   End If
Else

    If mfec.Text <> "__/__/____" Then
        xcont = 1
        xhh = Val(frm_espec.txt_hh.Text)
        xmm = Val(frm_espec.txt_mm.Text)
        data_lista.RecordSource = "Select * from lista where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and base =" & frm_espec.txt_base.Text & " and cod ='" & frm_espec.txt_cod.Text & "'"
        data_lista.Refresh
        If data_lista.Recordset.RecordCount > 0 Then
           MsgBox "Ya existe fecha", vbInformation, "Mensaje"
           data_lista.RecordSource = "selec * from lista where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and base =" & frm_espec.txt_base.Text & " and cod ='" & frm_espec.txt_cod.Text & "'"
           data_lista.Refresh
        Else
           data_fechas.Recordset.AddNew
           data_fechas.Recordset("base") = frm_espec.txt_base.Text
           data_fechas.Recordset("cod") = frm_espec.txt_cod.Text
           data_fechas.Recordset("desc") = frm_espec.txt_desc.Text
           data_fechas.Recordset("fecha") = Format(mfec.Text, "dd/mm/yyyy")
           If xhh > 9 Then
              If xmm > 9 Then
                 data_fechas.Recordset("descfec") = " H." + Trim(Str(xhh)) + ":" + Trim(Str(xmm)) + " " + Label5.Caption
              Else
                 data_fechas.Recordset("descfec") = " H." + Trim(Str(xhh)) + ":0" + Trim(Str(xmm)) + " " + Label5.Caption
              End If
           Else
              If xmm > 9 Then
                 data_fechas.Recordset("descfec") = " H.0" + Trim(Str(xhh)) + ":" + Trim(Str(xmm)) + " " + Label5.Caption
              Else
                 data_fechas.Recordset("descfec") = " H.0" + Trim(Str(xhh)) + ":0" + Trim(Str(xmm)) + " " + Label5.Caption
              End If
           End If
           data_fechas.Recordset.Update
           If frm_espec.Check1.value = 1 Then
              Xhaspac = Val(frm_espec.txt_cantp.Text) * 2
           Else
              Xhaspac = Val(frm_espec.txt_cantp.Text)
           End If
           Xquecantpac = Val(frm_espec.txt_mmpp.Text)
           Do While xcont <= Xhaspac
              If frm_espec.Check1.value = 1 Then
                 If Xlavuelta > 1 Then
                    Xlavuelta = 0
                    xmm = xmm + Xquecantpac
                    If xmm >= 60 Then
                       xmm = 0
                       xhh = xhh + 1
                    Else
                       xhh = xhh
                       xmm = xmm
                    End If
                 Else
                    data_lista.Recordset.AddNew
                    data_lista.Recordset("nro") = xcont
                    data_lista.Recordset("fecha") = Format(mfec.Text, "dd/mm/yyyy")
                    data_lista.Recordset("base") = frm_espec.txt_base.Text
                    data_lista.Recordset("cod") = frm_espec.txt_cod.Text
                    data_lista.Recordset("matric") = 0
                    data_lista.Recordset("nompac") = " "
                    data_lista.Recordset("tel") = " "
                    If xhh > 9 Then
                       If xmm > 9 Then
                          data_lista.Recordset("horacom") = Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                       Else
                          data_lista.Recordset("horacom") = Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                       End If
                    Else
                       If xmm > 9 Then
                          data_lista.Recordset("horacom") = "0" + Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                       Else
                          data_lista.Recordset("horacom") = "0" + Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                       End If
                    End If
                    data_lista.Recordset.Update
                    xcont = xcont + 1
                    Xlavuelta = Xlavuelta + 1
                 End If
              Else
                 data_lista.Recordset.AddNew
                 data_lista.Recordset("nro") = xcont
                 data_lista.Recordset("fecha") = Format(mfec.Text, "dd/mm/yyyy")
                 data_lista.Recordset("base") = frm_espec.txt_base.Text
                 data_lista.Recordset("cod") = frm_espec.txt_cod.Text
                 data_lista.Recordset("matric") = 0
                 data_lista.Recordset("nompac") = " "
                 data_lista.Recordset("tel") = " "
                 If xhh > 9 Then
                    If xmm > 9 Then
                       data_lista.Recordset("horacom") = Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                    Else
                       data_lista.Recordset("horacom") = Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                    End If
                 Else
                    If xmm > 9 Then
                       data_lista.Recordset("horacom") = "0" + Trim(Str(xhh)) + ":" + Trim(Str(xmm))
                    Else
                       data_lista.Recordset("horacom") = "0" + Trim(Str(xhh)) + ":0" + Trim(Str(xmm))
                    End If
                 End If
                 data_lista.Recordset.Update
                 xcont = xcont + 1
                 xmm = xmm + Xquecantpac
                 If xmm >= 60 Then
                    xmm = 0
                    xhh = xhh + 1
                 Else
                    xhh = xhh
                    xmm = xmm
                 End If
                 Xlavuelta = 0
              End If
           Loop
           xcont = 1
           Do While xcont <= Val(frm_espec.txt_espera.Text)
              data_lista.Recordset.AddNew
              data_lista.Recordset("nro") = xcont
              data_lista.Recordset("fecha") = Format(mfec.Text, "dd/mm/yyyy")
              data_lista.Recordset("base") = frm_espec.txt_base.Text
              data_lista.Recordset("cod") = frm_espec.txt_cod.Text
              data_lista.Recordset("matric") = 0
              data_lista.Recordset("nompac") = " "
    '          data_lista.Recordset("obs") = " "
              data_lista.Recordset("tel") = " "
              data_lista.Recordset("horacom") = ""
              data_lista.Recordset.Update
              xcont = xcont + 1
           Loop
           
           data_lista.RecordSource = "select * from lista where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and base =" & frm_espec.txt_base.Text & " and cod ='" & frm_espec.txt_cod.Text & "'"
           data_lista.Refresh
           
        End If
    Else
        MsgBox "No ingresó fecha", vbCritical, "Mensaje"
        mfec.SetFocus
    End If
End If

End Sub

Private Sub b_sele_Click()
'Dim XCol, Xlin As Double
'MsgBox "Función no habilitada!! Puede consultar socios desde la opción BUSCAR", vbInformation

End Sub

Private Sub cbosn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_telp.SetFocus
End If

End Sub

Private Sub cbotip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_agreg.SetFocus
End If

End Sub

Private Sub Check1_Click()
If XWeltipoU = "ADMINISTRADOR" Then
    If Check1.value = 1 Then
       If Combo1.ListIndex = -1 Then
          Check1.value = 0
          MsgBox "Seleccione primero el motivo de cancelación"
          Combo1.SetFocus
       Else
            data_fechas.RecordSource = "Select * from fechasesp where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cod ='" & Label2.Caption & "' and base =" & frm_espec.txt_base.Text
            data_fechas.Refresh
            If data_fechas.Recordset.RecordCount > 0 Then
               If IsNull(data_fechas.Recordset("codmed")) = True Then
                  data_fechas.Recordset.Edit
                  data_fechas.Recordset("codmed") = Combo1.ListIndex
                  data_fechas.Recordset.Update
               Else
                  If data_fechas.Recordset("codmed") = Combo1.ListIndex Then
                  Else
                     data_fechas.Recordset.Edit
                     data_fechas.Recordset("codmed") = Combo1.ListIndex
                     data_fechas.Recordset.Update
                  End If
               End If
            End If
        End If
    Else
       data_fechas.RecordSource = "Select * from fechasesp where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cod ='" & Label2.Caption & "' and base =" & frm_espec.txt_base.Text
       data_fechas.Refresh
       If data_fechas.Recordset.RecordCount > 0 Then
          If IsNull(data_fechas.Recordset("codmed")) = True Then
             data_fechas.Recordset.Edit
             data_fechas.Recordset("codmed") = -1
             data_fechas.Recordset.Update
          Else
             If data_fechas.Recordset("codmed") = -1 Then
             Else
                data_fechas.Recordset.Edit
                data_fechas.Recordset("codmed") = -1
                data_fechas.Recordset.Update
             End If
          End If
       End If
    End If
Else
    MsgBox "Usuario no autorizado"
    Check1.value = 0
End If


End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Command1_Click()
data_lista.Refresh

End Sub

Private Sub Command2_Click()
Dim Xmensajeme As String
Dim Xind, Xcant As Long
Dim Xnro As Long
Xcant = 0
For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
       Xcant = Xcant + 1
    End If
Next Xind
If Xcant = 1 Then
   Xmensajeme = MsgBox("Desea MODIFICAR el registro seleccionado?", vbInformation + vbYesNo, "Control")
   Xind = 0
   If Xmensajeme = vbYes Then
      For Xind = 1 To ListView1.ListItems.count
          ListView1.ListItems(Xind).Selected = True
          If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then

'       MsgBox "Chequeado"
'             MsgBox "ES:" &
             Xnro = ListView1.ListItems(Xind).Text
 '         Xmatme = lis1.SelectedItem.ListSubItems(9).Text
             
             data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Xnro
             data_lin2.Refresh
             If data_lin2.Recordset.RecordCount > 0 Then
                If t_nomp.Text = "" And t_convp.Text = "" And mfn.Text = "__/__/____" Then
                    data_lin2.Recordset.Edit
                    data_lin2.Recordset("cl_atrasoa") = 0
                    data_lin2.Recordset("cl_nom_sup") = "sin dato"
                    data_lin2.Recordset("cl_zona") = 0
                    data_lin2.Recordset("cl_nomcobr") = 0
                    data_lin2.Recordset("cl_descpag") = "SC"
                    data_lin2.Recordset("cl_fultmov") = Null
                    data_lin2.Recordset("cl_desc2") = "0"
                    data_lin2.Recordset("cl_atrasop") = -1
                    data_lin2.Recordset("cl_codconv") = "NR"
                    data_lin2.Recordset("cl_desc1") = WElusuario
                    data_lin2.Recordset("cl_val1") = 0
                    data_lin2.Recordset("cl_val2") = 0
                    data_lin2.Recordset("cl_val3") = 0
                    data_lin2.Recordset.Update
                Else
                    data_lin2.Recordset.Edit
                    data_lin2.Recordset("cl_atrasoa") = t_matp.Text
                    data_lin2.Recordset("cl_nom_sup") = t_nomp.Text
                    If t_cedp.Text <> "" Then
                       data_lin2.Recordset("cl_zona") = t_cedp.Text
                       data_lin2.Recordset("cl_nomcobr") = t_codcedp.Text
                    End If
                    If t_convp.Text <> "" Then
                       data_lin2.Recordset("cl_descpag") = t_convp.Text
                    End If
                    If mfn.Text <> "__/__/____" Then
                       data_lin2.Recordset("cl_fultmov") = mfn.Text
                    End If
                    data_lin2.Recordset("cl_desc2") = t_telp.Text
                    data_lin2.Recordset("cl_atrasop") = cbotip.ListIndex
                    If cbosn.ListIndex >= 0 Then
                       data_lin2.Recordset("cl_codconv") = cbosn.Text
                    Else
                       data_lin2.Recordset("cl_codconv") = "NR"
                    End If
                    If t_a.Text <> "" Then
                       data_lin2.Recordset("cl_val1") = t_a.Text
                    End If
                    If t_m.Text <> "" Then
                       data_lin2.Recordset("cl_val2") = t_m.Text
                    End If
                    If t_d.Text <> "" Then
                       data_lin2.Recordset("cl_val3") = t_d.Text
                    End If
                    data_lin2.Recordset("cl_desc1") = WElusuario
                    data_lin2.Recordset.Update
                End If
               t_matp.Text = ""
               t_nomp.Text = ""
               t_cedp.Text = ""
               t_codcedp.Text = ""
               t_convp.Text = ""
               mfn.Text = "__/__/____"
               t_telp.Text = ""
               cbotip.ListIndex = -1
               cbosn.ListIndex = -1
               Xcant = 0
                data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
                data_lin2.Refresh
                Dim Xcountt As Long
                Xcountt = 1
                If data_lin2.Recordset.RecordCount > 0 Then
                   ListView1.ListItems.Clear
                   data_lin2.Recordset.MoveFirst
                   Do While Not data_lin2.Recordset.EOF
                      If IsNull(data_lin2.Recordset("cl_nrovend")) = False Then
                         ListView1.ListItems.Add Xcountt, , data_lin2.Recordset("cl_nrovend")
                      Else
                         ListView1.ListItems.Add Xcountt, , "0"
                      End If
                      If IsNull(data_lin2.Recordset("cl_ruc")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_ruc")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                      End If
                      If IsNull(data_lin2.Recordset("cl_atrasoa")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_atrasoa")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                      End If
                      If IsNull(data_lin2.Recordset("cl_nom_sup")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_nom_sup")
                         If Len(data_lin2.Recordset("cl_nom_sup")) > 4 Then
                            Wopsed = Wopsed + 1
                         End If
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                      End If
                      If IsNull(data_lin2.Recordset("cl_zona")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_zona")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                      End If
                      If IsNull(data_lin2.Recordset("cl_descpag")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_descpag")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(data_lin2.Recordset("cl_fultmov")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_fultmov")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(data_lin2.Recordset("cl_desc2")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc2")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                        If IsNull(data_lin2.Recordset("cl_atrasop")) = False Then
                           If data_lin2.Recordset("cl_atrasop") = 0 Then
                              ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "META"
                           Else
                              If data_lin2.Recordset("cl_atrasop") = 1 Then
                                 ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "CONSULTA"
                              Else
                                 If data_lin2.Recordset("cl_atrasop") = 2 Then
                                    ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "RN (Recién Nacido)"
                                 Else
                                    ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                                 End If
                              End If
                           End If
                        Else
                           ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO REGISTRADO"
                        End If
                      If IsNull(data_lin2.Recordset("cl_codconv")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_codconv")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                      End If
                      If IsNull(data_lin2.Recordset("cl_desc1")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_desc1")
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                      End If
                      If IsNull(data_lin2.Recordset("cl_val1")) = False Then
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin2.Recordset("cl_val1") & " AÑOS " & data_lin2.Recordset("cl_val2") & " MESES " & data_lin2.Recordset("cl_val3") & " DIAS"
                      Else
                         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
                      End If
                      Xcountt = Xcountt + 1
                      data_lin2.Recordset.MoveNext
                   Loop
                Else
                   MsgBox "No existen datos"
                   b_cierra.SetFocus
                End If
             Else
                MsgBox "No se encuentra el registro a modificar"
             End If
          End If
      Next Xind
   End If
Else
   MsgBox "Debe existir un solo registro marcado para modificar"
   b_cierra_Click
End If


End Sub

Private Sub Command3_Click()
Dim Xmensajeme As String
Dim Xind, Xcant As Long
Dim Xnro As Long
Xmensajeme = MsgBox("Desea procesar los registros seleccionados como FALTA CON AVISO?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
'       MsgBox "Chequeado"
          Xnro = ListView1.ListItems(Xind).Text
''          Xnro = ListView1.SelectedItem.ListSubItems(0).Text
          data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Xnro
          data_lin2.Refresh
          If data_lin2.Recordset.RecordCount > 0 Then
             data_lin2.Recordset.Edit
             data_lin2.Recordset("cl_atrasoa") = 0
             data_lin2.Recordset("cl_nom_sup") = "sin dato"
             data_lin2.Recordset("cl_zona") = 0
             data_lin2.Recordset("cl_nomcobr") = 0
             data_lin2.Recordset("cl_descpag") = "SC"
             data_lin2.Recordset("cl_fultmov") = Null
             data_lin2.Recordset("cl_desc2") = "0"
             data_lin2.Recordset("cl_atrasop") = -1
             data_lin2.Recordset("cl_codconv") = "NR"
             data_lin2.Recordset("cl_numero") = 0
             data_lin2.Recordset.Update
             data_faltas.Recordset.AddNew
             data_faltas.Recordset("nom1") = ListView1.SelectedItem.ListSubItems(4).Text
             data_faltas.Recordset("ced") = Val(ListView1.SelectedItem.ListSubItems(2).Text)
             data_faltas.Recordset("fecha") = mfec.Text
             data_faltas.Recordset("codver") = 1
             data_faltas.Recordset("ape1") = Label2.Caption
             data_faltas.Recordset.Update
             
             data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
             data_lin2.Refresh
             MsgBox "Proceso terminado, QUEDARA UN LUGAR LIBRE, VUELVA A INGRESAR!", vbInformation
          End If
       End If
   Next Xind
   Unload Me
End If

End Sub

Private Sub Command4_Click()
Dim Xmensajeme As String
Dim Xind, Xcant As Long
Dim Xnro As Long
Xmensajeme = MsgBox("Desea procesar los registros seleccionados como FALTA SIN AVISO?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then

'       MsgBox "Chequeado"
'             MsgBox "ES:" &
           Xnro = ListView1.ListItems(Xind).Text
 '         Xmatme = lis1.SelectedItem.ListSubItems(9).Text
'          Xnro = ListView1.SelectedItem.ListSubItems(Xind).Text
          data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "' and cl_nrovend =" & Xnro
          data_lin2.Refresh
          If data_lin2.Recordset.RecordCount > 0 Then
             data_lin2.Recordset.Edit
             data_lin2.Recordset("cl_numero") = 2
             data_lin2.Recordset.Update
             data_lin2.RecordSource = "select * from mant_sol where cl_fnac =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and cl_grupo =" & frm_espec.txt_base.Text & " and cl_fax ='" & frm_espec.txt_cod.Text & "'"
             data_lin2.Refresh
             MsgBox "Proceso terminado"
'             Unload Me
          End If
       End If
   Next Xind
   Unload Me
End If

End Sub

Private Sub DBGrid1_AfterUpdate()
'data_lista.Refresh
txt_mat.SetFocus
DBGrid1.Refresh
End Sub

Private Sub Form_Load()
data_fechas.DatabaseName = App.Path & "\sapp.mdb"
data_lista.DatabaseName = App.Path & "\sapp.mdb"
data_cli.DatabaseName = App.Path & "\sapp.mdb"
data_lin2.DatabaseName = App.Path & "\sapp.mdb"
data_parsec.DatabaseName = App.Path & "\parse.mdb"
data_parsec.RecordSource = "parsec0"
data_parsec.Refresh
data_faltas.DatabaseName = App.Path & "\sapp.mdb"
data_faltas.RecordSource = "brou"
data_faltas.Refresh

Label2.Caption = WCodesp
Label3.Caption = WNomesp
'If WElusuario = "JFERNAN" Or WElusuario = "CLAUDIA" Or WElusuario = "AACUÑA" Or WElusuario = "MIKAELA" Or WElusuario = "SDOMINGUEZ" Then
If WElusuario = "JFERNAN" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "MIKAELA" Then
   b_crea.Visible = True
   Check1.Enabled = True
   Combo1.Enabled = True
Else
   b_crea.Visible = False
   Check1.Enabled = False
   Combo1.Enabled = False
End If
If WNomesp >= "PEDIATRIA " And WNomesp <= "PG" Then
   DBGrid1.Visible = False
   Frame1.Visible = True
   ListView1.Visible = True
   Label6.Visible = False
   txt_mat.Visible = False
   b_sele.Visible = False
   b_bussoc.Visible = False
Else
   DBGrid1.Visible = True
   Frame1.Visible = False
   ListView1.Visible = False
   Command3.Visible = False
   Command4.Visible = False
End If
   
End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_abrir.SetFocus
End If

End Sub

Private Sub mfec_LostFocus()
If mfec.Text <> "__/__/____" Then
   Label5.Caption = FormatDateTime(mfec.Text, vbLongDate)
End If

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub t_cedp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codcedp.SetFocus
End If

End Sub

Private Sub t_cedp_LostFocus()
If t_cedp.Text <> "" Then
   If t_cedp.Text <> 0 Then
      data_cli.RecordSource = "Select * from clientes where cl_cedula =" & t_cedp.Text
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         t_matp.Text = data_cli.Recordset("cl_codigo")
         If IsNull(data_cli.Recordset("cl_codced")) = False Then
            t_codcedp.Text = data_cli.Recordset("cl_codced")
         Else
            t_codcedp.Text = 0
         End If
         t_convp.Text = data_cli.Recordset("cl_codconv")
         t_nomp.Text = data_cli.Recordset("cl_apellid")
         If IsNull(data_cli.Recordset("cl_fnac")) = False Then
            mfn.Text = data_cli.Recordset("cl_fnac")
         Else
            mfn.Text = "__/__/____"
         End If
         If IsNull(data_cli.Recordset("cl_telefon")) = False Then
            t_telp.Text = data_cli.Recordset("cl_telefon")
         Else
            t_telp.Text = 0
         End If
         cbosn.ListIndex = 0
         cbotip.SetFocus
      Else
''         MsgBox "Socio no encontrado"
      End If
   End If
End If

End Sub

Private Sub t_codcedp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbosn.SetFocus
End If

End Sub

Private Sub t_convp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfn.SetFocus
End If

End Sub

Private Sub t_matp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then
   t_cedp.SetFocus
End If

End Sub

Private Sub t_matp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nomp.SetFocus
End If

End Sub

Private Sub t_matp_LostFocus()
If t_matp.Text <> "" Then
   If t_matp.Text <> 0 Then
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & t_matp.Text
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         t_cedp.Text = data_cli.Recordset("cl_cedula")
         If IsNull(data_cli.Recordset("cl_codced")) = False Then
            t_codcedp.Text = data_cli.Recordset("cl_codced")
         Else
            t_codcedp.Text = 0
         End If
         t_convp.Text = data_cli.Recordset("cl_codconv")
         t_nomp.Text = data_cli.Recordset("cl_apellid")
         If IsNull(data_cli.Recordset("cl_fnac")) = False Then
            mfn.Text = data_cli.Recordset("cl_fnac")
         Else
            mfn.Text = "__/__/____"
         End If
         If IsNull(data_cli.Recordset("cl_telefon")) = False Then
            t_telp.Text = data_cli.Recordset("cl_telefon")
         Else
            t_telp.Text = 0
         End If
         cbosn.ListIndex = 0
         cbotip.SetFocus
      Else
         MsgBox "Socio no encontrado"
      End If
   End If
End If




End Sub

Private Sub t_nomp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_convp.SetFocus
End If

End Sub

Private Sub t_telp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbotip.SetFocus
End If

End Sub

Private Sub txt_mat_GotFocus()
Command1_Click

DBGrid1.Refresh

End Sub

Private Sub txt_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_mat.Text <> "" Then
      b_sele_Click
   Else
      b_sele.SetFocus
   End If
End If

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

