VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmabm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "Mantenimiento de socios"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ForeColor       =   &H00000000&
   Icon            =   "frmabm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Data data_clientes 
      Caption         =   "data_clientes"
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
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      Height          =   2175
      Left            =   5760
      TabIndex        =   73
      Top             =   0
      Width           =   6015
      Begin VB.TextBox t_paemi 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   117
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox t_pmemi 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   116
         Top             =   1080
         Width           =   495
      End
      Begin VB.Data data_lin 
         Caption         =   "data_lin"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data data_cnvmut 
         Caption         =   "data_cnvmut"
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
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data data_factrep 
         Caption         =   "data_factrep"
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
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton b_histadm 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Historial Adm."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Picture         =   "frmabm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton btn_estadi 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Estadísticas"
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
         Left            =   4440
         Picture         =   "frmabm.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btn_fact 
         BackColor       =   &H00C0FFC0&
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
         Height          =   375
         Left            =   2760
         Picture         =   "frmabm.frx":0F56
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btn_histo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Historial Soc"
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
         Left            =   4440
         Picture         =   "frmabm.frx":14E0
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton btn_verdeu 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ver Deuda"
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
         Left            =   2760
         Picture         =   "frmabm.frx":1A6A
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   3240
         MouseIcon       =   "frmabm.frx":1FF4
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":257E
         ToolTipText     =   "Contiene notas de información (Haga doble click para ver)"
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   3240
         MouseIcon       =   "frmabm.frx":2E48
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":33D2
         ToolTipText     =   "Agregar notas de medicación (Economato)"
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Próxima emisión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
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
         TabIndex        =   88
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label labuap 
         BackColor       =   &H000000FF&
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
         Left            =   5160
         TabIndex        =   87
         Top             =   240
         Width           =   735
      End
      Begin VB.Label labump 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
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
         Left            =   4680
         TabIndex        =   86
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo mes pago:"
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
         Height          =   255
         Left            =   3000
         TabIndex        =   85
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label labuega 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
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
         Left            =   1200
         TabIndex        =   80
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label labuegm 
         BackColor       =   &H000000FF&
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
         Left            =   600
         TabIndex        =   79
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultima emisión generada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label labdeudap 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         TabIndex        =   77
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0FFC0&
         Caption         =   "$."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   480
         Width           =   735
      End
      Begin VB.Label labatra 
         BackColor       =   &H000000FF&
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
         Left            =   1800
         TabIndex        =   75
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ATRASO (Meses)"
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
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox t_queconv 
      Height          =   285
      Left            =   7680
      TabIndex        =   72
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9840
      TabIndex        =   70
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Data data_ctrlus 
      Caption         =   "data_ctrlus"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Búsqueda rápida por matrícula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   66
      Top             =   6600
      Width           =   5535
      Begin VB.TextBox txt_buscli 
         Alignment       =   1  'Right Justify
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
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4920
         Picture         =   "frmabm.frx":3C9C
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Matrícula:"
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
         TabIndex        =   67
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Data data_cobrador 
      Caption         =   "data_cobrador"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from cobrador order by cb_nombre"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_promo 
      Caption         =   "data_promo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from vende_func order by nombre"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_zonas 
      Caption         =   "data_zonas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from zonas order by zo_nombre"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   585
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   120
      TabIndex        =   56
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton b_autoafil 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6000
         Picture         =   "frmabm.frx":40DE
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Ver afiliaciones pendientes para autorizar"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton b_afil 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5160
         Picture         =   "frmabm.frx":4668
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Consultar afiliaciones pendientes"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   11040
         Picture         =   "frmabm.frx":4BF2
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Controlar ACTOS de ENFERMERIA"
         Top             =   240
         Width           =   495
      End
      Begin VB.Data data_abm 
         Caption         =   "data_abm"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton btn_busca 
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
         MouseIcon       =   "frmabm.frx":517C
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":5486
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Buscar datos"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton btn_cance 
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
         Left            =   3480
         MouseIcon       =   "frmabm.frx":5A10
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":5D1A
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cancelar acción"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton btn_baja 
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
         Left            =   2640
         MouseIcon       =   "frmabm.frx":62A4
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":65AE
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Baja de socio"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton btn_graba 
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
         Left            =   1800
         MouseIcon       =   "frmabm.frx":6B38
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":6E42
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Grabar datos"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton btn_modi 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         MouseIcon       =   "frmabm.frx":73CC
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":76D6
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Modificar registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton btn_alta 
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
         MouseIcon       =   "frmabm.frx":7C60
         MousePointer    =   99  'Custom
         Picture         =   "frmabm.frx":7F6A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Nuevo registro"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos Administrativos"
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
      ForeColor       =   &H000000FF&
      Height          =   5055
      Left            =   5760
      TabIndex        =   21
      Top             =   2280
      Width           =   6015
      Begin VB.ComboBox cbopromos 
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
         ItemData        =   "frmabm.frx":84F4
         Left            =   1560
         List            =   "frmabm.frx":84F6
         TabIndex        =   118
         Text            =   "cbopromos"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox t_ruta 
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
         Height          =   285
         Left            =   720
         TabIndex        =   107
         Top             =   2760
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc data_clicnv 
         Height          =   450
         Left            =   3720
         Top             =   -120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   794
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
         Caption         =   "data_clicnv"
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
      Begin VB.TextBox txt_dircob 
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
         Left            =   1800
         TabIndex        =   104
         Top             =   3120
         Width           =   4095
      End
      Begin VB.ComboBox cbosrv 
         BackColor       =   &H000000FF&
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
         Height          =   360
         ItemData        =   "frmabm.frx":84F8
         Left            =   4800
         List            =   "frmabm.frx":8502
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Data data_tarjetas 
         Caption         =   "data_tarjetas"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSMask.MaskEdBox txt_vence 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   55
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
      Begin VB.TextBox txt_nrotarj 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
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
         Left            =   4080
         TabIndex        =   53
         Top             =   4320
         Width           =   1815
      End
      Begin MSDBCtls.DBCombo cbotarj 
         Height          =   360
         Left            =   1200
         TabIndex        =   51
         Top             =   4320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txt_codemisor 
         Height          =   375
         Left            =   840
         TabIndex        =   50
         Top             =   4320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "99"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_codtarj 
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
         MaxLength       =   1
         TabIndex        =   48
         Top             =   3960
         Width           =   255
      End
      Begin VB.TextBox txt_cedtarj 
         Alignment       =   1  'Right Justify
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
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   47
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txt_nomtarj 
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
         Left            =   1320
         TabIndex        =   45
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox txt_diacob 
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
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   43
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox cbopago 
         BackColor       =   &H000000FF&
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
         Height          =   360
         ItemData        =   "frmabm.frx":850E
         Left            =   3480
         List            =   "frmabm.frx":8518
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2760
         Width           =   2295
      End
      Begin MSDBCtls.DBCombo cbonomcob 
         Bindings        =   "frmabm.frx":853E
         Height          =   360
         Left            =   2520
         TabIndex        =   39
         Top             =   1920
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "CB_NOMBRE"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txt_codcob 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   38
         Top             =   1920
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo cbonompro 
         Bindings        =   "frmabm.frx":855A
         Height          =   360
         Left            =   2520
         TabIndex        =   36
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "nombre"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txt_codpro 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   35
         Top             =   1440
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_fecbaj 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txt_fecing 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   1080
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
      Begin VB.Label labidpromo 
         Height          =   255
         Left            =   240
         TabIndex        =   119
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label37 
         BackColor       =   &H00000000&
         Caption         =   "Promoción:"
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
         TabIndex        =   114
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label28 
         BackColor       =   &H00000000&
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   106
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFC0C0&
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
         Left            =   4440
         TabIndex        =   108
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Dirección cobro:"
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
         TabIndex        =   103
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label34 
         BackColor       =   &H000000FF&
         Caption         =   "Servicio limitado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   101
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label labavis 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   120
         MouseIcon       =   "frmabm.frx":8573
         TabIndex        =   95
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label33 
         BackColor       =   &H00000000&
         Caption         =   "Vencimiento:"
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
         Left            =   3360
         TabIndex        =   54
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Caption         =   "Nro.Tarj:"
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
         Left            =   3120
         TabIndex        =   52
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label31 
         BackColor       =   &H00000000&
         Caption         =   "Emisor:"
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
         TabIndex        =   49
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label30 
         BackColor       =   &H00000000&
         Caption         =   "CI:Tit"
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
         Left            =   4080
         TabIndex        =   46
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label29 
         BackColor       =   &H00000000&
         Caption         =   "Titular:"
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
         TabIndex        =   44
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackColor       =   &H00000000&
         Caption         =   "Día de cobro:"
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
         TabIndex        =   42
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         Caption         =   "Forma pago:"
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
         Left            =   2160
         TabIndex        =   40
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackColor       =   &H00000000&
         Caption         =   "Cobrador"
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
         TabIndex        =   37
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "Promotor:"
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
         TabIndex        =   33
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Baja:"
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
         Left            =   3000
         TabIndex        =   31
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "Fecha Ingreso"
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
         TabIndex        =   29
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos del cliente"
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
      ForeColor       =   &H00008000&
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5535
      Begin VB.ComboBox cbotipoced 
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
         ItemData        =   "frmabm.frx":8AFD
         Left            =   120
         List            =   "frmabm.frx":8B0D
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   113
         Top             =   5400
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Rechazada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   112
         Top             =   5640
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Se niega"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   111
         ToolTipText     =   "Doble click para desmarcar"
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox t_rs 
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
         Height          =   285
         Left            =   1560
         TabIndex        =   110
         Top             =   1440
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mfcarta 
         Height          =   255
         Left            =   4440
         TabIndex        =   100
         Top             =   6000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Se recibe carta"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   99
         ToolTipText     =   "Doble click para desmarcar"
         Top             =   6000
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Aviso firmar carta"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   98
         ToolTipText     =   "Doble click para desmarcar"
         Top             =   6000
         Width           =   1575
      End
      Begin VB.TextBox t_otrocnv 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   97
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Data data_fecped 
         Caption         =   "data_fecped"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox t_correo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   120
         TabIndex        =   91
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox t_cel 
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
         MaxLength       =   12
         TabIndex        =   89
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txt_conmut 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   69
         Top             =   5280
         Width           =   2295
      End
      Begin VB.Data data_mutual 
         Caption         =   "data_mutual"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSMask.MaskEdBox txt_nac 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   65
         Top             =   1800
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
      Begin MSMask.MaskEdBox txt_codzon 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   34
         Top             =   3240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_matmut 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   28
         Top             =   4680
         Width           =   1575
      End
      Begin MSDBCtls.DBCombo cbomutual 
         Bindings        =   "frmabm.frx":8B33
         Height          =   330
         Left            =   600
         TabIndex        =   26
         Top             =   4680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "ca_nom"
         BoundColumn     =   "ca_nom"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cbosexo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmabm.frx":8B4D
         Left            =   3720
         List            =   "frmabm.frx":8B57
         TabIndex        =   24
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txt_telef 
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   23
         Top             =   3960
         Width           =   1695
      End
      Begin MSDBCtls.DBCombo cbolocalid 
         Bindings        =   "frmabm.frx":8B70
         Height          =   360
         Left            =   2040
         TabIndex        =   20
         Top             =   3240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         ListField       =   "ZO_NOMBRE"
         BoundColumn     =   "ZO_NOMBRE"
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
         MouseIcon       =   "frmabm.frx":8B89
      End
      Begin VB.TextBox txt_direcc2 
         Appearance      =   0  'Flat
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
         Left            =   120
         MaxLength       =   80
         TabIndex        =   18
         Top             =   2880
         Width           =   5295
      End
      Begin VB.TextBox txt_direcc1 
         Appearance      =   0  'Flat
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
         Left            =   120
         MaxLength       =   60
         TabIndex        =   17
         Top             =   2640
         Width           =   5295
      End
      Begin VB.TextBox txt_ced2 
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
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txt_ced 
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
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txt_apellid 
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
         Left            =   120
         MaxLength       =   60
         TabIndex        =   9
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txt_nomcnv 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   525
         Width           =   5295
      End
      Begin VB.TextBox txt_codcnv 
         BackColor       =   &H000000FF&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Tipo Documento:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label43 
         BackColor       =   &H00000000&
         Caption         =   "Razón social:"
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
         TabIndex        =   109
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label labmr 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   5040
         Width           =   5295
      End
      Begin VB.Label Label41 
         BackColor       =   &H00000000&
         Caption         =   "ACTIVO en otro Conv:"
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
         TabIndex        =   96
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Años/Meses/Días"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   93
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label labdias 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
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
         TabIndex        =   92
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Correo Electrónico:"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Conv. Mutual:"
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
         TabIndex        =   68
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label labunie 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
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
         Left            =   3840
         TabIndex        =   62
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "Mat.Mut:"
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
         Left            =   3000
         TabIndex        =   27
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "Mut:"
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
         TabIndex        =   25
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Teléfonos.......................Celular........................Sexo"
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
         TabIndex        =   22
         Top             =   3720
         Width           =   5295
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Localidad"
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
         TabIndex        =   19
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Dirección"
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
         TabIndex        =   16
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label labedad 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
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
         TabIndex        =   15
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Edad"
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
         Left            =   2640
         TabIndex        =   14
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Fec.Nac."
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
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Documento"
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
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Apellidos/Nombres:"
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
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Convenio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Doble click aquí para ver detalle del convenio"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label txt_mat 
      BackColor       =   &H000000FF&
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
      TabIndex        =   64
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label labestado 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      Left            =   4200
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Matrícula:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1065
      Left            =   0
      Picture         =   "frmabm.frx":8FDB
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   4725
   End
End
Attribute VB_Name = "frmabm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Option Explicit
'Funcion Api que obtiene información sobre el estado de Red
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

'Constantes para obtener la información
Private Const INTERNET_CONNECTION_MODEM As Long = &H1
Private Const INTERNET_CONNECTION_LAN As Long = &H2
Private Const INTERNET_CONNECTION_PROXY As Long = &H4
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Private dwflags As Long

      Dim MyForm As FRMSIZE
      Dim DesignX As Integer
      Dim DesignY As Integer

Private Sub b_afil_Click()

'If ControlUsuario(b_afil.Name) = 1 Then
   frm_afilpend.Show vbModal
'End If


End Sub

Private Sub b_autoafil_Click()
If ControlUsuario(b_autoafil.Name) = 1 Then
   frm_afilauto.Show vbModal
End If

End Sub

Private Sub b_histadm_Click()
'If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Then
   Xqueregi = 0
   frm_accadm.Show vbModal
'Else
'   MsgBox "Usuario no autorizado"
'End If
End Sub


Private Sub btn_alta_Click()
Borrar
txt_mat.Caption = data_parsec.Recordset("ultimo_soc") + 1
btn_cance.Enabled = True
btn_modi.Enabled = False
btn_baja.Enabled = False
btn_busca.Enabled = False
btn_fact.Enabled = False
btn_estadi.Enabled = False
btn_histo.Enabled = False
btn_verdeu.Enabled = False
btn_graba.Enabled = True
btn_alta.Enabled = False
Command2.Enabled = False
Frame4.Enabled = False
labestado.Caption = "ACTIVO"
Frame1.Enabled = True
Frame2.Enabled = True
txt_fecing.Text = Format(Date, "dd/mm/yyyy")
txt_codcnv.Enabled = True
txt_codcnv.SetFocus
XAlta = 1
data_parsec.Recordset.Edit
data_parsec.Recordset("ultimo_soc") = txt_mat.Caption
data_parsec.Recordset.Update

If WElusuario = "NELIDA" Or WElusuario = "JFERNAN" Or WElusuario = "MCOSTA" Or XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "ADMINISTRADOR" Or WElusuario = "JONATHAN" Or WElusuario = "MSANCHEZ" Then
   Frame2.Enabled = True
   txt_dircob.Enabled = True
Else
   Frame2.Enabled = False
   txt_dircob.Enabled = False
End If
cbotipoced.ListIndex = 0
Image3.Enabled = False
Image4.Enabled = False

End Sub

Private Sub btn_baja_Click()
Dim Fecbaja As Date

If cbopromos.Text = "Grupo de 3 o más" Then
   MsgBox "Socio con promoción, elimine primero la promoción y luego registre la baja.", vbCritical
Else
    If data_clientes.Recordset("estado") = 2 Or data_clientes.Recordset("estado") = 3 Then
       If IsNull(data_clientes.Recordset("fecha_baja")) = False Then
          MsgBox "Este socio ya está de Baja", vbCritical, "Mensaje"
       Else
          frm_deudasarq.Show vbModal
          Fecbaja = Date
          frm_baja.Show vbModal
       End If
    Else
       frm_deudasarq.Show vbModal
       Fecbaja = Date
       frm_baja.Show vbModal
    End If
End If

End Sub

Private Sub btn_busca_Click()

frm_busca.Show vbModal

End Sub

Private Sub btn_cance_Click()
On Error GoTo Alcance

If XAlta = 1 Then
   data_clientes.RecordSource = "Select * from clientes where cl_codigo =" & 25049
   data_clientes.Refresh
Else
   data_clientes.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Caption
   data_clientes.Refresh
End If

Borrar
XAlta = 0

If XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "ADMINISTRADOR" Then
   btn_modi.Enabled = True
   btn_baja.Enabled = True
Else
   If XWeltipoU = "USUARIOS" Then
      btn_modi.Enabled = True
      btn_baja.Enabled = False
   Else
      btn_modi.Enabled = False
      btn_baja.Enabled = False
   End If
End If
btn_cance.Enabled = False
btn_graba.Enabled = False
btn_busca.Enabled = True
btn_fact.Enabled = True
btn_estadi.Enabled = True
btn_histo.Enabled = True
btn_verdeu.Enabled = True
btn_alta.Enabled = True
txt_codcnv.Enabled = True
txt_nomcnv.Enabled = True
Command2.Enabled = True
quienes
Frame1.Enabled = False
Frame2.Enabled = False
txt_buscli.Text = ""
Frame4.Enabled = True
Image3.Enabled = True
Image4.Enabled = True
'data_clientes.Recordset.MoveLast
If data_clientes.Recordset.RecordCount > 0 Then
    If IsNull(data_clientes.Recordset("cl_fultvta")) = False Then
       If IsNull(data_clientes.Recordset("cl_tipocli")) = False Then
          Image1.Visible = True
       Else
          Image1.Visible = False
       End If
    Else
       Image1.Visible = False
    End If
    If Image1.Visible = False Then
       If IsNull(data_clientes.Recordset("cl_fultpag")) = False Then
          Image1.Visible = True
       End If
    End If
    If data_clientes.Recordset("estado") <> "" Then
       If data_clientes.Recordset("estado") = 2 Or data_clientes.Recordset("estado") = 3 Then
          labestado.Caption = "BAJA"
       Else
          labestado.Caption = "ACTIVO"
       End If
    Else
       If data_clientes.Recordset("fecha_baja") <> "" Then
          labestado.Caption = "BAJA"
       Else
          labestado.Caption = "ACTIVO"
       End If
    End If
    If data_clientes.Recordset("cl_codigo") <> "" Then
       txt_mat.Caption = data_clientes.Recordset("cl_codigo")
    Else
       txt_mat.Caption = ""
    End If
    If IsNull(data_clientes.Recordset("cl_codconv")) = True Then
       MsgBox "Verifique el convenio", vbCritical, "Mensaje"
       txt_codcnv.Text = ""
    Else
       txt_codcnv.Text = data_clientes.Recordset("cl_codconv")
    End If
    txt_nomcnv.Enabled = True
    If IsNull(data_clientes.Recordset("cl_nomconv")) = True Then
       txt_nomcnv.Text = ""
    Else
       txt_nomcnv.Text = data_clientes.Recordset("cl_nomconv")
    End If
    txt_nomcnv.Enabled = False
    If IsNull(data_clientes.Recordset("cl_apellid")) = False Then
       txt_apellid.Text = data_clientes.Recordset("cl_apellid")
    Else
       txt_apellid.Text = ""
    End If
    If IsNull(data_clientes.Recordset("cl_tipoced")) = False Then
       cbotipoced.ListIndex = data_clientes.Recordset("cl_tipoced")
       If data_clientes.Recordset("cl_tipoced") = 0 Then
          txt_ced2.Visible = True
          If data_clientes.Recordset("cl_cedula") <> "" Then
             txt_ced.Text = data_clientes.Recordset("cl_cedula")
          Else
             txt_ced.Text = ""
          End If
          If data_clientes.Recordset("cl_codced") <> "" Then
             txt_ced2.Text = data_clientes.Recordset("cl_codced")
          Else
             txt_ced2.Text = 0
          End If
       Else
          txt_ced2.Visible = False
          If data_clientes.Recordset("cl_cedula") <> "" Then
             txt_ced.Text = data_clientes.Recordset("cl_cedula")
          Else
             txt_ced.Text = ""
          End If
          txt_ced2.Text = 0
       End If
    Else
       cbotipoced.ListIndex = 0
       txt_ced2.Visible = True
       If data_clientes.Recordset("cl_cedula") <> "" Then
          txt_ced.Text = data_clientes.Recordset("cl_cedula")
       Else
          txt_ced.Text = ""
       End If
       If data_clientes.Recordset("cl_codced") <> "" Then
          txt_ced2.Text = data_clientes.Recordset("cl_codced")
       Else
          txt_ced2.Text = 0
       End If
    End If
    If IsNull(data_clientes.Recordset("cl_fnac")) = False Then
       txt_nac.Text = Format(data_clientes.Recordset("cl_fnac"), "dd/mm/yyyy")
       If Not IsDate(txt_nac.Text) Then
       Else
          CalculaEdad (txt_nac.Text)
       End If
    Else
       txt_nac.Text = "__/__/____"
       labedad.Caption = ""
       labunie.Caption = ""
       labdias.Caption = ""
    End If
    
    If data_clientes.Recordset("cl_ultmesp") <> "" Then
       labump.Caption = data_clientes.Recordset("cl_ultmesp")
    Else
       labump.Caption = ""
    End If
    If IsNull(data_clientes.Recordset("mesproxemi")) = False Then
       t_pmemi.Text = data_clientes.Recordset("mesproxemi")
       t_paemi.Text = data_clientes.Recordset("anoproxemi")
    Else
       t_pmemi.Text = 0
       t_paemi.Text = 0
    End If
    
    If IsNull(data_clientes.Recordset("cl_dpto")) = False Then
       t_cel.Text = data_clientes.Recordset("cl_dpto")
    Else
       t_cel.Text = ""
    End If
    If IsNull(data_clientes.Recordset("cl_referen")) = False Then
       t_correo.Text = data_clientes.Recordset("cl_referen")
    Else
       t_correo.Text = ""
    End If
    If data_clientes.Recordset("cl_ultanop") <> "" Then
       If data_clientes.Recordset("cl_ultanop") = 0 Then
          labuap.Caption = data_clientes.Recordset("cl_ultanop")
          Label7.Caption = ""
       Else
          labuap.Caption = data_clientes.Recordset("cl_ultanop")
          Label7.Caption = "/"
       End If
    Else
       labuap.Caption = ""
       Label7.Caption = ""
    End If
    If data_clientes.Recordset("cl_atrasoa") <> "" Then
       labatra.Caption = data_clientes.Recordset("cl_atrasoa")
    Else
       labatra.Caption = ""
    End If
    If data_clientes.Recordset("saldo_cc") <> "" Then
       labdeudap.Caption = data_clientes.Recordset("saldo_cc")
    Else
       labdeudap.Caption = ""
    End If
    If data_clientes.Recordset("cl_direcci") <> "" Then
       txt_direcc1.Text = data_clientes.Recordset("cl_direcci")
    Else
       txt_direcc1.Text = ""
    End If
    If data_clientes.Recordset("cl_entre") <> "" Then
       txt_direcc2.Text = data_clientes.Recordset("cl_entre")
    Else
       txt_direcc2.Text = ""
    End If
    If data_clientes.Recordset("cl_grupo") <> "" Then
       txt_codzon.Text = data_clientes.Recordset("cl_grupo")
    Else
       txt_codzon.Text = 0
    End If
    If data_clientes.Recordset("cl_zona") <> "" Then
       cbolocalid.Text = data_clientes.Recordset("cl_zona")
    Else
       cbolocalid.Text = ""
    End If
    If IsNull(data_clientes.Recordset("cl_codruta")) = False Then
       t_ruta.Text = data_clientes.Recordset("cl_codruta")
    Else
       t_ruta.Text = ""
    End If
    If data_clientes.Recordset("cl_sexo") = 2 Then
       cbosexo.Text = "FEMENINO"
    Else
       cbosexo.Text = "MASCULINO"
    End If
    If data_clientes.Recordset("cl_telefon") <> "" Then
       txt_telef.Text = data_clientes.Recordset("cl_telefon")
    Else
       txt_telef.Text = ""
    End If
    If data_clientes.Recordset("cl_dircobr") <> "" Then
       txt_dircob.Text = data_clientes.Recordset("cl_dircobr")
    Else
       txt_dircob.Text = ""
    End If
    If IsNull(data_clientes.Recordset("cl_nombre")) = False Then
       txt_conmut.Text = data_clientes.Recordset("cl_nombre")
    End If
    If data_clientes.Recordset("cl_socmnom") <> "" Then
       cbomutual.Text = data_clientes.Recordset("cl_socmnom")
    Else
       cbomutual.Text = ""
    End If
    If data_clientes.Recordset("cl_nrosocm") <> "" Then
       txt_matmut.Text = data_clientes.Recordset("cl_nrosocm")
    Else
       txt_matmut.Text = ""
    End If
    
    If data_clientes.Recordset("cl_fecing") <> "" Then
       txt_fecing.Text = Format(data_clientes.Recordset("cl_fecing"), "dd/mm/yyyy")
    Else
       txt_fecing.Text = "__/__/____"
    End If
    If data_clientes.Recordset("fecha_baja") <> "" Then
       txt_fecbaj.Text = Format(data_clientes.Recordset("fecha_baja"), "dd/mm/yyyy")
    Else
       txt_fecbaj.Text = "__/__/____"
    End If
    If data_clientes.Recordset("cl_nrovend") <> "" Then
       txt_codpro.Text = data_clientes.Recordset("cl_nrovend")
    Else
       txt_codpro.Text = ""
    End If
    If IsNull(data_clientes.Recordset("cl_ruc")) = False Then
       If data_clientes.Recordset("cl_ruc") <> "" Then
          t_otrocnv.Text = data_clientes.Recordset("cl_ruc")
       Else
          t_otrocnv.Text = ""
       End If
    Else
       t_otrocnv.Text = ""
    End If
    If data_clientes.Recordset("cl_nomvend") <> "" Then
       cbonompro.Text = data_clientes.Recordset("cl_nomvend")
    Else
       cbonompro.Text = ""
    End If
    If data_clientes.Recordset("cl_nrocobr") <> "" Then
       txt_codcob.Text = data_clientes.Recordset("cl_nrocobr")
    Else
       txt_codcob.Text = ""
    End If
    If data_clientes.Recordset("cl_nomcobr") <> "" Then
       cbonomcob.Text = data_clientes.Recordset("cl_nomcobr")
    Else
       cbonomcob.Text = ""
    End If
    If IsNull(data_clientes.Recordset("cl_descpag")) = True Then
       cbopago.Text = "Abono Mensual"
    Else
       If UCase(data_clientes.Recordset("cl_descpag")) = "DEBITO AUTOMATICO" Then
          cbopago.Text = "Debito Automatico"
       Else
          cbopago.Text = "Abono Mensual"
       End If
    End If
    If data_clientes.Recordset("cl_diacobr") <> "" Then
       txt_diacob.Text = data_clientes.Recordset("cl_diacobr")
    Else
       txt_diacob.Text = ""
    End If
    If data_clientes.Recordset("tit_tarj") <> "" Then
       txt_nomtarj.Text = data_clientes.Recordset("tit_tarj")
    Else
       txt_nomtarj.Text = ""
    End If
    If data_clientes.Recordset("cl_nrotarj") <> "" Then
       txt_nrotarj.Text = data_clientes.Recordset("cl_nrotarj")
    Else
       txt_nrotarj.Text = ""
    End If
    If data_clientes.Recordset("ci_tarj") <> "" Then
       txt_cedtarj.Text = data_clientes.Recordset("ci_tarj")
    Else
       txt_cedtarj.Text = ""
    End If
    If data_clientes.Recordset("codcitarj") <> "" Then
       txt_codtarj.Text = data_clientes.Recordset("codcitarj")
    Else
       txt_codtarj.Text = ""
    End If
    If data_clientes.Recordset("cl_tjemi_c") <> "" Then
       txt_codemisor.Text = data_clientes.Recordset("cl_tjemi_c")
    Else
       txt_codemisor.Text = ""
    End If
    If data_clientes.Recordset("cl_tjemi_n") <> "" Then
       cbotarj.Text = data_clientes.Recordset("cl_tjemi_n")
    Else
       cbotarj.Text = ""
    End If
    If data_clientes.Recordset("cl_tj_venc") <> "" Then
       txt_vence.Text = Format(data_clientes.Recordset("cl_tj_venc"), "dd/mm/yyyy")
    Else
       txt_vence.Text = "__/__/____"
    End If
    If IsNull(data_clientes.Recordset("cl_decuota")) = False Then
       If data_clientes.Recordset("cl_decuota") = 1 Then
          Option1.Value = True
       Else
          If data_clientes.Recordset("cl_decuota") = 2 Then
             Option2.Value = True
          Else
             If data_clientes.Recordset("cl_decuota") = 3 Then
                Option3.Value = True
             Else
                If data_clientes.Recordset("cl_decuota") = 4 Then
                   Option4.Value = True
                Else
                   Option1.Value = False
                   Option2.Value = False
                   Option3.Value = False
                   Option4.Value = False
                End If
             End If
          End If
       End If
    Else
       Option1.Value = False
       Option2.Value = False
       Option3.Value = False
       Option4.Value = False
    End If
    If IsNull(data_clientes.Recordset("fecha_reac")) = False Then
       mfcarta.Text = Format(data_clientes.Recordset("fecha_reac"), "dd/mm/yyyy")
    Else
       mfcarta.Text = "__/__/____"
    End If
    If IsNull(data_clientes.Recordset("saldo_chc2")) = False Then
       cbosrv.ListIndex = data_clientes.Recordset("saldo_chc2")
    Else
       cbosrv.ListIndex = -1
    End If
       
End If

Exit Sub

Alcance:
        If Err.Number = 444 Then
           MsgBox "Error, al cancelar"
           Unload Me
        Else
           MsgBox "Error al cancelar"
           Unload Me
        End If
        
End Sub

Private Sub btn_estadi_Click()
frm_estad.Show vbModal

End Sub

Private Sub btn_fact_Click()
Dim Nomar As String
Dim Reparar As String
'VALIDAR SOCIO SI CORRESPONDE
'Dim socio As String
'Dim validacionObj As Object
'Dim url_socios_en_validacion As String
'Dim responseEstado As Integer
'Dim responseText As String
'socio = txt_mat.Caption
Dim Xcodzoning As Integer
Dim Xlabase As Integer
Dim DeseaFacturar, textocorreo As String
Dim Xfechacartas As Date
Dim ValidadatosJ, XX As Integer
ValidadatosJ = 0
textocorreo = ""
Xfechacartas = Date - 150

''''''''''On Error GoTo Nosepuedefact

Dim Xrecconve As New ADODB.Recordset
Dim Xsqlstr As String
Dim VerificarCgal As Integer
VerificarCgal = 0
If Trim(txt_mat.Caption) <> "" Then
'   VerificarCgal = Verifica_Cgalicia()
'   If txt_codcnv.Text = "CASANR" Then
''      frm_autorizaesp.Show vbModal
'      MsgBox "Convenio no autorizado. No se puede facturar. Consulte con Padrón social.", vbCritical
'      btn_fact.Enabled = True
   'Else
       Verifica_datosJ_Fact
       If DatosVerificadosOk = 0 Then
          If Trim(txt_telef.Text) <> "" Then
             If IsNumeric(txt_telef.Text) = True Then
                If Len(txt_telef.Text) < 7 Then
                   ValidadatosJ = 1
                End If
             Else
                If txt_telef.Text <> "NO APLICA" Then
                   ValidadatosJ = 1
                End If
             End If
          Else
             ValidadatosJ = 1
          End If
          If Trim(t_cel.Text) <> "" Then
             If IsNumeric(t_cel.Text) = True Then
                If Len(t_cel.Text) < 7 Then
                   ValidadatosJ = 1
                End If
             Else
                If t_cel.Text <> "NO APLICA" Then
                   ValidadatosJ = 1
                End If
             End If
          Else
             ValidadatosJ = 1
          End If
          If Trim(t_correo.Text) <> "" Then
             If t_correo.Text <> "NO APLICA" Then
                For XX = 1 To Len(t_correo.Text)
                    If Mid(t_correo.Text, XX, 1) = "@" Then
                       textocorreo = "@"
                    Else
                       If Mid(t_correo.Text, XX, 1) = "." Then
                          If textocorreo = "@" Then
                             textocorreo = textocorreo + "."
                          End If
                       End If
                    End If
                Next
                If textocorreo = "@." Then
                Else
                   ValidadatosJ = 1
                End If
             End If
          Else
             ValidadatosJ = 1
          End If
          If Trim(cbomutual.Text) = "" Then
             ValidadatosJ = 1
          End If
       Else
          ValidadatosJ = 1
       End If
        If ValidadatosJ = 1 Then
           MsgBox "No se ha realizado validación de datos, VERIFIQUE!!", vbCritical
        Else
            data_parsec.DatabaseName = App.path & "\parse.mdb"
            data_parsec.RecordSource = "parsec0"
            data_parsec.Refresh
            Xlabase = data_parsec.Recordset("base")
            
            data_parsec.DatabaseName = App.path & "\mensa.mdb"
            data_parsec.RecordSource = "mensaje"
            data_parsec.Refresh
            If data_parsec.Recordset("base") <> Xlabase Then
               data_parsec.Recordset.Edit
               data_parsec.Recordset("base") = Xlabase
               data_parsec.Recordset.Update
               data_parsec.Refresh
            End If
            
            data_parsec.DatabaseName = App.path & "\parse.mdb"
            data_parsec.RecordSource = "parsec0"
            data_parsec.Refresh
            
            Nomar = App.path & "\factura.ldb"
            If Dir$(Nomar) <> "" Then
                MsgBox "Archivo de facturación abierto en otra CAJA, no se puede facturar!", vbCritical
                Reparar = MsgBox("Desea reparar el archivo de facturación?", vbInformation + vbYesNo)
                If Reparar = vbYes Then
                   data_factrep.DatabaseName = App.path & "\factura.mdb"
                   data_factrep.RecordSource = "lineas"
                   data_factrep.Refresh
                   MsgBox "Archivo reparado, vuelva a ingresar al sistema", vbInformation
                   End
                Else
                   Unload Me
                End If
            Else
                If cbosrv.Text = "SI" Then
                   MsgBox "ATENCION!! Socio con servicios RESTRINGIDOS! Estimado Funcionario NO dar servicio." & Chr(13) _
                   & "El hacerlo estará bajo su exclusiva responsabilidad." & Chr(13) & "El sistema no permitirá la continuidad de dicho servicio.", vbCritical, "SOCIOS"
                Else
                    If txt_codcnv.Text <> "" Then
                       If txt_codzon.Text <> "" Then
                          Xcodzoning = Val(txt_codzon.Text)
                       Else
                          Xcodzoning = 0
                       End If
                       If (Xcodzoning = 400 Or Xcodzoning = 401 Or Xcodzoning = 402 Or Xcodzoning = 403) And _
                          (txt_codcnv.Text = "CCNOS" Or txt_codcnv.Text = "CCNSAM") Then
                           Xestaok = 0
                       Else
                            If txt_codcnv.Text = "SMIN" Or txt_codcnv.Text = "SMINA" Or txt_codcnv.Text = "UNIVS" Or _
                               txt_codcnv.Text = "UNNSAM" Or txt_codcnv.Text = "HEVANO" Or txt_codcnv.Text = "EVNSAM" Or _
                               txt_codcnv.Text = "CCNOS" Or txt_codcnv.Text = "CCNSAM" Or txt_codcnv.Text = "GANOS" Or _
                               txt_codcnv.Text = "CASANO" Or txt_codcnv.Text = "CASNSA" Then
                               ConectarBD
                               ConbdSapp.Open
                               Xsqlstr = "Select * from linmmdd where cod_cli =" & Val(txt_mat.Caption) & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                               With Xrecconve
                                   .CursorLocation = adUseClient
                                   .CursorType = adOpenKeyset
                                   .LockType = adLockOptimistic
                                   .Open Xsqlstr, ConbdSapp, , , adCmdText
                               End With
                               If Xrecconve.RecordCount > 0 Then
                                  Xestaok = 0
                                  ConbdSapp.Close
                                  data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                  data_parsec.RecordSource = "mensaje"
                                  data_parsec.Refresh
                                  data_parsec.Recordset.Edit
                                  data_parsec.Recordset("text") = "RECUERDE! Confirmar socio con la mutualista."
                                  data_parsec.Recordset.Update
                                  XAlta = 17
                                  frm_mensajesvar.Show vbModal
                               Else
                                 ConbdSapp.Close
                                 If Option2.Value = True Then
                                 Else
                                    If Option1.Value = True Or Option3.Value = True Or Option4.Value = True Then
                                       data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                       data_parsec.RecordSource = "mensaje"
                                       data_parsec.Refresh
                                       data_parsec.Recordset.Edit
                                       If txt_codcnv.Text = "SMIN" Or txt_codcnv.Text = "SMINA" Then
                                          If WBase = 6 Or WBase = 17 Or WBase = 11 Or WBase = 16 Or WBase = 8 Or WBase = 10 Or WBase = 13 Or WBase = 12 Then
                                               data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                                               & "Fotocopia de Cédula de identidad vigente. " & _
                                               "RECUERDE! Confirmar socio con la mutualista."
                                          Else
                                               data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                                               & "Fotocopia de CI vigente. Comprobante de domicilio:" _
                                               & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumos," _
                                               & " a nombre del cliente y que sea del mes corriente o anterior." _
                                               & " RECUERDE! Confirmar socio con la mutualista."
                                          End If
                                       Else
                                          data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                                          & "Fotocopia de Cédula de identidad vigente." _
                                          & " RECUERDE! Confirmar socio con la mutualista."
                                       End If
                                       data_parsec.Recordset.Update
                                       data_parsec.Refresh
                                       frm_mensajesvar.Show vbModal
                                       
                                       ConectarBD
                                       ConbdSapp.Open
                                       Xsqlstr = "Select * from linmmdd where cod_cli =" & Val(txt_mat.Caption) & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                                       With Xrecconve
                                           .CursorLocation = adUseClient
                                           .CursorType = adOpenKeyset
                                           .LockType = adLockOptimistic
                                           .Open Xsqlstr, ConbdSapp, , , adCmdText
                                       End With
                                       If Xrecconve.RecordCount > 0 Then
                                          Xestaok = 0
                                       Else
                                          data_parsec.Recordset.Edit
                                          data_parsec.Recordset("text") = "ATENCION!!! Si no realiza carta mutual:" & Chr(13) & " No tendrá derecho a los servicios NO URGENTES."
                                          data_parsec.Recordset.Update
                                          frm_mensajesvar.Show vbModal
                                          Xestaok = 22
                                       End If
                                       ConbdSapp.Close
                                       data_parsec.DatabaseName = App.path & "\parse.mdb"
                                       data_parsec.RecordSource = "parsec0"
                                       data_parsec.Refresh
                                    Else
                                       data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                       data_parsec.RecordSource = "mensaje"
                                       data_parsec.Refresh
                                       data_parsec.Recordset.Edit
                                       If txt_codcnv.Text = "SMIN" Or txt_codcnv.Text = "SMINA" Then
                                          If WBase = 6 Or WBase = 17 Or WBase = 11 Or WBase = 16 Or WBase = 8 Or WBase = 10 Or WBase = 13 Or WBase = 12 Then
                                               data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                                               & "Fotocopia de Cédula de identidad vigente. " _
                                               & " RECUERDE! Confirmar socio con la mutualista."
                                          Else
                                               data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                                               & "Fotocopia CI vigente. Comprobante domicilio, que puede ser:" _
                                               & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumos," _
                                               & " a nombre del cliente y del mes corriente o anterior." _
                                               & " RECUERDE! Confirmar socio con la mutualista."
                                          End If
                                       Else
                                          data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & Chr(13) & "Documentación a presentar:" & Chr(13) _
                                          & "Fotocopia de Cédula de identidad vigente." _
                                          & " RECUERDE! Confirmar socio con la mutualista."
                                       End If
                                       data_parsec.Recordset.Update
                                       data_parsec.Refresh
                                       frm_mensajesvar.Show vbModal
                                       
                                       ConectarBD
                                       ConbdSapp.Open
                                       Xsqlstr = "Select * from linmmdd where cod_cli =" & Val(txt_mat.Caption) & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                                       With Xrecconve
                                           .CursorLocation = adUseClient
                                           .CursorType = adOpenKeyset
                                           .LockType = adLockOptimistic
                                           .Open Xsqlstr, ConbdSapp, , , adCmdText
                                       End With
                                       If Xrecconve.RecordCount > 0 Then
                                          Xestaok = 0
                                       Else
                                          XAlta = 19
                                          data_parsec.Recordset.Edit
                                          data_parsec.Recordset("text") = "ATENCION!!! Si no realiza carta mutual:" & Chr(13) & " No tendrá derecho a los servicios NO URGENTES."
                                          data_parsec.Recordset.Update
                                          frm_mensajesvar.Show vbModal
                                          If Xestaok = 19 Then
                                          Else
                                             Xestaok = 22
                                          End If
                                       End If
                                       ConbdSapp.Close
                                       data_parsec.DatabaseName = App.path & "\parse.mdb"
                                       data_parsec.RecordSource = "parsec0"
                                       data_parsec.Refresh
                                    
                                    End If
                                 End If
                               End If
                            End If
                       End If
                       ConectarBD
                       ConbdSapp.Open
                       Xsqlstr = "Select * from convenio where cnv_codigo ='" & txt_codcnv.Text & "' and cnv_fbaja is null and cnv_umpago not in (1)"
                       With Xrecconve
                           .CursorLocation = adUseClient
                           .CursorType = adOpenKeyset
                           .LockType = adLockOptimistic
                           .Open Xsqlstr, ConbdSapp, , , adCmdText
                       End With
                       data_parsec.DatabaseName = App.path & "\parse.mdb"
                       data_parsec.RecordSource = "parsec0"
                       data_parsec.Refresh
                       
                       If Xrecconve.RecordCount > 0 Then
                          If labestado.Caption = "BAJA" Then
                             If Image1.Visible = True Then
                                frmabm.btn_fact.Enabled = False
                                frmquefac.Show vbModal
                             Else
                                frmabm.btn_fact.Enabled = False
                                frm_facbaja.Show vbModal
                             End If
                          Else
                             frmabm.btn_fact.Enabled = False
                             frmquefac.Show vbModal
                          End If
                          If btn_fact.Enabled = True Then
                             btn_fact.SetFocus
                          Else
                             btn_estadi.SetFocus
                          End If
                       Else
                          MsgBox "El convenio figura de BAJA, no se puede facturar. Verifique con Administración al 097215419", vbCritical
                          txt_buscli.SetFocus
                       End If
                       If ConbdSapp.State = 1 Then
                          ConbdSapp.Close
                       End If
                    Else
                        MsgBox "Error en la ficha del socio al facturar. NO HAY CONVENIO!", vbCritical, "SAPP"
                        txt_buscli.SetFocus
                    End If
                End If
            End If
            
            data_parsec.DatabaseName = App.path & "\parse.mdb"
            data_parsec.RecordSource = "parsec0"
            data_parsec.Refresh
        End If
   'End If
Else
   MsgBox "No hay socio seleccionado"
End If
   
'VALIDAR SOCIO ONERROR
'Exit Sub

'Nosepuedefact:
'             If Err.Number = 3155 Then
'                MsgBox "Error al comenzar la factura, comunique a informática", vbInformation
'             Else
'                MsgBox "Error al iniciar la factura, comunique a informática " & Trim(str(Err.Number)) & " " & Err.Description, vbInformation
'             End If
             
End Sub

Private Sub btn_graba_Click()
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo As Long
Dim Devuelvecedula As Integer
Dim ValidadatosJ, XX As Integer
Dim textocorreo As String
textocorreo = ""
ValidadatosJ = 0

Devuelvecedula = 0

On Error GoTo Nograba

Xn1 = 2
Xn2 = 9
Xn3 = 8
Xn4 = 7
Xn5 = 6
Xn6 = 3
Xn7 = 4
Xpond = 10
If IsNumeric(txt_ced2.Text) = False Then
   txt_ced2.Text = 0
   Xtot = 0
Else
   If cbotipoced.ListIndex <= 0 Then
       Xcedtex = Trim(str(txt_ced.Text))
       Xlargo = Len(Xcedtex)
       If Xlargo = 6 Then
          Xcedtex = "0" & Trim(Xcedtex)
       End If
       Xced1 = Val(Mid(Trim(Xcedtex), 1, 1))
       Xced2 = Val(Mid(Xcedtex, 2, 1))
       Xced3 = Val(Mid(Xcedtex, 3, 1))
       Xced4 = Val(Mid(Xcedtex, 4, 1))
       Xced5 = Val(Mid(Xcedtex, 5, 1))
       Xced6 = Val(Mid(Xcedtex, 6, 1))
    '   If Xlargo = 6 Then
    '      Xced7 = 0
    '   Else
       Xced7 = Val(Mid(Xcedtex, 7, 1))
    '   End If
       Xced1 = Xced1 * Xn1
       Xced2 = Xced2 * Xn2
       Xced3 = Xced3 * Xn3
       Xced4 = Xced4 * Xn4
       Xced5 = Xced5 * Xn5
       Xced6 = Xced6 * Xn6
       Xced7 = Xced7 * Xn7
       Xtot = Xced1 + Xced2 + Xced3 + Xced4 + Xced5 + Xced6 + Xced7
       If Len(Trim(str(Xtot))) = 1 Then
          Xtottex = "0000" & Trim(str(Xtot))
       End If
       If Len(Trim(str(Xtot))) = 2 Then
          Xtottex = "000" & Trim(str(Xtot))
       End If
       If Len(Trim(str(Xtot))) = 3 Then
          Xtottex = "00" & Trim(str(Xtot))
       End If
       If Len(Trim(str(Xtot))) = 4 Then
          Xtottex = "0" & Trim(str(Xtot))
       End If
       Xtot = Val(Mid(Xtottex, 5, 1))
       If Xtot <> 0 Then
          Xtot = Xpond - Xtot
       Else
          Xtot = 0
       End If
   Else
       txt_ced2.Text = 0
       Xtot = 0
   End If
End If
If Trim(txt_telef.Text) = "no aplica" Then
   txt_telef.Text = "NO APLICA"
End If
If Trim(t_cel.Text) = "no aplica" Then
   t_cel.Text = "NO APLICA"
End If
If Trim(t_correo.Text) = "no aplica" Then
   t_correo.Text = "NO APLICA"
End If
If Trim(cbomutual.Text) = "no aplica" Then
   cbomutual.Text = "NO APLICA"
End If

'095434976 095673419
'   If Xtot = txt_ced2.Text Then
     If Xtot = txt_ced2.Text Then
        If txt_codcnv.Text <> "" Then
           If XAlta = 2 Then
              Verifica_datosJ
           Else
              DatosVerificadosOk = 0
           End If
           If DatosVerificadosOk <> 0 Then
              ValidadatosJ = 1
           End If
           If Trim(txt_telef.Text) <> "" Then
              If IsNumeric(txt_telef.Text) = True Then
                 If Len(txt_telef.Text) < 7 Then
                    ValidadatosJ = 1
                 End If
              Else
                 If txt_telef.Text <> "NO APLICA" Then
                    ValidadatosJ = 1
                 End If
              End If
           Else
              ValidadatosJ = 1
           End If
           If Trim(t_cel.Text) <> "" Then
              If IsNumeric(t_cel.Text) = True Then
                 If Len(t_cel.Text) < 7 Then
                    ValidadatosJ = 1
                 End If
              Else
                 If t_cel.Text <> "NO APLICA" Then
                    ValidadatosJ = 1
                 End If
              End If
           Else
              ValidadatosJ = 1
           End If
           If Trim(t_correo.Text) <> "" Then
              If t_correo.Text <> "NO APLICA" Then
                 For XX = 1 To Len(t_correo.Text)
                     If Mid(t_correo.Text, XX, 1) = "@" Then
                        textocorreo = "@"
                     Else
                        If Mid(t_correo.Text, XX, 1) = "." Then
                           If textocorreo = "@" Then
                              textocorreo = textocorreo + "."
                           End If
                        End If
                     End If
                 Next
                 If textocorreo = "@." Then
                 Else
                    ValidadatosJ = 1
                 End If
              End If
           Else
              ValidadatosJ = 1
           End If
           If Trim(cbomutual.Text) = "" Then
              ValidadatosJ = 1
           End If
           If ValidadatosJ = 1 Then
               MsgBox "No se puede grabar, verifique datos de:" & Chr(13) & _
               "Teléfono, Celular, Correo electrónico y Mutualista.", vbCritical
           Else
                If XAlta = 2 Then
                   If Cl_cedulaAnt = txt_ced.Text Then
                      Devuelvecedula = 0
                   Else
                      Devuelvecedula = Devuelve_ceduladoble()
                   End If
                   If Devuelvecedula = 0 Then
                       data_clientes.Recordset.Edit
                       If data_clientes.Recordset("estado") = 1 Or data_clientes.Recordset("estado") = 0 Then
                          data_clientes.Recordset("estado") = 1
                          labestado.Caption = "ACTIVO"
                       Else
                          If data_clientes.Recordset("fecha_baja") <> "" Then
                             labestado.Caption = "BAJA"
                             data_clientes.Recordset("estado") = 2
                          Else
                             labestado.Caption = "ACTIVO"
                             data_clientes.Recordset("estado") = 1
                          End If
                       End If
                       data_clientes.Recordset("cl_codigo") = txt_mat.Caption
                       data_clientes.Recordset("cl_codconv") = txt_codcnv.Text
                       data_clientes.Recordset("cl_nomconv") = Mid(txt_nomcnv.Text, 1, 30)
                       data_clientes.Recordset("cl_apellid") = txt_apellid.Text
                       If t_otrocnv.Text <> "" Then
                          If IsNull(data_clientes.Recordset("cl_ruc")) = False Then
                             If data_clientes.Recordset("cl_ruc") <> t_otrocnv.Text Then
                                data_clientes.Recordset("cl_ruc") = t_otrocnv.Text
                             End If
                          Else
                             data_clientes.Recordset("cl_ruc") = t_otrocnv.Text
                          End If
                       Else
                          If IsNull(data_clientes.Recordset("cl_ruc")) = False Then
                             data_clientes.Recordset("cl_ruc") = Null
                          End If
                       End If
                       data_clientes.Recordset("cl_tipoced") = cbotipoced.ListIndex
                       If txt_ced.Text <> "" Then
                          data_clientes.Recordset("cl_cedula") = txt_ced.Text
                          If txt_ced2.Text <> "" Then
                             data_clientes.Recordset("cl_codced") = txt_ced2.Text
                             data_clientes.Recordset("cl_cedula_t") = Trim(txt_ced.Text) & Trim(txt_ced2.Text)
                          Else
                             txt_ced2.Text = 0
                             data_clientes.Recordset("cl_codced") = 0
                          End If
                       Else
                          txt_ced.Text = 0
                          txt_ced2.Text = 0
                          data_clientes.Recordset("cl_cedula") = txt_ced.Text
                          data_clientes.Recordset("cl_codced") = txt_ced2.Text
                       End If
                       If t_ruta.Text <> "" Then
                          If IsNull(data_clientes.Recordset("cl_codruta")) = False Then
                             If t_ruta.Text <> data_clientes.Recordset("cl_codruta") Then
                                data_clientes.Recordset("cl_codruta") = t_ruta.Text
                             End If
                          Else
                             data_clientes.Recordset("cl_codruta") = t_ruta.Text
                          End If
                       Else
                          If IsNull(data_clientes.Recordset("cl_codruta")) = False Then
                             data_clientes.Recordset("cl_codruta") = Null
                          End If
                       End If
                       If txt_nac.Text <> "__/__/____" Then
                          data_clientes.Recordset("cl_fnac") = Format(txt_nac.Text, "dd/mm/yyyy")
                       End If
                       If labedad.Caption <> "" Then
                          data_clientes.Recordset("cl_edad") = labedad.Caption
                       Else
                          labedad.Caption = 0
                          data_clientes.Recordset("cl_edad") = labedad.Caption
                       End If
                        If labump.Caption <> "" Then
                           data_clientes.Recordset("cl_ultmesp") = labump.Caption
                        Else
                           labump.Caption = 0
                           data_clientes.Recordset("cl_ultmesp") = labump.Caption
                        End If
                        If labuap.Caption <> "" Then
                           data_clientes.Recordset("cl_ultanop") = Val(labuap.Caption)
                        Else
                           labuap.Caption = 0
                           data_clientes.Recordset("cl_ultanop") = labuap.Caption
                        End If
                        If labidpromo.Caption <> "" Then
                           data_clientes.Recordset("idpromos") = Val(labidpromo.Caption)
                        Else
                           data_clientes.Recordset("idpromos") = Null
                        End If
                        If t_pmemi.Text <> "" Then
                           data_clientes.Recordset("mesproxemi") = Val(t_pmemi.Text)
                           data_clientes.Recordset("anoproxemi") = Val(t_paemi.Text)
                        Else
                           data_clientes.Recordset("mesproxemi") = 0
                           data_clientes.Recordset("anoproxemi") = 0
                        End If
                        
                        If labatra.Caption <> "" Then
                           data_clientes.Recordset("cl_atrasoa") = labatra.Caption
                        Else
                           labatra.Caption = 0
                           data_clientes.Recordset("cl_atrasoa") = labatra.Caption
                        End If
                        If labdeudap.Caption <> "" Then
                           data_clientes.Recordset("saldo_cc") = labdeudap.Caption
                        Else
                           labdeudap.Caption = ""
                        End If
                        data_clientes.Recordset("cl_direcci") = txt_direcc1.Text
                        data_clientes.Recordset("cl_entre") = txt_direcc2.Text
                        If t_cel.Text <> "" Then
                           data_clientes.Recordset("cl_dpto") = t_cel.Text
                           data_clientes.Recordset("cl_celular_n") = Trim(t_cel.Text)
                        Else
                           data_clientes.Recordset("cl_dpto") = Null
                        End If
                        If t_correo.Text <> "" Then
                           data_clientes.Recordset("cl_referen") = Mid(t_correo.Text, 1, 120)
                        Else
                           data_clientes.Recordset("cl_referen") = Null
                        End If
                        If txt_codzon.Text <> "" Then
                           If cbolocalid.Text <> "" Then
                              data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                              data_clientes.Recordset("cl_zona") = Mid(cbolocalid.Text, 1, 25)
                           Else
                              cbolocalid.Text = "*TODOS"
                              data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                              data_clientes.Recordset("cl_zona") = Mid(cbolocalid.Text, 1, 25)
                           End If
                        Else
                           txt_codzon.Text = 999
                           cbolocalid.Text = "*TODOS"
                           data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                           data_clientes.Recordset("cl_zona") = cbolocalid.Text
                        End If
                        If cbosexo.Text = "FEMENINO" Then
                            data_clientes.Recordset("cl_sexo") = 2
                        Else
                            data_clientes.Recordset("cl_sexo") = 1
                        End If
                        data_clientes.Recordset("cl_telefon") = txt_telef.Text
                        data_clientes.Recordset("cl_dircobr") = txt_dircob.Text
                        data_clientes.Recordset("cl_nombre") = txt_conmut.Text
                        data_clientes.Recordset("cl_socmnom") = Mid(cbomutual.Text, 1, 25)
                        data_clientes.Recordset("cl_nrosocm") = txt_matmut.Text
                        If txt_fecing.Text = "__/__/____" Then
                           txt_fecing.Text = Date
                        End If
                        If txt_fecing.Text <> "__/__/____" Then
                           data_clientes.Recordset("cl_fecing") = Format(txt_fecing.Text, "dd/mm/yyyy")
                        End If
                        If txt_fecbaj.Text <> "__/__/____" Then
                           data_clientes.Recordset("fecha_baja") = Format(txt_fecbaj.Text, "dd/mm/yyyy")
                        End If
                        If txt_codpro.Text <> "" Then
                           If cbonompro.Text <> "" Then
                              data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                              data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                           Else
                              cbonompro.Text = "*TODOS"
                              data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                              data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                           End If
                        Else
                           txt_codpro.Text = 799
                           cbonompro.Text = "*TODOS"
                           data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                           data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                        End If
                        If txt_codcob.Text <> "" Then
                           If cbonomcob.Text <> "" Then
                              data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                              data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                           Else
                              cbonomcob.Text = "*TODOS"
                              data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                              data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                           End If
                        Else
                           txt_codcob.Text = 0
                           cbonomcob.Text = "*TODOS"
                           data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                           data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                        End If
                        If UCase(cbopago.Text) = "DEBITO AUTOMATICO" Then
                           data_clientes.Recordset("cl_forpago") = 2
                           data_clientes.Recordset("cl_descpag") = "Debito Automatico"
                        Else
                           data_clientes.Recordset("cl_forpago") = 1
                           data_clientes.Recordset("cl_descpag") = "Abono Mensual"
                        End If
                        data_clientes.Recordset("cl_diacobr") = txt_diacob.Text
                        data_clientes.Recordset("tit_tarj") = txt_nomtarj.Text
                        If txt_nrotarj.Text <> "" Then
                           data_clientes.Recordset("cl_nrotarj") = txt_nrotarj.Text
                        Else
                           data_clientes.Recordset("cl_nrotarj") = 0
                        End If
                        If txt_cedtarj.Text <> "" Then
                           data_clientes.Recordset("ci_tarj") = txt_cedtarj.Text
                        Else
                           txt_cedtarj.Text = 0
                           data_clientes.Recordset("ci_tarj") = txt_cedtarj.Text
                        End If
                        If txt_codtarj.Text <> "" Then
                           data_clientes.Recordset("codcitarj") = txt_codtarj.Text
                        Else
                           txt_codtarj.Text = 0
                           data_clientes.Recordset("codcitarj") = txt_codtarj.Text
                        End If
                        If txt_codemisor.Text <> "" Then
                           data_clientes.Recordset("cl_tjemi_c") = txt_codemisor.Text
                        Else
                           txt_codemisor.Text = 0
                           data_clientes.Recordset("cl_tjemi_c") = txt_codemisor.Text
                        End If
                        If cbotarj.Text <> "" Then
                           data_clientes.Recordset("cl_tjemi_n") = cbotarj.Text
                        Else
                           cbotarj.Text = ""
                           data_clientes.Recordset("cl_tjemi_n") = cbotarj.Text
                        End If
                        If txt_vence.Text <> "__/__/____" Then
                           data_clientes.Recordset("cl_tj_venc") = Format(txt_vence.Text, "dd/mm/yyyy")
                        End If
                        data_clientes.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
                        If Option1.Value = True Then
                           data_clientes.Recordset("cl_decuota") = 1
                           If mfcarta.Text <> "__/__/____" Then
                              data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                           Else
                              data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                           End If
                        Else
                           If Option2.Value = True Then
                              data_clientes.Recordset("cl_decuota") = 2
                              If mfcarta.Text <> "__/__/____" Then
                                 data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                              Else
                                 data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                              End If
                           Else
                              If Option3.Value = True Then
                                 data_clientes.Recordset("cl_decuota") = 3
                                 If mfcarta.Text <> "__/__/____" Then
                                    data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                                 Else
                                    data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                                 End If
                              Else
                                 If Option4.Value = True Then
                                    data_clientes.Recordset("cl_decuota") = 4
                                    If mfcarta.Text <> "__/__/____" Then
                                       data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                                    Else
                                       data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                                    End If
                                 Else
                                    If Option5.Value = True Then
                                       data_clientes.Recordset("cl_decuota") = 5
                                       If mfcarta.Text <> "__/__/____" Then
                                          data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                                       Else
                                          data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                                       End If
                                    Else
                                       data_clientes.Recordset("cl_decuota") = Null
                                       If IsNull(data_clientes.Recordset("fecha_reac")) = False Then
                                          data_clientes.Recordset("fecha_reac") = Null
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                        data_clientes.Recordset("saldo_chc2") = cbosrv.ListIndex
                        'VALIDA TEL CEL MAIL MUT SOCIO EN MODIFICACION'
                        'armo json y envio datos del cliente a servicio
                        'consume alta cliente (solo para validarlo)
                        'me fijo si hay errores, si hay errores imprimo, si esta todo bien continuanormal
                        'FIN TEL CEL MAIL MUT SOCIO EN MODIFICACION'
                             
                        data_clientes.Recordset.Update
                        Image3.Enabled = True
                        Image4.Enabled = True
                         'registro cambio en historial despues del update (asincrono)
        
                        txt_nomcnv.Enabled = True
                        vercli
                        Actualiza_Magik
                        frm_abmm.Show vbModal
                        XAlta = 0
                        If XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "ADMINISTRADOR" Then
                           btn_modi.Enabled = True
                           btn_baja.Enabled = True
                        Else
                           If XWeltipoU = "USUARIOS" Then
                              btn_modi.Enabled = True
                              btn_baja.Enabled = False
                           Else
                              btn_modi.Enabled = False
                              btn_baja.Enabled = False
                           End If
                        End If
                        btn_cance.Enabled = False
                        btn_busca.Enabled = True
                        btn_fact.Enabled = True
                        btn_estadi.Enabled = True
                        btn_histo.Enabled = True
                        btn_verdeu.Enabled = True
                        btn_graba.Enabled = False
                        btn_alta.Enabled = True
                        Command2.Enabled = True
                        quienes
                        Frame4.Enabled = True
                        txt_codcnv.Enabled = False
                        txt_nomcnv.Enabled = True
                        txt_nomcnv.Enabled = False
                        
                        Frame1.Enabled = False
                        Frame2.Enabled = False
                   Else
                        MsgBox "La cédula ingresada, ya existe en otro paciente, verifique!", vbCritical
                   End If
                    
                End If
                If XAlta = 1 Then
                   If txt_ced.Text <> "" Then
                      data_clientes.RecordSource = "Select * from clientes where cl_cedula =" & txt_ced.Text
                      data_clientes.Refresh
                      If txt_ced.Text > 0 Then
                         'Data1.Recordset.FindFirst "cl_cedula =" & txt_ced.Text
                         Data1.RecordSource = "Select * from clientes where cl_cedula =" & txt_ced.Text
                         Data1.Refresh
                         If Data1.Recordset.RecordCount > 0 Then
                            MsgBox "Ya existe socio con ésta cédula, Verifique o presione el botón de Cancelar", vbInformation, "Clientes"
                            txt_apellid.SetFocus
                         Else
                            data_clientes.Recordset.AddNew
                           data_clientes.Recordset("estado") = 1
                        '   labestado.Caption = "ACTIVO"
                           data_clientes.Recordset("cl_codigo") = txt_mat.Caption
                           data_clientes.Recordset("cl_codconv") = txt_codcnv.Text
                           data_clientes.Recordset("cl_nomconv") = Mid(txt_nomcnv.Text, 1, 30)
                           data_clientes.Recordset("cl_apellid") = txt_apellid.Text
                           If t_otrocnv.Text <> "" Then
                              data_clientes.Recordset("cl_ruc") = t_otrocnv.Text
                           End If
                           data_clientes.Recordset("cl_tipoced") = cbotipoced.ListIndex
                           If txt_ced.Text <> "" Then
                              data_clientes.Recordset("cl_cedula") = txt_ced.Text
                              If txt_ced2.Text <> "" Then
                                 data_clientes.Recordset("cl_codced") = txt_ced2.Text
                                 data_clientes.Recordset("cl_cedula_t") = Trim(txt_ced.Text) & Trim(txt_ced2.Text)
                              Else
                                 txt_ced2.Text = 0
                                 data_clientes.Recordset("cl_codced") = 0
                              End If
                           Else
                              txt_ced.Text = 0
                              txt_ced2.Text = 0
                              data_clientes.Recordset("cl_cedula") = txt_ced.Text
                              data_clientes.Recordset("cl_codced") = txt_ced2.Text
                           End If
                           If txt_nac.Text <> "__/__/____" Then
                              data_clientes.Recordset("cl_fnac") = Format(txt_nac.Text, "dd/mm/yyyy")
                           End If
                           If labedad.Caption <> "" Then
                              data_clientes.Recordset("cl_edad") = labedad.Caption
                           Else
                              labedad.Caption = 0
                              data_clientes.Recordset("cl_edad") = labedad.Caption
                           End If
                           If t_ruta.Text <> "" Then
                              data_clientes.Recordset("cl_codruta") = t_ruta.Text
                           End If
                            If labump.Caption <> "" Then
                               data_clientes.Recordset("cl_ultmesp") = labump.Caption
                            Else
                               labump.Caption = 0
                               data_clientes.Recordset("cl_ultmesp") = labump.Caption
                            End If
                            If labuap.Caption <> "" Then
                               data_clientes.Recordset("cl_ultanop") = labuap.Caption
                            Else
                               labuap.Caption = 0
                               data_clientes.Recordset("cl_ultanop") = labuap.Caption
                            End If
                            If labatra.Caption <> "" Then
                               data_clientes.Recordset("cl_atrasoa") = labatra.Caption
                            Else
                               labatra.Caption = 0
                               data_clientes.Recordset("cl_atrasoa") = labatra.Caption
                            End If
                            If labdeudap.Caption <> "" Then
                               data_clientes.Recordset("saldo_cc") = labdeudap.Caption
                            Else
                               labdeudap.Caption = ""
                            End If
                            If labidpromo.Caption <> "" Then
                               data_clientes.Recordset("idpromos") = Val(labidpromo.Caption)
                            Else
                               data_clientes.Recordset("idpromos") = 0
                            End If
                            If t_pmemi.Text <> "" Then
                               data_clientes.Recordset("mesproxemi") = Val(t_pmemi.Text)
                               data_clientes.Recordset("anoproxemi") = Val(t_paemi.Text)
                            Else
                               data_clientes.Recordset("mesproxemi") = 0
                               data_clientes.Recordset("anoproxemi") = 0
                            End If
                            
                            data_clientes.Recordset("cl_direcci") = txt_direcc1.Text
                            data_clientes.Recordset("cl_entre") = txt_direcc2.Text
                            If t_cel.Text <> "" Then
                               data_clientes.Recordset("cl_dpto") = t_cel.Text
                               data_clientes.Recordset("cl_celular_n") = Trim(t_cel.Text)
                            Else
                               data_clientes.Recordset("cl_dpto") = Null
                            End If
                            If t_correo.Text <> "" Then
                               data_clientes.Recordset("cl_referen") = Mid(t_correo.Text, 1, 120)
                            Else
                               data_clientes.Recordset("cl_referen") = Null
                            End If
                            If txt_codzon.Text <> "" Then
                               If cbolocalid.Text <> "" Then
                                  data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                                  data_clientes.Recordset("cl_zona") = cbolocalid.Text
                               Else
                                  cbolocalid.Text = "*TODOS"
                                  data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                                  data_clientes.Recordset("cl_zona") = cbolocalid.Text
                               End If
                            Else
                               txt_codzon.Text = 999
                               cbolocalid.Text = "*TODOS"
                               data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                               data_clientes.Recordset("cl_zona") = cbolocalid.Text
                            End If
                            If cbosexo.Text = "FEMENINO" Then
                                data_clientes.Recordset("cl_sexo") = 2
                            Else
                                data_clientes.Recordset("cl_sexo") = 1
                            End If
                            data_clientes.Recordset("cl_telefon") = txt_telef.Text
                            data_clientes.Recordset("cl_dircobr") = txt_dircob.Text
                            data_clientes.Recordset("cl_nombre") = txt_conmut.Text
                            data_clientes.Recordset("cl_socmnom") = Mid(cbomutual.Text, 1, 25)
                            data_clientes.Recordset("cl_nrosocm") = txt_matmut.Text
                            If txt_fecing.Text = "__/__/____" Then
                               txt_fecing.Text = Date
                            End If
                            If txt_fecing.Text <> "__/__/____" Then
                               data_clientes.Recordset("cl_fecing") = Format(txt_fecing.Text, "dd/mm/yyyy")
                            End If
                            If txt_fecbaj.Text <> "__/__/____" Then
                               data_clientes.Recordset("fecha_baja") = Format(txt_fecbaj.Text, "dd/mm/yyyy")
                            End If
                            If txt_codpro.Text <> "" Then
                               If cbonompro.Text <> "" Then
                                  data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                                  data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                               Else
                                  cbonompro.Text = "*TODOS"
                                  data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                                  data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                               End If
                            Else
                               txt_codpro.Text = 799
                               cbonompro.Text = "*TODOS"
                               data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                               data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                            End If
                            If txt_codcob.Text <> "" Then
                               If cbonomcob.Text <> "" Then
                                  data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                                  data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                               Else
                                  cbonomcob.Text = "*TODOS"
                                  data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                                  data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                               End If
                            Else
                               txt_codcob.Text = 0
                               cbonomcob.Text = "*TODOS"
                               data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                               data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                            End If
                            If cbopago.Text = "Debito Automatico" Then
                               data_clientes.Recordset("cl_forpago") = 2
                               data_clientes.Recordset("cl_descpag") = "Debito Automatico"
                            Else
                               data_clientes.Recordset("cl_forpago") = 1
                               data_clientes.Recordset("cl_descpag") = "Abono Mensual"
                            End If
                            data_clientes.Recordset("cl_diacobr") = txt_diacob.Text
                            data_clientes.Recordset("tit_tarj") = txt_nomtarj.Text
                            If txt_nrotarj.Text <> "" Then
                               data_clientes.Recordset("cl_nrotarj") = txt_nrotarj.Text
                            Else
                               data_clientes.Recordset("cl_nrotarj") = 0
                            End If
                            If txt_cedtarj.Text <> "" Then
                               data_clientes.Recordset("ci_tarj") = txt_cedtarj.Text
                            Else
                               txt_cedtarj.Text = 0
                               data_clientes.Recordset("ci_tarj") = txt_cedtarj.Text
                            End If
                            If txt_codtarj.Text <> "" Then
                               data_clientes.Recordset("codcitarj") = txt_codtarj.Text
                            Else
                               txt_codtarj.Text = 0
                               data_clientes.Recordset("codcitarj") = txt_codtarj.Text
                            End If
                            If txt_codemisor.Text <> "" Then
                               data_clientes.Recordset("cl_tjemi_c") = txt_codemisor.Text
                            Else
                               txt_codemisor.Text = 0
                               data_clientes.Recordset("cl_tjemi_c") = txt_codemisor.Text
                            End If
                            If cbotarj.Text <> "" Then
                               data_clientes.Recordset("cl_tjemi_n") = cbotarj.Text
                            Else
                               cbotarj.Text = ""
                               data_clientes.Recordset("cl_tjemi_n") = cbotarj.Text
                            End If
                            If txt_vence.Text <> "__/__/____" Then
                               data_clientes.Recordset("cl_tj_venc") = Format(txt_vence.Text, "dd/mm/yyyy")
                            End If
                            data_clientes.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
                            If Option1.Value = True Then
                               data_clientes.Recordset("cl_decuota") = 1
                               If mfcarta.Text <> "__/__/____" Then
                                  data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                               Else
                                  data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                               End If
                            Else
                               If Option2.Value = True Then
                                  data_clientes.Recordset("cl_decuota") = 2
                                  If mfcarta.Text <> "__/__/____" Then
                                     data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                                  Else
                                     data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                                  End If
                               Else
                                  If Option3.Value = True Then
                                     data_clientes.Recordset("cl_decuota") = 3
                                     If mfcarta.Text <> "__/__/____" Then
                                        data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                                     Else
                                        data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                                     End If
                                  Else
                                     If Option4.Value = True Then
                                        data_clientes.Recordset("cl_decuota") = 4
                                        If mfcarta.Text <> "__/__/____" Then
                                           data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                                        Else
                                           data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                                        End If
                                     Else
                                        If Option5.Value = True Then
                                           data_clientes.Recordset("cl_decuota") = 5
                                           If mfcarta.Text <> "__/__/____" Then
                                              data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                                           Else
                                              data_clientes.Recordset("fecha_reac") = Format(Date, "dd/mm/yyyy")
                                           End If
                                        Else
                                           data_clientes.Recordset("cl_decuota") = Null
                                           If IsNull(data_clientes.Recordset("fecha_reac")) = False Then
                                              data_clientes.Recordset("fecha_reac") = Null
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                            data_clientes.Recordset("saldo_chc2") = cbosrv.ListIndex
                            
                            'VALIDA TEL CEL MAIL MUT SOCIO EN ALTA'
                            'armo json y envio datos del cliente a servicio
                            'consume alta cliente (solo para validar)
                            'me fijo si hay errores en validacion, si hay errores imprimo, si esta todo bien continuanormal
                            
                            'FIN VALIDA TEL CEL MAIL MUT SOCIO EN ALTA'
                            
                            data_clientes.Recordset.Update
                            altaValidacionDatos
                            altaValidacionDatosabm
                            Image3.Enabled = True
                            Image4.Enabled = True
                            data_clientes.Refresh
                            txt_nomcnv.Enabled = True
                            vercli
                            Actualiza_Magik
                            txt_nomcnv.Enabled = False
                            data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & txt_mat.Caption
                            data_abm.Refresh
                            data_abm.Recordset.AddNew
                            data_abm.Recordset("cl_codigo") = txt_mat.Caption
                            If XWeltipoU = "USUARIOS ADM" Then
                               data_abm.Recordset("cl_motivo") = "ALTA CON PROMOTOR"
                            Else
                               data_abm.Recordset("cl_motivo") = "ALTA DE FICHA"
                            End If
                            data_abm.Recordset("desc") = "ALTA"
                            data_abm.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                            data_abm.Recordset("hora") = Format(Time, "HH:mm")
                            data_abm.Recordset("usuario") = WElusuario
                            data_abm.Recordset("convenio") = txt_codcnv.Text
                            data_abm.Recordset("base") = data_parsec.Recordset("base")
                            data_abm.Recordset.Update
                            XAlta = 0
                            If XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "ADMINISTRADOR" Then
                               btn_modi.Enabled = True
                               btn_baja.Enabled = True
                            Else
                               If XWeltipoU = "USUARIOS" Then
                                  btn_modi.Enabled = True
                                  btn_baja.Enabled = False
                               Else
                                  btn_modi.Enabled = False
                                  btn_baja.Enabled = False
                               End If
                            End If
                            btn_busca.Enabled = True
                            btn_cance.Enabled = False
                            btn_fact.Enabled = True
                            btn_estadi.Enabled = True
                            btn_histo.Enabled = True
                            btn_verdeu.Enabled = True
                            btn_graba.Enabled = False
                            btn_alta.Enabled = True
                            Command2.Enabled = True
                            Frame4.Enabled = True
                            Frame1.Enabled = False
                            Frame2.Enabled = False
                         End If
                      Else
    '                     MsgBox "LA CEDULA NO PUEDE SER CERO, VERIFIQUE!!", vbCritical
                         Command1_Click
                      End If
                   Else
                      MsgBox "CEDULA EN BLANCO, VERIFIQUE O INGRESE CERO!", vbCritical, "Clientes"
                      txt_apellid.SetFocus
                   End If
                End If
           End If
        Else
            MsgBox "Verifique el Convenio", vbCritical, "Mantenimiento"
            txt_codcnv.SetFocus
        End If
    Else
        MsgBox "Error en dígito verificador de cédula", vbInformation
    End If
Exit Sub


Nograba:
        If Err.Number = 3197 Then
           MsgBox "No se modificaron datos."
        Else
           MsgBox "Error al grabar datos, verifique!", vbInformation
        End If
        
        
End Sub

Private Sub btn_histo_Click()
frm_histo.Show vbModal


End Sub

Private Sub btn_modi_Click()
Dim Xresp As String
Dim XFB As Variant
Dim Xotrodato As String
If XWeltipoU = "USUARIOS" Or XWeltipoU = "USUARIOS FARM" Or XWeltipoU = "ADM FARMACIA" Then
   XAlta = 2
   btn_modi.Enabled = False
   btn_cance.Enabled = True
   btn_baja.Enabled = False
   btn_busca.Enabled = False
   btn_fact.Enabled = False
   btn_estadi.Enabled = False
   btn_histo.Enabled = False
   btn_verdeu.Enabled = False
   btn_graba.Enabled = True
   btn_alta.Enabled = False
   Command2.Enabled = False
   Frame4.Enabled = False
   Frame1.Enabled = True
   Frame2.Enabled = False
   txt_codcnv.Enabled = False
   txt_nomcnv.Enabled = False
   txt_apellid.SetFocus
   Label9.Enabled = False
   Option1.Enabled = False
   Option2.Enabled = False
   Option4.Enabled = False
   Option5.Enabled = False
   
Else
    Label9.Enabled = True
    Option1.Enabled = True
    Option2.Enabled = True
    Option4.Enabled = True
    Option5.Enabled = True
    If txt_codcnv.Text <> "" Then
       ControlproxEmi
    End If
    If IsNull(data_clientes.Recordset("cl_fultvta")) = False Then
       If IsNull(data_clientes.Recordset("cl_tipocli")) = False Then
          Xotrodato = MsgBox("Contiene datos de usuario, desea modificarlos ?", vbCritical + vbYesNo, "SAPP")
          If Xotrodato = vbYes Then
             data_clientes.Recordset.Edit
             data_clientes.Recordset("cl_fultvta") = Null
             data_clientes.Recordset("cl_tipocli") = Null
             data_clientes.Recordset("cl_celular") = Null
             data_clientes.Recordset("cl_fultpag") = Null
             data_clientes.Recordset.Update
             data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & txt_mat.Caption
             data_abm.Refresh
             data_abm.Recordset.AddNew
             data_abm.Recordset("cl_codigo") = txt_mat.Caption
             data_abm.Recordset("cl_motivo") = "BORRA DATOS"
             data_abm.Recordset("desc") = "MODIF"
             data_abm.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
             data_abm.Recordset("hora") = Format(Time, "HH:mm")
             data_abm.Recordset("usuario") = WElusuario
             data_abm.Recordset("convenio") = txt_codcnv.Text
             data_abm.Recordset("base") = data_parsec.Recordset("base")
             data_abm.Recordset.Update
    
             Image1.Visible = False
          End If
       End If
    Else
       If IsNull(data_clientes.Recordset("cl_fultvta")) = False Then
          Xotrodato = MsgBox("Contiene datos de usuario, desea modificarlos ?", vbCritical + vbYesNo, "SAPP")
          If Xotrodato = vbYes Then
             data_clientes.Recordset.Edit
             data_clientes.Recordset("cl_fultpag") = Null
             data_clientes.Recordset.Update
             data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & txt_mat.Caption
             data_abm.Refresh
             data_abm.Recordset.AddNew
             data_abm.Recordset("cl_codigo") = txt_mat.Caption
             data_abm.Recordset("cl_motivo") = "BORRA DATOS"
             data_abm.Recordset("desc") = "MODIF"
             data_abm.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
             data_abm.Recordset("hora") = Format(Time, "HH:mm")
             data_abm.Recordset("usuario") = WElusuario
             data_abm.Recordset("convenio") = txt_codcnv.Text
             data_abm.Recordset("base") = data_parsec.Recordset("base")
             data_abm.Recordset.Update
             Image1.Visible = False
          End If
       End If
    End If
    If data_clientes.Recordset("estado") = 2 Or data_clientes.Recordset("estado") = 3 Then
       Xresp = MsgBox("SOCIO DE BAJA, desea REACTIVARLO?", vbYesNo + vbExclamation, "Reactivar Socio")
       If Xresp = vbYes Then
          data_clientes.Recordset.Edit
          data_clientes.Recordset("estado") = 1
          data_clientes.Recordset("fecha_baja") = Null
          labestado.Caption = "ACTIVO"
          txt_fecbaj.Text = "__/__/____"
          data_clientes.Recordset.Update
          data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & txt_mat.Caption
          data_abm.Refresh
          data_abm.Recordset.AddNew
          data_abm.Recordset("cl_codigo") = txt_mat.Caption
          data_abm.Recordset("cl_motivo") = "REACTIVACION"
          data_abm.Recordset("desc") = "MODIF"
          data_abm.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
          data_abm.Recordset("hora") = Format(Time, "HH:mm")
          data_abm.Recordset("usuario") = WElusuario
          data_abm.Recordset("convenio") = txt_codcnv.Text
          data_abm.Recordset("base") = data_parsec.Recordset("base")
          data_abm.Recordset.Update
       
       End If
    Else
       If data_clientes.Recordset("fecha_baja") <> "" Then
          Xresp = MsgBox("SOCIO DE BAJA, desea REACTIVARLO?", vbYesNo + vbExclamation, "Reactivar Socio")
          If Xresp = vbYes Then
             data_clientes.Recordset.Edit
             data_clientes.Recordset("estado") = 1
             data_clientes.Recordset("fecha_baja") = Null
             labestado.Caption = "ACTIVO"
             txt_fecbaj.Text = "__/__/____"
             data_clientes.Recordset.Update
          End If
       End If
    End If
    XAlta = 2
    btn_modi.Enabled = False
    btn_cance.Enabled = True
    btn_baja.Enabled = False
    btn_busca.Enabled = False
    btn_fact.Enabled = False
    btn_estadi.Enabled = False
    btn_histo.Enabled = False
    btn_verdeu.Enabled = False
    btn_graba.Enabled = True
    btn_alta.Enabled = False
    Command2.Enabled = False
    Frame4.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    txt_codcnv.Enabled = True
    txt_apellid.SetFocus
'          labestado.Caption = "ACTIVO"
    
'    If IsNull(data_clientes.Recordset("estado")) = False Then
    If labestado.Caption = "ACTIVO" Then
       If data_clientes.Recordset("estado") = 1 Then
          If txt_codcnv.Text <> "" Then
             ControlproxEmi
          End If
       End If
    End If
    If t_pmemi.Text <> "" Then
       If t_pmemi.Text > 0 Then
          If txt_codzon.Text <> "" Then
             If txt_codzon.Text <> 999 Then
                If txt_codcob.Text <> "" Then
                   If txt_codcob.Text = 0 Then
                      Consulta_cobZon
                   End If
                Else
                   Consulta_cobZon
                End If
             End If
          End If
       End If
    End If
End If
Image3.Enabled = False
Image4.Enabled = False
If Trim(txt_ced.Text) = "" Then
   Cl_cedulaAnt = 0
Else
   If txt_ced.Text > 0 Then
      Cl_cedulaAnt = txt_ced.Text
   Else
      Cl_cedulaAnt = 0
   End If
End If

End Sub

Private Sub btn_verdeu_Click()
Dim Xquedes As String
Xquedes = MsgBox("Desea ver el TOTAL de la deuda?", vbExclamation + vbYesNo, "Deudas")
If Xquedes = vbYes Then
   frm_veodeudab.Show vbModal
Else
   frm_veodeudac.Show vbModal
End If

End Sub

Private Sub cbolocalid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_telef.SetFocus
End If

End Sub

Private Sub cbolocalid_LostFocus()
If cbolocalid.Text <> "" Then
   data_zonas.Recordset.FindFirst "zo_nombre = '" & UCase(cbolocalid.Text) & "'"
   If Not data_zonas.Recordset.NoMatch Then
      cbolocalid.Text = data_zonas.Recordset("zo_nombre")
      txt_codzon.Text = data_zonas.Recordset("zo_grupo")
   Else
      MsgBox "Error en la zona digitada", vbCritical
'      cbolocalid.SetFocus
   End If
Else
   MsgBox "Ingrese ZONA", vbCritical
'   cbolocalid.SetFocus
End If

End Sub

Private Sub cbomutual_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 Then
   txt_matmut.SetFocus
End If

End Sub

Private Sub cbomutual_LostFocus()
If cbomutual.Text <> "" Then
   data_mutual.Recordset.FindFirst "ca_nom ='" & cbomutual.Text & "'"
   If Not data_mutual.Recordset.NoMatch Then
   Else
      MsgBox "Mutualista no encontrada, Verifique!!"
      cbomutual.SetFocus
   End If
End If

End Sub

Private Sub cbonomcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_ruta.Enabled = True Then
      t_ruta.SetFocus
   Else
      cbopago.SetFocus
   End If
End If

End Sub

Private Sub cbonomcob_LostFocus()
If cbonomcob.Text <> "" Then
   data_cobrador.Recordset.FindFirst "cb_nombre ='" & UCase(cbonomcob.Text) & "'"
   If Not data_cobrador.Recordset.NoMatch Then
      txt_codcob.Text = data_cobrador.Recordset("cb_numero")
      cbonomcob.Text = data_cobrador.Recordset("cb_nombre")
      cbopago.SetFocus
   Else
      MsgBox "Error al digitar Cobrador", vbCritical
      cbonomcob.SetFocus
   End If
Else
   cbonomcob.SetFocus
End If

End Sub

Private Sub cbonompro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codcob.SetFocus
End If

End Sub

Private Sub cbonompro_LostFocus()
If cbonompro.Text <> "" Then
   data_promo.Recordset.FindFirst "nombre ='" & UCase(cbonompro.Text) & "'"
   If Not data_promo.Recordset.NoMatch Then
      txt_codpro.Text = data_promo.Recordset("idfunc")
      cbonompro.Text = data_promo.Recordset("nombre")
   Else
      txt_codpro.Text = 799
      cbonompro.Text = "*TODOS"
      txt_codpro.SetFocus
   End If
Else
   txt_codpro.Text = 799
   cbonompro.Text = "*TODOS"
   txt_codpro.SetFocus
End If

End Sub

Private Sub cbopago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_diacob.SetFocus
End If

End Sub

Private Sub cbopromos_Click()
Dim Ingreso As Date

If cbopromos.Text = "Grupo de 3 o más" Then
   MsgBox "Recuerde ingresar RUTA (la cédula de titular) a los dependientes del grupo", vbInformation
Else
   If cbopromos.Text = "Pago anual" Then
      If txt_fecing.Text <> "__/__/____" Then
         Ingreso = CDate(txt_fecing.Text) + 334
         If Year(txt_fecing.Text) = Year(Date) Then
            If Month(Ingreso) = 12 Then
               t_pmemi.Text = 1
               t_paemi.Text = Year(Ingreso) + 1
            Else
               t_pmemi.Text = Month(Ingreso) + 1
               t_paemi.Text = Year(Ingreso)
            End If
         End If
      End If
   End If
End If

End Sub

Private Sub cbopromos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_ruta.SetFocus
End If

End Sub

Private Sub cbopromos_LostFocus()
If cbopromos.Text <> "" Then
   BuscaPromos
Else
   labidpromo.Caption = 0
End If

End Sub

Private Sub cbosexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_correo.SetFocus
End If

End Sub

Private Sub cbotarj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nrotarj.SetFocus
End If

End Sub

Private Sub cbotarj_LostFocus()
If cbotarj.Text <> "" Then
   data_tarjetas.Recordset.FindFirst "nombre ='" & UCase(cbotarj.Text) & "'"
   If Not data_tarjetas.Recordset.NoMatch Then
      txt_codemisor.Text = data_tarjetas.Recordset("numero")
      cbotarj.Text = data_tarjetas.Recordset("nombre")
      txt_nrotarj.SetFocus
   Else
      MsgBox "Error al digitar Emisor", vbCritical
      cbotarj.SetFocus
   End If
Else
   txt_nrotarj.SetFocus
End If

End Sub

Private Sub cbotipoced_Click()
If cbotipoced.ListIndex = 0 Then
   txt_ced2.Visible = True
   If XAlta = 2 Then
      txt_ced2.Text = data_clientes.Recordset("cl_codced")
   End If
Else
   txt_ced2.Visible = False
   txt_ced2.Text = 0
End If

End Sub

Private Sub cbotipoced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ced.SetFocus
End If

End Sub

Private Sub Command1_Click()
'    MsgBox "Verifique el Convenio", vbCritical, "Mantenimiento"
'    txt_codcnv.SetFocus

                        data_clientes.Recordset.AddNew
                       data_clientes.Recordset("estado") = 1
                    '   labestado.Caption = "ACTIVO"
                       data_clientes.Recordset("cl_codigo") = txt_mat.Caption
                       data_clientes.Recordset("cl_codconv") = txt_codcnv.Text
                       data_clientes.Recordset("cl_nomconv") = Mid(txt_nomcnv.Text, 1, 30)
                       data_clientes.Recordset("cl_apellid") = txt_apellid.Text
                       If t_otrocnv.Text <> "" Then
                          data_clientes.Recordset("cl_ruc") = t_otrocnv.Text
                       End If
                       If cbotipoced.ListIndex >= 0 Then
                          data_clientes.Recordset("cl_tipoced") = cbotipoced.ListIndex
                       Else
                          data_clientes.Recordset("cl_tipoced") = 0
                       End If
                       If txt_ced.Text <> "" Then
                          data_clientes.Recordset("cl_cedula") = txt_ced.Text
                          If txt_ced2.Text <> "" Then
                             data_clientes.Recordset("cl_codced") = txt_ced2.Text
                          Else
                             txt_ced2.Text = 0
                             data_clientes.Recordset("cl_codced") = 0
                          End If
                       Else
                          txt_ced.Text = 0
                          txt_ced2.Text = 0
                          data_clientes.Recordset("cl_cedula") = txt_ced.Text
                          data_clientes.Recordset("cl_codced") = txt_ced2.Text
                       End If
                       If txt_nac.Text <> "__/__/____" Then
                          data_clientes.Recordset("cl_fnac") = Format(txt_nac.Text, "dd/mm/yyyy")
                       End If
                       If labedad.Caption <> "" Then
                          data_clientes.Recordset("cl_edad") = labedad.Caption
                       Else
                          labedad.Caption = 0
                          data_clientes.Recordset("cl_edad") = labedad.Caption
                       End If
'                       If data_clientes.Recordset("cl_uniedad") <> "" Then
'                          data_clientes.Recordset("cl_uniedad") = labunie.Caption
'                        Else
'                          labunie.Caption = "A"
'                          data_clientes.Recordset("cl_uniedad") = labunie.Caption
'                        End If
                        If labump.Caption <> "" Then
                           data_clientes.Recordset("cl_ultmesp") = labump.Caption
                        Else
                           labump.Caption = 0
                           data_clientes.Recordset("cl_ultmesp") = labump.Caption
                        End If
                        If labuap.Caption <> "" Then
                           data_clientes.Recordset("cl_ultanop") = labuap.Caption
                        Else
                           labuap.Caption = 0
                           data_clientes.Recordset("cl_ultanop") = labuap.Caption
                        End If
                        If labatra.Caption <> "" Then
                           data_clientes.Recordset("cl_atrasoa") = labatra.Caption
                        Else
                           labatra.Caption = 0
                           data_clientes.Recordset("cl_atrasoa") = labatra.Caption
                        End If
                        If labdeudap.Caption <> "" Then
                           data_clientes.Recordset("saldo_cc") = labdeudap.Caption
                        Else
                           labdeudap.Caption = ""
                        End If
                        data_clientes.Recordset("cl_direcci") = txt_direcc1.Text
                        data_clientes.Recordset("cl_entre") = txt_direcc2.Text
                        If t_cel.Text <> "" Then
                           data_clientes.Recordset("cl_dpto") = t_cel.Text
                        Else
                           data_clientes.Recordset("cl_dpto") = Null
                        End If
                        If t_correo.Text <> "" Then
                           data_clientes.Recordset("cl_referen") = t_correo.Text
                        Else
                           data_clientes.Recordset("cl_referen") = Null
                        End If
                        If txt_codzon.Text <> "" Then
                           If cbolocalid.Text <> "" Then
                              data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                              data_clientes.Recordset("cl_zona") = cbolocalid.Text
                           Else
                              cbolocalid.Text = "*TODOS"
                              data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                              data_clientes.Recordset("cl_zona") = cbolocalid.Text
                           End If
                        Else
                           txt_codzon.Text = 999
                           cbolocalid.Text = "*TODOS"
                           data_clientes.Recordset("cl_grupo") = txt_codzon.Text
                           data_clientes.Recordset("cl_zona") = cbolocalid.Text
                        End If
                        If cbosexo.Text = "FEMENINO" Then
                            data_clientes.Recordset("cl_sexo") = 2
                        Else
                            data_clientes.Recordset("cl_sexo") = 1
                        End If
                        data_clientes.Recordset("cl_telefon") = txt_telef.Text
                        data_clientes.Recordset("cl_dircobr") = txt_dircob.Text
                        data_clientes.Recordset("cl_nombre") = txt_conmut.Text
                        data_clientes.Recordset("cl_socmnom") = cbomutual.Text
                        data_clientes.Recordset("cl_nrosocm") = txt_matmut.Text
                        If labidpromo.Caption <> "" Then
                           data_clientes.Recordset("idpromos") = Val(labidpromo.Caption)
                        Else
                           data_clientes.Recordset("idpromos") = 0
                        End If
                        If t_pmemi.Text <> "" Then
                           data_clientes.Recordset("mesproxemi") = Val(t_pmemi.Text)
                           data_clientes.Recordset("anoproxemi") = Val(t_paemi.Text)
                        Else
                           data_clientes.Recordset("mesproxemi") = 0
                           data_clientes.Recordset("anoproxemi") = 0
                        End If
                        
                        If txt_fecing.Text <> "__/__/____" Then
                           data_clientes.Recordset("cl_fecing") = Format(txt_fecing.Text, "dd/mm/yyyy")
                        End If
                        If txt_fecbaj.Text <> "__/__/____" Then
                           data_clientes.Recordset("fecha_baja") = Format(txt_fecbaj.Text, "dd/mm/yyyy")
                        End If
                        If txt_codpro.Text <> "" Then
                           If cbonompro.Text <> "" Then
                              data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                              data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                           Else
                              cbonompro.Text = "*TODOS"
                              data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                              data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                           End If
                        Else
                           txt_codpro.Text = 799
                           cbonompro.Text = "*TODOS"
                           data_clientes.Recordset("cl_nrovend") = txt_codpro.Text
                           data_clientes.Recordset("cl_nomvend") = cbonompro.Text
                        End If
                        If txt_codcob.Text <> "" Then
                           If cbonomcob.Text <> "" Then
                              data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                              data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                           Else
                              cbonomcob.Text = "*TODOS"
                              data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                              data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                           End If
                        Else
                           txt_codcob.Text = 0
                           cbonomcob.Text = "*TODOS"
                           data_clientes.Recordset("cl_nrocobr") = txt_codcob.Text
                           data_clientes.Recordset("cl_nomcobr") = Mid(cbonomcob.Text, 1, 25)
                        End If
                        If cbopago.Text = "Debito Automatico" Then
                           data_clientes.Recordset("cl_forpago") = 2
                           data_clientes.Recordset("cl_descpag") = "Debito Automatico"
                        Else
                           data_clientes.Recordset("cl_forpago") = 1
                           data_clientes.Recordset("cl_descpag") = "Abono Mensual"
                        End If
                        data_clientes.Recordset("cl_diacobr") = txt_diacob.Text
                        data_clientes.Recordset("tit_tarj") = txt_nomtarj.Text
                        If txt_nrotarj.Text <> "" Then
                           data_clientes.Recordset("cl_nrotarj") = txt_nrotarj.Text
                        Else
                           data_clientes.Recordset("cl_nrotarj") = 0
                        End If
                        If t_ruta.Text <> "" Then
                           data_clientes.Recordset("cl_codruta") = t_ruta.Text
                        End If
                        If txt_cedtarj.Text <> "" Then
                           data_clientes.Recordset("ci_tarj") = txt_cedtarj.Text
                        Else
                           txt_cedtarj.Text = 0
                           data_clientes.Recordset("ci_tarj") = txt_cedtarj.Text
                        End If
                        If txt_codtarj.Text <> "" Then
                           data_clientes.Recordset("codcitarj") = txt_codtarj.Text
                        Else
                           txt_codtarj.Text = 0
                           data_clientes.Recordset("codcitarj") = txt_codtarj.Text
                        End If
                        If txt_codemisor.Text <> "" Then
                           data_clientes.Recordset("cl_tjemi_c") = txt_codemisor.Text
                        Else
                           txt_codemisor.Text = 0
                           data_clientes.Recordset("cl_tjemi_c") = txt_codemisor.Text
                        End If
                        If cbotarj.Text <> "" Then
                           data_clientes.Recordset("cl_tjemi_n") = cbotarj.Text
                        Else
                           cbotarj.Text = ""
                           data_clientes.Recordset("cl_tjemi_n") = cbotarj.Text
                        End If
                        If txt_vence.Text <> "__/__/____" Then
                           data_clientes.Recordset("cl_tj_venc") = Format(txt_vence.Text, "dd/mm/yyyy")
                        End If
                        data_clientes.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
                        If Option1.Value = True Then
                           data_clientes.Recordset("cl_decuota") = 1
                           If mfcarta.Text <> "__/__/____" Then
                              data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                           End If
                        Else
                           If Option2.Value = True Then
                              data_clientes.Recordset("cl_decuota") = 2
                              If mfcarta.Text <> "__/__/____" Then
                                 data_clientes.Recordset("fecha_reac") = Format(mfcarta.Text, "dd/mm/yyyy")
                              End If
                           Else
                              data_clientes.Recordset("cl_decuota") = 0
                           End If
                        End If
                        data_clientes.Recordset("saldo_chc2") = cbosrv.ListIndex
                        data_clientes.Recordset.Update
                        data_clientes.Refresh
                        data_clientes.Recordset.FindFirst "cl_codigo =" & txt_mat.Caption
                        txt_nomcnv.Enabled = True
                        vercli
                        txt_nomcnv.Enabled = False
                        data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & txt_mat.Caption
                        data_abm.Refresh
                        data_abm.Recordset.AddNew
                        data_abm.Recordset("cl_codigo") = txt_mat.Caption
                        If XWeltipoU = "USUARIOS ADM" Then
                           data_abm.Recordset("cl_motivo") = "ALTA CON PROMOTOR"
                        Else
                           data_abm.Recordset("cl_motivo") = "ALTA DE FICHA"
                        End If
                        data_abm.Recordset("desc") = "ALTA"
                        data_abm.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                        data_abm.Recordset("hora") = Format(Time, "HH:mm")
                        data_abm.Recordset("usuario") = WElusuario
                        data_abm.Recordset("convenio") = txt_codcnv.Text
                        data_abm.Recordset("base") = data_parsec.Recordset("base")
                        data_abm.Recordset.Update
                        XAlta = 0
                        If XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "ADMINISTRADOR" Then
                           btn_modi.Enabled = True
                           btn_baja.Enabled = True
                        Else
                           If XWeltipoU = "USUARIOS" Then
                              btn_modi.Enabled = True
                              btn_baja.Enabled = False
                           Else
                              btn_modi.Enabled = False
                              btn_baja.Enabled = False
                           End If
                        End If
                        btn_busca.Enabled = True
                        btn_cance.Enabled = False
                        btn_fact.Enabled = True
                        btn_estadi.Enabled = True
                        btn_histo.Enabled = True
                        btn_verdeu.Enabled = True
                        btn_graba.Enabled = False
                        btn_alta.Enabled = True
                        Command2.Enabled = True
                        Frame4.Enabled = True
                        Frame1.Enabled = False
                        Frame2.Enabled = False

End Sub

Private Sub Command2_Click()
frmabm.MousePointer = 11
If frm_ctrlactenf.Visible = True Then
   frmabm.MousePointer = 0
   MsgBox "Ya está abierto!", vbCritical
Else
   frm_ctrlactenf.Show
End If
frmabm.MousePointer = 0

End Sub


Private Sub Form_Load()
Dim Xs As String
     Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      DesignX = 800
      DesignY = 600
      RePosForm = True
      DoResize = False
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips
      Xpixels = Screen.Width / Xtwips

      ScaleFactorX = (Xpixels / DesignX)
      ScaleFactorY = (Ypixels / DesignY)
      ScaleMode = 1  ' twips
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height
      MyForm.Width = Me.Width
With Image2
   .Left = 0
   .Top = 0
   .Height = Me.Height
   .Width = Me.Width
End With
'cl_codruta
data_cnvmut.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_clicnv.ConnectionString = "dsn=" & Xconexrmt
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.RecordSource = "clientes"
'Data1.Refresh
data_abm.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_abm.RecordSource = "Select top 50 * from abmsocio"
'data_abm.Refresh

'data_clicnv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_clientes.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cobrador.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_mutual.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_parsec.DatabaseName = App.path & "\parse.mdb"
data_parsec.RecordSource = "parsec0"
data_parsec.Refresh

data_promo.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_promo.RecordSource = "select * from vende_func order by nombre"
data_promo.Refresh

'select * from vendedor order by vn_nombre

data_tarjetas.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_tarjetas.RecordSource = "creditos"
data_tarjetas.Refresh

data_zonas.Connect = "odbc;dsn=" & Xconexrmt & ";"
'''data_clientes.RecordSource = "clientes"
'''data_clientes.Refresh
data_mutual.RecordSource = "select * from ca_adm order by ca_nom"
data_mutual.Refresh
CargaPromos

Xs = "SI"
'vercli
'Borrar
If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "USUARIOS DESP" Then
   btn_modi.Enabled = True
   btn_baja.Enabled = True
   If WElusuario = "MCOSTA" Or WElusuario = "JFERNAN" Or WElusuario = "MPEREZ" Or XWeltipoU = "ADMINISTRADOR" Then
'      WElusuario = "MPEREZ" Or WElusuario = "JONATHAN" Or WElusuario = "MSANCHEZ" Then
   
   Else
      Option1.Enabled = False
      Option2.Enabled = False
      mfcarta.Enabled = False
      cbosrv.Enabled = False
   End If
Else
   If XWeltipoU = "USUARIOS" Or XWeltipoU = "ADM FARMACIA" Then
      btn_modi.Enabled = True
      btn_baja.Enabled = False
      data_clicnv.RecordSource = "Select * from convenio where cnv_emite <>'" & Trim(Xs) & "'"
      data_clicnv.Refresh
      Option1.Enabled = False
      Option2.Enabled = False
      mfcarta.Enabled = False
      cbosrv.Enabled = False
   Else
      Option1.Enabled = False
      Option2.Enabled = False
      mfcarta.Enabled = False
      cbosrv.Enabled = False
      btn_modi.Enabled = False
      btn_baja.Enabled = False
      data_clicnv.RecordSource = "Select * from convenio where cnv_emite <>'" & Trim(Xs) & "'"
      data_clicnv.Refresh
   End If
End If

If ControlUsuario("Datos tarjetas") = 1 Then
   btn_histo.Enabled = True
   Label29.Visible = True
   Label30.Visible = True
   Label31.Visible = True
   Label32.Visible = True
   Label33.Visible = True
   txt_nomtarj.Visible = True
   txt_cedtarj.Visible = True
   txt_codtarj.Visible = True
   txt_codemisor.Visible = True
   cbotarj.Visible = True
   txt_nrotarj.Visible = True
   txt_vence.Visible = True
Else
   Label29.Visible = False
   Label30.Visible = False
   Label31.Visible = False
   Label32.Visible = False
   Label33.Visible = False
   txt_nomtarj.Visible = False
   txt_cedtarj.Visible = False
   txt_codtarj.Visible = False
   txt_codemisor.Visible = False
   cbotarj.Visible = False
   txt_nrotarj.Visible = False
   txt_vence.Visible = False
   
   btn_histo.Enabled = False
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

Private Sub Form_Terminate()
Unload Me
'frmabm.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me

End Sub


Public Function Borrar()
Image1.Visible = False
txt_mat.Caption = ""
labestado.Caption = ""
labatra.Caption = ""
labump.Caption = ""
labuap.Caption = ""
Label7.Caption = ""
txt_codcnv.Enabled = True
txt_nomcnv.Enabled = True
txt_codcnv.Text = ""
txt_nomcnv.Text = ""
txt_codcnv.Enabled = False
txt_nomcnv.Enabled = False
txt_apellid.Text = ""
txt_ced.Text = ""
txt_ced2.Text = ""
cbotipoced.ListIndex = -1
txt_nac.Text = "__/__/____"
t_cel.Text = ""
t_correo.Text = ""
labedad.Caption = ""
labunie.Caption = ""
labdias.Caption = ""
txt_direcc1.Text = ""
txt_direcc2.Text = ""
txt_codzon.Text = ""
'cbolocalid.Text = ""
txt_telef.Text = ""
cbosexo.Text = ""
t_rs.Text = ""
txt_dircob.Text = ""
txt_conmut.Text = ""
cbomutual.Text = ""
txt_matmut.Text = ""
labdeudap.Caption = ""
txt_fecing.Text = "__/__/____"
txt_fecbaj.Text = "__/__/____"
txt_codpro.Text = ""
t_ruta.Text = ""
'cbonompro.Text = ""
txt_codcob.Text = ""
'cbonomcob.Text = ""
cbopago.ListIndex = 0
txt_diacob.Text = ""
txt_nomtarj.Text = ""
txt_cedtarj.Text = ""
txt_codtarj.Text = ""
txt_codemisor.Text = ""
cbotarj.Text = ""
txt_nrotarj.Text = ""
txt_vence.Text = "__/__/____"
t_otrocnv.Text = ""
Option1.Value = False
Option2.Value = False
mfcarta.Text = "__/__/____"
cbosrv.ListIndex = -1
labidpromo.Caption = ""
cbopromos.Text = ""
t_pmemi.Text = ""
t_paemi.Text = ""


End Function

Private Sub Image1_Click()
If IsNull(data_clientes.Recordset("cl_fultvta")) = True Then
   If IsNull(data_clientes.Recordset("cl_fultpag")) = False Then
      Image1.ToolTipText = "FECHA NACIMIENTO: " & data_clientes.Recordset("cl_fultpag")
   End If
Else
   If IsNull(data_clientes.Recordset("cl_fultpag")) = False Then
      Image1.ToolTipText = "FECHA:" & data_clientes.Recordset("cl_fultvta") & " CAT:" & data_clientes.Recordset("cl_celular") & " Nro." & data_clientes.Recordset("cl_tipocli") & " NACIMIENTO: " & data_clientes.Recordset("cl_fultpag")
   Else
      Image1.ToolTipText = "FECHA:" & data_clientes.Recordset("cl_fultvta") & " CAT:" & data_clientes.Recordset("cl_celular") & " Nro." & data_clientes.Recordset("cl_tipocli")
   End If
End If

End Sub

Private Sub Image3_DblClick()
If Trim(txt_mat.Caption) <> "" Then
   frm_notas_med.Show vbModal
Else
   MsgBox "No hay matrícula seleccionada."
End If

End Sub

Private Sub Image4_DblClick()
frm_notas_med.Show vbModal

End Sub

Private Sub Label9_Click()
If txt_codcnv.Text <> "" Then
   Xconv = txt_codcnv.Text
   frm_buscacnv.Show vbModal
End If

End Sub

Private Sub mfcarta_GotFocus()
mfcarta.Text = Date
End Sub

Private Sub Option1_DblClick()
If Option1.Value = True Then
   Option1.Value = False
   mfcarta.Enabled = True
   mfcarta.Text = "__/__/____"

End If

End Sub

Private Sub Option2_DblClick()
If Option2.Value = True Then
   Option2.Value = False
   mfcarta.Enabled = True
   mfcarta.Text = "__/__/____"
End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
   mfcarta.Enabled = True
Else
   mfcarta.Text = "__/__/____"
   mfcarta.Enabled = False
End If

End Sub

Private Sub Option3_DblClick()
If Option3.Value = True Then
   Option3.Value = False
   mfcarta.Text = "__/__/____"
   
End If
End Sub

Private Sub Option4_DblClick()
If Option4.Value = True Then
   Option4.Value = False
   mfcarta.Text = "__/__/____"
End If

End Sub

Private Sub Option5_DblClick()
If Option5.Value = True Then
   Option5.Value = False
   mfcarta.Text = "__/__/____"
End If

End Sub

Private Sub t_cel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 Then
   cbosexo.SetFocus
End If

End Sub

Private Sub t_cel_LostFocus()
If Trim(t_cel.Text) = "" Then
   MsgBox "Falta ingresar dato de celular.", vbCritical
Else
   t_cel.Text = Trim(t_cel.Text)
End If

End Sub

Private Sub t_correo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   cbomutual.SetFocus
End If

End Sub

Private Sub t_correo_LostFocus()
If Trim(t_correo.Text) = "" Then
   MsgBox "Falta ingresar dato de correo.", vbCritical
Else
   t_correo.Text = Trim(t_correo.Text)
End If
If Trim(t_correo.Text) = "no aplica" Then
   t_correo.Text = UCase(t_correo.Text)
End If

End Sub

Private Sub t_ruta_Change()
If Not IsNumeric(t_ruta.Text) And _
 t_ruta.Text <> "" Then
 Beep
 MsgBox "Se debe ingresar solo números en RUTA o vacío"
  t_ruta.Text = ""
  t_ruta.SetFocus
End If

End Sub

Private Sub t_ruta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbopago.SetFocus
End If

End Sub

Private Sub t_ruta_LostFocus()
Dim txtced As String
txtced = ""

If t_ruta.Text <> "" Then
   If cbopromos.Text = "Grupo de 3 o más" Then
      If txt_ced.Text <> "" Then
         txtced = Trim(txt_ced.Text) & Trim(txt_ced2.Text)
         If txtced = t_ruta.Text Then
            MsgBox "El titular no lleva ruta", vbExclamation
            t_ruta.Text = ""
         End If
      Else
         MsgBox "Debe ingresar número de cédula", vbCritical
      End If
   Else
      If Val(t_ruta.Text) = Val(txt_mat.Caption) Then
         MsgBox "El número de ruta NO PUEDE ser igual al número de cliente. Si es titular, no lleva RUTA.", vbCritical
         t_ruta.Text = ""
      Else
         VerPromoCli (t_ruta.Text)
         If Xconvprom <> "" Then
            VerPromosiono (Xconvprom)
         End If
      End If
   End If
   
End If

End Sub

Private Sub txt_apellid_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If txt_nomcnv.Text = "" Then
      MsgBox "No seleccionó convenio, VERIFIQUE!", vbCritical
      txt_codcnv.SetFocus
   Else
      cbotipoced.SetFocus
'      txt_ced.SetFocus
   End If
End If

End Sub

Private Sub txt_apellid_LostFocus()
If XAlta = 1 Then
   Data1.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Caption
   Data1.Refresh
'   Data1.Recordset.FindFirst "cl_codigo =" & txt_mat.Caption
   If Data1.Recordset.RecordCount > 0 Then
      MsgBox "Ya existe la matrícula", vbCritical, "Mensaje"
      btn_cance_Click
   End If
End If

End Sub

Private Sub txt_buscli_KeyPress(KeyAscii As Integer)
Dim Xgpoconv As String
Dim Xfecvolver As Date
Dim RsP As New ADODB.Recordset
Dim SqlP As String

On Error GoTo Quepasaal

If txt_buscli.Text <> "" Then
    If KeyAscii = 13 Then
       
       frmabm.MousePointer = 11
'       data_clientes.Recordset.MoveFirst
'       data_clientes.Recordset.FindFirst "cl_codigo =" & txt_buscli.Text
'       If Not data_clientes.Recordset.NoMatch Then
        data_clientes.RecordSource = "Select * from clientes where cl_codigo =" & txt_buscli.Text
        data_clientes.Refresh
        
        If data_clientes.Recordset.RecordCount > 0 Then
           Consulta_Notas (txt_buscli.Text)
            txt_buscli.Text = ""
            Borrar
            If IsNull(data_clientes.Recordset("cl_fultvta")) = False Then
               If IsNull(data_clientes.Recordset("cl_tipocli")) = False Then
                  If data_clientes.Recordset("cl_tipocli") = 1 Or data_clientes.Recordset("cl_tipocli") = 2 Then
                     Image1.Visible = True
                  Else
                     Image1.Visible = True
                  End If
               Else
                  Image1.Visible = False
               End If
            Else
               Image1.Visible = False
            End If
            If Image1.Visible = False Then
               If IsNull(data_clientes.Recordset("cl_fultpag")) = False Then
                  Image1.Visible = True
               End If
            End If
            If data_clientes.Recordset("estado") = 2 Or data_clientes.Recordset("estado") = 3 Then
               labestado.Caption = "BAJA"
            Else
               If data_clientes.Recordset("fecha_baja") <> "" Then
                  labestado.Caption = "BAJA"
               Else
                  labestado.Caption = "ACTIVO"
               End If
            End If
            txt_mat.Caption = data_clientes.Recordset("cl_codigo")
            If IsNull(data_clientes.Recordset("cl_apellid")) = False Then
               txt_apellid.Text = data_clientes.Recordset("cl_apellid")
            End If
            If IsNull(data_clientes.Recordset("cl_codconv")) = True Then
               MsgBox "Verifique el convenio", vbCritical, "Mensaje"
               txt_codcnv.Text = ""
               Xgpoconv = ""
            Else
               txt_codcnv.Text = data_clientes.Recordset("cl_codconv")
               data_cnvmut.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_codcnv.Text & "'"
               data_cnvmut.Refresh
               If data_cnvmut.Recordset.RecordCount > 0 Then
                  If IsNull(data_cnvmut.Recordset("cnv_entre")) = False Then
                     If Trim(data_cnvmut.Recordset("cnv_entre")) <> "" Then
                        If Val(data_cnvmut.Recordset("cnv_cuenta")) = Val(txt_mat.Caption) Then
                           t_rs.Text = data_cnvmut.Recordset("cnv_entre")
                        Else
                           t_rs.Text = ""
                        End If
                     Else
                        t_rs.Text = ""
                     End If
                  Else
                     t_rs.Text = ""
                  End If
                  If IsNull(data_cnvmut.Recordset("cnv_grupo")) = False Then
                     If Trim(data_cnvmut.Recordset("cnv_grupo")) <> "" Then
                        Xgpoconv = data_cnvmut.Recordset("cnv_grupo")
                     Else
                        Xgpoconv = ""
                     End If
                  Else
                     Xgpoconv = ""
                  End If
               Else
                  Xgpoconv = ""
               End If
               data_cnvmut.Recordset.Close
            End If
            txt_nomcnv.Enabled = True
            If IsNull(data_clientes.Recordset("cl_nomconv")) = True Then
               txt_nomcnv.Text = ""
            Else
               txt_nomcnv.Text = data_clientes.Recordset("cl_nomconv")
            End If
            txt_nomcnv.Enabled = False
            If IsNull(data_clientes.Recordset("cl_tipoced")) = False Then
               cbotipoced.ListIndex = data_clientes.Recordset("cl_tipoced")
               If data_clientes.Recordset("cl_tipoced") = 0 Then
                  txt_ced2.Visible = True
                  If data_clientes.Recordset("cl_cedula") <> "" Then
                     txt_ced.Text = data_clientes.Recordset("cl_cedula")
                  Else
                     txt_ced.Text = ""
                  End If
                  If data_clientes.Recordset("cl_codced") <> "" Then
                     txt_ced2.Text = data_clientes.Recordset("cl_codced")
                  Else
                     txt_ced2.Text = 0
                  End If
               Else
                  txt_ced2.Visible = False
                  If data_clientes.Recordset("cl_cedula") <> "" Then
                     txt_ced.Text = data_clientes.Recordset("cl_cedula")
                  Else
                     txt_ced.Text = ""
                  End If
                  txt_ced2.Text = 0
               End If
            Else
               cbotipoced.ListIndex = 0
               txt_ced2.Visible = True
               If data_clientes.Recordset("cl_cedula") <> "" Then
                  txt_ced.Text = data_clientes.Recordset("cl_cedula")
               Else
                  txt_ced.Text = ""
               End If
               If data_clientes.Recordset("cl_codced") <> "" Then
                  txt_ced2.Text = data_clientes.Recordset("cl_codced")
               Else
                  txt_ced2.Text = 0
               End If
            End If
            If IsNull(data_clientes.Recordset("cl_ruc")) = True Then
               t_otrocnv.Text = ""
            Else
               t_otrocnv.Text = data_clientes.Recordset("cl_ruc")
            End If
            If IsNull(data_clientes.Recordset("cl_codruta")) = True Then
               t_ruta.Text = ""
            Else
               t_ruta.Text = data_clientes.Recordset("cl_codruta")
            End If
            If IsNull(data_clientes.Recordset("cl_dpto")) = False Then
               t_cel.Text = data_clientes.Recordset("cl_dpto")
            Else
               t_cel.Text = ""
            End If
            If IsNull(data_clientes.Recordset("cl_referen")) = False Then
               t_correo.Text = data_clientes.Recordset("cl_referen")
            Else
               t_correo.Text = ""
            End If
            If data_clientes.Recordset("cl_fnac") <> "" Then
               txt_nac.Text = Format(data_clientes.Recordset("cl_fnac"), "dd/mm/yyyy")
            Else
               txt_nac.Text = "__/__/____"
               labedad.Caption = 0
               labunie.Caption = 0
               labdias.Caption = 0
            End If
            If Not IsDate(txt_nac.Text) Then
'   MsgBox "Digite una fecha válida"
            Else
                CalculaEdad (txt_nac.Text)
            End If
            If data_clientes.Recordset("cl_ultmesp") <> "" Then
               labump.Caption = data_clientes.Recordset("cl_ultmesp")
            Else
               labump.Caption = ""
            End If
            If data_clientes.Recordset("cl_ultanop") <> "" Then
               If data_clientes.Recordset("cl_ultanop") = 0 Then
                  labuap.Caption = data_clientes.Recordset("cl_ultanop")
               Else
                  labuap.Caption = data_clientes.Recordset("cl_ultanop")
               End If
            Else
               labuap.Caption = ""
            End If
            If data_clientes.Recordset("cl_atrasoa") <> "" Then
               labatra.Caption = data_clientes.Recordset("cl_atrasoa")
            Else
               labatra.Caption = ""
            End If
            If data_clientes.Recordset("saldo_cc") <> "" Then
               labdeudap.Caption = data_clientes.Recordset("saldo_cc")
            Else
               labdeudap.Caption = ""
            End If
            If data_clientes.Recordset("cl_direcci") <> "" Then
               txt_direcc1.Text = data_clientes.Recordset("cl_direcci")
            Else
               txt_direcc1.Text = ""
            End If
            If data_clientes.Recordset("cl_entre") <> "" Then
               txt_direcc2.Text = data_clientes.Recordset("cl_entre")
            Else
               txt_direcc2.Text = ""
            End If
            If data_clientes.Recordset("cl_grupo") <> "" Then
               txt_codzon.Text = data_clientes.Recordset("cl_grupo")
            Else
               txt_codzon.Text = 0
            End If
            If data_clientes.Recordset("cl_zona") <> "" Then
               cbolocalid.Text = data_clientes.Recordset("cl_zona")
            Else
               cbolocalid.Text = ""
            End If
            If data_clientes.Recordset("cl_sexo") = 2 Then
               cbosexo.Text = "FEMENINO"
            Else
               cbosexo.Text = "MASCULINO"
            End If
            If data_clientes.Recordset("cl_telefon") <> "" Then
               txt_telef.Text = data_clientes.Recordset("cl_telefon")
            Else
               txt_telef.Text = ""
            End If
            If data_clientes.Recordset("cl_dircobr") <> "" Then
               txt_dircob.Text = data_clientes.Recordset("cl_dircobr")
            Else
               txt_dircob.Text = ""
            End If
            If IsNull(data_clientes.Recordset("cl_nombre")) = False Then
                txt_conmut.Text = data_clientes.Recordset("cl_nombre")
            End If
            If data_clientes.Recordset("cl_socmnom") <> "" Then
               cbomutual.Text = data_clientes.Recordset("cl_socmnom")
            Else
               cbomutual.Text = ""
            End If
            If data_clientes.Recordset("cl_nrosocm") <> "" Then
               txt_matmut.Text = data_clientes.Recordset("cl_nrosocm")
            Else
               txt_matmut.Text = ""
            End If
            
            If data_clientes.Recordset("cl_fecing") <> "" Then
               txt_fecing.Text = Format(data_clientes.Recordset("cl_fecing"), "dd/mm/yyyy")
            Else
               txt_fecing.Text = "__/__/____"
            End If
            If data_clientes.Recordset("fecha_baja") <> "" Then
               txt_fecbaj.Text = Format(data_clientes.Recordset("fecha_baja"), "dd/mm/yyyy")
            Else
               txt_fecbaj.Text = "__/__/____"
            End If
            If data_clientes.Recordset("cl_nrovend") <> "" Then
               txt_codpro.Text = data_clientes.Recordset("cl_nrovend")
            Else
               txt_codpro.Text = ""
            End If
            If data_clientes.Recordset("cl_nomvend") <> "" Then
               cbonompro.Text = data_clientes.Recordset("cl_nomvend")
            Else
               cbonompro.Text = ""
            End If
            If IsNull(data_clientes.Recordset("idpromos")) = False Then
               labidpromo.Caption = data_clientes.Recordset("idpromos")
               If Val(labidpromo.Caption) > 0 Then
                  BuscaPromosId
               Else
                  cbopromos.Text = ""
               End If
            Else
               labidpromo.Caption = 0
               cbopromos.Text = ""
            End If
            If IsNull(data_clientes.Recordset("mesproxemi")) = False Then
               t_pmemi.Text = data_clientes.Recordset("mesproxemi")
               t_paemi.Text = data_clientes.Recordset("anoproxemi")
            Else
               t_pmemi.Text = 0
               t_paemi.Text = 0
            End If
            
            If data_clientes.Recordset("cl_nrocobr") <> "" Then
               txt_codcob.Text = data_clientes.Recordset("cl_nrocobr")
            Else
               txt_codcob.Text = ""
            End If
            If data_clientes.Recordset("cl_nomcobr") <> "" Then
               cbonomcob.Text = data_clientes.Recordset("cl_nomcobr")
            Else
               cbonomcob.Text = ""
            End If
            If IsNull(data_clientes.Recordset("cl_descpag")) = True Then
               cbopago.Text = "Abono Mensual"
            Else
               If UCase(data_clientes.Recordset("cl_descpag")) = "DEBITO AUTOMATICO" Then
                  cbopago.Text = "Debito Automatico"
               Else
                  cbopago.Text = "Abono Mensual"
               End If
            End If
            If data_clientes.Recordset("cl_diacobr") <> "" Then
               txt_diacob.Text = data_clientes.Recordset("cl_diacobr")
            Else
               txt_diacob.Text = ""
            End If
            If data_clientes.Recordset("tit_tarj") <> "" Then
               txt_nomtarj.Text = data_clientes.Recordset("tit_tarj")
            Else
               txt_nomtarj.Text = ""
            End If
            If data_clientes.Recordset("cl_nrotarj") <> "" Then
               txt_nrotarj.Text = data_clientes.Recordset("cl_nrotarj")
            Else
               txt_nrotarj.Text = ""
            End If
            If data_clientes.Recordset("ci_tarj") <> "" Then
               txt_cedtarj.Text = data_clientes.Recordset("ci_tarj")
            Else
               txt_cedtarj.Text = ""
            End If
            If data_clientes.Recordset("codcitarj") <> "" Then
               txt_codtarj.Text = data_clientes.Recordset("codcitarj")
            Else
               txt_codtarj.Text = ""
            End If
            If data_clientes.Recordset("cl_tjemi_c") <> "" Then
               txt_codemisor.Text = data_clientes.Recordset("cl_tjemi_c")
            Else
               txt_codemisor.Text = ""
            End If
            If data_clientes.Recordset("cl_tjemi_n") <> "" Then
               cbotarj.Text = data_clientes.Recordset("cl_tjemi_n")
            Else
               cbotarj.Text = ""
            End If
            If data_clientes.Recordset("cl_tj_venc") <> "" Then
               txt_vence.Text = Format(data_clientes.Recordset("cl_tj_venc"), "dd/mm/yyyy")
            Else
               txt_vence.Text = "__/__/____"
            End If
            
            If IsNull(data_clientes.Recordset("cl_decuota")) = False Then
               If data_clientes.Recordset("cl_decuota") = 1 Then
                  Option1.Value = True
               Else
                  If data_clientes.Recordset("cl_decuota") = 2 Then
                     Option2.Value = True
                  Else
                     If data_clientes.Recordset("cl_decuota") = 3 Then
                        Option3.Value = True
                     Else
                        If data_clientes.Recordset("cl_decuota") = 4 Then
                           Option4.Value = True
                        Else
                           If data_clientes.Recordset("cl_decuota") = 5 Then
                              Option5.Value = True
                           Else
                              Option1.Value = False
                              Option2.Value = False
                              Option3.Value = False
                              Option4.Value = False
                              Option5.Value = False
                           End If
                        End If
                     End If
                  End If
               End If
            Else
               Option1.Value = False
               Option2.Value = False
               Option3.Value = False
               Option4.Value = False
               Option5.Value = False
            End If
            If IsNull(data_clientes.Recordset("fecha_reac")) = False Then
               mfcarta.Text = Format(data_clientes.Recordset("fecha_reac"), "dd/mm/yyyy")
            Else
               mfcarta.Text = "__/__/____"
            End If
            If IsNull(data_clientes.Recordset("saldo_chc2")) = False Then
               cbosrv.ListIndex = data_clientes.Recordset("saldo_chc2")
            Else
               cbosrv.ListIndex = -1
            End If
            labmr.Caption = ""
            Dim Xtex As String
            Dim Xqfees As Date
            Xqfees = Date + 1
            Xtex = ""
            If txt_mat.Caption <> "" Then
               If CBool(Online()) = True Then
                  data_fecped.Connect = "ODBC;DSN=" & Xconexrmt & ";"
                  data_fecped.RecordSource = "Select * from t_fechas where mat_pac =" & txt_mat.Caption & " and cdate(fecha) >=#" & Format(Xqfees, "yyyy/mm/dd") & "#"
                  data_fecped.Refresh
                  If data_fecped.Recordset.RecordCount > 0 Then
                     Do While Not data_fecped.Recordset.EOF
                        If Xtex = "" Then
                           Xtex = "Anotado para: " & data_fecped.Recordset("especial") & " DIA:" & data_fecped.Recordset("fecha") & " H." & data_fecped.Recordset("hora") & " BASE:" & data_fecped.Recordset("base")
                        Else
                           Xtex = Xtex & Chr(13) & "Anotado para: " & data_fecped.Recordset("especial") & " DIA:" & data_fecped.Recordset("fecha") & " H." & data_fecped.Recordset("hora") & " BASE:" & data_fecped.Recordset("base")
                        End If
                        data_fecped.Recordset.MoveNext
                     Loop
                  End If
                  If Xtex <> "" Then
                     labavis.Caption = Xtex
                  Else
                     labavis.Caption = ""
                  End If
                  data_fecped.Recordset.Close
               Else
                  labavis.Caption = ""
               End If
            End If
            Veoladeuda (txt_mat.Caption)
            If cbopromos.Text = "Grupo de 3 o más" Then
               VerPromoCliNew
            Else
               VerPromocion (txt_mat.Caption)
            End If
       Else
            MsgBox "Matrícula no encontrada", vbInformation, "Búsqueda"
            txt_buscli.SelLength = Len(txt_buscli.Text)
            Borrar
       End If
       Dim Xelaviso As String
       Xgpoconv = ""
       If Xgpoconv <> "" Then
          
          Xelaviso = labavis.Caption
          labavis.Caption = ""
          If Xgpoconv = "CCOU" Or Xgpoconv = "SMI" Or Xgpoconv = "UNIVERSAL" Or _
             Xgpoconv = "H.EVANGELICO" Or Xgpoconv = "CASA DE GALICIA" Then
             ConectarBD
             ConbdSapp.Open
             SqlP = "Select * from prestamo where nom1 ='" & Trim(str(txt_mat.Caption)) & "' and nomc ='" & "MEDICO DE REFERENCIA" & "' order by fecing DESC"
             With RsP
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open SqlP, ConbdSapp, , , adCmdText
             End With
             If RsP.RecordCount > 0 Then
                RsP.MoveFirst
                labmr.Caption = "Med.Ref:" & RsP("desccar") & " FECHA:" & RsP("fecing")
             End If
             ConbdSapp.Close
             If Val(labedad.Caption) = 0 Then
                If Val(labunie.Caption) = 0 Then
                   If Val(labdias.Caption) > 0 And Val(labdias.Caption) <= 10 Then
                      data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190001
                      data_lin.Refresh
                      If data_lin.Recordset.RecordCount > 0 Then
                         labavis.Caption = "METAS: " & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " " & data_lin.Recordset("base")
                      Else
                         labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CAPTACIÓN RECIÉN NACIDO-"
                      End If
                      data_lin.Recordset.Close
                   End If
                Else
                   If Val(labunie.Caption) <= 11 Then
                      data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190003 & " order by fecha"
                      data_lin.Refresh
                      If data_lin.Recordset.RecordCount > 0 Then
                         data_lin.Recordset.MoveFirst
                         labavis.Caption = "METAS: "
                         Do While Not data_lin.Recordset.EOF
                            labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " " & data_lin.Recordset("base")
                            data_lin.Recordset.MoveNext
                         Loop
                         data_lin.Recordset.MovePrevious
                         Xfecvolver = data_lin.Recordset("fecha") + 36
                         labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                      Else
                         labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.1ER.AÑO DE VIDA-"
                      End If
                      data_lin.Recordset.Close
                   End If
                End If
             Else
                If Val(labedad.Caption) = 1 And Val(labunie.Caption) >= 0 Then
                   data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190003 & " order by fecha"
                   data_lin.Refresh
                   If data_lin.Recordset.RecordCount > 0 Then
                      data_lin.Recordset.MoveFirst
                      labavis.Caption = "METAS: "
                      Do While Not data_lin.Recordset.EOF
                         labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                         data_lin.Recordset.MoveNext
                      Loop
                      data_lin.Recordset.MovePrevious
                      Xfecvolver = data_lin.Recordset("fecha") + 36
                      labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                   Else
                      labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.1er.AÑO DE VIDA-"
                   End If
                   data_lin.Recordset.Close
                Else
                   If Val(labedad.Caption) = 2 And Val(labunie.Caption) >= 0 Then
                      data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190004 & " order by fecha"
                      data_lin.Refresh
                      If data_lin.Recordset.RecordCount > 0 Then
                         data_lin.Recordset.MoveFirst
                         labavis.Caption = "METAS: "
                         Do While Not data_lin.Recordset.EOF
                            labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                            data_lin.Recordset.MoveNext
                         Loop
                         data_lin.Recordset.MovePrevious
                         Xfecvolver = data_lin.Recordset("fecha") + 91
                         labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                      Else
                         labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.2do.AÑO DE VIDA-"
                      End If
                      data_lin.Recordset.Close
                   Else
                      'cambiar el codigo de facturación
                      If Val(labedad.Caption) = 3 And Val(labunie.Caption) >= 0 Then
                         data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190004 & " order by fecha"
                         data_lin.Refresh
                         If data_lin.Recordset.RecordCount > 0 Then
                            data_lin.Recordset.MoveFirst
                            labavis.Caption = "METAS: "
                            Do While Not data_lin.Recordset.EOF
                               labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                               data_lin.Recordset.MoveNext
                            Loop
                            data_lin.Recordset.MovePrevious
                            Xfecvolver = data_lin.Recordset("fecha") + 122
                            labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                         Else
                            labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.3er.AÑO DE VIDA-"
                         End If
                         data_lin.Recordset.Close
                      Else
                         'cambiar el codigo de facturación
                         If Val(labedad.Caption) = 4 And Val(labunie.Caption) >= 0 Then
                            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190030 & " order by fecha"
                            data_lin.Refresh
                            If data_lin.Recordset.RecordCount > 0 Then
                               data_lin.Recordset.MoveFirst
                               labavis.Caption = "METAS: "
                               Do While Not data_lin.Recordset.EOF
                                  labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                  data_lin.Recordset.MoveNext
                               Loop
                               data_lin.Recordset.MovePrevious
                               Xfecvolver = data_lin.Recordset("fecha") + 183
                               labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                            Else
                               labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.4to.AÑO DE VIDA-"
                            End If
                            data_lin.Recordset.Close
                         Else
                            'cambiar el codigo de facturación
                            If Val(labedad.Caption) = 5 And Val(labunie.Caption) >= 0 Then
                               data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190031 & " order by fecha"
                               data_lin.Refresh
                               If data_lin.Recordset.RecordCount > 0 Then
                                  data_lin.Recordset.MoveFirst
                                  labavis.Caption = "METAS: "
                                  Do While Not data_lin.Recordset.EOF
                                     labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                     data_lin.Recordset.MoveNext
                                  Loop
                                  data_lin.Recordset.MovePrevious
                                  Xfecvolver = data_lin.Recordset("fecha") + 183
                                  labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                               Else
                                  labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.5to.AÑO DE VIDA-"
                               End If
                               data_lin.Recordset.Close
                            Else
                               If Val(labedad.Caption) >= 15 And Val(labedad.Caption) <= 100 Then
                                  If cbosexo.Text = "FEMENINO" Then
                                     data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod =" & 190010 & " order by fecha"
                                     data_lin.Refresh
                                     If data_lin.Recordset.RecordCount > 0 Then
                                        data_lin.Recordset.MoveFirst
                                        labavis.Caption = "METAS: "
                                        Do While Not data_lin.Recordset.EOF
                                           labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                           data_lin.Recordset.MoveNext
                                        Loop
                                        data_lin.Recordset.MovePrevious
                                        Xfecvolver = data_lin.Recordset("fecha") + 365
                                        labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                                     Else
                                        labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -PESQUISA V.DOMESTICA-"
                                     End If
                                     data_lin.Recordset.Close
                                  End If
                               Else
                                  If Val(labedad.Caption) >= 12 And Val(labedad.Caption) <= 19 Then
                                     data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod in (190011,190012)"
                                     data_lin.Refresh
                                     If data_lin.Recordset.RecordCount > 0 Then
                                        data_lin.Recordset.MoveFirst
                                        labavis.Caption = "METAS: "
                                        Do While Not data_lin.Recordset.EOF
                                           labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                           data_lin.Recordset.MoveNext
                                        Loop
                                     Else
                                        labavis.Caption = "METAS: " & "DEBE REALIZAR META 2 -MÉDICO DE REF. 12 A 19AÑOS"
                                     End If
                                     data_lin.Recordset.Close
                                  Else
                                     If Val(labedad.Caption) >= 45 And Val(labedad.Caption) <= 64 Then
                                        data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod in (190013,190014)"
                                        data_lin.Refresh
                                        If data_lin.Recordset.RecordCount > 0 Then
                                           data_lin.Recordset.MoveFirst
                                           labavis.Caption = "METAS: "
                                           Do While Not data_lin.Recordset.EOF
                                              labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                              data_lin.Recordset.MoveNext
                                           Loop
                                        Else
                                           labavis.Caption = "METAS: " & "DEBE REALIZAR META 2 -MÉDICO DE REF. 45 A 64AÑOS"
                                        End If
                                        data_lin.Recordset.Close
                                        If Val(labedad.Caption) > 50 Then
                                           data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod in (30063,30067)"
                                           data_lin.Refresh
                                           If data_lin.Recordset.RecordCount > 0 Then
                                              data_lin.Recordset.MoveFirst
                                              Do While Not data_lin.Recordset.EOF
                                                 labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                                 data_lin.Recordset.MoveNext
                                              Loop
                                           Else
                                              labavis.Caption = labavis.Caption & vbCrLf & "FALTA FECATEST"
                                           End If
                                           data_lin.Recordset.Close
                                        End If
                                     Else
                                        If Val(labedad.Caption) >= 65 And Val(labedad.Caption) <= 74 Then
                                           data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod in (190018,190019,190023,190027) order by fecha"
                                           data_lin.Refresh
                                           If data_lin.Recordset.RecordCount > 0 Then
                                              data_lin.Recordset.MoveFirst
                                              labavis.Caption = "METAS: "
                                              Do While Not data_lin.Recordset.EOF
                                                 labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                                 data_lin.Recordset.MoveNext
                                              Loop
                                              data_lin.Recordset.MovePrevious
                                              Xfecvolver = data_lin.Recordset("fecha") + 275
                                              labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                                           Else
                                              labavis.Caption = "METAS: " & "DEBE REALIZAR META 3 -MÉDICO DE REF. 65 A 74AÑOS"
                                           End If
                                           data_lin.Recordset.Close
                                        Else
                                           If Val(labedad.Caption) >= 75 And Val(labedad.Caption) <= 115 Then
                                              data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & txt_mat.Caption & " and cod_prod in (190020,190021) order by fecha"
                                              data_lin.Refresh
                                              If data_lin.Recordset.RecordCount > 0 Then
                                                 data_lin.Recordset.MoveFirst
                                                 labavis.Caption = "METAS: "
                                                 Do While Not data_lin.Recordset.EOF
                                                    labavis.Caption = labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                                    data_lin.Recordset.MoveNext
                                                 Loop
                                                 data_lin.Recordset.MovePrevious
                                                 Xfecvolver = data_lin.Recordset("fecha") + 91
                                                 labavis.Caption = labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                                              Else
                                                 labavis.Caption = "METAS: " & "DEBE REALIZAR META 3 -MÉDICO DE REF. >75 AÑOS"
                                              End If
                                              data_lin.Recordset.Close
                                           Else

                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
             End If
          Else
          
          End If
          If Trim(labavis.Caption) = "" Then
             labavis.Caption = Xelaviso
          Else
             labavis.Caption = labavis.Caption & vbCrLf & Xelaviso
          End If
       End If
       
       frmabm.MousePointer = 0
    End If
Else
    If KeyAscii = 13 Then
       btn_busca.SetFocus
    End If
End If

Exit Sub

Quepasaal:
          If Err.Number = 3157 Then
             MsgBox "Error al ingresar"
          Else
             MsgBox "Error al conectar"
          End If
          
End Sub


Private Sub txt_ced_GotFocus()
   If txt_nomcnv.Text = "" Then
      MsgBox "No seleccionó convenio, VERIFIQUE!", vbCritical
   End If
   
End Sub

Private Sub txt_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_ced2.Visible = True Then
      txt_ced2.SetFocus
   Else
      txt_nac.SetFocus
   End If
End If

End Sub

Private Sub txt_ced_LostFocus()
If XAlta = 1 Then
    If txt_ced.Text <> "" Then
       If txt_ced.Text > 0 Then
'          Data1.Recordset.FindFirst "cl_cedula =" & txt_ced.Text
          Data1.RecordSource = "Select * from clientes where cl_cedula =" & txt_ced.Text
          Data1.Refresh
          If Data1.Recordset.RecordCount > 0 Then
             MsgBox "Ya existe socio con ésta cédula", vbInformation, "Clientes"
             btn_cance_Click
             'txt_apellid.SetFocus
          End If
       End If
    End If
End If

End Sub

Private Sub txt_ced2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nac.SetFocus
End If

End Sub

Private Sub txt_cedtarj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codtarj.SetFocus
End If

End Sub

Private Sub txt_codcnv_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txt_apellid.SetFocus
End If


End Sub

Private Sub txt_codcnv_LostFocus()
Dim Xsicnv As String
Xsicnv = "SI"
If Trim(txt_codcnv.Text) <> "" Then
   ControlproxEmi
   If WElusuario = "MCOSTA" Or XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Or WElusuario = "MARIAROSA" Or WElusuario = "MSANCHEZ" Then
      data_clicnv.RecordSource = "Select * from convenio where cnv_codigo ='" & Trim(txt_codcnv.Text) & "' and cnv_umpago not in (1)"
      data_clicnv.Refresh
      If IsNull(data_clicnv.Recordset("cnv_fbaja")) = False Then
         MsgBox "Convenio de BAJA, Verifique!!", vbInformation
      End If
   Else
      data_clicnv.RecordSource = "Select * from convenio where cnv_alta ='" & Trim(Xsicnv) & "' and cnv_fbaja is null and cnv_codigo ='" & Trim(txt_codcnv.Text) & "' and cnv_umpago not in (1)"
      data_clicnv.Refresh
   End If
'   data_clicnv.Recordset.FindFirst "cnv_codigo = '" & Trim(txt_codcnv.Text) & "'"
'   If Not data_clicnv.Recordset.NoMatch Then
   If data_clicnv.Recordset.RecordCount > 0 Then
      txt_codcnv.Text = data_clicnv.Recordset("cnv_codigo")
      txt_nomcnv.Text = data_clicnv.Recordset("cnv_desc")
      If IsNull(data_clicnv.Recordset("cnv_entre")) = False Then
         If Trim(data_clicnv.Recordset("cnv_entre")) <> "" Then
            If Val(data_clicnv.Recordset("cnv_cuenta")) = Val(txt_mat.Caption) Then
               t_rs.Text = data_clicnv.Recordset("cnv_entre")
            Else
               t_rs.Text = ""
            End If
         Else
            t_rs.Text = ""
         End If
      Else
         t_rs.Text = ""
      End If
      txt_apellid.SetFocus
   Else
      Xconv = txt_codcnv.Text
      t_rs.Text = ""
      frm_buscacnv.Show vbModal
   End If
Else
   Xconv = txt_codcnv.Text
   frm_buscacnv.Show vbModal
End If

End Sub

Private Sub txt_codcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbonomcob.SetFocus
End If

End Sub

Private Sub txt_codcob_LostFocus()
If txt_codcob.Text <> "" Then
   data_cobrador.Recordset.FindFirst "cb_numero =" & txt_codcob.Text
   If Not data_cobrador.Recordset.NoMatch Then
      txt_codcob.Text = data_cobrador.Recordset("cb_numero")
      cbonomcob.Text = data_cobrador.Recordset("cb_nombre")
      cbopago.SetFocus
   Else
      cbonomcob.SetFocus
   End If
Else
   txt_codcob.Text = 0
   cbonomcob.Text = "*TODOS"
   cbonomcob.SetFocus
End If

End Sub

Private Sub txt_codemisor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbotarj.SetFocus
End If

End Sub

Private Sub txt_codemisor_LostFocus()
If txt_codemisor.Text <> "" Then
   data_tarjetas.Recordset.FindFirst "numero =" & txt_codemisor.Text
   If Not data_tarjetas.Recordset.NoMatch Then
      txt_codemisor.Text = data_tarjetas.Recordset("numero")
      cbotarj.Text = data_tarjetas.Recordset("nombre")
      txt_cedtarj.SetFocus
   Else
      cbotarj.SetFocus
   End If
Else
   cbotarj.SetFocus
End If

End Sub

Private Sub txt_codpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbonompro.SetFocus
End If

End Sub

Private Sub txt_codpro_LostFocus()
If txt_codpro.Text <> "" Then
   data_promo.Recordset.FindFirst "idfunc =" & txt_codpro.Text
   If Not data_promo.Recordset.NoMatch Then
      cbonompro.Text = data_promo.Recordset("nombre")
   Else
      txt_codpro.Text = 799
      cbonompro.Text = "*TODOS"
      cbonompro.SetFocus
   End If
Else
   txt_codpro.Text = 799
   cbonompro.Text = "*TODOS"
   cbonompro.SetFocus
End If

End Sub

Private Sub txt_codtarj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codemisor.SetFocus
End If

End Sub

Private Sub txt_codzon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbolocalid.SetFocus
End If

End Sub

Private Sub txt_codzon_LostFocus()
If txt_codzon.Text <> "" Then
   data_zonas.Recordset.FindFirst "zo_grupo =" & txt_codzon.Text
   If Not data_zonas.Recordset.NoMatch Then
      cbolocalid.Text = data_zonas.Recordset("zo_nombre")
      txt_telef.SetFocus
   Else
      MsgBox "No existe zona", vbCritical, "Mensaje"
'      cbolocalid.SetFocus
   End If
End If

End Sub

Private Sub txt_conmut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Frame2.Enabled = True Then
      txt_fecing.SetFocus
   Else
      btn_graba.SetFocus
   End If
End If

End Sub

Private Sub txt_diacob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cbopago.Text = "Debito Automatico" Then
      txt_nomtarj.SetFocus
   Else
      btn_graba.SetFocus
   End If
End If
End Sub

Private Sub txt_direcc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_direcc2.SetFocus
End If

End Sub

Private Sub txt_direcc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codzon.SetFocus
End If
End Sub

Private Sub txt_fecbaj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codpro.SetFocus
End If

End Sub

Private Sub txt_fecing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codpro.SetFocus
End If

End Sub

Private Sub txt_fecing_LostFocus()
If txt_fecing.Text = "__/__/____" Then
   MsgBox "Verifique FECHA de ingreso, no puede ser vacía!", vbCritical
   txt_fecing.Text = Date
End If

End Sub

Private Sub txt_matmut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_conmut.SetFocus
End If

End Sub

Private Sub txt_nac_KeyPress(KeyAscii As Integer)
Dim Xed As Long
If KeyAscii = 13 Then
   If txt_nac.Text <> "__/__/____" Then
   
   Else
      labedad.Caption = 0
      labunie.Caption = 0
      labdias.Caption = 0
   End If
   txt_direcc1.SetFocus
End If

If Not IsDate(txt_nac.Text) Then
'   MsgBox "Digite una fecha válida"
Else
   CalculaEdad (txt_nac.Text)
End If


End Sub

Private Sub txt_nomc_Change()

End Sub

Private Sub txt_nomc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 Then
   txt_ced.SetFocus
End If

End Sub

Private Sub txt_nomcnv_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'If KeyAscii = 13 Then
'   txt_apellid.SetFocus
'End If
End Sub

Private Sub txt_nomtarj_KeyPress(KeyAscii As Integer)
Dim XX As String
If KeyAscii = 13 Then
   txt_cedtarj.SetFocus
End If

End Sub

Private Sub txt_nrotarj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_vence.SetFocus
End If

End Sub

Private Sub txt_telef_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   t_cel.SetFocus
End If

End Sub

Private Sub txt_telef_LostFocus()
If Trim(txt_telef.Text) = "" Then
   MsgBox "Falta ingresar datos en teléfono.", vbCritical
Else
   txt_telef.Text = Trim(txt_telef.Text)
End If

      
End Sub

Private Sub txt_vence_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btn_graba.SetFocus
End If

End Sub

Public Function vercli()
If IsNull(data_clientes.Recordset("cl_fultvta")) = False Then
   If IsNull(data_clientes.Recordset("cl_tipocli")) = False Then
      Image1.Visible = True
   Else
      Image1.Visible = False
   End If
Else
   Image1.Visible = False
End If
If Image1.Visible = False Then
   If IsNull(data_clientes.Recordset("cl_fultpag")) = False Then
      Image1.Visible = True
   End If
End If
If data_clientes.Recordset("estado") <> "" Then
   If data_clientes.Recordset("estado") = 2 Or data_clientes.Recordset("estado") = 3 Then
      labestado.Caption = "BAJA"
   Else
      labestado.Caption = "ACTIVO"
   End If
Else
   If data_clientes.Recordset("fecha_baja") <> "" Then
      labestado.Caption = "BAJA"
   Else
      labestado.Caption = "ACTIVO"
   End If
End If
If data_clientes.Recordset("cl_codigo") <> "" Then
   txt_mat.Caption = data_clientes.Recordset("cl_codigo")
Else
   txt_mat.Caption = ""
End If
If IsNull(data_clientes.Recordset("cl_codconv")) = True Then
   MsgBox "Verifique el convenio", vbCritical, "Mensaje"
   txt_codcnv.Text = ""
Else
   txt_codcnv.Text = data_clientes.Recordset("cl_codconv")
End If
txt_nomcnv.Enabled = True
If IsNull(data_clientes.Recordset("cl_nomconv")) = True Then
   txt_nomcnv.Text = ""
Else
   txt_nomcnv.Text = data_clientes.Recordset("cl_nomconv")
End If
txt_nomcnv.Enabled = False
If data_clientes.Recordset("cl_apellid") <> "" Then
   txt_apellid.Text = data_clientes.Recordset("cl_apellid")
Else
   txt_apellid.Text = ""
End If
data_cnvmut.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_codcnv.Text & "'"
data_cnvmut.Refresh
If data_cnvmut.Recordset.RecordCount > 0 Then
   If IsNull(data_cnvmut.Recordset("cnv_entre")) = False Then
      If Trim(data_cnvmut.Recordset("cnv_entre")) <> "" Then
         If Val(data_cnvmut.Recordset("cnv_cuenta")) = Val(txt_mat.Caption) Then
            t_rs.Text = data_cnvmut.Recordset("cnv_entre")
         Else
            t_rs.Text = ""
         End If
      Else
         t_rs.Text = ""
      End If
   Else
      t_rs.Text = ""
   End If
End If
If IsNull(data_clientes.Recordset("cl_ruc")) = False Then
   If data_clientes.Recordset("cl_ruc") <> "" Then
      t_otrocnv.Text = data_clientes.Recordset("cl_ruc")
   Else
      t_otrocnv.Text = ""
   End If
Else
   t_otrocnv.Text = ""
End If
If IsNull(data_clientes.Recordset("cl_tipoced")) = False Then
   cbotipoced.ListIndex = data_clientes.Recordset("cl_tipoced")
   If data_clientes.Recordset("cl_tipoced") = 0 Then
      txt_ced2.Visible = True
      If data_clientes.Recordset("cl_cedula") <> "" Then
         txt_ced.Text = data_clientes.Recordset("cl_cedula")
      Else
         txt_ced.Text = ""
      End If
      If data_clientes.Recordset("cl_codced") <> "" Then
         txt_ced2.Text = data_clientes.Recordset("cl_codced")
      Else
         txt_ced2.Text = 0
      End If
   Else
      txt_ced2.Visible = False
      If data_clientes.Recordset("cl_cedula") <> "" Then
         txt_ced.Text = data_clientes.Recordset("cl_cedula")
      Else
         txt_ced.Text = ""
      End If
      txt_ced2.Text = 0
   End If
Else
   cbotipoced.ListIndex = 0
   txt_ced2.Visible = True
   If data_clientes.Recordset("cl_cedula") <> "" Then
      txt_ced.Text = data_clientes.Recordset("cl_cedula")
   Else
      txt_ced.Text = ""
   End If
   If data_clientes.Recordset("cl_codced") <> "" Then
      txt_ced2.Text = data_clientes.Recordset("cl_codced")
   Else
      txt_ced2.Text = 0
   End If
End If
If IsNull(data_clientes.Recordset("cl_dpto")) = False Then
   t_cel.Text = data_clientes.Recordset("cl_dpto")
Else
   t_cel.Text = ""
End If
If IsNull(data_clientes.Recordset("cl_codruta")) = False Then
   t_ruta.Text = data_clientes.Recordset("cl_codruta")
Else
   t_ruta.Text = ""
End If
If IsNull(data_clientes.Recordset("cl_referen")) = False Then
   t_correo.Text = data_clientes.Recordset("cl_referen")
Else
   t_correo.Text = ""
End If
If IsNull(data_clientes.Recordset("cl_fnac")) = False Then
   txt_nac.Text = Format(data_clientes.Recordset("cl_fnac"), "dd/mm/yyyy")
   If Not IsDate(txt_nac.Text) Then
   Else
      CalculaEdad (txt_nac.Text)
   End If
Else
   txt_nac.Text = "__/__/____"
   labedad.Caption = ""
   labunie.Caption = ""
   labdias.Caption = ""
End If

If data_clientes.Recordset("cl_ultmesp") <> "" Then
   labump.Caption = data_clientes.Recordset("cl_ultmesp")
Else
   labump.Caption = ""
End If
If data_clientes.Recordset("cl_ultanop") <> "" Then
   If data_clientes.Recordset("cl_ultanop") = 0 Then
      labuap.Caption = data_clientes.Recordset("cl_ultanop")
      Label7.Caption = ""
   Else
      Label7.Caption = "/"
      labuap.Caption = data_clientes.Recordset("cl_ultanop")
   End If
Else
   labuap.Caption = ""
   Label7.Caption = ""
End If
If data_clientes.Recordset("cl_atrasoa") <> "" Then
   labatra.Caption = data_clientes.Recordset("cl_atrasoa")
Else
   labatra.Caption = ""
End If
If data_clientes.Recordset("saldo_cc") <> "" Then
   labdeudap.Caption = data_clientes.Recordset("saldo_cc")
Else
   labdeudap.Caption = ""
End If
If data_clientes.Recordset("cl_direcci") <> "" Then
   txt_direcc1.Text = data_clientes.Recordset("cl_direcci")
Else
   txt_direcc1.Text = ""
End If
If data_clientes.Recordset("cl_entre") <> "" Then
   txt_direcc2.Text = data_clientes.Recordset("cl_entre")
Else
   txt_direcc2.Text = ""
End If
If data_clientes.Recordset("cl_grupo") <> "" Then
   txt_codzon.Text = data_clientes.Recordset("cl_grupo")
Else
   txt_codzon.Text = 0
End If
If data_clientes.Recordset("cl_zona") <> "" Then
   cbolocalid.Text = data_clientes.Recordset("cl_zona")
Else
   cbolocalid.Text = ""
End If
If data_clientes.Recordset("cl_telefon") <> "" Then
   txt_telef.Text = data_clientes.Recordset("cl_telefon")
Else
   txt_telef.Text = ""
End If
If data_clientes.Recordset("cl_dircobr") <> "" Then
   txt_dircob.Text = data_clientes.Recordset("cl_dircobr")
Else
   txt_dircob.Text = ""
End If
If IsNull(data_clientes.Recordset("cl_nombre")) = False Then
   txt_conmut.Text = data_clientes.Recordset("cl_nombre")
End If
If data_clientes.Recordset("cl_socmnom") <> "" Then
   cbomutual.Text = data_clientes.Recordset("cl_socmnom")
Else
   cbomutual.Text = ""
End If
If data_clientes.Recordset("cl_nrosocm") <> "" Then
   txt_matmut.Text = data_clientes.Recordset("cl_nrosocm")
Else
   txt_matmut.Text = ""
End If

If data_clientes.Recordset("cl_sexo") = 2 Then
   cbosexo.Text = "FEMENINO"
Else
   cbosexo.Text = "MASCULINO"
End If
If data_clientes.Recordset("cl_fecing") <> "" Then
   txt_fecing.Text = Format(data_clientes.Recordset("cl_fecing"), "dd/mm/yyyy")
Else
   txt_fecing.Text = "__/__/____"
End If
If data_clientes.Recordset("fecha_baja") <> "" Then
   txt_fecbaj.Text = Format(data_clientes.Recordset("fecha_baja"), "dd/mm/yyyy")
Else
   txt_fecbaj.Text = "__/__/____"
End If
If data_clientes.Recordset("cl_nrovend") <> "" Then
   txt_codpro.Text = data_clientes.Recordset("cl_nrovend")
Else
   txt_codpro.Text = ""
End If
If data_clientes.Recordset("cl_nomvend") <> "" Then
   cbonompro.Text = data_clientes.Recordset("cl_nomvend")
Else
   cbonompro.Text = ""
End If
If IsNull(data_clientes.Recordset("idpromos")) = False Then
   labidpromo.Caption = data_clientes.Recordset("idpromos")
   If Val(labidpromo.Caption) > 0 Then
      BuscaPromosId
   Else
      cbopromos.Text = ""
   End If
Else
   labidpromo.Caption = 0
   cbopromos.Text = ""
End If
If IsNull(data_clientes.Recordset("mesproxemi")) = False Then
   t_pmemi.Text = data_clientes.Recordset("mesproxemi")
   t_paemi.Text = data_clientes.Recordset("anoproxemi")
Else
   t_pmemi.Text = 0
   t_paemi.Text = 0
End If
If data_clientes.Recordset("cl_nrocobr") <> "" Then
   txt_codcob.Text = data_clientes.Recordset("cl_nrocobr")
Else
   txt_codcob.Text = ""
End If
If data_clientes.Recordset("cl_nomcobr") <> "" Then
   cbonomcob.Text = data_clientes.Recordset("cl_nomcobr")
Else
   cbonomcob.Text = ""
End If
If IsNull(data_clientes.Recordset("cl_descpag")) = True Then
   cbopago.Text = "Abono Mensual"
Else
   If UCase(data_clientes.Recordset("cl_descpag")) = "DEBITO AUTOMATICO" Then
      cbopago.Text = "Debito Automatico"
   Else
      cbopago.Text = "Abono Mensual"
   End If
End If
If data_clientes.Recordset("cl_diacobr") <> "" Then
   txt_diacob.Text = data_clientes.Recordset("cl_diacobr")
Else
   txt_diacob.Text = ""
End If
If data_clientes.Recordset("tit_tarj") <> "" Then
   txt_nomtarj.Text = data_clientes.Recordset("tit_tarj")
Else
   txt_nomtarj.Text = ""
End If
If data_clientes.Recordset("cl_nrotarj") <> "" Then
   txt_nrotarj.Text = data_clientes.Recordset("cl_nrotarj")
Else
   txt_nrotarj.Text = ""
End If
If data_clientes.Recordset("ci_tarj") <> "" Then
   txt_cedtarj.Text = data_clientes.Recordset("ci_tarj")
Else
   txt_cedtarj.Text = ""
End If
If data_clientes.Recordset("codcitarj") <> "" Then
   txt_codtarj.Text = data_clientes.Recordset("codcitarj")
Else
   txt_codtarj.Text = ""
End If
If data_clientes.Recordset("cl_tjemi_c") <> "" Then
   txt_codemisor.Text = data_clientes.Recordset("cl_tjemi_c")
Else
   txt_codemisor.Text = ""
End If
If data_clientes.Recordset("cl_tjemi_n") <> "" Then
   cbotarj.Text = data_clientes.Recordset("cl_tjemi_n")
Else
   cbotarj.Text = ""
End If
If data_clientes.Recordset("cl_tj_venc") <> "" Then
   txt_vence.Text = Format(data_clientes.Recordset("cl_tj_venc"), "dd/mm/yyyy")
Else
   txt_vence.Text = "__/__/____"
End If
If IsNull(data_clientes.Recordset("cl_decuota")) = False Then
   If data_clientes.Recordset("cl_decuota") = 1 Then
      Option1.Value = True
   Else
      If data_clientes.Recordset("cl_decuota") = 2 Then
         Option2.Value = True
      Else
         If data_clientes.Recordset("cl_decuota") = 3 Then
            Option3.Value = True
         Else
            If data_clientes.Recordset("cl_decuota") = 4 Then
               Option4.Value = True
            Else
               If data_clientes.Recordset("cl_decuota") = 5 Then
                  Option5.Value = True
               Else
                  Option1.Value = False
                  Option2.Value = False
                  Option3.Value = False
                  Option4.Value = False
                  Option5.Value = False
               End If
            End If
         End If
      End If
   End If
End If

If IsNull(data_clientes.Recordset("fecha_reac")) = False Then
   mfcarta.Text = Format(data_clientes.Recordset("fecha_reac"), "dd/mm/yyyy")
Else
   mfcarta.Text = "__/__/____"
End If
If IsNull(data_clientes.Recordset("saldo_chc2")) = False Then
   cbosrv.ListIndex = data_clientes.Recordset("saldo_chc2")
Else
   cbosrv.ListIndex = -1
End If

   

End Function

Public Function quienes()
If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Then
   btn_modi.Enabled = True
   btn_baja.Enabled = True
Else
   If XWeltipoU = "USUARIOS" Then
      btn_modi.Enabled = True
      btn_baja.Enabled = False
   Else
      btn_modi.Enabled = False
      btn_baja.Enabled = False
   End If
End If

End Function

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

   If Month(FNaci) = 1 Or Month(FNaci) = 3 Or Month(FNaci) = 5 Or _
      Month(FNaci) = 7 Or Month(FNaci) = 8 Or Month(FNaci) = 10 Or _
      Month(FNaci) = 12 Then
      newday = Day(CDate(FAct)) + 31
   Else
      If Month(FNaci) = 2 Then
         newday = Day(CDate(FAct)) + 28
      Else
         newday = Day(CDate(FAct)) + 30
      End If
   End If
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
   labedad.Caption = Anios
   If Month(Date) = Month(FNaci) Then
      If Day(Date) > Day(FNaci) Then
         Meses = Meses
      Else
         If Day(Date) = Day(FNaci) Then
            Meses = 0
         Else
            Meses = 11
         End If
      End If
   End If
   labunie.Caption = Meses
   labdias.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   labedad.Caption = 0
   labunie.Caption = 0
   labdias.Caption = 0
End If

End Sub

Public Sub Veoladeuda(ByVal Xmatricula As Long)

Dim Xsubt As Double
Dim Xcant As Long
Dim Xmes, Xano As Integer
Dim Xsqldeuda As String
Dim Xrecdeuda As New ADODB.Recordset
Xcant = 0
Xsubt = 0
Xmes = 0
Xano = 0
ConectarBDDeuda
ConbdSappDeu.Open
             
Xsqldeuda = "Select * from deudas where cliente =" & Xmatricula & " and fecha_pago is null order by ano,mes"
With Xrecdeuda
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqldeuda, ConbdSappDeu, , , adCmdText
End With

'Set Basesapp = Sesionsapp.OpenDatabase(App.Path & "\sapp.mdb")

'Set Recdeudas = Basesapp.OpenRecordset("Select * from deudas where cliente =" & Xmatricula & " and fecha_pago is null order by ano,mes")

If Xrecdeuda.RecordCount > 0 Then
   Xrecdeuda.MoveFirst
   Do While Not Xrecdeuda.EOF
      If Xrecdeuda("mes") = 0 Then
         Xsubt = Xsubt + Xrecdeuda("total")
      Else
         Xsubt = Xsubt + Xrecdeuda("total")
         If Xmes = 0 Then
            Xmes = Xrecdeuda("mes")
            Xano = Xrecdeuda("ano")
         End If
         Xcant = Xcant + 1
      End If
      Xrecdeuda.MoveNext
   Loop
   labump.Caption = Xmes
   labuap.Caption = Xano
   labatra.Caption = Xcant
   labdeudap.Caption = Format(Xsubt, "0.00")
   Xrecdeuda.Close
   
   Xsqldeuda = "Select * from deudas where cliente =" & Xmatricula & " and mes not in (0) and fecha_pago is not null order by fecha DESC limit 1"
   With Xrecdeuda
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqldeuda, ConbdSappDeu, , , adCmdText
   End With
   If Xrecdeuda.RecordCount > 0 Then
      Xmes = Xrecdeuda("mes")
      Xano = Xrecdeuda("ano")
      labump.Caption = Xmes
      labuap.Caption = Xano
   Else
      Xmes = 0
      Xano = 0
      labump.Caption = Xmes
      labuap.Caption = Xano
   End If
   Xrecdeuda.Close
   ConbdSappDeu.Close
Else
   labatra.Caption = 0
   labdeudap.Caption = 0
   Xrecdeuda.Close
   Xsqldeuda = "Select * from deudas where cliente =" & Xmatricula & " and mes not in (0) and fecha_pago is not null order by fecha DESC limit 1"
   With Xrecdeuda
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqldeuda, ConbdSappDeu, , , adCmdText
   End With
   If Xrecdeuda.RecordCount > 0 Then
      Xmes = Xrecdeuda("mes")
      Xano = Xrecdeuda("ano")
      labump.Caption = Xmes
      labuap.Caption = Xano
   Else
      Xmes = 0
      Xano = 0
      labump.Caption = Xmes
      labuap.Caption = Xano
   End If
   Xrecdeuda.Close
   ConbdSappDeu.Close
End If

End Sub

Private Function Lan() As Boolean
   
   Call InternetGetConnectedState(dwflags, 0&)
   Lan = dwflags And INTERNET_CONNECTION_LAN
End Function
Private Function Online() As Boolean
   Online = InternetGetConnectedState(0&, 0&)
End Function

Public Function ConectarBDDeuda()
ConbdSappDeu.ConnectionString = "dsn=" & Xconexrmt

End Function


Public Sub VerPromocion(ByVal Xmatricula As Long)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & Xmatricula
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   Label42.Caption = "Tiene promo X" & Xrecclii.RecordCount
   t_ruta.Enabled = True
Else
   Label42.Caption = ""
   t_ruta.Enabled = True
End If
Xrecclii.Close
ConbdSapp.Close


End Sub

Public Sub VerPromosiono(ByVal Xcate As String)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from convenio where cnv_codigo ='" & Trim(Xcate) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cnv_grupo")) = False Then
      If Xrecclii("cnv_grupo") <> "" Then
         If Xrecclii("cnv_grupo") = "SEMM" Or Xrecclii("cnv_grupo") = "CASH" Or _
            Xrecclii("cnv_grupo") = "CPS" Or Xrecclii("cnv_grupo") = "CASMU" Then
            MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
         Else
            If Xcate = "CCNOS" Or Xcate = "SMIN" Or Xcate = "HEVANO" Or Xcate = "CASANO" Or Xcate = "GANOS" Or _
               Xcate = "UNIVS" Or Xcate = "SMINR" Then
               MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
            End If
         End If
      Else
         If Xcate = "EMEFES" Or Xcate = "SAFES" Or Xcate = "SPFES" Or Xcate = "SPFFES" Then
            MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
         Else
            If IsNull(Xrecclii("cnv_precio")) = True Then
               MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
            Else
              If Xrecclii("cnv_precio") <= 0 Then
                 MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
              End If
            End If
         End If
      End If
   Else
         If Xcate = "EMEFES" Or Xcate = "SAFES" Or Xcate = "SPFES" Or Xcate = "SPFFES" Then
            MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
         Else
            If IsNull(Xrecclii("cnv_precio")) = True Then
               MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
            Else
              If Xrecclii("cnv_precio") <= 0 Then
                 MsgBox "Verifique si la categoría del titular puede ingresar en la promoción!", vbInformation
              End If
            End If
         End If
   
   End If
End If
         
Xrecclii.Close
ConbdSapp.Close


End Sub

Public Sub VerPromoCli(ByVal Xmatricula As Long)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select cl_codruta,cl_codigo,cl_codconv,estado from clientes where cl_codigo =" & Xmatricula & " and estado =" & 1
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   Xconvprom = Xrecclii("cl_codconv")
Else
   Xconvprom = "PART"
   MsgBox "No se encuentra socio activo para ésta promoción. VERIFIQUE!!", vbCritical
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub CargaPromos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from promocion_gpo"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
cbopromos.Clear
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      cbopromos.AddItem Xrecclii("descrip")
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub BuscaPromos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from promocion_gpo where descrip ='" & cbopromos.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labidpromo.Caption = Xrecclii("id")
Else
   MsgBox "No se encuentra promoción. Verifique!", vbCritical
   labidpromo.Caption = 0
   cbopromos.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub BuscaPromosId()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from promocion_gpo where id =" & Val(labidpromo.Caption)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   cbopromos.Text = Xrecclii("descrip")
Else
   MsgBox "No se encuentra promoción. Verifique!", vbCritical
   cbopromos.Text = ""
   labidpromo.Caption = 0
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub ControlproxEmi()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xmmpemi, Xaapemi As Integer
Dim Xarmoelmesd, Xarmoelmesh As String


If Month(Date) > 9 Then
   Xarmoelmesh = "01/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
Else
   Xarmoelmesh = "01/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
End If

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
If Month(Date) = 12 Then
   Xmmpemi = 1
   Xaapemi = Year(Date) + 1
Else
   Xmmpemi = Month(Date) + 1
   Xaapemi = Year(Date)
End If

Xsqlpromo = "Select * from convenio where cnv_codigo ='" & txt_codcnv.Text & "' and cnv_emite ='" & "SI" & "' and cnv_cant_r in (1,2) and cnv_hasta >='" & Format(Date, "yyyy/mm/dd") & "'"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
'SI ES >25 DEL MES/AÑO ACTUAL
   If txt_fecing.Text = "__/__/____" Then
      txt_fecing.Text = Date
   End If
   If Day(txt_fecing.Text) > 25 And Year(txt_fecing.Text) = Xaapemi Then
        If t_pmemi.Text <> "" Then
           If Val(t_pmemi.Text) <= 0 Then
              If Xmmpemi = 12 Then
                 t_pmemi.Text = 1
                 t_paemi.Text = Xaapemi + 1
              Else
                 t_pmemi.Text = Xmmpemi + 1
                 t_paemi.Text = Xaapemi
              End If
           Else
              If t_pmemi.Text > 9 Then
                 Xarmoelmesd = "01/" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              Else
                 Xarmoelmesd = "01/0" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              End If
              If Format(Xarmoelmesd, "yyyy/mm/dd") <= Format(Xarmoelmesh, "yyyy/mm/dd") Then
                 If Xmmpemi = 12 Then
                    t_pmemi.Text = 1
                    t_paemi.Text = Xaapemi + 1
                 Else
                    t_pmemi.Text = Xmmpemi + 1
                    t_paemi.Text = Xaapemi
                 End If
              End If
           End If
        Else
           If Xmmpemi = 12 Then
              t_pmemi.Text = 1
              t_paemi.Text = Xaapemi + 1
           Else
              t_pmemi.Text = Xmmpemi + 1
              t_paemi.Text = Xaapemi
           End If
        End If
   Else
        If t_pmemi.Text <> "" Then
           If Val(t_pmemi.Text) <= 0 Then
              t_pmemi.Text = Xmmpemi
              t_paemi.Text = Xaapemi
           Else
              If t_pmemi.Text > 9 Then
                 Xarmoelmesd = "01/" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              Else
                 Xarmoelmesd = "01/0" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              End If
              If Format(Xarmoelmesd, "yyyy/mm/dd") <= Format(Xarmoelmesh, "yyyy/mm/dd") Then
                 t_pmemi.Text = Xmmpemi
                 t_paemi.Text = Xaapemi
              End If
           End If
        Else
           t_pmemi.Text = Xmmpemi
           t_paemi.Text = Xaapemi
        End If
   End If
Else
   t_pmemi.Text = 0
   t_paemi.Text = 0
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub VerPromoCliNew()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim cedruta As String
If txt_ced.Text <> "" Then
   cedruta = Trim(txt_ced.Text) & Trim(txt_ced2.Text)
Else
   cedruta = "0"
End If
If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
If t_ruta.Text = "" Then
   Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & Val(cedruta)
Else
   Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & t_ruta.Text
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   Label42.Caption = "Grupo: " & Xrecclii.RecordCount + 1
   t_ruta.Enabled = True
Else
   Label42.Caption = ""
   t_ruta.Enabled = True
End If
Xrecclii.Close
ConbdSapp.Close

End Sub


Public Sub Consulta_cobZon()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xelcob As Integer
Xelcob = 0

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from zonas where zo_grupo =" & txt_codzon.Text
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("zo_cob")) = False Then
      Xelcob = Xrecclii("zo_cob")
      If Xelcob > 0 Then
         MsgBox "Cobrador sugerido para la zona: " & Xelcob, vbInformation
      End If
   End If
End If
Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_Notas(ByVal XmatNotas As Long)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from notas_med where matricula =" & XmatNotas
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   Image4.Visible = True
   Image3.Visible = False
Else
   Image3.Visible = True
   Image4.Visible = False
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Function Devuelve_ceduladoble() As Integer

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
If Trim(txt_ced.Text) <> "" Then
   If txt_ced.Text > 0 Then
      Xsqlpromo = "Select * from clientes where cl_cedula =" & txt_ced.Text
      With Xrecclii
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Xsqlpromo, ConbdSapp, , , adCmdText
      End With
      If Xrecclii.RecordCount >= 1 Then
         Devuelve_ceduladoble = 1
      Else
         Devuelve_ceduladoble = 0
      End If
      Xrecclii.Close
   Else
      Devuelve_ceduladoble = 0
   End If
Else
   Devuelve_ceduladoble = 0
End If

ConbdSapp.Close

End Function



Public Sub Verifica_datosJ()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Fecha_Datos As Date
Fecha_Datos = Date - 35

DatosVerificadosOk = 1
If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from clientes_history where cl_codigo =" & Val(txt_mat.Caption) & " and fecha_modif >='" & Format(Fecha_Datos, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount <= 0 Then
   XorigenDatos = 1
Else
   DatosVerificadosOk = 0
End If

Xrecclii.Close
ConbdSapp.Close

If DatosVerificadosOk = 1 Then
   frm_valida_datos_socio.Show vbModal
End If

End Sub

Public Sub altaValidacionDatos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from clientes_history where cl_codigo =" & Val(txt_mat.Caption)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xrecclii.AddNew
Xrecclii("cl_dpto") = Trim(t_cel.Text)
Xrecclii("cl_referen") = Trim(t_correo.Text)
Xrecclii("cl_socmnom") = cbomutual.Text
Xrecclii("cl_telefon") = Trim(txt_telef.Text)
Xrecclii("fecha_modif") = Date
Xrecclii("origen") = "FICHA " & data_parsec.Recordset("base")
Xrecclii("usuario") = WElusuario
Xrecclii("cl_codigo") = Val(txt_mat.Caption)
Xrecclii.Update


Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub altaValidacionDatosabm()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from abmsocio where cl_codigo =" & Val(txt_mat.Caption)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xrecclii.AddNew
Xrecclii("usuario") = WElusuario
Xrecclii("fecha") = Date
Xrecclii("hora") = Format(Time, "HH:mm")
Xrecclii("cl_codigo") = Val(txt_mat.Caption)
Xrecclii("desc") = "MODIF"
Xrecclii("cl_motivo") = "VALIDACION DATOS"
Xrecclii("convenio") = txt_codcnv.Text
Xrecclii("base") = data_parsec.Recordset("base")
Xrecclii.Update


Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub Verifica_datosJ_Fact()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Fecha_Datos As Date
Fecha_Datos = Date - 35

DatosVerificadosOk = 1

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from clientes_history where cl_codigo =" & Val(txt_mat.Caption) & " and fecha_modif >='" & Format(Fecha_Datos, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount <= 0 Then
   XorigenDatos = 4
Else
   DatosVerificadosOk = 0
End If

Xrecclii.Close
ConbdSapp.Close

If DatosVerificadosOk = 1 Then
   frm_valida_datos_socio.Show vbModal
End If

End Sub


Public Sub Control_proxemi_guardar()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xmmpemi, Xaapemi As Integer
Dim Xarmoelmesd, Xarmoelmesh As String


If Month(Date) > 9 Then
   Xarmoelmesh = "01/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
Else
   Xarmoelmesh = "01/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
End If

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
If Month(Date) = 12 Then
   Xmmpemi = 1
   Xaapemi = Year(Date) + 1
Else
   Xmmpemi = Month(Date) + 1
   Xaapemi = Year(Date)
End If

Xsqlpromo = "Select * from convenio where cnv_codigo ='" & txt_codcnv.Text & "' and cnv_emite ='" & "SI" & "' and cnv_cant_r in (1,2) and cnv_hasta >='" & Format(Date, "yyyy/mm/dd") & "'"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
'SI ES >25 DEL MES/AÑO ACTUAL
   If txt_fecing.Text = "__/__/____" Then
      txt_fecing.Text = Date
   End If
   If Day(txt_fecing.Text) > 25 And Year(txt_fecing.Text) = Xaapemi Then
        If t_pmemi.Text <> "" Then
           If Val(t_pmemi.Text) <= 0 Then
              If Xmmpemi = 12 Then
                 t_pmemi.Text = 1
                 t_paemi.Text = Xaapemi + 1
              Else
                 t_pmemi.Text = Xmmpemi + 1
                 t_paemi.Text = Xaapemi
              End If
           Else
              If t_pmemi.Text > 9 Then
                 Xarmoelmesd = "01/" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              Else
                 Xarmoelmesd = "01/0" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              End If
              If Format(Xarmoelmesd, "yyyy/mm/dd") <= Format(Xarmoelmesh, "yyyy/mm/dd") Then
                 If Xmmpemi = 12 Then
                    t_pmemi.Text = 1
                    t_paemi.Text = Xaapemi + 1
                 Else
                    t_pmemi.Text = Xmmpemi + 1
                    t_paemi.Text = Xaapemi
                 End If
              End If
           End If
        Else
           If Xmmpemi = 12 Then
              t_pmemi.Text = 1
              t_paemi.Text = Xaapemi + 1
           Else
              t_pmemi.Text = Xmmpemi + 1
              t_paemi.Text = Xaapemi
           End If
        End If
   Else
        If t_pmemi.Text <> "" Then
           If Val(t_pmemi.Text) <= 0 Then
              t_pmemi.Text = Xmmpemi
              t_paemi.Text = Xaapemi
           Else
              If t_pmemi.Text > 9 Then
                 Xarmoelmesd = "01/" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              Else
                 Xarmoelmesd = "01/0" & Trim(str(t_pmemi.Text)) & "/" & Trim(str(t_paemi.Text))
              End If
              If Format(Xarmoelmesd, "yyyy/mm/dd") <= Format(Xarmoelmesh, "yyyy/mm/dd") Then
                 t_pmemi.Text = Xmmpemi
                 t_paemi.Text = Xaapemi
              End If
           End If
        Else
           t_pmemi.Text = Xmmpemi
           t_paemi.Text = Xaapemi
        End If
   End If
Else
   t_pmemi.Text = 0
   t_paemi.Text = 0
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Function Verifica_Cgalicia() As Integer

Dim XsqlpromoF As String
Dim XreccliiAvisoF As New ADODB.Recordset
Dim Minutos As String
Dim CedulaBusca As String
Dim Hayautorizacion As Integer
Dim EsUrgente As String

Hayautorizacion = 0
ConectarAvisoF
ConbdSappAvisoF.Open

If frmabm.txt_mat.Caption <> "" Then
   XsqlpromoF = "Select * from Codigos_aut where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and socio =" & Val(frmabm.txt_mat.Caption) & " and codaut ='" & "C.GALICIA" & "' and horafin <='" & Format(Time, "HH:mm") & "' and modulo ='" & "FACTURACION" & "'"
Else
   XsqlpromoF = "Select * from Codigos_aut where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cuando ='" & "0-0" & "' and codaut ='" & "C.GALICIA" & "' and horafin <='" & Format(Time, "HH:mm") & "' and modulo ='" & "FACTURACION" & "'"
End If
With XreccliiAvisoF
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
End With
If XreccliiAvisoF.RecordCount > 0 Then
   Hayautorizacion = 1
End If
XreccliiAvisoF.Close

If Hayautorizacion = 1 Then
   Verifica_Cgalicia = 0
Else
    XsqlpromoF = "Select * from convenio where cnv_codigo ='" & Trim(frmabm.txt_codcnv.Text) & "' and cnv_grupo in ('CASA DE GALICIA')"
    With XreccliiAvisoF
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
    End With
    If XreccliiAvisoF.RecordCount > 0 Then
       EsUrgente = MsgBox("Es COBRANZA DE CUOTAS?", vbCritical + vbYesNo, "Facturación")
       If EsUrgente = vbYes Then
          Verifica_Cgalicia = 0
          CgalDesde = 0
          XreccliiAvisoF.Close
          XsqlpromoF = "Select * from Codigos_aut where fecha ='" & Format(Date, "yyyy-mm-dd") & "'"
          With XreccliiAvisoF
              .CursorLocation = adUseClient
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
          End With
          Minutos = DateAdd("n", 5, Time)
          XreccliiAvisoF.AddNew
          XreccliiAvisoF("fecha") = Date
          XreccliiAvisoF("usuario") = WElusuario
          XreccliiAvisoF("codaut") = "URGENCIA CGAL"
          XreccliiAvisoF("hora") = Format(Time, "HH:mm")
          XreccliiAvisoF("horafin") = Format(Minutos, "HH:mm")
          XreccliiAvisoF("socio") = Val(frmabm.txt_mat.Caption)
          XreccliiAvisoF("modulo") = "FACTURACION"
          XreccliiAvisoF("usuario_caja") = WElusuario
          XreccliiAvisoF("contacto") = frmabm.t_cel.Text & "//" & frmabm.txt_telef.Text
          XreccliiAvisoF("observa") = Mid(frmabm.cbolocalid.Text, 1, 140)
          XreccliiAvisoF("cuando") = frmabm.txt_ced.Text & "-" & frmabm.txt_ced2.Text
          XreccliiAvisoF("base") = frmabm.data_parsec.Recordset("base")
          XreccliiAvisoF("aviso") = 1
          XreccliiAvisoF.Update
       Else
          Verifica_Cgalicia = 1
          CgalDesde = 2
       End If
    Else
       Verifica_Cgalicia = 0
       CgalDesde = 0
    End If
    XreccliiAvisoF.Close
    If Verifica_Cgalicia = 1 Then
        XsqlpromoF = "Select * from Codigos_aut where fecha ='" & Format(Date, "yyyy-mm-dd") & "'"
        With XreccliiAvisoF
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
        End With
        Minutos = DateAdd("n", 5, Time)
        XreccliiAvisoF.AddNew
        XreccliiAvisoF("fecha") = Date
        XreccliiAvisoF("usuario") = WElusuario
        XreccliiAvisoF("codaut") = "C.GALICIA"
        XreccliiAvisoF("hora") = Format(Time, "HH:mm")
        XreccliiAvisoF("horafin") = Format(Minutos, "HH:mm")
        XreccliiAvisoF("socio") = Val(frmabm.txt_mat.Caption)
        XreccliiAvisoF("modulo") = "FACTURACION"
        XreccliiAvisoF("usuario_caja") = WElusuario
        XreccliiAvisoF("contacto") = frmabm.t_cel.Text & "//" & frmabm.txt_telef.Text
        XreccliiAvisoF("observa") = Mid(frmabm.cbolocalid.Text, 1, 140)
        XreccliiAvisoF("cuando") = frmabm.txt_ced.Text & "-" & frmabm.txt_ced2.Text
        XreccliiAvisoF("base") = frmabm.data_parsec.Recordset("base")
        XreccliiAvisoF.Update
        XreccliiAvisoF.Close
    End If
End If

ConbdSappAvisoF.Close

End Function


Public Sub Actualiza_Magik()
Dim Datos As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim mutualista As String
Dim ConvenioNuevo As String
Dim XX As Integer

On Error GoTo ErrMagik

mutualista = ""
ConvenioNuevo = ""

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from convenio where cnv_codigo ='" & data_clientes.Recordset("cl_codconv") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cnv_grupo")) = False Then
      mutualista = Xrecclii("cnv_grupo")
   End If
Else
   mutualista = ""
End If
Xrecclii.Close
ConbdSapp.Close

If mutualista = "UNIVERSAL" Or mutualista = "CCOU" Or mutualista = "H.EVANGELICO" Or mutualista = "SMI" Then
   If data_clientes.Recordset.RecordCount > 0 Then
      Open App.path & "\magik.xml" For Output As #1
      Datos = ""
      Datos = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:mgk=""MgkWebEvo3"">"
      Datos = Datos & vbCrLf & "<soapenv:Header/>"
      Datos = Datos & vbCrLf & "<soapenv:Body>"
      Datos = Datos & vbCrLf
      Datos = Datos & "<mgk:wsAFIPersonaContratos.Execute>"
      Datos = Datos & vbCrLf & "<mgk:Sdtafipersonacontratosws>"
      Datos = Datos & vbCrLf & "<mgk:PerIdExterno>" & Int(data_clientes.Recordset("cl_codigo")) & "</mgk:PerIdExterno>"
      Datos = Datos & vbCrLf & "<mgk:PaisDocIdentidad>URY^URUGUAY</mgk:PaisDocIdentidad>"
      Datos = Datos & vbCrLf & "<mgk:TipoDocIdentidad>1^CEDULA</mgk:TipoDocIdentidad>"
      Datos = Datos & vbCrLf & "<mgk:PerCI>" & Int(data_clientes.Recordset("cl_cedula")) & "</mgk:PerCI>"
      Datos = Datos & vbCrLf & "<mgk:PerCIDg>" & Int(data_clientes.Recordset("cl_codced")) & "</mgk:PerCIDg>"
'      If Len(Trim(data_clientes.Recordset("cl_apellid"))) > 30 Then
         Datos = Datos & vbCrLf & "<mgk:PerNom1>" & "*," & "</mgk:PerNom1>"
'      Else
'         Datos = Datos & vbCrLf & "<mgk:PerNom1>" & Mid(Trim(data_clientes.Recordset("cl_apellid")), 1, 20) & "</mgk:PerNom1>"
'      End If
'      If Len(Trim(data_clientes.Recordset("cl_apellid"))) > 30 Then
         Datos = Datos & vbCrLf & "<mgk:PerApe1>" & Mid(data_clientes.Recordset("cl_apellid"), 1, 20) & "</mgk:PerApe1>"
'      Else
'         Datos = Datos & vbCrLf & "<mgk:PerApe1>NO APLICA</mgk:PerApe1>"
'      End If
      If IsNull(data_clientes.Recordset("cl_fnac")) = False Then
         Datos = Datos & vbCrLf & "<mgk:PerFNac>" & Format(data_clientes.Recordset("cl_fnac"), "yyyy-mm-dd") & "</mgk:PerFNac>"
      Else
         Datos = Datos & vbCrLf & "<mgk:PerFNac>2000-12-01</mgk:PerFNac>"
      End If
      If IsNull(data_clientes.Recordset("cl_sexo")) = False Then
         If data_clientes.Recordset("cl_sexo") = 1 Then
            Datos = Datos & vbCrLf & "<mgk:PerSexo>M</mgk:PerSexo>"
         Else
            Datos = Datos & vbCrLf & "<mgk:PerSexo>M</mgk:PerSexo>"
         End If
      End If
      Datos = Datos & vbCrLf & "<mgk:PerEstCiv>5</mgk:PerEstCiv>"
      If IsNull(data_clientes.Recordset("cl_fecing")) = False Then
         Datos = Datos & vbCrLf & "<mgk:PerFIng>" & Format(data_clientes.Recordset("cl_fecing"), "yyyy-mm-dd") & "</mgk:PerFIng>"
      Else
         Datos = Datos & vbCrLf & "<mgk:PerFIng>2000-12-01</mgk:PerFIng>"
      End If
      Datos = Datos & vbCrLf & "<mgk:PerNHCli>" & Int(data_clientes.Recordset("cl_codigo")) & "</mgk:PerNHCli>"
      Datos = Datos & vbCrLf & "<mgk:PerMat>" & Int(data_clientes.Recordset("cl_codigo")) & "</mgk:PerMat>"
      Datos = Datos & vbCrLf & "<mgk:Permail>NO APLICA</mgk:Permail>"
      If IsNull(data_clientes.Recordset("cl_direcci")) = False Then
         Datos = Datos & vbCrLf & "<mgk:Direccion>" & data_clientes.Recordset("cl_direcci") & "</mgk:Direccion>"
      Else
         Datos = Datos & vbCrLf & "<mgk:Direccion>NO APLICA</mgk:Direccion>"
      End If
      If IsNull(data_clientes.Recordset("cl_zona")) = False Then
         Datos = Datos & vbCrLf & "<mgk:LocalidadDomAte>UY-CA-" & Mid(Trim(data_clientes.Recordset("cl_zona")), 1, 3) & "^" & data_clientes.Recordset("cl_zona") & "</mgk:LocalidadDomAte>"
      Else
         Datos = Datos & vbCrLf & "<mgk:LocalidadDomAte>UY-CA-SAL^SALINAS</mgk:LocalidadDomAte>"
      End If
      Datos = Datos & vbCrLf & "<mgk:DeparDomAte>UY-CA^CANELONES</mgk:DeparDomAte>"
      Datos = Datos & vbCrLf & "<mgk:PaisDomAte>UY^URUGUAY</mgk:PaisDomAte>"
      Datos = Datos & vbCrLf & "<mgk:ColTelefonos>"
      Datos = Datos & vbCrLf & "<mgk:ColTelefonosItem>"
      Datos = Datos & vbCrLf & "<mgk:DomTelTpoCod>1</mgk:DomTelTpoCod>"
      If IsNull(data_clientes.Recordset("cl_telefon")) = False Then
         Datos = Datos & vbCrLf & "<mgk:DomicTel>" & data_clientes.Recordset("cl_telefon") & "</mgk:DomicTel>"
      Else
         Datos = Datos & vbCrLf & "<mgk:DomicTel>NO APLICA</mgk:DomicTel>"
      End If
      Datos = Datos & vbCrLf & "<mgk:DomTelTpoCod>3</mgk:DomTelTpoCod>"
      If IsNull(data_clientes.Recordset("cl_dpto")) = False Then
         Datos = Datos & vbCrLf & "<mgk:DomicTel>" & Trim(data_clientes.Recordset("cl_dpto")) & "</mgk:DomicTel>"
      Else
         Datos = Datos & vbCrLf & "<mgk:DomicTel>NO APLICA</mgk:DomicTel>"
      End If
      Datos = Datos & vbCrLf & "</mgk:ColTelefonosItem>"
      Datos = Datos & vbCrLf & "</mgk:ColTelefonos>"
      Datos = Datos & vbCrLf & "<mgk:ColContratos>"
      Datos = Datos & vbCrLf & "<mgk:ColContratosItem>"
      Datos = Datos & vbCrLf & "<mgk:Servicio>" & mutualista & "</mgk:Servicio>"
      For XX = 1 To Len(data_clientes.Recordset("cl_nomconv"))
          If Mid(data_clientes.Recordset("cl_nomconv"), XX, 1) = ">" Or Mid(data_clientes.Recordset("cl_nomconv"), XX, 1) = "<" Then
             ConvenioNuevo = ConvenioNuevo & " "
          Else
             ConvenioNuevo = ConvenioNuevo & Mid(data_clientes.Recordset("cl_nomconv"), XX, 1)
          End If
      Next
      
      Datos = Datos & vbCrLf & "<mgk:Categoria>" & Int(data_clientes.Recordset("id")) & "^" & Mid(Trim(ConvenioNuevo), 1, 30) & "</mgk:Categoria>"
      Datos = Datos & vbCrLf & "<mgk:PerContFIni>1987-01-01</mgk:PerContFIni>"
      Datos = Datos & vbCrLf & "<mgk:PerContVig>2030-01-01</mgk:PerContVig>"
      Datos = Datos & vbCrLf & "<mgk:PerContFVen>2030-01-01</mgk:PerContFVen>"
      Datos = Datos & vbCrLf & "<mgk:PerContFBaj></mgk:PerContFBaj>"
      Datos = Datos & vbCrLf & "</mgk:ColContratosItem>"
      Datos = Datos & vbCrLf & "</mgk:ColContratos>"
      Datos = Datos & vbCrLf & "</mgk:Sdtafipersonacontratosws>"
      Datos = Datos & vbCrLf & "</mgk:wsAFIPersonaContratos.Execute>"
      Print #1, Datos
      Datos = ""
      Datos = ""
      Datos = Datos & "</soapenv:Body>"
      Datos = Datos & vbCrLf & "</soapenv:Envelope>"
      Print #1, Datos
      Close #1

      Dim xmlResponse As MSXML2.DOMDocument30
      Dim strSoap As String
      Dim strSOAPAction As String
      Dim strWsdl As String
      Dim FicheroXML As String
      FicheroXML = App.path & "\magik.xml"
      Dim StrTextoxml As String
      Dim StrTextoError As String
      Dim Strlinea As String
      Open FicheroXML For Input As #1
      StrTextoError = ""
      Do While Not EOF(1)
         Line Input #1, Strlinea
         StrTextoxml = StrTextoxml + Strlinea
      Loop
      Close #1
      strSoap = StrTextoxml
      strSOAPAction = "wsAFIPersonaContratos"
      strWsdl = "http://192.168.10.182:8080/MAGIK_Servicios/servlet/com.mgkwebevo3.afiliaciones.awsafipersonacontratos?wsdl"
'''      strWsdl = "http://192.168.10.183:8080/MAGIK_Servicios/servlet/com.mgkwebevo3.afiliaciones.awsafipersonacontratos?wsdl"
      
      If InvokeWebService(strSoap, strSOAPAction, strWsdl, xmlResponse) Then
         StrTextoError = xmlResponse.XML
      Else
         StrTextoError = "Error"
      End If
      If Trim(StrTextoError) = "Error" Then
         MsgBox "Hubo un error en el envío a MAGIK, vuelva a modificar", vbCritical
      Else
         Dim valorBuscado As String
         Dim oNode As IXMLDOMNode
         Set oNode = xmlResponse.selectSingleNode("//Sdtadmbitinterror//CodError")
         If Not oNode Is Nothing Then
            valorBuscado = oNode.Text
         Else
            valorBuscado = ""
         End If
         If Trim(valorBuscado) = "0" Then
   '         MsgBox "Carga exitosa"
         Else
            MsgBox "Error en la carga del archivo XML a Magik, vuelva a modificar.", vbCritical
         End If
      End If
      Set xmlResponse = Nothing
   End If
End If

'MsgBox "Terminado"
Exit Sub

ErrMagik:
        If Err.Number = 3155 Then
           MsgBox "Error al enviar XML : " & Err.Description
        Else
           MsgBox "Error al enviar XML : " & Err.Description
        End If
End Sub
