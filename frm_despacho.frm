VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_largador 
   BackColor       =   &H00FF8080&
   Caption         =   "Largador"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11205
   Icon            =   "frm_despacho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11205
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton b_covid 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10680
      Picture         =   "frm_despacho.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   132
      ToolTipText     =   "Datos para seguimiento de COVID"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos para CMT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frm_despacho.frx":6814
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Enviar datos"
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
      Left            =   7080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   "Envía datos a padrón social para modificar en la ficha del socio"
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ver mapa..."
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
      Left            =   5280
      MouseIcon       =   "frm_despacho.frx":6D9E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "Buscar dirección en google maps"
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Data data_hist 
      Caption         =   "data_hist"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton b_hist 
      BackColor       =   &H000080FF&
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
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Data data_mov 
      Caption         =   "data_mov"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data data_llasql 
      Caption         =   "data_llasql"
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Data data_cons 
      Caption         =   "data_cons"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_convbus 
      Caption         =   "data_convbus"
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
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_ref 
      Caption         =   "data_ref"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport crl 
      Left            =   9480
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_lla2 
      Caption         =   "data_lla2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PARSEC0"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7800
      MaxLength       =   5
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton b_cancel 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cancelar llamado"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   7560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   76
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data data_clib 
      Caption         =   "data_clib"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Datos Largador..."
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
      Height          =   2775
      Left            =   120
      TabIndex        =   40
      Top             =   4800
      Width           =   10935
      Begin VB.Data data_aut 
         Caption         =   "data_aut"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox txt_lugar 
         Height          =   315
         ItemData        =   "frm_despacho.frx":70A8
         Left            =   8160
         List            =   "frm_despacho.frx":70AA
         TabIndex        =   128
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
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
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mhtrassol 
         Height          =   255
         Left            =   6600
         TabIndex        =   113
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mftrassol 
         Height          =   255
         Left            =   5400
         TabIndex        =   112
         Top             =   840
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
      Begin VB.CommandButton c_aft 
         BackColor       =   &H0080C0FF&
         Caption         =   "AFT"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Autorización final de traslado"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Data data_med2 
         Caption         =   "data_med2"
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
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data data_llamod 
         Caption         =   "data_llamod"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_codmed2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6120
         TabIndex        =   109
         Top             =   2040
         Width           =   735
      End
      Begin MSDBCtls.DBCombo dbcbomed2 
         Bindings        =   "frm_despacho.frx":70AC
         DataSource      =   "data_med2"
         Height          =   345
         Left            =   3240
         TabIndex        =   108
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0000FFFF&
         Caption         =   "Traslado sol. por TERCEROS"
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
         Left            =   7440
         TabIndex        =   106
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txt_codmedtra 
         Height          =   375
         Left            =   6720
         TabIndex        =   98
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_salca 
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
         Left            =   4440
         TabIndex        =   97
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txt_queb 
         Height          =   285
         Left            =   120
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frm_despacho.frx":70C4
         Left            =   5040
         List            =   "frm_despacho.frx":70D7
         TabIndex        =   82
         Text            =   "Combo1"
         Top             =   480
         Width           =   2295
      End
      Begin VB.Data data_imp 
         Caption         =   "data_imp"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "inflla"
         Top             =   -120
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txt_codmed 
         Alignment       =   2  'Center
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
         TabIndex        =   75
         Top             =   1800
         Width           =   855
      End
      Begin MSMask.MaskEdBox txt_hortd 
         Height          =   255
         Left            =   3120
         TabIndex        =   54
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_movtra 
         Alignment       =   2  'Center
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
         Left            =   9960
         MaxLength       =   6
         TabIndex        =   73
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txt_enzona 
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
         Left            =   6240
         TabIndex        =   71
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txt_enca 
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
         Left            =   2640
         TabIndex        =   69
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txt_trassal 
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
         Left            =   840
         MaxLength       =   5
         TabIndex        =   67
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox cbotras 
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
         ItemData        =   "frm_despacho.frx":710B
         Left            =   5400
         List            =   "frm_despacho.frx":710D
         TabIndex        =   64
         Top             =   1560
         Width           =   2775
      End
      Begin MSDBCtls.DBCombo dbcbomed 
         Bindings        =   "frm_despacho.frx":710F
         DataSource      =   "data_med"
         Height          =   360
         Left            =   1200
         TabIndex        =   62
         Top             =   1560
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.ComboBox cbocolfin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frm_despacho.frx":7126
         Left            =   9240
         List            =   "frm_despacho.frx":7136
         TabIndex        =   60
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt_diag 
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
         Left            =   1560
         MaxLength       =   70
         TabIndex        =   58
         ToolTipText     =   "Diagnóstico final"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txt_demora 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9720
         MaxLength       =   5
         TabIndex        =   56
         Top             =   480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mtd 
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
         Left            =   1920
         TabIndex        =   53
         Top             =   840
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
      Begin VB.TextBox txt_horlle 
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
         Left            =   8760
         MaxLength       =   5
         TabIndex        =   51
         Top             =   480
         Width           =   615
      End
      Begin MSMask.MaskEdBox mllegada 
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
         Left            =   7560
         TabIndex        =   50
         Top             =   480
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
      Begin VB.TextBox txt_horsal 
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
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   48
         Top             =   480
         Width           =   615
      End
      Begin MSMask.MaskEdBox msalida 
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
         Left            =   3120
         TabIndex        =   47
         Top             =   480
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
      Begin VB.TextBox txt_horasig 
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
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   45
         Top             =   480
         Width           =   615
      End
      Begin MSMask.MaskEdBox mfecasig 
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
         Left            =   1200
         TabIndex        =   44
         Top             =   480
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
      Begin VB.TextBox txt_movil 
         Alignment       =   2  'Center
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
         Left            =   120
         MaxLength       =   4
         TabIndex        =   42
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label40 
         BackColor       =   &H0080FFFF&
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
         Left            =   4320
         TabIndex        =   130
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label labnomchof 
         BackColor       =   &H00C0C000&
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
         Left            =   7440
         TabIndex        =   127
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label labcodchof 
         Height          =   255
         Left            =   960
         TabIndex        =   114
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Solicita trasl."
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
         TabIndex        =   111
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Médico del traslado:"
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
         Left            =   1200
         TabIndex        =   107
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label47 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sale C.A:"
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
         Left            =   3480
         TabIndex        =   96
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MOVIL EN BASE?"
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
         Left            =   5040
         TabIndex        =   81
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Móvil:"
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
         Left            =   9240
         TabIndex        =   72
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0FFC0&
         Caption         =   "En Zona"
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
         Left            =   5400
         TabIndex        =   70
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0FFC0&
         Caption         =   "En C.Asis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   68
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salida:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Lugar......:"
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
         Left            =   7080
         TabIndex        =   65
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackColor       =   &H0080FFFF&
         Caption         =   "Traslado:"
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
         Left            =   4320
         MouseIcon       =   "frm_despacho.frx":7158
         MousePointer    =   99  'Custom
         TabIndex        =   63
         ToolTipText     =   "Haciendo CLICK AQUI puede ver un detalle para cada opción de traslados"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Medico:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Color final:"
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
         Left            =   6960
         TabIndex        =   59
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Diag. Final:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "DEMORA:"
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
         Left            =   9720
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFC0&
         Caption         =   "T/D....:"
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
         Left            =   1200
         TabIndex        =   52
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Llegada..."
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
         Left            =   7560
         TabIndex        =   49
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salida..."
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
         Left            =   3120
         TabIndex        =   46
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fecha/Hora"
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
         Left            =   1200
         TabIndex        =   43
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MOVIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frm_despacho.frx":7462
         MousePointer    =   99  'Custom
         TabIndex        =   41
         ToolTipText     =   "Doble click para cancelar largada de móvil"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Datos del llamado"
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
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   10935
      Begin VB.TextBox t_timbre 
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   136
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cbotimbre 
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
         Height          =   315
         ItemData        =   "frm_despacho.frx":776C
         Left            =   3360
         List            =   "frm_despacho.frx":7776
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox chcovid 
         BackColor       =   &H008080FF&
         Caption         =   "Covid-19"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   131
         ToolTipText     =   "Sospecha COVID-19"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox txt_locali 
         Height          =   315
         ItemData        =   "frm_despacho.frx":7782
         Left            =   7320
         List            =   "frm_despacho.frx":7784
         TabIndex        =   129
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Data data_chof 
         Caption         =   "data_chof"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "No modificar datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8520
         TabIndex        =   126
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_despacho.frx":7786
         Left            =   8160
         List            =   "frm_despacho.frx":7790
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txt_boleta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         TabIndex        =   123
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_costo 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   121
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chtmut 
         BackColor       =   &H0080FFFF&
         Caption         =   "Tiene ticket mutual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7800
         TabIndex        =   119
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton b_cmt 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         Picture         =   "frm_despacho.frx":77A6
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Pasar a CMT"
         Top             =   2640
         Width           =   495
      End
      Begin VB.Data data_histant 
         Caption         =   "data_histant"
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
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data data_aran 
         Caption         =   "data_aran"
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
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox t_codced 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   115
         Top             =   720
         Width           =   375
      End
      Begin VB.Data data_u 
         Caption         =   "data_u"
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
         RecordSource    =   ""
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data data_deuda 
         Caption         =   "data_deuda"
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
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frm_despacho.frx":7BE8
         Left            =   6840
         List            =   "frm_despacho.frx":7BF2
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txt_quien 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   100
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   10440
         Picture         =   "frm_despacho.frx":7C01
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Cargar otros datos para 911"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Acto de Enfermería"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   87
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txt_obs 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   8160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         ToolTipText     =   "Para el caso de CERTIFICACIONES ingresar AQUI fechas de licencia."
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Data data_usu 
         Caption         =   "data_usu"
         Connect         =   "Access"
         DatabaseName    =   "C:\Windows\usapp.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "usuarioact"
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox cbobase 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   10200
         TabIndex        =   30
         Top             =   2160
         Width           =   615
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
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
         RecordSource    =   ""
         Top             =   3960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data data_lla 
         Caption         =   "data_lla"
         Connect         =   "Access"
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
         Width           =   2895
      End
      Begin VB.TextBox txt_mot 
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
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   3480
         Width           =   5655
      End
      Begin VB.ComboBox cboed 
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
         ItemData        =   "frm_despacho.frx":818B
         Left            =   9600
         List            =   "frm_despacho.frx":8198
         TabIndex        =   38
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_edad 
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
         Left            =   9000
         MaxLength       =   3
         TabIndex        =   37
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox cbocolor 
         BackColor       =   &H00008000&
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
         ItemData        =   "frm_despacho.frx":81AF
         Left            =   8520
         List            =   "frm_despacho.frx":81C5
         TabIndex        =   35
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txt_ante 
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
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   33
         Top             =   4080
         Width           =   5655
      End
      Begin VB.ComboBox cbozona 
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
         ItemData        =   "frm_despacho.frx":81F6
         Left            =   8160
         List            =   "frm_despacho.frx":81F8
         TabIndex        =   29
         ToolTipText     =   "ZONAS: 1=Costa; 2=Norte, 3=Tala, 4=Traslados SEMESA y Llamados Universal LAS PIEDRAS, 5=San Jacinto, 6=ARM"
         Top             =   2160
         Width           =   1095
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
         Height          =   285
         Left            =   3720
         MaxLength       =   35
         TabIndex        =   27
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txt_direc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2160
         Width           =   6015
      End
      Begin VB.TextBox txt_ced 
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
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt_nomcat 
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
         Left            =   2760
         MaxLength       =   45
         TabIndex        =   21
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txt_cat 
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
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txt_nomb 
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
         Left            =   3840
         MaxLength       =   70
         TabIndex        =   18
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txt_mat 
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
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txt_usua 
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
         Height          =   285
         Left            =   8160
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txt_hora 
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
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin MSMask.MaskEdBox mfecha 
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
         TabIndex        =   11
         Top             =   360
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label52 
         BackColor       =   &H00FF0000&
         Caption         =   "TIMBRE?"
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
         Left            =   2400
         TabIndex        =   134
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label labanthor 
         BackColor       =   &H00808000&
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
         Left            =   6360
         TabIndex        =   133
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label51 
         BackColor       =   &H00FF0000&
         Caption         =   "Tipo Boleta:"
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
         Left            =   7080
         TabIndex        =   124
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FF0000&
         Caption         =   "Boleta:"
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
         Left            =   4920
         TabIndex        =   122
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FF0000&
         Caption         =   "COSTO:"
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
         TabIndex        =   120
         ToolTipText     =   "Si el campo de COSTO está sin habilitar es porque el llamado ya está FACTURADO"
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label labcmt 
         BackColor       =   &H0080FFFF&
         Caption         =   "PASADO A CMT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frm_despacho.frx":81FA
         MousePointer    =   99  'Custom
         TabIndex        =   116
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FF0000&
         Caption         =   "Sexo:"
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
         Left            =   6240
         TabIndex        =   101
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label48 
         Height          =   255
         Left            =   1920
         TabIndex        =   99
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label46 
         Height          =   255
         Left            =   8400
         TabIndex        =   95
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label45 
         Height          =   255
         Left            =   6600
         TabIndex        =   94
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label44 
         Height          =   255
         Left            =   5640
         TabIndex        =   93
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label43 
         Height          =   255
         Left            =   4680
         TabIndex        =   92
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label42 
         Height          =   255
         Left            =   3720
         TabIndex        =   91
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label41 
         Height          =   255
         Left            =   2880
         TabIndex        =   90
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FF0000&
         Caption         =   "H. Grabado:"
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
         Left            =   9720
         TabIndex        =   84
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
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
         Left            =   9720
         TabIndex        =   83
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FF0000&
         Caption         =   "OBSERV.:"
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
         Left            =   6960
         TabIndex        =   79
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Motivo de Consulta:"
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
         TabIndex        =   74
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
         Caption         =   "Edad:"
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
         Left            =   8400
         TabIndex        =   36
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label15 
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
         Left            =   7320
         TabIndex        =   34
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF0000&
         Caption         =   "Antecedentes:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF0000&
         Caption         =   "BASE:"
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
         Left            =   9480
         TabIndex        =   31
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF0000&
         Caption         =   "Zona:"
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
         Left            =   7320
         TabIndex        =   28
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF0000&
         Caption         =   "Teléf.:"
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
         Left            =   2760
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         Caption         =   "Dirección y Referencia"
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
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "Cédula:"
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
         Left            =   120
         MouseIcon       =   "frm_despacho.frx":8504
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Categoría"
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
         TabIndex        =   19
         ToolTipText     =   "Haga doble click aquí para ver derechos del convenio"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Left            =   3000
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Matrícula"
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
         TabIndex        =   15
         ToolTipText     =   "Haga doble click AQUI para ver los datos de la matrícula ingresada"
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Usuario:"
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
         Left            =   7320
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha/Hora:"
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
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Nro."
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
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprime"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime pantalla actual de llamado"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton b_buscar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Buscar"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton b_pend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Pendientes"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "También puede presionar F1 para ver pendientes o En curso"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton b_grabar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Grabar"
      Enabled         =   0   'False
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modificar"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Nuevo"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label39 
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
      Left            =   1680
      TabIndex        =   86
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Largador:"
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
      Left            =   120
      TabIndex        =   85
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4800
      Picture         =   "frm_despacho.frx":8A8E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frm_largador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
   nSize As Long) As Long


Private Sub Text18_Change()

End Sub

Private Sub b_buscar_Click()
frm_busllamado.Show vbModal

End Sub

Private Sub b_cancel_Click()
frm_cancella.Show vbModal

End Sub

Private Sub b_cancela_Click()
'Dim Xdeseacan As String
'Xdeseacan = MsgBox("SEGURO QUE DESEA CANCELAR LOS DATOS?", vbExclamation + vbYesNo)
'If Xdeseacan = vbYes Then

On Error GoTo AlCancelas

    b_cmt.Enabled = True
'    labcmt.Visible = False
    
    If XAlta = 1 Then
        XllevacostoAp = 0
        data_hist.Recordset.AddNew
        If txt_nro.Text <> "" Then
           data_hist.Recordset("idllamado") = txt_nro.Text
        Else
           data_hist.Recordset("idllamado") = 0
        End If
        data_hist.Recordset("fecha") = Date
        data_hist.Recordset("hora") = Format(Time, "HH:mm")
        data_hist.Recordset("usuario") = WElusuario
        data_hist.Recordset("accion") = "CANCELA NUEVO"
        If txt_cat.Text <> "" Then
           data_hist.Recordset("categ") = txt_cat.Text
        Else
           data_hist.Recordset("categ") = "AABB"
        End If
        If Trim(cbocolor.Text) <> "" Then
           data_hist.Recordset("claveini") = cbocolor.Text
        Else
           data_hist.Recordset("claveini") = "AABB"
        End If
        data_hist.Recordset.Update
       
       Frame1.Enabled = False
       Frame2.Enabled = False
'       data_lla.Recordset.CancelUpdate
       borra_ya
       XAlta = 0
       b_nuevo.Enabled = True
       b_modif.Enabled = True
       b_imp.Enabled = True
       b_buscar.Enabled = True
       b_grabar.Enabled = False
       b_covid.Enabled = True
       b_cancel.Enabled = True
       b_cancela.Enabled = False
       b_hist.Enabled = True
       Command2.Enabled = True
       Command3.Enabled = True
       If WDespa = 1 Then
          b_pend.Enabled = True
       Else
          b_pend.Enabled = True
       End If
'       igualar_lla
       Xdeudasi = 0
       txt_costo.Enabled = True
    Else
       If txt_nro.Text <> "" Then
          guarda_Alcancelar
       End If
       Frame1.Enabled = False
       Frame2.Enabled = False
       XAlta = 0
       b_nuevo.Enabled = True
       b_modif.Enabled = True
       b_imp.Enabled = True
       b_buscar.Enabled = True
       b_grabar.Enabled = False
       b_cancel.Enabled = True
       b_cancela.Enabled = False
       b_hist.Enabled = True
       b_covid.Enabled = True
       Command2.Enabled = True
       Command3.Enabled = True
       If WDespa = 1 Then
          b_pend.Enabled = True
       Else
          b_pend.Enabled = True
       End If
       If data_lla.Recordset.RecordCount > 0 Then
          igualar_sin
       End If
       Xdeudasi = 0
       txt_costo.Enabled = True
    End If
    XAlta = 3
    chcovid.Enabled = False
'End If

Exit Sub

AlCancelas:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ALCANCE ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ALCANCE ERR:" & Err.Number
      End If

End Sub

Private Sub b_cmt_Click()
Dim Xdeseacmt As String
Dim MensajeAtsocio As String
Dim MensajeRepite As String

On Error GoTo Alcmt

Xdeseacmt = ""
'If cbocolor.Text = "NEGRO" Then
'   MsgBox "No puede pasar a CMT"
'Else
   Xdeseacmt = MsgBox("Desea pasar el llamado a CONSULTA MEDICA TELEFONICA (CMT)?", vbInformation + vbYesNo)
   If Xdeseacmt = vbYes Then
      If txt_movil.Text = 0 Or txt_movil.Text = "" Then
         MensajeAtsocio = MsgBox("DESEA DERIVAR EL LLAMADO PARA ATENCIÓN AL SOCIO?", vbInformation + vbYesNo, "CMT Despacho")
         If MensajeAtsocio = vbYes Then
            MensajeRepite = MsgBox("LA CONSULTA ES POR REPETICIÓN DE MEDICACION?", vbExclamation + vbYesNo, "Despacho")
            If MensajeRepite = vbYes Then
               Grabar_RepeticionMed
            End If
            Grabar_CmtAt
         Else
            frm_seleccmt.Show vbModal
         End If
         
      Else
         MsgBox "Ya tiene móvil asignado"
      End If
      b_cancela_Click
   End If
'End If

Exit Sub

Alcmt:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ALCMT ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ALCMT ERR:" & Err.Number
      End If


End Sub

Private Sub b_covid_Click()
Dim SiCMT As String
SiCMT = MsgBox("Desea visualizar datos de COVID?", vbInformation + vbYesNo)
If SiCMT = vbYes Then
    If txt_nro.Text <> "" Then
       frm_seguicovid.Show vbModal
    Else
       MsgBox "No seleccionó el llamado"
    End If
Else
    If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS DESP" Then
       frm_cmt.Show vbModal
    End If
End If

End Sub

Private Sub b_grabar_Click()
Dim Xhd, xha, xmd, xma, xdemh, xdemm, Xcolu, Xestapen, XconCostoAPOk As Integer
Dim Xrespdes, Xrespdes2 As String
Dim SioNocovid As String
Dim XmatCtrol As Long
Dim ConfirmaCosto As String
ConfirmaCosto = ""
XconCostoAPOk = 0


If cbocolor.Text = "VERDE" Or cbocolor.Text = "AZUL" Or cbocolor.Text = "" Then
    If XWeltipoU <> "USUARIOS DESP" Then ' NO Largador
       If txt_mat.Text <> "" Then
          If txt_mat.Text > 0 Then
             XmatCtrol = Val(txt_mat.Text)
          Else
             XmatCtrol = 0
          End If
       Else
          XmatCtrol = 0
       End If
       If Check1.Value = 1 Then
          XmatCtrol = 0
       End If
       If txt_cat.Text = "MSP" Or txt_cat.Text = "911B" Or txt_cat.Text = "CAAMEP" Or _
          txt_cat.Text = "CERSEM" Or txt_cat.Text = "UDEMM" Or txt_cat.Text = "MUCAMT" Or _
          txt_cat.Text = "MUCATA" Or Val(cbozona.Text) = 6 Or Val(cbozona.Text) = 4 Or _
          txt_cat.Text = "UCM" Then
          XmatCtrol = 0
       End If
       If XmatCtrol > 0 Then
          Verifica_datosJ
       Else
          DatosVerificadosOk = 0
       End If
    End If
'continuallamado:
    'FIN VALIDAR SOCIO SI CORRESPONDE
End If

' envio SMS traslado: consume backend asincrono, el SMS lo manda el backend
On Error Resume Next
Dim url_ws_traslado As String
Dim obj As Object
If cbotras.Text <> "" And cbotras.Text <> "NO" And txt_enca.Text <> "" Then 'traslado, en centro asistencial
    url_ws_traslado = getValorFromParametro("url_enviosmstraslado")
    url_ws_traslado = url_ws_traslado & "/" & txt_nro.Text
    Set obj = consumirServicioAsync("GET", url_ws_traslado, "")
End If
' envio SMS

''''On Error GoTo Quepasa

Dim Mensajecovid As String

If cbozona.Text = "" Then
   cbozona.ListIndex = 0
End If
If cbocolor.Text = "" Then
   cbocolor.ListIndex = 0
End If
If Trim(txt_ced.Text) = "" Then
   txt_ced.Text = 0
End If
If Trim(txt_nomb.Text) = "" Then
   txt_nomb.Text = "NN"
End If

If Trim(t_codced.Text) = "" Then
   t_codced.Text = 0
End If
If Verificar_digito() = t_codced.Text Or Verificar_digito() = 999 Then
    If XAlta = 1 Then
       If Verificar_siTieneL = 99 Then
           Lleva_timbre
           Llamado_Ap
           If txt_cat.Text = "APS" Then
              If cbocolor.Text = "VERDE" Or cbocolor.Text = "AMARILLO" Or cbocolor.Text = "AZUL" Or cbocolor.Text = "CELESTE" Then
                 MsgBox "ATENCION! Socio no habilitado para servicios no urgentes! POR AGRESIONES! Consulte con administración!", vbCritical
                 b_cancela_Click
              End If
           End If
           If DatosVerificadosOk = 1 Then
              MsgBox "ATENCION " & Welnombredu & " No ha confirmado datos. Se registrará en historial.", vbCritical
              altaValidacionDatosabm
           End If
           data_hist.Recordset.AddNew
           data_hist.Recordset("idllamado") = txt_nro.Text
           data_hist.Recordset("fecha") = Date
           data_hist.Recordset("hora") = Format(Time, "HH:mm")
           data_hist.Recordset("usuario") = WElusuario
           data_hist.Recordset("accion") = "NUEVO LLAMADO"
           data_hist.Recordset("categ") = txt_cat.Text
           data_hist.Recordset("claveini") = cbocolor.Text
           data_hist.Recordset.Update
           
           data_lla.Recordset.AddNew
           controlasiesta
        '   VerificarCgal = Cgalicia()
           If XllevacostoAp = 9 Or XllevacostoAp = 8 Then
              If XllevacostoAp = 8 Then
                 Grabar_llamadoAp
              Else
                 ConfirmaCosto = MsgBox("Confirma el ingreso del llamado de AP con costo?", vbExclamation + vbYesNo, "Despacho")
                 If ConfirmaCosto = vbYes Then
                    Graba_CostoAp
                    Grabar_llamadoAp
                 Else
                    XconCostoAPOk = 9
                 End If
              End If
           End If
           If Xhayregistros = 9 Or XconCostoAPOk = 9 Then
              If XconCostoAPOk = 9 Then
                 MsgBox "Llamado no aceptado.", vbCritical
                 b_cancela_Click
                 Xhayregistros = 0
              Else
                 frm_mensajedesp.Show vbModal
                 Xhayregistros = 0
                 Unload Me
              End If
           Else
               If cbozona.Text <> "" Then
                  If cbocolor.Text <> "" Then
                        data_lla.Recordset("nrolla") = txt_nro.Text
                        data_lla.Recordset("segui_covid") = chcovid.Value
                        data_lla.Recordset("nro") = txt_nro.Text
                        data_lla.Recordset("fecha") = Format(mfecha.Text, "dd/mm/yyyy")
                        data_lla.Recordset("hora") = Format(txt_hora.Text, "HH:mm")
                        data_lla.Recordset("activo") = Format(Time, "HH:mm:ss")
                        data_lla.Recordset("usuario") = txt_usua.Text
                        data_lla.Recordset("nomodif") = Check1.Value
                        data_lla.Recordset("timbre") = cbotimbre.ListIndex
                        If Trim(t_timbre.Text) <> "" Then
                           data_lla.Recordset("valor_timbre") = t_timbre.Text
                        End If
                        If txt_mat.Text = "" Then
                           data_lla.Recordset("matric") = 0
                        Else
                           data_lla.Recordset("matric") = txt_mat.Text
                        End If
                        data_lla.Recordset("nombre") = txt_nomb.Text
                        If txt_edad = "" Then
                           data_lla.Recordset("edad") = 0
                        Else
                           data_lla.Recordset("edad") = txt_edad.Text
                        End If
                        If cboed.ListIndex = 0 Then
                           data_lla.Recordset("unied") = 3
                        Else
                           If cboed.ListIndex = 1 Then
                              data_lla.Recordset("unied") = 2
                           Else
                              If cboed.ListIndex = 2 Then
                                 data_lla.Recordset("unied") = 1
                              Else
                                 data_lla.Recordset("unied") = 3
                              End If
                           End If
                        End If
                        If txt_cat.Text = "" Then
                           txt_cat.Text = "AAABBB"
                        End If
                        data_lla.Recordset("categ") = txt_cat.Text
                        data_lla.Recordset("nomcat") = txt_nomcat.Text
                        If txt_ced.Text = "" Then
                           data_lla.Recordset("ci") = 0
                        Else
                           data_lla.Recordset("ci") = txt_ced.Text
                        End If
                        data_lla.Recordset("direcc") = "S/D"
                        data_lla.Recordset("telef") = txt_tel.Text
                        If cbozona.Text <> "" Then
                           data_lla.Recordset("codzon") = Val(cbozona.Text)
                        Else
                           data_lla.Recordset("codzon") = 1
                        End If
                        If cbobase.Text = "" Then
                           data_lla.Recordset("base") = 0
                        Else
                           data_lla.Recordset("base") = cbobase.Text
                        End If
                        data_lla.Recordset("referen") = txt_direc.Text
                        If txt_ante.Text <> "" Then
                           data_lla.Recordset("motcon") = txt_ante.Text  'motivo de consulta que no va mas (100) pasa como antecedentes
                        End If
                        data_lla.Recordset("obsmot") = txt_mot.Text
                        data_lla.Recordset("realiza") = chtmut.Value
                        If cbocolor.ListIndex = 0 Then
                           data_lla.Recordset("codmot") = "V"
                           data_lla.Recordset("descol") = "VERDE"
                        Else
                           If cbocolor.ListIndex = 1 Then
                              data_lla.Recordset("codmot") = "A"
                              data_lla.Recordset("descol") = "AMARILLO"
                           Else
                              If cbocolor.ListIndex = 2 Then
                                 data_lla.Recordset("codmot") = "R"
                                 data_lla.Recordset("descol") = "ROJO"
                              Else
                                 If cbocolor.ListIndex = 3 Then
                                    data_lla.Recordset("codmot") = "C"
                                    data_lla.Recordset("descol") = "CELESTE"
                                 Else
                                    If cbocolor.ListIndex = 4 Then
                                       data_lla.Recordset("codmot") = "Z"
                                       data_lla.Recordset("descol") = "AZUL"
                                    Else
                                       If cbocolor.ListIndex = 5 Then
                                          data_lla.Recordset("codmot") = "N"
                                          data_lla.Recordset("descol") = "NEGRO"
                                       Else
                                          data_lla.Recordset("codmot") = "V"
                                          data_lla.Recordset("descol") = "VERDE"
                                          cbocolor.ListIndex = 0
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                        If txt_movil.Text <> "" Then
                           data_lla.Recordset("movilpas") = txt_movil.Text
                           If txt_movil.Text <> 0 Then
                              If XWeltipoU = "USUARIOS DESP" Or XWeltipoU = "ADMINISTRADOR" Then
                                 data_lla.Recordset("timdes") = Trim(WElusuario)
                              End If
                              data_lla.Recordset("pend") = 1
                              If mtd.Text = "__/__/____" Then
                                 data_lla.Recordset("pend") = 1
                              Else
                                 data_lla.Recordset("fec_rea") = Format(mtd.Text, "dd/mm/yyyy")
                                 data_lla.Recordset("pend") = 2
                              End If
                              If txt_hortd.Text = "__:__" Then
                                 data_lla.Recordset("pend") = 1
                              Else
                                 data_lla.Recordset("hor_rea") = txt_hortd.Text
                                 data_lla.Recordset("pend") = 2
                              End If
                           Else
                              data_lla.Recordset("pend") = 0
                           End If
                        Else
                           data_lla.Recordset("pend") = 0
                           data_lla.Recordset("movilpas") = 0
                        End If
                        If mfecasig.Text = "__/__/____" Then
                           data_lla.Recordset("pend") = 0
                        Else
                           data_lla.Recordset("fecpas") = Format(mfecasig.Text, "dd/mm/yyyy")
                        End If
                        If txt_horasig.Text <> "" Then
                           data_lla.Recordset("horpas") = txt_horasig.Text
                        Else
                           data_lla.Recordset("pend") = 0
                        End If
                        If msalida.Text = "__/__/____" Then
                        Else
                           data_lla.Recordset("fecsali") = Format(msalida.Text, "dd/mm/yyyy")
                        End If
                        data_lla.Recordset("horsali") = txt_horsal.Text
                        If mllegada.Text = "__/__/____" Then
                        Else
                           data_lla.Recordset("fec_llega") = Format(mllegada.Text, "dd/mm/yyyy")
                        End If
                        data_lla.Recordset("hor_llega") = txt_horlle.Text
                        data_lla.Recordset("diag") = txt_diag.Text
                        If cbocolfin.ListIndex = 0 Then
                           data_lla.Recordset("colormot") = "V"
                        Else
                           If cbocolfin.ListIndex = 1 Then
                              data_lla.Recordset("colormot") = "A"
                           Else
                              If cbocolfin.ListIndex = 2 Then
                                 data_lla.Recordset("colormot") = "R"
                              Else
                                 If cbocolfin.ListIndex = 3 Then
                                    data_lla.Recordset("colormot") = "N"
                                 End If
                              End If
                           End If
                        End If
                        If txt_codmed.Text = "" Then
                           data_lla.Recordset("codmed") = 0
                        Else
                           data_lla.Recordset("codmed") = txt_codmed.Text
                        End If
                        data_lla.Recordset("nommed") = dbcbomed.Text
                        data_lla.Recordset("trasla") = cbotras.ListIndex
                        data_lla.Recordset("lugar") = Mid(txt_lugar.Text, 1, 35)
                        data_lla.Recordset("hsald") = txt_trassal.Text
                        data_lla.Recordset("hllega") = txt_enca.Text
                        data_lla.Recordset("hzona") = txt_enzona.Text
                        data_lla.Recordset("obs") = txt_obs.Text
                        If txt_movtra.Text = "" Then
                           data_lla.Recordset("movtras") = 0
                        Else
                           If IsNumeric(txt_movtra.Text) = False Then
                              data_lla.Recordset("movtras") = 0
                           Else
                              data_lla.Recordset("movtras") = txt_movtra.Text
                           End If
                        End If
                        data_lla.Recordset("totdem") = txt_demora.Text
                        If txt_salca.Text = "" Then
                        Else
                           data_lla.Recordset("hor_cance") = txt_salca.Text
                        End If
                        If Combo1.Visible = True Then
                           If Combo1.Text <> "" Then
                              data_lla.Recordset("dcobr") = Combo1.Text
                           Else
                              data_lla.Recordset("dcobr") = ""
                           End If
                        End If
                        If txt_locali.Text <> "" Then
                           data_lla.Recordset("motmov") = txt_locali.Text
                        End If
                        If txt_queb.Text = "" Then
                           txt_queb.Text = 0
                        End If
                        data_lla.Recordset("ncobr") = txt_queb.Text
                        If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                           If Label41.Caption = "" Then
                              Label41.Caption = 0
                           End If
                           If Label42.Caption = "" Then
                              Label42.Caption = 0
                           End If
                           If Label43.Caption = "" Then
                              Label43.Caption = 0
                           End If
                           If Label44.Caption = "" Then
                              Label44.Caption = 0
                           End If
                           If Label45.Caption = "" Then
                              Label45.Caption = 0
                           End If
                           If Label46.Caption = "" Then
                              Label46.Caption = -1
                           End If
                           If Label48.Caption = "" Then
                              Label48.Caption = 0
                           End If
                           If txt_quien.Text <> "" Then
                              data_lla.Recordset("motcance") = txt_quien.Text
                           End If
                           data_lla.Recordset("mm") = Label41.Caption
                           data_lla.Recordset("thh") = Label42.Caption
                           data_lla.Recordset("tmm") = Label43.Caption
                           data_lla.Recordset("pasado") = Label44.Caption
                           data_lla.Recordset("ano") = Label45.Caption
                           data_lla.Recordset("mes") = Label46.Caption
                           data_lla.Recordset("timsi") = Trim(str(Val(Label48.Caption)))
                        End If
                        If txt_costo.Text <> "" Then
                           If txt_costo.Text > 0 Then
                              data_lla.Recordset("mes") = txt_costo.Text
                           End If
                        End If
                        If txt_boleta.Text <> "" Then
                           If txt_boleta.Text > 0 Then
                              data_lla.Recordset("ano") = txt_boleta.Text
                           End If
                        End If
                        If txt_codmedtra.Text <> "" Then
                           data_lla.Recordset("movil_rea") = txt_codmedtra.Text
                        Else
                           data_lla.Recordset("movil_rea") = 0
                        End If
                        data_lla.Recordset("enfer") = Check2.Value ' actos de enfermería
                        data_lla.Recordset("hh") = Combo3.ListIndex
                        If cbocolor.ListIndex = 4 Then
                           If UCase(txt_locali.Text) = "SAUCE" Then
                              data_lla.Recordset("mm") = 1
                           Else
                              If UCase(txt_locali.Text) = "TOLEDO" Then
                                 data_lla.Recordset("mm") = 2
                              Else
                                 If UCase(txt_locali.Text) = "CASARINO" Then
                                    data_lla.Recordset("mm") = 3
                                 Else
                                    If UCase(txt_locali.Text) = "SUAREZ" Then
                                       data_lla.Recordset("mm") = 4
                                    Else
                                       If UCase(txt_locali.Text) = "BARROS BLANCOS" Then
                                          data_lla.Recordset("mm") = 5
                                       Else
                                          If UCase(txt_locali.Text) = "PANDO" Then
                                             data_lla.Recordset("mm") = 6
                                          Else
                                             data_lla.Recordset("mm") = 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                        If Xdeudasi = 9 Or Wopscob = 1 And (cbocolor.Text = "VERDE" Or cbocolor.Text = "AZUL" Or cbocolor.Text = "NEGRO") And (Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Or Val(cbozona.Text) = 5 Or Val(cbozona.Text) = 6) And _
                           (txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Or txt_cat.Text = "UNIVS" Or _
                           txt_cat.Text = "UNNSAM" Or txt_cat.Text = "HEVANO" Or txt_cat.Text = "EVNSAM" Or _
                           txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM" Or txt_cat.Text = "GANOS" Or _
                           txt_cat.Text = "CASANO" Or txt_cat.Text = "CASNSA") Then
                           If Wopscob = 1 Then
                              
                              MsgBox "No tiene derechos en servicios no urgentes sin firmar carta. El llamado se guardará como CANCELADO.", vbCritical
                              data_lla.Recordset("fec_cance") = Format(Date, "dd/mm/yyyy")
                              data_lla.Recordset("hor_cance") = Format(Time, "HH:mm")
                              data_lla.Recordset("cancela") = 1
                              data_lla.Recordset("pend") = 2
                              data_lla.Recordset("motcance") = "NO REALIZA CARTA"
                              data_lla.Recordset("movilpas") = 99
                              data_lla.Recordset("timdes") = Trim(WElusuario)
                              data_lla.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                              data_lla.Recordset("hor_rea") = Format(Time, "HH:mm")
                              data_lla.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                              data_lla.Recordset("horpas") = Format(Time, "HH:mm")
                              data_lla.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                              data_lla.Recordset("horsali") = Format(Time, "HH:mm")
                              data_lla.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                              data_lla.Recordset("hor_llega") = Format(Time, "HH:mm")
                              data_lla.Recordset("diag") = "CANCELADO NO REALIZA CARTA"
                              data_lla.Recordset("editando") = 1
                              data_lla.Recordset.Update
                              GrabaNosapp
                              borra_ya
                              despuesdegraba
                              Xdeudasi = 0
                              Check1.Value = 0
                           Else
                              If Xdeudasi = 9 Then
                                 MsgBox "No tiene autorización. Se guardará el llamado como CANCELADO por DEUDAS.", vbInformation
                                 data_lla.Recordset("fec_cance") = Format(Date, "dd/mm/yyyy")
                                 data_lla.Recordset("hor_cance") = Format(Time, "HH:mm")
                                 data_lla.Recordset("cancela") = 1
                                 data_lla.Recordset("pend") = 2
                                 data_lla.Recordset("motcance") = "MOROSO"
                                 data_lla.Recordset("movilpas") = 99
                                 data_lla.Recordset("timdes") = Trim(WElusuario)
                                 data_lla.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                                 data_lla.Recordset("hor_rea") = Format(Time, "HH:mm")
                                 data_lla.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                                 data_lla.Recordset("horpas") = Format(Time, "HH:mm")
                                 data_lla.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                                 data_lla.Recordset("horsali") = Format(Time, "HH:mm")
                                 data_lla.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                                 data_lla.Recordset("hor_llega") = Format(Time, "HH:mm")
                                 data_lla.Recordset("diag") = "CANCELADO AUT.POR DEUDAS"
                                 data_lla.Recordset("editando") = 1
                                 data_lla.Recordset.Update
                                 borra_ya
                                 despuesdegraba
                                 Xdeudasi = 0
                                 Check1.Value = 0
                              End If
                           End If
                        Else
                           data_lla.Recordset.Update
                           XllevacostoAp = 0
                           If chcovid.Value = 1 Then
                              frm_seguimdesp.Show vbModal
                           End If
                           
                           If Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                              If Wopscob = 1 Then
                                 GrabaNosapp
                              End If
                           End If
                           If Check4.Value = 1 Then
                              data_llamod.Recordset.AddNew
                              data_llamod.Recordset("nro") = txt_nro.Text
                              data_llamod.Recordset("fecha") = mfecha.Text
                              data_llamod.Recordset("pasado") = Check4.Value
                              data_llamod.Recordset("movilpas") = txt_codmed2.Text
                              If dbcbomed.Text <> "" Then
                                 data_llamod.Recordset("nommed") = dbcbomed2.Text
                              End If
                              data_llamod.Recordset("hora") = Format(Time, "HH:mm")
                              data_llamod.Recordset("usuario") = WElusuario
                              data_llamod.Recordset.Update
                              data_llamod.Refresh
                           End If
                           If mftrassol.Text <> "__/__/____" And mhtrassol.Text <> "__:__" Then
                              data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                              data_llamod.Refresh
                              If data_llamod.Recordset.RecordCount > 0 Then
                                 data_llamod.Recordset.Edit
                                 data_llamod.Recordset("fec_llega") = mftrassol.Text
                                 data_llamod.Recordset("hor_llega") = mhtrassol.Text
                                 data_llamod.Recordset.Update
                              Else
                                 data_llamod.Recordset.AddNew
                                 data_llamod.Recordset("nro") = txt_nro.Text
                                 data_llamod.Recordset("fecha") = mfecha.Text
                                 data_llamod.Recordset("fec_llega") = mftrassol.Text
                                 data_llamod.Recordset("hor_llega") = mhtrassol.Text
                                 data_llamod.Recordset.Update
                              End If
                           End If
                           If Combo2.ListIndex >= 0 Then
                              data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                              data_llamod.Refresh
                              If data_llamod.Recordset.RecordCount > 0 Then
                                 If IsNull(data_llamod.Recordset("telef")) = True Then
                                    data_llamod.Recordset.Edit
                                    data_llamod.Recordset("telef") = Combo2.Text
                                    data_llamod.Recordset.Update
                                 Else
                                    If data_llamod.Recordset("telef") <> Combo2.Text Then
                                       data_llamod.Recordset.Edit
                                       data_llamod.Recordset("telef") = Combo2.Text
                                       data_llamod.Recordset.Update
                                    End If
                                 End If
                              Else
                                 data_llamod.Recordset.AddNew
                                 data_llamod.Recordset("nro") = txt_nro.Text
                                 data_llamod.Recordset("fecha") = mfecha.Text
                                 data_llamod.Recordset("telef") = Combo2.Text
                                 data_llamod.Recordset.Update
                              End If
                           Else
                              data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                              data_llamod.Refresh
                              If data_llamod.Recordset.RecordCount > 0 Then
                                 If IsNull(data_llamod.Recordset("telef")) = False Then
                                    data_llamod.Recordset.Edit
                                    data_llamod.Recordset("telef") = Null
                                    data_llamod.Recordset.Update
                                 End If
                              End If
                           End If
                           If labcodchof.Caption <> "" Then
                              data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                              data_llamod.Refresh
                              If data_llamod.Recordset.RecordCount > 0 Then
                                 data_llamod.Recordset.Edit
                                 data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                                 data_llamod.Recordset.Update
                              Else
                                 data_llamod.Recordset.AddNew
                                 data_llamod.Recordset("nro") = txt_nro.Text
                                 data_llamod.Recordset("fecha") = mfecha.Text
                                 data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                                 data_llamod.Recordset.Update
                              End If
                           End If
                           If t_codced.Text <> "" Then
                              data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                              data_llamod.Refresh
                              If data_llamod.Recordset.RecordCount > 0 Then
                                 If IsNull(data_llamod.Recordset("mes")) = False Then
                                    If data_llamod.Recordset("mes") <> t_codced.Text Then
                                       data_llamod.Recordset.Edit
                                       data_llamod.Recordset("mes") = t_codced.Text
                                       data_llamod.Recordset.Update
                                    End If
                                 Else
                                    data_llamod.Recordset.Edit
                                    data_llamod.Recordset("mes") = t_codced.Text
                                    data_llamod.Recordset.Update
                                 End If
                              Else
                                 data_llamod.Recordset.AddNew
                                 data_llamod.Recordset("mes") = t_codced.Text
                                 data_llamod.Recordset("nro") = txt_nro.Text
                                 data_llamod.Recordset("fecha") = mfecha.Text
                                 data_llamod.Recordset.Update
                              End If
                           Else
                              data_llamod.Recordset.AddNew
                              data_llamod.Recordset("mes") = 0
                              data_llamod.Recordset("nro") = txt_nro.Text
                              data_llamod.Recordset("fecha") = mfecha.Text
                              data_llamod.Recordset.Update
                           End If
                           Xdeudasi = 0
                           Check1.Value = 0
                           historial
                           despuesdegraba
                        
                        End If
        '                Graba_aft
                        If txt_codmed2.Text = "" Then
                           txt_codmed2.Text = 0
                        End If
                        If XAlta = 1 Then
                           If txt_ced.Text <> "" And txt_mot.Text <> "" Then
                              If Val(txt_ced.Text) > 0 Then
                                   data_cons.Recordset.AddNew
                                   data_cons.Recordset("mat") = txt_ced.Text
                                   data_cons.Recordset("motivo") = txt_mot.Text
                                   data_cons.Recordset("fecha") = Date
                                   data_cons.Recordset.Update
                              End If
                           End If
                        End If
                        XAlta = 3
                  Else
                    MsgBox "Ingrese COLOR de llamado", vbCritical, "Mensaje"
                    cbocolor.SetFocus
                  End If
               Else
                  MsgBox "Ingrese Zona", vbCritical, "Mensaje"
                  cbozona.SetFocus
               End If
        ''''aqui
           End If
       Else
           MsgBox "ATENCION! Ya existe un llamado pendiente con esta cédula.", vbCritical
       End If
    Else
       If cbocolor.Text = "" Then
          cbocolor.ListIndex = 0
       End If
       If cbozona.Text <> "" Then
          If cbocolor.Text <> "" Then
    ''         ControlCosto
            If data_lla.Recordset("nrolla") = txt_nro.Text Then
                data_lla.Recordset.Edit
                
                data_lla.Recordset("editando") = 1
                data_lla.Recordset("usuario_edit") = Null
                data_lla.Recordset("fecha") = Format(mfecha.Text, "dd/mm/yyyy")
                data_lla.Recordset("hora") = txt_hora.Text
                data_lla.Recordset("usuario") = txt_usua.Text
                data_lla.Recordset("segui_covid") = chcovid.Value
                data_lla.Recordset("timbre") = cbotimbre.ListIndex
                If Trim(t_timbre.Text) <> "" Then
                   data_lla.Recordset("valor_timbre") = t_timbre.Text
                End If
                data_lla.Recordset("nomodif") = Check1.Value
                If Label3.Caption <> data_lla.Recordset("activo") Then
                   data_lla.Recordset("activo") = Format(Label3.Caption, "HH:mm:ss")
                End If
                If txt_mat.Text = "" Then
                   data_lla.Recordset("matric") = 0
                Else
                   data_lla.Recordset("matric") = txt_mat.Text
                End If
                data_lla.Recordset("nombre") = txt_nomb.Text
                If txt_edad = "" Then
                   data_lla.Recordset("edad") = 0
                Else
                   data_lla.Recordset("edad") = txt_edad.Text
                End If
                If cboed.ListIndex = 0 Then
                   data_lla.Recordset("unied") = 3
                Else
                   If cboed.ListIndex = 1 Then
                      data_lla.Recordset("unied") = 2
                   Else
                      If cboed.ListIndex = 2 Then
                         data_lla.Recordset("unied") = 1
                      Else
                         data_lla.Recordset("unied") = 3
                      End If
                   End If
                End If
                If txt_cat.Text = "" Then
                   txt_cat.Text = "AAABBB"
                End If
                data_lla.Recordset("categ") = txt_cat.Text
                data_lla.Recordset("nomcat") = txt_nomcat.Text
                If txt_ced.Text = "" Then
                   data_lla.Recordset("ci") = 0
                Else
                   data_lla.Recordset("ci") = txt_ced.Text
                End If
                data_lla.Recordset("telef") = txt_tel.Text
                If cbozona.ListIndex >= 0 Then
                   data_lla.Recordset("codzon") = Val(cbozona.Text)
                Else
                   data_lla.Recordset("codzon") = 1
                End If
                If cbobase.Text = "" Then
                   data_lla.Recordset("base") = 0
                Else
                   data_lla.Recordset("base") = cbobase.Text
                End If
                data_lla.Recordset("referen") = txt_direc.Text
                data_lla.Recordset("realiza") = chtmut.Value
                If txt_ante.Text <> "" Then
                   data_lla.Recordset("motcon") = txt_ante.Text
                End If
                data_lla.Recordset("obsmot") = txt_mot.Text
                If cbocolor.ListIndex = 0 Then
                   data_lla.Recordset("codmot") = "V"
                   data_lla.Recordset("descol") = "VERDE"
                Else
                   If cbocolor.ListIndex = 1 Then
                      data_lla.Recordset("codmot") = "A"
                      data_lla.Recordset("descol") = "AMARILLO"
                   Else
                      If cbocolor.ListIndex = 2 Then
                         data_lla.Recordset("codmot") = "R"
                         data_lla.Recordset("descol") = "ROJO"
                      Else
                         If cbocolor.ListIndex = 3 Then
                            data_lla.Recordset("codmot") = "C"
                            data_lla.Recordset("descol") = "CELESTE"
                         Else
                            If cbocolor.ListIndex = 4 Then
                               data_lla.Recordset("codmot") = "Z"
                               data_lla.Recordset("descol") = "AZUL"
                            Else
                               If cbocolor.ListIndex = 5 Then
                                  data_lla.Recordset("codmot") = "N"
                                  data_lla.Recordset("descol") = "NEGRO"
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
                If chcovid.Value = 1 Then
                   data_lla.Recordset("pend") = 2
                Else
                    If txt_movil.Text <> "" Then
                       data_lla.Recordset("movilpas") = txt_movil.Text
                       If txt_movil.Text <> 0 Then
                          data_lla.Recordset("pend") = 1
                          If XWeltipoU = "USUARIOS DESP" Or XWeltipoU = "ADMINISTRADOR" Then
                             data_lla.Recordset("timdes") = Trim(WElusuario)
                          End If
                          If mtd.Text = "__/__/____" Then
                             data_lla.Recordset("fec_rea") = Null
                             data_lla.Recordset("pend") = 1
                             Xestapen = 1
                          Else
                             data_lla.Recordset("fec_rea") = Format(mtd.Text, "dd/mm/yyyy")
                             If cbocolfin.Text = "" Then
                                data_lla.Recordset("pend") = 1
                                Xestapen = 1
                             Else
                                data_lla.Recordset("pend") = 2
                                Xestapen = 2
                             End If
                          End If
                          If txt_hortd.Text = "__:__" Then
                             data_lla.Recordset("hor_rea") = ""
                             data_lla.Recordset("pend") = 1
                             Xestapen = 1
                          Else
                             data_lla.Recordset("hor_rea") = txt_hortd.Text
                             If cbocolfin.Text = "" Then
                                data_lla.Recordset("pend") = 1
                                Xestapen = 1
                             Else
                                data_lla.Recordset("pend") = 2
                                Xestapen = 2
                             End If
                          End If
                       Else
                          If data_lla.Recordset("pend") = 4 Then
                          Else
                             data_lla.Recordset("pend") = 0
                          End If
                       End If
                    Else
                       If data_lla.Recordset("pend") <> 4 Then
                          data_lla.Recordset("pend") = 0
                          data_lla.Recordset("movilpas") = 0
                       End If
                    End If
                End If
    'aft
                If Label40.Caption <> "" Then
                   If IsNull(data_lla.Recordset("aft")) = False Then
                      If Trim(XaftC) <> "" Then
                         If data_lla.Recordset("aft") <> XaftC Then
                            data_lla.Recordset("aft") = XaftC
                         End If
                      End If
                   Else
                      If Trim(XaftC) <> "" Then
                         data_lla.Recordset("aft") = XaftC
                      End If
                   End If
                End If
                If mfecasig.Text = "__/__/____" Then
                   data_lla.Recordset("fecpas") = Null
                   If data_lla.Recordset("pend") <> 4 Then
                      data_lla.Recordset("pend") = 0
                   End If
                   Xestapen = 0
                Else
                   data_lla.Recordset("fecpas") = Format(mfecasig.Text, "dd/mm/yyyy")
                End If
                If txt_horasig.Text <> "" Then
                   data_lla.Recordset("horpas") = txt_horasig.Text
                Else
                   data_lla.Recordset("horpas") = ""
                   If data_lla.Recordset("pend") <> 4 Then
                      data_lla.Recordset("pend") = 0
                   End If
                   Xestapen = 0
                End If
                If msalida.Text = "__/__/____" Then
                   data_lla.Recordset("fecsali") = Null
                   If Xestapen = 2 Then
                      data_lla.Recordset("pend") = 1
                   End If
                Else
                   data_lla.Recordset("fecsali") = Format(msalida.Text, "dd/mm/yyyy")
                End If
                data_lla.Recordset("horsali") = txt_horsal.Text
                If mllegada.Text = "__/__/____" Then
                   data_lla.Recordset("fec_llega") = Null
                   If Xestapen = 2 Then
                      data_lla.Recordset("pend") = 1
                   End If
                Else
                   data_lla.Recordset("fec_llega") = Format(mllegada.Text, "dd/mm/yyyy")
                End If
                data_lla.Recordset("hor_llega") = txt_horlle.Text
                If mtd.Text = "__/__/____" Then
                   data_lla.Recordset("fec_rea") = Null
                   If Xestapen = 2 Then
                      data_lla.Recordset("pend") = 1
                   End If
                Else
                   data_lla.Recordset("fec_rea") = Format(mtd.Text, "dd/mm/yyyy")
                End If
                If txt_hortd.Text = "__:__" Then
                   data_lla.Recordset("hor_rea") = ""
                Else
                   data_lla.Recordset("hor_rea") = txt_hortd.Text
                End If
                data_lla.Recordset("diag") = txt_diag.Text
                If cbocolfin.ListIndex = 0 Then
                   data_lla.Recordset("colormot") = "V"
                Else
                   If cbocolfin.ListIndex = 1 Then
                      data_lla.Recordset("colormot") = "A"
                   Else
                      If cbocolfin.ListIndex = 2 Then
                         data_lla.Recordset("colormot") = "R"
                      Else
                         If cbocolfin.ListIndex = 3 Then
                            data_lla.Recordset("colormot") = "N"
                         End If
                      End If
                   End If
                End If
                If txt_codmed.Text = "" Then
                   data_lla.Recordset("codmed") = 0
                Else
                   data_lla.Recordset("codmed") = txt_codmed.Text
                End If
                If Xestapen >= 1 Then
                   If txt_obs.Text <> "" Then
                      txt_obs.Text = txt_obs.Text + "-"
                   Else
                      txt_obs.Text = "-"
                   End If
                End If
                data_lla.Recordset("obs") = txt_obs.Text
                data_lla.Recordset("nommed") = dbcbomed.Text
                data_lla.Recordset("trasla") = cbotras.ListIndex
                data_lla.Recordset("lugar") = Mid(txt_lugar.Text, 1, 35)
                data_lla.Recordset("hsald") = txt_trassal.Text
                data_lla.Recordset("hllega") = txt_enca.Text
                data_lla.Recordset("hzona") = txt_enzona.Text
                If txt_movtra.Text = "" Then
                   data_lla.Recordset("movtras") = 0
                Else
                   If IsNumeric(txt_movtra.Text) = False Then
                      data_lla.Recordset("movtras") = 0
                   Else
                      data_lla.Recordset("movtras") = txt_movtra.Text
                   End If
                End If
                If mllegada.Text <> "__/__/____" Then
                   If txt_horlle.Text <> "" Then
                      If txt_hora.Text <> "" Then
                         If Format(mfecha.Text, "dd/mm/yyyy") = Format(mllegada.Text, "dd/mm/yyyy") Then
                            Xhd = Val(Mid(txt_hora.Text, 1, 2))
                            xmd = Val(Mid(txt_hora.Text, 4, 2))
                            xha = Val(Mid(txt_horlle.Text, 1, 2))
                            xma = Val(Mid(txt_horlle.Text, 4, 2))
                            xdemh = xha - Xhd
                            xdemm = xma - xmd
                            If xdemh = 0 Then
                               xdemm = xma - xmd
                            Else
                               If xdemh = 1 Then
                                  If xdemm >= 0 Then
                                     xdemh = 1
                                  Else
                                     xdemh = 0
                                     xdemm = xdemm + 60
                                  End If
                               Else
                                  If xdemm >= 0 Then
                                     xdemh = xdemh
                                  Else
                                     xdemh = xdemh - 1
                                     xdemm = xdemm + 60
                                  End If
                               End If
                            End If
                         Else
                            Xhd = Val(Mid(txt_hora.Text, 1, 2))
                            xmd = Val(Mid(txt_hora.Text, 4, 2))
                            xha = Val(Mid(txt_horlle.Text, 1, 2))
                            xma = Val(Mid(txt_horlle.Text, 4, 2))
                            xdemh = xha - Xhd
                            xdemm = xma - xmd
                            If xdemh = 0 Then
        '                       xdemh = 24
                               If xdemm >= 0 Then
                                  xdemh = xdemh
                               Else
        '                          xdemh = xdemh - 1
        '                          xdemm = xmd - xma
                                  xdemm = xdemm + 60
                               End If
                            Else
                               If xdemh < 0 Then
                                  xdemh = xdemh + 24
                                  If xdemm >= 0 Then
                                     xdemh = xdemh + 1
                                  Else
                                     xdemh = xdemh - 1
                                     xdemm = xdemm + 60
                                  End If
                               Else
                                  Xhd = Xhd + 24
                                  xdemh = Xhd
                                  If xdemm >= 0 Then
                                     xdemh = xdemh
                                  Else
                                     xdemh = xdemh - 1
                                     xdemm = xdemm + 60
                                  End If
                               End If
                            End If
                         End If
                         txt_demora.Text = Format(Trim(str(xdemh)) + ":" + Trim(str(xdemm)), "HH:mm")
                         data_lla.Recordset("totdem") = txt_demora.Text
                      End If
                   End If
                End If
                If mtd.Text <> "__/__/____" Then
                   If txt_hortd.Text <> "__:__" Then
                      If txt_horlle.Text <> "" Then
                         If Format(mtd.Text, "dd/mm/yyyy") = Format(mllegada.Text, "dd/mm/yyyy") Then
                            Xhd = Val(Mid(txt_horlle.Text, 1, 2))
                            xmd = Val(Mid(txt_horlle.Text, 4, 2))
                            xha = Val(Mid(txt_hortd.Text, 1, 2))
                            xma = Val(Mid(txt_hortd.Text, 4, 2))
                            xdemh = xha - Xhd
                            xdemm = xma - xmd
                            If xdemh = 0 Then
                               xdemm = xmd - xma
                            Else
                               If xdemh = 1 Then
                                  If xdemm >= 0 Then
                                     xdemh = 1
                                  Else
                                     xdemh = 0
                                     xdemm = xdemm + 60
                                  End If
                               Else
                                  If xdemm >= 0 Then
                                     xdemh = xdemh
                                  Else
                                     xdemh = xdemh - 1
                                     xdemm = xdemm + 60
                                  End If
                               End If
                            End If
                         Else
                            Xhd = Val(Mid(txt_horlle.Text, 1, 2))
                            xmd = Val(Mid(txt_horlle.Text, 4, 2))
                            xha = Val(Mid(txt_hortd.Text, 1, 2))
                            xma = Val(Mid(txt_hortd.Text, 4, 2))
                            xdemh = xha - Xhd
                            xdemm = xma - xmd
                            If xdemh = 0 Then
                         '      xdemh = 24
                               If xdemm >= 0 Then
                                  xdemh = xdemh + 1
                               Else
                         '         xdemh = xdemh - 1
                                  xdemm = xdemm + 60
                               End If
                            Else
                               If xdemh < 0 Then
                                  xdemh = Xhd - xha
                                  xdemh = xdemh + 24
                                  If xdemm >= 0 Then
                                     xdemh = xdemh + 1
                                  Else
                                     xdemh = xdemh - 1
                                     xdemm = xdemm + 60
                                  End If
                               Else
                                  Xhd = Xhd + 24
                                  xdemh = Xhd
                                  If xdemm >= 0 Then
                                     xdemh = xdemh
                                  Else
                                     xdemh = xdemh - 1
                                     xdemm = xdemm + 60
                                  End If
                               End If
                            End If
                         End If
                         Text2.Text = Format(Trim(str(xdemh)) + ":" + Trim(str(xdemm)), "HH:mm")
                      End If
                   End If
                End If
                If txt_movil.Text <> "" Then
                   If txt_movil.Text > 0 And txt_horasig.Text <> "" Then
                      If XWeltipoU = "USUARIOS DESP" Then
                         If IsNull(data_lla.Recordset("timdes")) = False Then
                         Else
                            data_lla.Recordset("timdes") = WElusuario
                         End If
                      End If
                   End If
                End If
                If Combo1.Text <> "" Then
                   data_lla.Recordset("dcobr") = Combo1.Text
                Else
                   data_lla.Recordset("dcobr") = ""
                End If
                data_lla.Recordset("enfer") = Check2.Value
                If txt_locali.Text <> "" Then
                   data_lla.Recordset("motmov") = txt_locali.Text
                End If
                If txt_queb.Text = "" Then
                   txt_queb.Text = 0
                End If
                data_lla.Recordset("ncobr") = txt_queb.Text
                If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                   If Label41.Caption = "" Then
                      Label41.Caption = 0
                   End If
                   If Label42.Caption = "" Then
                      Label42.Caption = 0
                   End If
                   If Label43.Caption = "" Then
                      Label43.Caption = 0
                   End If
                   If Label44.Caption = "" Then
                      Label44.Caption = 0
                   End If
                   If Label45.Caption = "" Then
                      Label45.Caption = 0
                   End If
                   If Label46.Caption = "" Then
                      Label46.Caption = -1
                   End If
                   If Label48.Caption = "" Then
                      Label48.Caption = 0
                   End If
                   If txt_quien.Text <> "" Then
                      If IsNull(data_lla.Recordset("cancela")) = True Then
                         data_lla.Recordset("motcance") = txt_quien.Text
                      Else
                         If data_lla.Recordset("cancela") <> 1 Then
                            data_lla.Recordset("motcance") = txt_quien.Text
                         Else
                            data_lla.Recordset("obs") = data_lla.Recordset("obs") + " " + data_lla.Recordset("motcance")
                            data_lla.Recordset("motcance") = txt_quien.Text
                         End If
                      End If
                   End If
                   data_lla.Recordset("mm") = Label41.Caption
                   data_lla.Recordset("thh") = Label42.Caption
                   data_lla.Recordset("tmm") = Label43.Caption
                   data_lla.Recordset("pasado") = Label44.Caption
                   data_lla.Recordset("ano") = Label45.Caption
                   If IsNull(data_lla.Recordset("mes")) = True Then
                      data_lla.Recordset("mes") = Label46.Caption
                   Else
                      If data_lla.Recordset("mes") <= 10 Then
                         data_lla.Recordset("mes") = Label46.Caption
                      End If
                   End If
                   data_lla.Recordset("timsi") = Label48.Caption
                End If
                If txt_costo.Text <> "" Then
                   If txt_costo.Text > 0 Then
                      data_lla.Recordset("mes") = txt_costo.Text
                   Else
                      data_lla.Recordset("mes") = 0
                   End If
                Else
                   data_lla.Recordset("mes") = 0
                End If
                If txt_boleta.Text <> "" Then
                   If txt_boleta.Text > 0 Then
                      data_lla.Recordset("ano") = txt_boleta.Text
                   End If
                End If
                If txt_salca.Text = "" Then
                Else
                   If IsNull(data_lla.Recordset("cancela")) = True Then
                      data_lla.Recordset("hor_cance") = txt_salca.Text
                   End If
                End If
                If txt_codmedtra.Text <> "" Then
                   data_lla.Recordset("movil_rea") = txt_codmedtra.Text
                Else
                   data_lla.Recordset("movil_rea") = 0
                End If
                data_lla.Recordset("hh") = Combo3.ListIndex
                If cbocolor.ListIndex = 4 Then
                   If UCase(txt_locali.Text) = "SAUCE" Then
                      data_lla.Recordset("mm") = 1
                   Else
                      If UCase(txt_locali.Text) = "TOLEDO" Then
                         data_lla.Recordset("mm") = 2
                      Else
                         If UCase(txt_locali.Text) = "CASARINO" Then
                            data_lla.Recordset("mm") = 3
                         Else
                            If UCase(txt_locali.Text) = "SUAREZ" Then
                               data_lla.Recordset("mm") = 4
                            Else
                               If UCase(txt_locali.Text) = "BARROS BLANCOS" Then
                                  data_lla.Recordset("mm") = 5
                               Else
                                  If UCase(txt_locali.Text) = "PANDO" Then
                                     data_lla.Recordset("mm") = 6
                                  Else
                                     data_lla.Recordset("mm") = 1
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
                data_lla.Recordset.Update
                If chcovid.Value = 1 Then
                   SioNocovid = MsgBox("Desea modificar fecha de próximo seguimiento?", vbYesNo + vbInformation, "Seguimiento Codiv-19")
                   If SioNocovid = vbYes Then
                      frm_seguimdesp.Show vbModal
                   End If
                End If
                
                Check1.Value = 0
                data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                data_llamod.Refresh
                If txt_codmed2.Text = "" Then
                   txt_codmed2.Text = 0
                End If
                If data_llamod.Recordset.RecordCount > 0 Then
                   If IsNull(data_llamod.Recordset("fecha")) = False Then
                      If Format(data_llamod.Recordset("fecha"), "yyyy/mm/dd") = Format(mfecha.Text, "yyyy/mm/dd") Then
                      Else
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("fecha") = mfecha.Text
                         data_llamod.Recordset.Update
                      End If
                   Else
                      data_llamod.Recordset.Edit
                      data_llamod.Recordset("fecha") = mfecha.Text
                      data_llamod.Recordset.Update
                   End If
                   If Combo2.ListIndex >= 0 Then
                      If IsNull(data_llamod.Recordset("telef")) = True Then
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("telef") = Combo2.Text
                         data_llamod.Recordset.Update
                      Else
                         If data_llamod.Recordset("telef") <> Combo2.Text Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("telef") = Combo2.Text
                            data_llamod.Recordset.Update
                         End If
                      End If
                   Else
                      If IsNull(data_llamod.Recordset("telef")) = False Then
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("telef") = Null
                         data_llamod.Recordset.Update
                      End If
                   End If
                   If t_codced.Text <> "" Then
                      If IsNull(data_llamod.Recordset("mes")) = False Then
                         If data_llamod.Recordset("mes") <> t_codced.Text Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("mes") = t_codced.Text
                            data_llamod.Recordset.Update
                         End If
                      Else
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("mes") = t_codced.Text
                         data_llamod.Recordset.Update
                      End If
                   Else
                      data_llamod.Recordset.AddNew
                      data_llamod.Recordset("mes") = 0
                      data_llamod.Recordset.Update
                   End If
                   If mftrassol.Text <> "__/__/____" And mhtrassol.Text <> "__:__" Then
                      If IsNull(data_llamod.Recordset("fec_llega")) = False Then
                         If Format(data_llamod.Recordset("fec_llega"), "dd/mm/yyyy") <> Format(mftrassol.Text, "dd/mm/yyyy") Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("fec_llega") = mftrassol.Text
                            data_llamod.Recordset.Update
                         End If
                         If Format(data_llamod.Recordset("hor_llega"), "HH:mm") <> Format(mhtrassol.Text, "HH:mm") Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("hor_llega") = mhtrassol.Text
                            data_llamod.Recordset.Update
                         End If
                      Else
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("fec_llega") = mftrassol.Text
                         data_llamod.Recordset("hor_llega") = mhtrassol.Text
                         data_llamod.Recordset.Update
                      End If
                   Else
                      If IsNull(data_llamod.Recordset("fec_llega")) = False Then
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("fec_llega") = Null
                         data_llamod.Recordset("hor_llega") = Null
                         data_llamod.Recordset.Update
                      End If
                   End If
                   If IsNull(data_llamod.Recordset("pasado")) = False Then
                      If data_llamod.Recordset("pasado") = Check4.Value Then
                      Else
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("pasado") = Check4.Value
                         data_llamod.Recordset.Update
                      End If
                   Else
                      data_llamod.Recordset.Edit
                      data_llamod.Recordset("pasado") = Check4.Value
                      data_llamod.Recordset.Update
                   End If
                   If IsNull(data_llamod.Recordset("movilpas")) = False Then
                      If data_llamod.Recordset("movilpas") = txt_codmed2.Text Then
                      Else
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("movilpas") = txt_codmed2.Text
                         data_llamod.Recordset("nommed") = dbcbomed2.Text
                         data_llamod.Recordset.Update
                      End If
                   Else
                      If txt_codmed2.Text = "" Then
                      Else
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("movilpas") = txt_codmed2.Text
                         data_llamod.Recordset("nommed") = dbcbomed2.Text
                         data_llamod.Recordset.Update
                      End If
                   End If
                   If labcodchof.Caption <> "" Then
                      If IsNull(data_llamod.Recordset("movil_rea")) = False Then
                         If data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption) Then
                         Else
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                            data_llamod.Recordset.Update
                         End If
                      Else
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                         data_llamod.Recordset.Update
                      End If
                   End If
                Else
                    data_llamod.Recordset.AddNew
                    data_llamod.Recordset("nro") = txt_nro.Text
                    data_llamod.Recordset("fecha") = mfecha.Text
                    data_llamod.Recordset("pasado") = Check4.Value
                    If Combo2.ListIndex >= 0 Then
                       data_llamod.Recordset("telef") = Combo2.Text
                    End If
                    If t_codced.Text <> "" Then
                       data_llamod.Recordset("mes") = t_codced.Text
                    Else
                       data_llamod.Recordset("mes") = 0
                    End If
                    If txt_codmed2.Text <> "" Then
                       data_llamod.Recordset("movilpas") = txt_codmed2.Text
                    End If
                    If dbcbomed2.Text <> "" Then
                       data_llamod.Recordset("nommed") = dbcbomed2.Text
                    End If
                    data_llamod.Recordset("hora") = Format(Time, "HH:mm")
                    data_llamod.Recordset("usuario") = WElusuario
                    If labcodchof.Caption <> "" Then
                       data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                    End If
                    data_llamod.Recordset.Update
                End If
                data_hist.Recordset.AddNew
                data_hist.Recordset("idllamado") = txt_nro.Text
                data_hist.Recordset("fecha") = Date
                data_hist.Recordset("hora") = Format(Time, "HH:mm")
                data_hist.Recordset("usuario") = WElusuario
                data_hist.Recordset("accion") = "MODIF. LLAMADO"
                data_hist.Recordset("categ") = txt_cat.Text
                data_hist.Recordset("claveini") = cbocolor.Text
                data_hist.Recordset.Update
                Xestapen = 0
                XAlta = 3
                historial
                despuesdegraba
            Else
                MsgBox "Error al grabar modificación, reingrese datos y vuelva a grabar!", vbCritical
                b_cancela_Click
            End If
        
          Else
            MsgBox "Ingrese COLOR de llamado", vbCritical, "Mensaje"
            cbocolor.SetFocus
          End If
       Else
          MsgBox "Ingrese Zona", vbCritical, "Mensaje"
          cbozona.SetFocus
       End If
    End If
Else
    MsgBox "ERROR en la cédula, VERIFIQUE!!", vbCritical
End If
chcovid.Enabled = False

Exit Sub

errorws:
    LogError Err.Number, Err.Description, "btgrabar_Click llamado despacho", Erl

'Quepasa:
'        If Err.Number = 3155 Then
'           MsgBox "Error 3155 al intentar grabar el registro, Verifique DATOS e intente grabar nuevamente o presione CANCELAR", vbInformation, "Mensaje"
'           b_grabar.Enabled = True
'        Else
'           If Err.Number = 3197 Then
'              b_grabar.Enabled = False
'           Else
'              MsgBox "ERROR: " & str(Err.Number) & Err.Description & " COMUNIQUE A INFORMATICA, y presione botón CANCELAR", vbInformation, "Mensaje"
'              b_grabar.Enabled = True
'           End If
'        End If
        
End Sub

Private Sub b_hist_Click()
'Dim Xmensahist As String
'Xmensahist = MsgBox("Desea ver el historial administrativo?", vbInformation + vbYesNo)
'If Xmensahist = vbYes Then
'   Wveohistoadmd = 8
'   frm_accadm.Show vbModal
'Else
'   Wveohistoadmd = 0
'   If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS DESP" Or XWeltipoU = "USUARIOS RECEP" Then
frm_histdesp.Show vbModal
'   Else
'      MsgBox "Usuario sin permisos"
'   End If
'End If

End Sub

Private Sub b_imp_Click()
'If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Then
   frm_largador.PrintForm
'Else
'   MsgBox "USUARIO NO AUTORIZADO A IMPRIMIR", vbCritical
   
'End If
End Sub

Private Sub b_modif_Click()
Dim XpuedeMod, XpuedeMod2 As Integer
Dim Xusuarioedit As String
Xusuarioedit = ""
XpuedeMod = 1
XpuedeMod2 = 0

On Error GoTo Almod

'If ControlUsuario("Utilitarios despacho") = 1 Then

    Frame1.Enabled = True
    
    XcolorAntAzul = "S"
    
    If WDespa <> 1 Then
       If XWeltipoU = "USUARIOS DESP" Then
          Frame2.Enabled = True
          If txt_cat.Text = "911" Or UCase(txt_cat.Text) = "911B" Then
             Command1.Enabled = True
          Else
             Command1.Enabled = False
          End If
          If txt_movil.Text = "" Then
             If Frame2.Visible = True Then
                txt_movil.SetFocus
             End If
          Else
             If txt_movil.Text > 0 Then
                If Frame2.Visible = True Then
                   txt_diag.SetFocus
                End If
             Else
                If Frame2.Visible = True Then
                   txt_movil.SetFocus
                End If
             End If
          End If
       Else
          Frame2.Enabled = False
          If txt_cat.Text = "911" Or UCase(txt_cat.Text) = "911B" Then
             Command1.Enabled = True
          Else
             Command1.Enabled = False
          End If
          
          txt_obs.SetFocus
       End If
    Else
       txt_nomb.SetFocus
    End If
    
    If txt_nro.Text <> "" Then
       If Val(txt_nro.Text) > 0 Then
            data_lla.RecordSource = "Select * from llamado where nrolla =" & txt_nro.Text
            data_lla.Refresh
            If data_lla.Recordset.RecordCount > 0 Then
               If IsNull(data_lla.Recordset("editando")) = False Then
                  If data_lla.Recordset("editando") = 0 Then
                     XpuedeMod2 = 9
                     If IsNull(data_lla.Recordset("usuario_edit")) = False Then
                        Xusuarioedit = data_lla.Recordset("usuario_edit")
                     Else
                        Xusuarioedit = ""
                     End If
                  Else
                     data_lla.Recordset.Edit
                     data_lla.Recordset("editando") = 0
                     data_lla.Recordset("usuario_edit") = WElusuario
                     data_lla.Recordset.Update
                     XpuedeMod2 = 0
                  End If
               Else
                  data_lla.Recordset.Edit
                  data_lla.Recordset("editando") = 0
                  data_lla.Recordset("usuario_edit") = WElusuario
                  data_lla.Recordset.Update
                  XpuedeMod2 = 0
               End If
               If IsNull(data_lla.Recordset("totend")) = False Then
                  If data_lla.Recordset("totend") = "FACT" Then
                     XpuedeMod = 9
                  Else
                     XpuedeMod = 1
                  End If
               Else
                  XpuedeMod = 1
               End If
               If XpuedeMod = 9 Or XpuedeMod2 = 9 Then
'                  If ControlUsuario("Modifica Despacho") = 1 And XpuedeMod2 = 0 Then
                  If ControlUsuario("Modifica Despacho") = 1 Then
                       borra_ya
                       igualar_sin
                     '  MsgBox "ATENCION! Este llamado ya fue facturado. No modifique datos de facturación.", vbCritical
                       If IsNull(data_lla.Recordset("cancela")) = False Then
                          If data_lla.Recordset("cancela") = 1 Then
                             MsgBox "El llamado figura CANCELADO, verifique!!", vbCritical
                          End If
                       End If
                       XAlta = 0
                       b_nuevo.Enabled = False
                       b_modif.Enabled = False
                       b_imp.Enabled = False
                       b_buscar.Enabled = False
                       b_grabar.Enabled = True
                       b_cancel.Enabled = False
                       b_cancela.Enabled = True
                       b_pend.Enabled = False
                       txt_nro.Enabled = False
                       mfecha.Enabled = False
                       txt_hora.Enabled = False
                       txt_usua.Enabled = False
                       b_hist.Enabled = False
                       Command2.Enabled = False
                       Command3.Enabled = False
                       If cbocolor.Text = "ROJO" Or cbocolor.Text = "AMARILLO" Then
                          If WElusuario = "MARTINC" Then
                             cbocolor.Enabled = True
                          Else
                             cbocolor.Enabled = False
                          End If
                       Else
                          cbocolor.Enabled = True
                       End If
                       If data_lla.Recordset("pend") = 0 Then
                          b_cmt.Enabled = True
                       Else
                          b_cmt.Enabled = False
                       End If
                       If IsNull(data_lla.Recordset("totend")) = False Then
                          If data_lla.Recordset("totend") = "FACT" Then
                             chtmut.Enabled = False
                             txt_costo.Enabled = False
                          Else
                             chtmut.Enabled = True
                             txt_costo.Enabled = True
                          End If
                       Else
                          chtmut.Enabled = True
                          txt_costo.Enabled = True
                       End If
            
                       XaftC = ""
                       If Label40.Caption <> "" Then
                       Else
                          If cbotras.ListIndex = 1 Or cbotras.ListIndex = 2 Or cbotras.ListIndex = 5 Then
                             If cbocolor.Text = "ROJO" Or cbozona.ListIndex = 4 Or cbozona.ListIndex >= 6 Then
                                Label40.Caption = "AFT:1987"
                                XaftC = "1987"
                             Else
                                If txt_hortd.Text <> "__:__" Then
                                   If Format(txt_hortd.Text, "HH:mm") >= "21:00" Then
                                      Label40.Caption = "AFT:1987"
                                      XaftC = "1987"
                                   Else
                                      If Format(txt_hortd.Text, "HH:mm") <= "09:00" Then
                                         Label40.Caption = "AFT:1987"
                                         XaftC = "1987"
                                      End If
                                   End If
                                End If
                             End If
                          Else
                             If cbotras.ListIndex <= 0 Then
                                c_aft.Enabled = False
                                Label40.Caption = ""
                                XaftC = ""
                             Else
                                Label40.Caption = "AFT:1987"
                                XaftC = "1987"
                             End If
                          End If
                       End If
                       If txt_codmed.Text <> "" Then
                          If txt_codmed.Text = 959 Then
                             Label40.Caption = "AFT:1987"
                             XaftC = "1987"
                          End If
                       End If
                       If txt_cat.Text = "MSP" Then
                          Label40.Caption = "AFT:1987"
                          XaftC = "1987"
                       End If
                       If cbocolor.Text <> "" Then
                          XcolorAntAzul = cbocolor.Text
                       Else
                          XcolorAntAzul = "S"
                       End If
                  Else
                     If XpuedeMod2 = 9 Then
                        MsgBox "El llamado lo está editando el usuario." & Xusuarioedit, vbCritical
'                        If ControlUsuario("Despachador Edita") = 1 Then
'                           MsgBox "Verifique y reitente! Se habilitará para grabar.", vbInformation
                        b_cancela_Click
                        
'                        End If
                     Else
                        MsgBox "El llamado ya fue facturado, no se puede modificar.", vbCritical
                        b_cancela_Click
                     End If
                  End If
               Else
                   borra_ya
                   igualar_sin
                   If IsNull(data_lla.Recordset("cancela")) = False Then
                      If data_lla.Recordset("cancela") = 1 Then
                         MsgBox "El llamado figura CANCELADO, verifique!!", vbCritical
                      End If
                   End If
                   XAlta = 0
                   b_nuevo.Enabled = False
                   b_modif.Enabled = False
                   b_imp.Enabled = False
                   b_buscar.Enabled = False
                   b_grabar.Enabled = True
                   b_cancel.Enabled = False
                   b_cancela.Enabled = True
                   b_pend.Enabled = False
                   txt_nro.Enabled = False
                   mfecha.Enabled = False
                   txt_hora.Enabled = False
                   txt_usua.Enabled = False
                   b_hist.Enabled = False
                   Command2.Enabled = False
                   Command3.Enabled = False
                   If cbocolor.Text = "ROJO" Or cbocolor.Text = "AMARILLO" Then
                      If WElusuario = "MARTINC" Then
                         cbocolor.Enabled = True
                      Else
                         cbocolor.Enabled = False
                      End If
                   Else
                      cbocolor.Enabled = True
                   End If
                   If data_lla.Recordset("pend") = 0 Then
                      b_cmt.Enabled = True
                   Else
                      b_cmt.Enabled = False
                   End If
                   If IsNull(data_lla.Recordset("totend")) = False Then
                      If data_lla.Recordset("totend") = "FACT" Then
                         chtmut.Enabled = False
                         txt_costo.Enabled = False
                      Else
                         chtmut.Enabled = True
                         txt_costo.Enabled = True
                      End If
                   Else
                      chtmut.Enabled = True
                      txt_costo.Enabled = True
                   End If
        
                   XaftC = ""
                   If Label40.Caption <> "" Then
                   Else
                      If cbotras.ListIndex = 1 Or cbotras.ListIndex = 2 Or cbotras.ListIndex = 5 Then
                         If cbocolor.Text = "ROJO" Or cbozona.ListIndex = 4 Or cbozona.ListIndex >= 6 Then
                            Label40.Caption = "AFT:1987"
                            XaftC = "1987"
                         Else
                            If txt_hortd.Text <> "__:__" Then
                               If Format(txt_hortd.Text, "HH:mm") >= "21:00" Then
                                  Label40.Caption = "AFT:1987"
                                  XaftC = "1987"
                               Else
                                  If Format(txt_hortd.Text, "HH:mm") <= "09:00" Then
                                     Label40.Caption = "AFT:1987"
                                     XaftC = "1987"
                                  End If
                               End If
                            End If
                         End If
                      Else
                         If cbotras.ListIndex <= 0 Then
                            c_aft.Enabled = False
                            Label40.Caption = ""
                            XaftC = ""
                         Else
                            Label40.Caption = "AFT:1987"
                            XaftC = "1987"
                         End If
                      End If
                   End If
                   If txt_codmed.Text <> "" Then
                      If txt_codmed.Text = 959 Then
                         Label40.Caption = "AFT:1987"
                         XaftC = "1987"
                      End If
                   End If
                   If txt_cat.Text = "MSP" Then
                      Label40.Caption = "AFT:1987"
                      XaftC = "1987"
                   End If
                   If cbocolor.Text <> "" Then
                      XcolorAntAzul = cbocolor.Text
                   Else
                      XcolorAntAzul = "S"
                   End If
               End If
            End If
            chcovid.Enabled = True
        End If
    Else
        MsgBox "Verifique!! No hay llamado seleccionado para modificar", vbCritical
        Unload Me
    End If
'Else
'    MsgBox "Usuario sin permisos de modificar.", vbInformation
'End If

Exit Sub

Almod:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ERR:" & Err.Number
      End If

End Sub

Private Sub b_nuevo_Click()
Dim Xlahorctrf As String
'If frm_menu.data_parse.Recordset("base") = 29 Then
'   MsgBox "Es"
'End If
Xlahorctrf = Format(Time, "HH:mm")

Frame1.Enabled = True
If Frame2.Visible = False Then
   Frame2.Visible = True
End If

Frame2.Enabled = True
borra_ya

Command1.Enabled = False
Command1.Enabled = False
txt_locali.Visible = True
b_covid.Enabled = False

'data_lla.Refresh
'''data_lla.Recordset.MoveLast
'''txt_nro.Text = data_lla.Recordset("nrolla") + 1
txt_nro.Text = data_par.Recordset("tasaintmn") + 1
data_par.Recordset.Edit
data_par.Recordset("tasaintmn") = txt_nro.Text
data_par.Recordset.Update
'094428513

'''If IsNull(data_lla.Recordset("nrolla")) = False Then
'''   txt_nro.Text = data_lla.Recordset("nrolla") + 1
'''Else
'''   txt_nro.Text = 10000
'''End If

mfecha.Text = Format(Date, "dd/mm/yyyy")
txt_hora.Text = Format(Time, "HH:mm")
txt_usua.Text = WElusuario
Xwhorarec = Format(Time, "HH:mm:ss")

mfecha.Enabled = False
txt_hora.Enabled = False
txt_usua.Enabled = False

XAlta = 1
b_nuevo.Enabled = False
b_modif.Enabled = False
b_imp.Enabled = False
b_buscar.Enabled = False
b_grabar.Enabled = True
b_cancel.Enabled = False
b_cancela.Enabled = True
b_pend.Enabled = False
b_hist.Enabled = False
Command2.Enabled = False
Command3.Enabled = False

cbocolor.Enabled = True
cbocolor.BackColor = &HFFFFFF
txt_costo.Text = 0
txt_boleta.Text = 0
cbozona.ListIndex = 0
cbobase.Text = 0

c_aft.Enabled = False
chcovid.Enabled = True

txt_ced.SetFocus

End Sub

Private Sub b_pend_Click()
   Dim Xquehace As String
'If XWeltipoU = "USUARIOS DESP" Then
'   If frm_pendlla.Visible = True Then
'      MsgBox "Ya está abierto"
'   Else
'      frm_pendlla.Show
'   End If
'Else
   Xquehace = MsgBox("DESEA VER SOLO LLAMADOS CODIGO AZUL?", vbInformation + vbYesNo)
   If Xquehace = vbYes Then
      If frm_pendllap.Visible = True Then
         MsgBox "Ya está abierto"
      Else
         frm_pendllap.Show
      End If
   Else
      If frm_pendlla.Visible = True Then
         MsgBox "Ya está abierto"
      Else
         frm_pendlla.Show
      End If
   End If
'End If


End Sub

Private Sub c_aft_Click()
Dim Xrespdes, Xrespdes2 As String
On Error GoTo Enaft

If Xhayregistros = 9 Then
   txt_nro.Text = data_par.Recordset("tasaintmn") + 1
   data_par.Recordset.Edit
   data_par.Recordset("tasaintmn") = txt_nro.Text
   data_par.Recordset.Update
   mfecha.Text = Format(Date, "dd/mm/yyyy")
   txt_hora.Text = Format(Time, "HH:mm")
   txt_usua.Text = WElusuario
   Xwhorarec = Format(Time, "HH:mm:ss")
       If cbozona.Text <> "" Then
          If cbocolor.Text <> "" Then
                data_lla.Recordset("nrolla") = txt_nro.Text
                data_lla.Recordset("nro") = txt_nro.Text
                data_lla.Recordset("editando") = 1
                data_lla.Recordset("fecha") = Format(mfecha.Text, "dd/mm/yyyy")
                data_lla.Recordset("hora") = Format(txt_hora.Text, "HH:mm")
                data_lla.Recordset("activo") = Format(Time, "HH:mm:ss")
                data_lla.Recordset("usuario") = txt_usua.Text
                If txt_mat.Text = "" Then
                   data_lla.Recordset("matric") = 0
                Else
                   data_lla.Recordset("matric") = txt_mat.Text
                End If
                data_lla.Recordset("nombre") = txt_nomb.Text
                If txt_edad = "" Then
                   data_lla.Recordset("edad") = 0
                Else
                   data_lla.Recordset("edad") = txt_edad.Text
                End If
                If cboed.ListIndex = 0 Then
                   data_lla.Recordset("unied") = 3
                Else
                   If cboed.ListIndex = 1 Then
                      data_lla.Recordset("unied") = 2
                   Else
                      If cboed.ListIndex = 2 Then
                         data_lla.Recordset("unied") = 1
                      Else
                         data_lla.Recordset("unied") = 3
                      End If
                   End If
                End If
                If txt_cat.Text = "" Then
                   txt_cat.Text = "AAABBB"
                End If
                data_lla.Recordset("categ") = txt_cat.Text
                data_lla.Recordset("nomcat") = txt_nomcat.Text
                If txt_ced.Text = "" Then
                   data_lla.Recordset("ci") = 0
                Else
                   data_lla.Recordset("ci") = txt_ced.Text
                End If
                data_lla.Recordset("direcc") = "S/D"
                data_lla.Recordset("telef") = txt_tel.Text
                If cbozona.ListIndex >= 0 Then
                   data_lla.Recordset("codzon") = Val(cbozona.Text)
                Else
                   data_lla.Recordset("codzon") = 1
                End If
                If cbobase.Text = "" Then
                   data_lla.Recordset("base") = 0
                Else
                   data_lla.Recordset("base") = cbobase.Text
                End If
                data_lla.Recordset("referen") = txt_direc.Text
                If txt_ante.Text <> "" Then
                   data_lla.Recordset("motcon") = txt_ante.Text  'motivo de consulta que no va mas (100) pasa como antecedentes
                End If
                data_lla.Recordset("obsmot") = txt_mot.Text
                data_lla.Recordset("realiza") = chtmut.Value
                If cbocolor.ListIndex = 0 Then
                   data_lla.Recordset("codmot") = "V"
                   data_lla.Recordset("descol") = "VERDE"
                Else
                   If cbocolor.ListIndex = 1 Then
                      data_lla.Recordset("codmot") = "A"
                      data_lla.Recordset("descol") = "AMARILLO"
                   Else
                      If cbocolor.ListIndex = 2 Then
                         data_lla.Recordset("codmot") = "R"
                         data_lla.Recordset("descol") = "ROJO"
                      Else
                         If cbocolor.ListIndex = 3 Then
                            data_lla.Recordset("codmot") = "C"
                            data_lla.Recordset("descol") = "CELESTE"
                         Else
                            If cbocolor.ListIndex = 4 Then
                               data_lla.Recordset("codmot") = "Z"
                               data_lla.Recordset("descol") = "AZUL"
                            Else
                               If cbocolor.ListIndex = 5 Then
                                  data_lla.Recordset("codmot") = "N"
                                  data_lla.Recordset("descol") = "NEGRO"
                               Else
                                  data_lla.Recordset("codmot") = "V"
                                  data_lla.Recordset("descol") = "VERDE"
                                  cbocolor.ListIndex = 0
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
                If txt_movil.Text <> "" Then
                   data_lla.Recordset("movilpas") = txt_movil.Text
                   If txt_movil.Text <> 0 Then
                      If XWeltipoU = "USUARIOS DESP" Or XWeltipoU = "ADMINISTRADOR" Then
                         data_lla.Recordset("timdes") = Trim(WElusuario)
                      End If
                      data_lla.Recordset("pend") = 1
                      If mtd.Text = "__/__/____" Then
                         data_lla.Recordset("pend") = 1
                      Else
                         data_lla.Recordset("fec_rea") = Format(mtd.Text, "dd/mm/yyyy")
                         data_lla.Recordset("pend") = 2
                      End If
                      If txt_hortd.Text = "__:__" Then
                         data_lla.Recordset("pend") = 1
                      Else
                         data_lla.Recordset("hor_rea") = txt_hortd.Text
                         data_lla.Recordset("pend") = 2
                      End If
                   Else
                      data_lla.Recordset("pend") = 0
                   End If
                Else
                   data_lla.Recordset("pend") = 0
                   data_lla.Recordset("movilpas") = 0
                End If
                If mfecasig.Text = "__/__/____" Then
                   data_lla.Recordset("pend") = 0
                Else
                   data_lla.Recordset("fecpas") = Format(mfecasig.Text, "dd/mm/yyyy")
                End If
                If txt_horasig.Text <> "" Then
                   data_lla.Recordset("horpas") = txt_horasig.Text
                Else
                   data_lla.Recordset("pend") = 0
                End If
                If msalida.Text = "__/__/____" Then
                Else
                   data_lla.Recordset("fecsali") = Format(msalida.Text, "dd/mm/yyyy")
                End If
                data_lla.Recordset("horsali") = txt_horsal.Text
                If mllegada.Text = "__/__/____" Then
                Else
                   data_lla.Recordset("fec_llega") = Format(mllegada.Text, "dd/mm/yyyy")
                End If
                data_lla.Recordset("hor_llega") = txt_horlle.Text
                data_lla.Recordset("diag") = txt_diag.Text
                If cbocolfin.ListIndex = 0 Then
                   data_lla.Recordset("colormot") = "V"
                Else
                   If cbocolfin.ListIndex = 1 Then
                      data_lla.Recordset("colormot") = "A"
                   Else
                      If cbocolfin.ListIndex = 2 Then
                         data_lla.Recordset("colormot") = "R"
                      Else
                         If cbocolfin.ListIndex = 3 Then
                            data_lla.Recordset("colormot") = "N"
                         End If
                      End If
                   End If
                End If
                If txt_codmed.Text = "" Then
                   data_lla.Recordset("codmed") = 0
                Else
                   data_lla.Recordset("codmed") = txt_codmed.Text
                End If
                data_lla.Recordset("nommed") = dbcbomed.Text
                data_lla.Recordset("trasla") = cbotras.ListIndex
                data_lla.Recordset("lugar") = txt_lugar.Text
                data_lla.Recordset("hsald") = txt_trassal.Text
                data_lla.Recordset("hllega") = txt_enca.Text
                data_lla.Recordset("hzona") = txt_enzona.Text
                data_lla.Recordset("obs") = txt_obs.Text
                If txt_movtra.Text = "" Then
                   data_lla.Recordset("movtras") = 0
                Else
                   If IsNumeric(txt_movtra.Text) = False Then
                      data_lla.Recordset("movtras") = 0
                   Else
                      data_lla.Recordset("movtras") = txt_movtra.Text
                   End If
                End If
                data_lla.Recordset("totdem") = txt_demora.Text
                If Combo1.Visible = True Then
                   If Combo1.Text <> "" Then
                      data_lla.Recordset("dcobr") = Combo1.Text
                   Else
                      data_lla.Recordset("dcobr") = ""
                   End If
                End If
                If txt_locali.Text <> "" Then
                   data_lla.Recordset("motmov") = txt_locali.Text
                End If
                If txt_queb.Text = "" Then
                   txt_queb.Text = 0
                End If
                data_lla.Recordset("ncobr") = txt_queb.Text
                If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                   If Label41.Caption = "" Then
                      Label41.Caption = 0
                   End If
                   If Label42.Caption = "" Then
                      Label42.Caption = 0
                   End If
                   If Label43.Caption = "" Then
                      Label43.Caption = 0
                   End If
                   If Label44.Caption = "" Then
                      Label44.Caption = 0
                   End If
                   If Label45.Caption = "" Then
                      Label45.Caption = 0
                   End If
                   If Label46.Caption = "" Then
                      Label46.Caption = -1
                   End If
                   If Label48.Caption = "" Then
                      Label48.Caption = 0
                   End If
                   If txt_quien.Text <> "" Then
                      data_lla.Recordset("motcance") = txt_quien.Text
                   End If
                   data_lla.Recordset("mm") = Label41.Caption
                   data_lla.Recordset("thh") = Label42.Caption
                   data_lla.Recordset("tmm") = Label43.Caption
                   data_lla.Recordset("pasado") = Label44.Caption
                   data_lla.Recordset("ano") = Label45.Caption
                   data_lla.Recordset("mes") = Label46.Caption
                   data_lla.Recordset("timsi") = Trim(str(Val(Label48.Caption)))
                End If
                If txt_costo.Text <> "" Then
                   If txt_costo.Text > 0 Then
                      data_lla.Recordset("mes") = txt_costo.Text
                   End If
                End If
                If txt_boleta.Text <> "" Then
                   If txt_boleta.Text > 0 Then
                      data_lla.Recordset("ano") = txt_boleta.Text
                   End If
                End If
                If txt_codmedtra.Text <> "" Then
                   data_lla.Recordset("movil_rea") = txt_codmedtra.Text
                Else
                   data_lla.Recordset("movil_rea") = 0
                End If
                data_lla.Recordset("enfer") = Check2.Value ' actos de enfermería
                data_lla.Recordset("hh") = Combo3.ListIndex
                If cbocolor.ListIndex = 4 Then
                   If UCase(txt_locali.Text) = "SAUCE" Then
                      data_lla.Recordset("mm") = 1
                   Else
                      If UCase(txt_locali.Text) = "TOLEDO" Then
                         data_lla.Recordset("mm") = 2
                      Else
                         If UCase(txt_locali.Text) = "CASARINO" Then
                            data_lla.Recordset("mm") = 3
                         Else
                            If UCase(txt_locali.Text) = "SUAREZ" Then
                               data_lla.Recordset("mm") = 4
                            Else
                               If UCase(txt_locali.Text) = "BARROS BLANCOS" Then
                                  data_lla.Recordset("mm") = 5
                               Else
                                  If UCase(txt_locali.Text) = "PANDO" Then
                                     data_lla.Recordset("mm") = 6
                                  Else
                                     data_lla.Recordset("mm") = 1
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
                If Xdeudasi = 9 Then
                   Dim Xcodaut2 As String
                   Xcodaut2 = InputBox("Ingrese CODIGO DE AUTORIZACION PARA PODER GRABAR", "Código autorización")
                   If Xcodaut2 <> "" Then
                      data_lla.Recordset("direcc") = Xcodaut2
                      data_lla.Recordset.Update
                      If Check4.Value = 1 Then
                         data_llamod.Recordset.AddNew
                         data_llamod.Recordset("nro") = txt_nro.Text
                         data_llamod.Recordset("fecha") = mfecha.Text
                         data_llamod.Recordset("pasado") = Check4.Value
                         data_llamod.Recordset("movilpas") = txt_codmed2.Text
                         If dbcbomed.Text <> "" Then
                            data_llamod.Recordset("nommed") = dbcbomed2.Text
                         End If
                         data_llamod.Recordset("hora") = Format(Time, "HH:mm")
                         data_llamod.Recordset("usuario") = WElusuario
                         If t_codced.Text <> "" Then
                            data_llamod.Recordset("mes") = t_codced.Text
                         Else
                            data_llamod.Recordset("mes") = 0
                         End If
                         data_llamod.Recordset.Update
                         data_llamod.Refresh
                      End If
                      If mftrassol.Text <> "__/__/____" And mhtrassol.Text <> "__:__" Then
                         data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                         data_llamod.Refresh
                         If data_llamod.Recordset.RecordCount > 0 Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("fec_llega") = mftrassol.Text
                            data_llamod.Recordset("hor_llega") = mhtrassol.Text
                            data_llamod.Recordset.Update
                         Else
                            data_llamod.Recordset.AddNew
                            data_llamod.Recordset("nro") = txt_nro.Text
                            data_llamod.Recordset("fecha") = mfecha.Text
                            data_llamod.Recordset("fec_llega") = mftrassol.Text
                            data_llamod.Recordset("hor_llega") = mhtrassol.Text
                            data_llamod.Recordset.Update
                         End If
                      End If
                      If Combo2.ListIndex >= 0 Then
                         data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                         data_llamod.Refresh
                         If data_llamod.Recordset.RecordCount > 0 Then
                            If IsNull(data_llamod.Recordset("telef")) = True Then
                               data_llamod.Recordset.Edit
                               data_llamod.Recordset("telef") = Combo2.Text
                               data_llamod.Recordset.Update
                            Else
                               If data_llamod.Recordset("telef") <> Combo2.Text Then
                                  data_llamod.Recordset.Edit
                                  data_llamod.Recordset("telef") = Combo2.Text
                                  data_llamod.Recordset.Update
                               End If
                            End If
                         Else
                            data_llamod.Recordset.AddNew
                            data_llamod.Recordset("nro") = txt_nro.Text
                            data_llamod.Recordset("fecha") = mfecha.Text
                            data_llamod.Recordset("telef") = Combo2.Text
                            data_llamod.Recordset.Update
                         End If
                      Else
                         data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                         data_llamod.Refresh
                         If data_llamod.Recordset.RecordCount > 0 Then
                            If IsNull(data_llamod.Recordset("telef")) = False Then
                               data_llamod.Recordset.Edit
                               data_llamod.Recordset("telef") = Null
                               data_llamod.Recordset.Update
                            End If
                         End If
                      End If
                      If labcodchof.Caption <> "" Then
                         data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                         data_llamod.Refresh
                         If data_llamod.Recordset.RecordCount > 0 Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                            data_llamod.Recordset.Update
                         Else
                            data_llamod.Recordset.AddNew
                            data_llamod.Recordset("nro") = txt_nro.Text
                            data_llamod.Recordset("fecha") = mfecha.Text
                            data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                            data_llamod.Recordset.Update
                         End If
                      End If
                      If t_codced.Text <> "" Then
                         data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                         data_llamod.Refresh
                         If data_llamod.Recordset.RecordCount > 0 Then
                            If IsNull(data_llamod.Recordset("mes")) = False Then
                               If data_llamod.Recordset("mes") <> t_codced.Text Then
                                  data_llamod.Recordset.Edit
                                  data_llamod.Recordset("mes") = t_codced.Text
                                  data_llamod.Recordset.Update
                               End If
                            Else
                               data_llamod.Recordset.Edit
                               data_llamod.Recordset("mes") = t_codced.Text
                               data_llamod.Recordset.Update
                            End If
                         Else
                            data_llamod.Recordset.AddNew
                            data_llamod.Recordset("mes") = t_codced.Text
                            data_llamod.Recordset("nro") = txt_nro.Text
                            data_llamod.Recordset("fecha") = mfecha.Text
                            data_llamod.Recordset.Update
                         End If
                      Else
                         data_llamod.Recordset.AddNew
                         data_llamod.Recordset("mes") = 0
                         data_llamod.Recordset("nro") = txt_nro.Text
                         data_llamod.Recordset("fecha") = mfecha.Text
                         data_llamod.Recordset.Update
                      End If
                      data_hist.Recordset.AddNew
                      data_hist.Recordset("nro") = txt_nro.Text
                      data_hist.Recordset("fecha") = Date
                      data_hist.Recordset("hora") = Format(Time, "HH:mm")
                      data_hist.Recordset("usuario") = WElusuario
                      data_hist.Recordset("base") = data_par.Recordset("base")
                      data_hist.Recordset("timdes") = "NUEVO"
                      data_hist.Recordset("referen") = txt_direc.Text
                      data_hist.Recordset("obs") = txt_obs.Text
                      data_hist.Recordset("obsmot") = txt_mot.Text
                      If txt_mat.Text <> "" Then
                         data_hist.Recordset("matric") = txt_mat.Text
                      Else
                         data_hist.Recordset("matric") = 0
                      End If
                      If Label3.Caption <> "" Then
                         data_hist.Recordset("nomcat") = Label3.Caption
                      End If
                      data_hist.Recordset("descol") = cbocolor.Text
                      If txt_tel.Text <> "" Then
                         data_hist.Recordset("telef") = txt_tel.Text
                      End If
                      If txt_movil.Text <> "" Then
                         data_hist.Recordset("movilpas") = txt_movil.Text
                      End If
                      If txt_horasig.Text <> "" Then
                         data_hist.Recordset("horpas") = txt_horasig.Text
                      End If
                      If txt_horsal.Text <> "" Then
                         data_hist.Recordset("horsali") = txt_horsal.Text
                      End If
                      If txt_horlle.Text <> "" Then
                         data_hist.Recordset("hor_llega") = txt_horlle.Text
                      End If
                      If txt_diag.Text <> "" Then
                         data_hist.Recordset("diag") = txt_diag.Text
                      End If
                      If txt_codmed.Text <> "" Then
                         data_hist.Recordset("realiza") = txt_codmed.Text
                      End If
                      If txt_nomb.Text <> "" Then
                         data_hist.Recordset("nombre") = txt_nomb.Text
                      End If
                      If txt_cat.Text <> "" Then
                         data_hist.Recordset("categ") = txt_cat.Text
                      End If
                      data_hist.Recordset("trasla") = cbotras.ListIndex
                      If txt_edad.Text <> "" Then
                         data_hist.Recordset("edad") = txt_edad.Text
                      End If
                      data_hist.Recordset.Update
    '                  borra_ya
                      Xdeudasi = 0
                   Else
                      b_cancela_Click
                   End If
                Else
                   data_lla.Recordset.Update
                   If Check4.Value = 1 Then
                      data_llamod.Recordset.AddNew
                      data_llamod.Recordset("nro") = txt_nro.Text
                      data_llamod.Recordset("fecha") = mfecha.Text
                      data_llamod.Recordset("pasado") = Check4.Value
                      data_llamod.Recordset("movilpas") = txt_codmed2.Text
                      If dbcbomed.Text <> "" Then
                         data_llamod.Recordset("nommed") = dbcbomed2.Text
                      End If
                      data_llamod.Recordset("hora") = Format(Time, "HH:mm")
                      data_llamod.Recordset("usuario") = WElusuario
                      data_llamod.Recordset.Update
                      data_llamod.Refresh
                   End If
                   If mftrassol.Text <> "__/__/____" And mhtrassol.Text <> "__:__" Then
                      data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                      data_llamod.Refresh
                      If data_llamod.Recordset.RecordCount > 0 Then
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("fec_llega") = mftrassol.Text
                         data_llamod.Recordset("hor_llega") = mhtrassol.Text
                         data_llamod.Recordset.Update
                      Else
                         data_llamod.Recordset.AddNew
                         data_llamod.Recordset("nro") = txt_nro.Text
                         data_llamod.Recordset("fecha") = mfecha.Text
                         data_llamod.Recordset("fec_llega") = mftrassol.Text
                         data_llamod.Recordset("hor_llega") = mhtrassol.Text
                         data_llamod.Recordset.Update
                      End If
                   End If
                   If Combo2.ListIndex >= 0 Then
                      data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                      data_llamod.Refresh
                      If data_llamod.Recordset.RecordCount > 0 Then
                         If IsNull(data_llamod.Recordset("telef")) = True Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("telef") = Combo2.Text
                            data_llamod.Recordset.Update
                         Else
                            If data_llamod.Recordset("telef") <> Combo2.Text Then
                               data_llamod.Recordset.Edit
                               data_llamod.Recordset("telef") = Combo2.Text
                               data_llamod.Recordset.Update
                            End If
                         End If
                      Else
                         data_llamod.Recordset.AddNew
                         data_llamod.Recordset("nro") = txt_nro.Text
                         data_llamod.Recordset("fecha") = mfecha.Text
                         data_llamod.Recordset("telef") = Combo2.Text
                         data_llamod.Recordset.Update
                      End If
                   Else
                      data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                      data_llamod.Refresh
                      If data_llamod.Recordset.RecordCount > 0 Then
                         If IsNull(data_llamod.Recordset("telef")) = False Then
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("telef") = Null
                            data_llamod.Recordset.Update
                         End If
                      End If
                   End If
                   If labcodchof.Caption <> "" Then
                      data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                      data_llamod.Refresh
                      If data_llamod.Recordset.RecordCount > 0 Then
                         data_llamod.Recordset.Edit
                         data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                         data_llamod.Recordset.Update
                      Else
                         data_llamod.Recordset.AddNew
                         data_llamod.Recordset("nro") = txt_nro.Text
                         data_llamod.Recordset("fecha") = mfecha.Text
                         data_llamod.Recordset("movil_rea") = Val(labcodchof.Caption)
                         data_llamod.Recordset.Update
                      End If
                   End If
                   If t_codced.Text <> "" Then
                      data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
                      data_llamod.Refresh
                      If data_llamod.Recordset.RecordCount > 0 Then
                         If IsNull(data_llamod.Recordset("mes")) = False Then
                            If data_llamod.Recordset("mes") <> t_codced.Text Then
                               data_llamod.Recordset.Edit
                               data_llamod.Recordset("mes") = t_codced.Text
                               data_llamod.Recordset.Update
                            End If
                         Else
                            data_llamod.Recordset.Edit
                            data_llamod.Recordset("mes") = t_codced.Text
                            data_llamod.Recordset.Update
                         End If
                      Else
                         data_llamod.Recordset.AddNew
                         data_llamod.Recordset("mes") = t_codced.Text
                         data_llamod.Recordset("nro") = txt_nro.Text
                         data_llamod.Recordset("fecha") = mfecha.Text
                         data_llamod.Recordset.Update
                      End If
                   Else
                      data_llamod.Recordset.AddNew
                      data_llamod.Recordset("mes") = 0
                      data_llamod.Recordset("nro") = txt_nro.Text
                      data_llamod.Recordset("fecha") = mfecha.Text
                      data_llamod.Recordset.Update
                   End If
                   data_hist.Recordset.AddNew
                   data_hist.Recordset("nro") = txt_nro.Text
                   data_hist.Recordset("fecha") = Date
                   data_hist.Recordset("hora") = Format(Time, "HH:mm")
                   data_hist.Recordset("usuario") = WElusuario
                   data_hist.Recordset("base") = data_par.Recordset("base")
                   data_hist.Recordset("timdes") = "NUEVO"
                   data_hist.Recordset("referen") = txt_direc.Text
                   data_hist.Recordset("obs") = txt_obs.Text
                   data_hist.Recordset("obsmot") = txt_mot.Text
                   If txt_mat.Text <> "" Then
                      data_hist.Recordset("matric") = txt_mat.Text
                   Else
                      data_hist.Recordset("matric") = 0
                   End If
                   If Label3.Caption <> "" Then
                      data_hist.Recordset("nomcat") = Label3.Caption
                   End If
                   data_hist.Recordset("descol") = cbocolor.Text
                   If txt_tel.Text <> "" Then
                      data_hist.Recordset("telef") = txt_tel.Text
                   End If
                   If txt_movil.Text <> "" Then
                      data_hist.Recordset("movilpas") = txt_movil.Text
                   End If
                   If txt_horasig.Text <> "" Then
                      data_hist.Recordset("horpas") = txt_horasig.Text
                   End If
                   If txt_horsal.Text <> "" Then
                      data_hist.Recordset("horsali") = txt_horsal.Text
                   End If
                   If txt_horlle.Text <> "" Then
                      data_hist.Recordset("hor_llega") = txt_horlle.Text
                   End If
                   If txt_diag.Text <> "" Then
                      data_hist.Recordset("diag") = txt_diag.Text
                   End If
                   If txt_codmed.Text <> "" Then
                      data_hist.Recordset("realiza") = txt_codmed.Text
                   End If
                   If txt_nomb.Text <> "" Then
                      data_hist.Recordset("nombre") = txt_nomb.Text
                   End If
                   If txt_cat.Text <> "" Then
                      data_hist.Recordset("categ") = txt_cat.Text
                   End If
                   data_hist.Recordset("trasla") = cbotras.ListIndex
                   If txt_edad.Text <> "" Then
                      data_hist.Recordset("edad") = txt_edad.Text
                   End If
                   data_hist.Recordset.Update
                   Xdeudasi = 0
                End If
                If txt_codmed2.Text = "" Then
                   txt_codmed2.Text = 0
                End If
                If XAlta = 1 Then
                   If txt_ced.Text <> "" And txt_mot.Text <> "" Then
                      data_cons.Recordset.AddNew
                      data_cons.Recordset("mat") = txt_ced.Text
                      data_cons.Recordset("motivo") = txt_mot.Text
                      data_cons.Recordset("fecha") = Date
                      data_cons.Recordset.Update
                   End If
                End If
                XAlta = 3
                Check1.Value = 0
                historial
                despuesdegraba
          Else
            MsgBox "Ingrese COLOR de llamado", vbCritical, "Mensaje"
            cbocolor.SetFocus
          End If
       Else
          MsgBox "Ingrese Zona", vbCritical, "Mensaje"
          cbozona.SetFocus
       End If
   Xhayregistros = 0
Else
''   frm_aft.Show vbModal
   Dim Xelaftes As String
   Xelaftes = InputBox("Ingrese código de AFT:", "Despacho")
   If Trim(Xelaftes) <> "" Then
      XaftC = Xelaftes
      If IsNull(data_lla.Recordset("aft")) = False Then
         If data_lla.Recordset("aft") <> Xelaftes Then
            data_lla.Recordset.Edit
            data_lla.Recordset("editando") = 1
            data_lla.Recordset("aft") = Xelaftes
            data_lla.Recordset.Update
         End If
      Else
         data_lla.Recordset.Edit
         data_lla.Recordset("editando") = 1
         data_lla.Recordset("aft") = Xelaftes
         data_lla.Recordset.Update
      End If
      Label40.Caption = "AFT:" & Xelaftes
   Else
      Label40.Caption = ""
      If IsNull(data_lla.Recordset("aft")) = False Then
         data_lla.Recordset.Edit
         data_lla.Recordset("editando") = 1
         data_lla.Recordset("aft") = Null
         data_lla.Recordset.Update
      End If
   End If

End If

Exit Sub

Enaft:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ENAFT ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ENAFT ERR:" & Err.Number
      End If

End Sub

Private Sub cbobase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mot.SetFocus
End If

End Sub

Private Sub cbocolfin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   dbcbomed.SetFocus
End If

End Sub

Private Sub cbocolor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mot.SetFocus
End If

End Sub

Private Sub cbocolor_LostFocus()
Dim Xdescuento As Double
Dim CambiaColor As String

If cbocolor.Text = "" Then
'   MsgBox "Ingrese color del llamado", vbInformation, "Mensaje"
'   cbocolor.SetFocus
   cbocolor.ListIndex = 0
Else
   If cbocolor.Text = "VERDE" Then
      cbocolor.BackColor = &HC000&
   Else
      If cbocolor.Text = "ROJO" Then
         cbocolor.BackColor = &HFF&
      Else
         If cbocolor.Text = "AMARILLO" Then
            cbocolor.BackColor = &HFFFF&
         Else
            If cbocolor.Text = "CELESTE" Then
               cbocolor.BackColor = &HFFFF00
            Else
               If cbocolor.Text = "AZUL" Then
                  cbocolor.BackColor = &HC00000
               Else
                  If cbocolor.Text = "NEGRO" Then
                     cbocolor.BackColor = &H80000006
                  Else
                     cbocolor.BackColor = &HFFFFFF
                  End If
               End If
            End If
         End If
      End If
   End If
'&H00C00000&
End If
Dim Xcodestudio As Long
Dim XImp, Ximpest, Ximppart As Double
XImp = 0
Ximpest = 0

If txt_cat.Text = "SAMCB" Or cbobase.Text > 0 Or cbozona.ListIndex >= 4 Then
   XImp = 0
   txt_costo.Text = 0
Else
   If cbocolor.Text = "" Then
   Else
      If cbocolor.Text = "VERDE" Or cbocolor.Text = "CELESTE" Then
         If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
            Xcodestudio = 10014
         Else
            Xcodestudio = 10002
         End If
      Else
         If cbocolor.Text = "AMARILLO" Then
            If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
               Xcodestudio = 10013
            Else
               Xcodestudio = 10004
            End If
         Else
            If cbocolor.Text = "ROJO" Then
               If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                  Xcodestudio = 10012
               Else
                  Xcodestudio = 10006
               End If
            Else
               If cbocolor.Text = "AZUL" Then
                  Xcodestudio = 14004
               Else
                  If cbocolor.Text = "NEGRO" Then
                     If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                        Xcodestudio = 10012
                     Else
                        Xcodestudio = 10016
                     End If
                  Else
                     Xcodestudio = 10016
                  End If
               End If
            End If
         End If
      End If
      If txt_cat.Text = "MSP" Then
         If cbocolor.Text = "ROJO" Then
            Xcodestudio = 90017
         Else
            If cbocolor.Text = "AMARILLO" Then
               Xcodestudio = 90018
            Else
               Xcodestudio = 90019
            End If
         End If
      End If
      If txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "CERDGI" Or _
         txt_cat.Text = "CERADU" Or txt_cat.Text = "CERHEV" Or txt_cat.Text = "CERMAT" Or txt_cat.Text = "CERSEV" Or txt_cat.Text = "CERVIS" Then
'''''''UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and base =" & 0 & " and movilpas <>" & 99
         Xcodestudio = 10008
      End If
'Desde acá recuperar
      If txt_cat.Text <> "" Then
         data_aran.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
         data_aran.Refresh
         If data_aran.Recordset.RecordCount > 0 Then
            If IsNull(data_aran.Recordset("cnv_aran")) = False Then
               Xop1 = data_aran.Recordset("cnv_aran")
            Else
               Xop1 = 0
            End If
         Else
            Xop1 = 0
         End If
      Else
         Xop1 = 0
      End If
      data_aran.RecordSource = "Select * from estudios where codest =" & Xcodestudio
      data_aran.Refresh
      If data_aran.Recordset.RecordCount > 0 Then
         If IsNull(data_aran.Recordset("cons")) = False Then
            Ximpest = data_aran.Recordset("cons")
            Ximppart = data_aran.Recordset("part")
         Else
            Ximpest = 0
            Ximppart = 0
         End If
         data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & Xop1 & " and id_serv =" & Xcodestudio
         data_aran.Refresh
         If data_aran.Recordset.RecordCount > 0 Then
            If data_aran.Recordset("prec_serv") > 0 Then
               XImp = data_aran.Recordset("prec_serv")
            Else
               If data_aran.Recordset("por_serv") = 100 Then
                  XImp = 0
               Else
                  If data_aran.Recordset("por_serv") = 0 Then
                     XImp = Ximpest
                  Else
                     Xdescuento = data_aran.Recordset("por_serv") * Ximpest / 100
                     XImp = Ximpest - Xdescuento
                  End If
               End If
            End If
         Else
            XImp = Ximppart
         End If
         If txt_cat.Text = "PART" Then
            XImp = Ximppart
         End If
         If chtmut.Value = 1 Then
            XImp = 0
         End If
      Else
         Ximpest = 0
         Ximppart = 0
     
      End If
         If txt_cat.Text <> "PART" Then
            data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
            data_convbus.Refresh
            If data_convbus.Recordset.RecordCount > 0 Then
               If IsNull(data_convbus.Recordset("cnv_colrec")) = False Then
                  If txt_cat.Text = "SA" Or txt_cat.Text = "SAF" Or txt_cat.Text = "CCOMS" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "SEGAM" Or _
                     txt_cat.Text = "EMERN" Or txt_cat.Text = "EMERG" Or txt_cat.Text = "EMERNE" Or txt_cat.Text = "CAAM" Or _
                     txt_cat.Text = "SAP" Or txt_cat.Text = "VIV19" Or txt_cat.Text = "VIV20" Or txt_cat.Text = "CAAMEP" Or _
                     txt_cat.Text = "MSP" Or txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "APNORE" Or _
                     txt_cat.Text = "SJ01" Or data_convbus.Recordset("cnv_colrec") = "M" Or data_convbus.Recordset("cnv_colrec") = "R" Or _
                     Mid(txt_cat.Text, 1, 4) = "TALA" Or cbozona.Text = 3 Or txt_cat.Text = "SEMM" Or txt_cat.Text = "SEMM1" Then
                     XImp = 0
                  End If
               Else
                  If txt_cat.Text = "SA" Or txt_cat.Text = "SAF" Or txt_cat.Text = "CCOMS" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "SEGAM" Or _
                     txt_cat.Text = "EMERN" Or txt_cat.Text = "EMERG" Or txt_cat.Text = "EMERNE" Or txt_cat.Text = "CAAM" Or txt_cat.Text = "911" Or _
                     txt_cat.Text = "SAP" Or txt_cat.Text = "VIV19" Or txt_cat.Text = "VIV20" Or txt_cat.Text = "CAAMEP" Or txt_cat.Text = "911B" Or _
                     txt_cat.Text = "MSP" Or txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "APNORE" Or txt_cat.Text = "CASH" Or _
                     txt_cat.Text = "SJ01" Or Mid(txt_cat.Text, 1, 4) = "TALA" Or cbozona.Text = 3 Then
                     XImp = 0
                  End If
               End If
            Else
               If txt_cat.Text = "SA" Or txt_cat.Text = "SAF" Or txt_cat.Text = "CCOMS" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "SEGAM" Or _
                  txt_cat.Text = "EMERN" Or txt_cat.Text = "EMERG" Or txt_cat.Text = "EMERNE" Or txt_cat.Text = "CAAM" Or txt_cat.Text = "911" Or _
                  txt_cat.Text = "SAP" Or txt_cat.Text = "VIV19" Or txt_cat.Text = "VIV20" Or txt_cat.Text = "CAAMEP" Or txt_cat.Text = "911B" Or _
                  txt_cat.Text = "MSP" Or txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "APNORE" Or txt_cat.Text = "CASH" Or _
                  txt_cat.Text = "SJ01" Or Mid(txt_cat.Text, 1, 4) = "TALA" Or cbozona.Text = 3 Then
                  XImp = 0
               End If
            End If
         End If
         If txt_cat.Text = "MUCAFL" Or txt_cat.Text = "MUCAMA" Or txt_cat.Text = "MUCAMI" Or txt_cat.Text = "MUCAMM" Or txt_cat.Text = "MUCAMP" Or _
            txt_cat.Text = "MUCAMS" Or txt_cat.Text = "MUCAMT" Or txt_cat.Text = "MUCATA" Or txt_cat.Text = "SOLEME" Or _
            txt_cat.Text = "CAAMEP" Or txt_cat.Text = "SOLAF" Or txt_cat.Text = "SOLAMB" Or txt_cat.Text = "SOC" Or txt_cat.Text = "CPS" Then
            XImp = 0
         End If
         
         If txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "CERDGI" Or _
            txt_cat.Text = "CERADU" Or txt_cat.Text = "CERHEV" Or txt_cat.Text = "CERMAT" Or txt_cat.Text = "CERSEV" Or txt_cat.Text = "CERVIS" Then
            XImp = 0
            If txt_costo.Enabled = True Then
               txt_costo.Text = 0
            End If
         End If
   End If
End If
If cbozona.Text = "4" Or cbozona.Text = "5" Or cbozona.Text = "6" Then
   XImp = 0
   txt_costo.Text = 0
End If

If txt_costo.Text <> "" Then
   If Val(txt_costo.Text) > 0 Then
   Else
      txt_costo.Text = Format(XImp, "Standard")
   End If
Else
  txt_costo.Text = Format(XImp, "Standard")
End If

If XAlta <> 1 Then
   If XAlta = 0 Then
      If XcolorAntAzul = "AZUL" Then
         CambiaColor = MsgBox("Desea reclasificar el llamado con nuevo horario?", vbInformation + vbYesNo, "Despacho")
         If CambiaColor = vbYes Then
            mfecha.Text = Format(Date, "dd/mm/yyyy")
            txt_hora.Text = Format(Time, "HH:mm")
            Label3.Caption = Format(Time, "HH:mm:ss")
         End If
      End If
   End If
End If

End Sub

Private Sub cboed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cat.SetFocus
End If

End Sub

Private Sub cboed_LostFocus()
If cboed.Text = "" Then
'   MsgBox "Ingrese datos de edad", vbInformation, "Mensaje"
'   cboed.SetFocus
   cboed.ListIndex = 0
Else
   If UCase(cboed.Text) = "A" Or cboed.Text = "Años" Then
      cboed.ListIndex = 0
   Else
      If UCase(cboed.Text) = "M" Or cboed.Text = "Meses" Then
         cboed.ListIndex = 1
      Else
         If UCase(cboed.Text) = "D" Or cboed.Text = "Días" Then
            cboed.ListIndex = 2
         Else
            cboed.ListIndex = 0
         End If
      End If
   End If
End If
End Sub

Private Sub cbotras_Click()
If cbotras.ListIndex > 0 Then
   If XAlta = 1 Then
      c_aft.Enabled = False
   Else
      c_aft.Enabled = True
   End If
Else
   c_aft.Enabled = False
End If

End Sub

Private Sub cbotras_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_lugar.SetFocus
End If

End Sub

Private Sub cbozona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cbocolor.Enabled = True Then
      cbocolor.SetFocus
   Else
      txt_mot.SetFocus
   End If
End If

End Sub

Private Sub cbozona_LostFocus()
If cbozona.Text = "" Then
'   MsgBox "Ingrese datos de zona", vbInformation, "Mensaje"
'   cbozona.SetFocus
   cbozona.ListIndex = 0
End If

End Sub




Private Sub chcovid_Click()
If chcovid.Value = 1 Then
   If txt_ced.Text <> "" Then
      If txt_ced.Text <> 0 Then
         Consulta_cedcovid
      End If
   End If
End If

End Sub

Private Sub Command1_Click()
frm_ops911.Show vbModal

End Sub




Private Sub Command2_Click()
'If App.PrevInstance = True Then
'   MsgBox "Ya está abierta la ventana de mapas"
'Else
'   frm_mapas.Show vbModal
'End If

End Sub

Private Sub Command3_Click()
Dim Xresponda As String
Xresponda = MsgBox("Desea enviar datos del paciente al sector de Padrón Social?", vbInformation + vbYesNo, "Despacho")
If Xresponda = vbYes Then
   frm_largador.MousePointer = 11
   If txt_mat.Text <> "" Then
      If txt_mat.Text > 0 Then
         Data1.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Text
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            If IsNull(Data1.Recordset("cl_fultvta")) = False Then
               MsgBox "EL REGISTRO YA TIENE UN AVISO DE MODIFICACION", vbCritical, "DESPACHO"
            Else
               Data1.Recordset.Edit
               Data1.Recordset("cl_fultvta") = mfecha.Text
               Data1.Recordset("cl_tipocli") = 2
               Data1.Recordset("cl_celular") = "DE DESPACHO"
               Data1.Recordset.Update
            End If
            MsgBox "Proceso terminado, se envió el registro a padrón social"
         Else
            MsgBox "No se encuentra esa matrícula"
         End If
      Else
         MsgBox "No se puede consultar una matrícula CERO"
         
      End If
   Else
      MsgBox "No se puede consultar una matrícula en BLANCO"
      
   End If
   frm_largador.MousePointer = 0
End If

End Sub


Private Sub Command5_Click()
Dim Veragenda As String
Veragenda = ""

If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS DESP" Or XWeltipoU = "USUARIOS RECEP" Then
   If ControlUsuario("Utilitarios despacho") = 1 Then
      Veragenda = MsgBox("Desea ver la agenda de CMT?", vbExclamation + vbYesNo, "Despacho")
      If Veragenda = vbYes Then
         frm_seleccmt.Show vbModal
      Else
         frm_cmt.Show vbModal
      End If
   Else
      frm_cmt.Show vbModal
   End If
Else
'   frm_cmt.Show vbModal

'   MsgBox "Usuario no autorizado"
End If

End Sub


Private Sub dbcbomed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbotras.SetFocus
End If

End Sub

Private Sub dbcbomed_LostFocus()
If dbcbomed.Text <> "" Then
   data_med.Recordset.FindFirst "med_nombre ='" & dbcbomed.Text & "'"
   If Not data_med.Recordset.NoMatch Then
      txt_codmed.Text = data_med.Recordset("med_cod")
   Else
      MsgBox "El médico ingresado NO EXISTE, VERIFIQUE!!", vbCritical, "Despacho"
      txt_codmed.Text = 0
   End If
Else
   txt_codmed.Text = 0
End If
End Sub

Private Sub DBCombo1_Click(Area As Integer)

End Sub

Private Sub dbcbomed2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_trassal.SetFocus
End If

End Sub

Private Sub dbcbomed2_LostFocus()
If dbcbomed2.Text <> "" Then
   data_med2.Recordset.FindFirst "med_nombre ='" & dbcbomed2.Text & "'"
   If Not data_med2.Recordset.NoMatch Then
      txt_codmed2.Text = data_med2.Recordset("med_cod")
   Else
      MsgBox "El médico ingresado NO EXISTE, VERIFIQUE!!", vbCritical, "Despacho"
      txt_codmed2.Text = 0
   End If
Else
   txt_codmed2.Text = 0
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   If XWeltipoU = "USUARIOS RECEP" Then
      frm_pendllap.Show vbModal
   Else
      frm_pendlla.Show vbModal
   End If
End If
If KeyCode = vbKeyEscape Then
   If b_cancela.Enabled = True Then
      b_cancela_Click
   End If
   Unload Me
End If

End Sub

Private Sub Form_Load()
Dim Xlafecdes As Date

On Error GoTo Aliniciar
Xentrantes = 0
Dim Xquepc As Integer
Xquepc = 0
Xlafecdes = Date - 60
cbozona.AddItem "1"
cbozona.AddItem "2"
cbozona.AddItem "3"
cbozona.AddItem "4" ' traslados coordinados y univ.las piedras
cbozona.AddItem "5" ' llamados san jacinto
cbozona.AddItem "6" ' RM
cbozona.AddItem "7" ' Traslados fuera de zona

'get_Usuario

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
''data_aft.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_aut.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_u.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_aran.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_histant.DatabaseName = App.path & "\ante.mdb"

data_med2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med2.RecordSource = "Select * from medicos order by med_nombre"
data_med2.Refresh

data_llamod.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_llamod.RecordSource = "Select * from resplla where nrolla =" & 674
data_llamod.Refresh
data_deuda.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_chof.Connect = "odbc;dsn=" & Xconexrmt & ";"

XAlta = 3

Dim Xregbusca As New ADODB.Recordset
Dim XsqlCons As String
txt_lugar.Clear
txt_locali.Clear
ConectarBD
ConbdSapp.Open
XsqlCons = "Select * from sociedad order by soc_nombre"
With Xregbusca
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open XsqlCons, ConbdSapp, , , adCmdText
End With
If Xregbusca.RecordCount > 0 Then
   Xregbusca.MoveFirst
   Do While Not Xregbusca.EOF
      txt_lugar.AddItem Xregbusca("soc_nombre")
      Xregbusca.MoveNext
   Loop
   
End If
Xregbusca.Close
XsqlCons = "Select * from zonas order by zo_nombre"
With Xregbusca
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open XsqlCons, ConbdSapp, , , adCmdText
End With
If Xregbusca.RecordCount > 0 Then
   Xregbusca.MoveFirst
   Do While Not Xregbusca.EOF
      txt_locali.AddItem Xregbusca("zo_nombre")
      Xregbusca.MoveNext
   Loop
End If

ConbdSapp.Close

If LCase(get_Usuario) = "largador" Then
   data_par.DatabaseName = App.path & "\largador\parse.mdb"
Else
   If LCase(get_Usuario) = "recepcionsur" Then
      data_par.DatabaseName = App.path & "\sur\parse.mdb"
   Else
      If LCase(get_Usuario) = "recepcionnorte" Then
         data_par.DatabaseName = App.path & "\norte\parse.mdb"
      Else
         data_par.DatabaseName = App.path & "\parse.mdb"
         Xquepc = 8
      End If
   End If
End If

data_par.Refresh
   

'data_lla2.DatabaseName = App.Path & "\llamado.mdb"
data_lla2.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_lla2.RecordSource = "select * from llamado where fecha >=#" & Format(Xlafecdes, "yyyy/mm/dd") & "# order by fecha"
'data_lla2.Refresh

'data_hist.DatabaseName = App.path & "\abmdesp.mdb"
data_hist.Connect = "odbc;dsn=" & Xconexrmt & ";"
If txt_nro.Text <> "" Then
   data_hist.RecordSource = "select * from abmdespa where idllamado =" & txt_nro.Text
Else
   data_hist.RecordSource = "select * from abmdespa where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
End If
data_hist.Refresh
  
'data_hist.RecordSource = "abm"
'data_hist.Refresh

data_mov.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_mov.RecordSource = "moviles"
data_mov.Refresh

data_clib.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_clib.RecordSource = "clientes"
'data_clib.Refresh
    
data_convbus.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_convbus.RecordSource = "convenio"
'data_convbus.Refresh
    
data_ref.DatabaseName = App.path & "\referen.mdb"
data_ref.RecordSource = "Select * from referen where mat =" & 1
data_ref.Refresh
    
data_cons.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cons.RecordSource = "Select * from consmas where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
data_cons.Refresh
    
carga_trasl

data_llasql.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_med.RecordSource = "Select * from medicos order by med_nombre"
    data_med.Refresh
'    data_lla.DatabaseName = App.Path & "\llamado.mdb"
    data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_lla.RecordSource = "select * from llamado where fecha >=#" & Format(Xlafecdes, "yyyy-mm-dd") & "# order by fecha DESC"
    data_lla.Refresh
    data_imp.DatabaseName = App.path & "\informes.mdb"
    data_imp.Refresh
    crl.ReportFileName = App.path & "\infllamado.rpt"
    'dbcbomed.RowSource = data_med
    dbcbomed.ListField = "med_nombre"
    dbcbomed.BoundColumn = "med_nombre"
    
    dbcbomed2.ListField = "med_nombre"
    dbcbomed2.BoundColumn = "med_nombre"
    
    If Xquepc = 8 Then
       If data_par.Recordset("base") = 19 Then
          igualar_lla
       Else
       End If
    End If
    
    If WDespa = 1 Then
       Frame2.Visible = False
       frm_largador.Height = 5750
       b_pend.Enabled = True
       frm_largador.Caption = "Recepción"
    Else
       b_pend.Enabled = True
       Frame2.Visible = True
       frm_largador.Caption = "Largador"
    End If
       
       If cbocolor.Text = "VERDE" Then
          cbocolor.BackColor = &HC000&
       Else
          If cbocolor.Text = "ROJO" Then
             cbocolor.BackColor = &HFF&
          Else
             If cbocolor.Text = "AMARILLO" Then
                cbocolor.BackColor = &HFFFF&
             Else
                If cbocolor.Text = "CELESTE" Then
                   cbocolor.BackColor = &HFFFF00
                Else
                   If cbocolor.Text = "AZUL" Then
                      cbocolor.BackColor = &HC00000
                   Else
                      If cbocolor.Text = "NEGRO" Then
                         cbocolor.BackColor = &H80000006
                      Else
                         cbocolor.BackColor = &HFFFFFF
                      End If
                   End If
                End If
             End If
          End If
       End If

Exit Sub

XAlta = 3
Aliniciar:
          If Err.Number = 3021 Then
             MsgBox "No se encuentran llamados recientes ALINICIAR ", vbInformation
          Else
             MsgBox "Error al cargar la ventana inicial, " & Err.Number & " " & Err.Description
          End If
          
          
End Sub

Private Sub Form_Resize()
With Image1
   .Left = 0
   .Top = 0
   .Width = Me.Width
   .Height = Me.Height
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Cerrarlo As String

If XAlta = 1 Or XAlta = 0 Then
   Cerrarlo = MsgBox("Está editando información, Desea cerrar igual sin grabar?", vbCritical + vbYesNo)
   If Cerrarlo = vbYes Then
      XAlta = 3
      b_cancela_Click
      Unload Me
   Else
      Cancel = 1
   End If
Else
    XAlta = 3
    Unload Me
    
End If

End Sub

Private Sub labcmt_Click()
b_cancela_Click

frm_cmt.Show vbModal

End Sub

Private Sub Label17_DblClick()
frm_cancella2.Show vbModal

End Sub

Private Sub Label26_Click()
frm_vertrasl.Show vbModal

End Sub

Private Sub Label5_DblClick()
Dim Xq As Integer
Dim Xelcodigoaut, Xlapersona As String
Dim Xaltaanterior As Integer
Dim Xladat, Xhoy As Date
Dim Xrecconve As New ADODB.Recordset
Dim Xsqlstr As String
Dim Xcodzoning As Integer
Dim MensajeClave3 As String
Xaltaanterior = XAlta

If txt_mat.Text <> "" Then
   If txt_mat.Text <> 0 Then
      data_clib.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Text
      data_clib.Refresh
      If data_clib.Recordset.RecordCount > 0 Then
         If IsNull(data_clib.Recordset("estado")) = False Then
            If data_clib.Recordset("estado") = 2 Or data_clib.Recordset("estado") = 3 Then
               If IsNull(data_clib.Recordset("fecha_baja")) = False Then
                  MsgBox "ATENCION!! SOCIO FIGURA DE BAJA con FECHA: " & Format(data_clib.Recordset("fecha_baja"), "dd/mm/yyyy"), vbCritical, "Mensaje"
               Else
                  MsgBox "ATENCION!! SOCIO FIGURA DE BAJA", vbCritical, "Mensaje"
               End If
            End If
         End If
         txt_nomb.SetFocus
         txt_nomb.Text = data_clib.Recordset("cl_apellid")
         If IsNull(data_clib.Recordset("cl_codced")) = False Then
            t_codced.Text = Int(data_clib.Recordset("cl_codced"))
         End If
         If txt_mat.Text >= 0 And txt_mat.Text <= 99999998 Then
            data_ref.RecordSource = "Select * from referen where mat =" & txt_mat.Text
            data_ref.Refresh
            If data_ref.Recordset.RecordCount > 0 Then
               If IsNull(data_ref.Recordset("refmat")) = False Then
                  If txt_direc.Text <> "" Then
                     txt_direc.Text = data_ref.Recordset("refmat")
                  Else
                     txt_direc.Text = data_ref.Recordset("refmat")
                  End If
               Else
               End If
            Else
               If IsNull(data_clib.Recordset("cl_direcci")) = False Then
                  If IsNull(data_clib.Recordset("cl_zona")) = False Then
                     If data_clib.Recordset("cl_zona") <> "*TODOS" Then
                        txt_direc.Text = data_clib.Recordset("cl_direcci") & "--" & data_clib.Recordset("cl_zona")
                     Else
                        txt_direc.Text = data_clib.Recordset("cl_direcci")
                     End If
                  Else
                     txt_direc.Text = data_clib.Recordset("cl_direcci")
                  End If
               End If
            End If
            If IsNull(data_clib.Recordset("cl_edad")) = False Then
               txt_edad.Text = data_clib.Recordset("cl_edad")
               If IsNull(data_clib.Recordset("cl_uniedad")) = False Then
                  If data_clib.Recordset("cl_uniedad") = "A" Then
                     cboed.ListIndex = 0
                  Else
                     If data_clib.Recordset("cl_uniedad") = "M" Then
                        cboed.ListIndex = 1
                     Else
                        If data_clib.Recordset("cl_uniedad") = "D" Then
                           cboed.ListIndex = 2
                        Else
                           cboed.ListIndex = 0
                        End If
                     End If
                  End If
               Else
                  cboed.ListIndex = 0
               End If
            Else
               txt_edad.Text = 0
               cboed.ListIndex = 0
            End If
         Else
            txt_direc.Text = ""
         End If
         If IsNull(data_clib.Recordset("cl_codconv")) = False Then
            txt_cat.Text = data_clib.Recordset("cl_codconv")
         End If
         If IsNull(data_clib.Recordset("cl_nomconv")) = False Then
            txt_nomcat.Text = data_clib.Recordset("cl_nomconv")
         End If
         If IsNull(data_clib.Recordset("cl_cedula")) = False Then
            txt_ced.Text = Int(data_clib.Recordset("cl_cedula"))
         End If
         If IsNull(data_clib.Recordset("cl_dpto")) = False Then
            If IsNull(data_clib.Recordset("cl_telefon")) = False Then
               txt_tel.Text = data_clib.Recordset("cl_dpto") & "//" & data_clib.Recordset("cl_telefon")
            Else
               txt_tel.Text = data_clib.Recordset("cl_dpto")
            End If
         Else
            If IsNull(data_clib.Recordset("cl_telefon")) = False Then
               txt_tel.Text = data_clib.Recordset("cl_telefon")
            End If
         End If
         If IsNull(data_clib.Recordset("cl_zona")) = False Then
            If data_clib.Recordset("cl_zona") <> "*TODOS" Then
               txt_locali.Text = data_clib.Recordset("cl_zona")
            Else
               txt_locali.Text = ""
            End If
         Else
            txt_locali.Text = ""
         End If
         data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & UCase(txt_cat.Text) & "' and cnv_umpago not in (1)"
         data_convbus.Refresh
         If data_convbus.Recordset.RecordCount > 0 Then
            txt_cat.Text = data_convbus.Recordset("cnv_codigo")
            txt_nomcat.Text = data_convbus.Recordset("cnv_desc")
            If IsNull(data_convbus.Recordset("cnv_fbaja")) = False Then
               MsgBox "ATENCION!! El convenio figura de BAJA, comuníquese con Administración al 097215419.", vbCritical
               MsgBox "Se ingresará cómo categoría PARTICULAR"
               txt_cat.Text = "PART"
               txt_nomcat.Text = "PARTICULARES"
            End If
         Else
            MsgBox "Convenio no encontrado.", vbCritical, "Mensaje"
            txt_cat.Text = "PART"
            txt_nomcat.Text = "PARTICULARES"
            frm_buscnvlla.Show vbModal
         End If
         
         If IsNull(data_clib.Recordset("cl_sexo")) = False Then
            If data_clib.Recordset("cl_sexo") = 1 Then
               Combo3.ListIndex = 0
            Else
               Combo3.ListIndex = 1
            End If
         Else
            Combo3.ListIndex = 0
         End If
      
          Wxquepreg = 0
          Wopszond = ""
          Xop4 = 0
          Xop5 = 0
          If txt_mat.Text <> "" Then
             Xhab = txt_mat.Text
          Else
             Xhab = 0
          End If
          
          If Check1.Value = 1 Then
              data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and mes >" & 0 & " and fecha_pago is null order by ano,mes"
              data_deuda.Refresh
              If data_deuda.Recordset.RecordCount > 0 Then
                 data_deuda.Recordset.MoveLast
                 If data_deuda.Recordset.RecordCount > 2 Then
                    Xop4 = data_deuda.Recordset("mes")
                    Xop5 = data_deuda.Recordset("ano")
                    Xq = 9
                    Wxquepreg = 2 'Deuda por cuota
                 End If
              End If
          Else
              If Trim(txt_cat.Text) = "" Then
                 data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & "AABB" & "'"
                 data_convbus.Refresh
              Else
                 data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & UCase(txt_cat.Text) & "' and cnv_sindeuda in (1) and cnv_fbaja is null"
                 data_convbus.Refresh
              End If
              If data_convbus.Recordset.RecordCount > 0 Then
                 Xq = 0
              Else
                   If data_clib.Recordset.RecordCount > 0 Then
                      Xhoy = Date
                      Xq = 0
                     data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanciacion" & "' and mes =" & 0
                     data_deuda.Refresh
                     If data_deuda.Recordset.RecordCount > 0 Then
                        data_deuda.Recordset.MoveFirst
                        Do While Not data_deuda.Recordset.EOF
                           If IsNull(data_deuda.Recordset("nro_superv")) = False Then
                              Xladat = data_deuda.Recordset("fecha") + data_deuda.Recordset("nro_superv")
                           Else
                              Xladat = data_deuda.Recordset("fecha") + 15
                           End If
                           If Format(Xladat, "yyyy/mm/dd") < Format(Xhoy, "yyyy/mm/dd") Then
                              Xq = 9
                              Wxquepreg = 1 'es deuda por servicio
                           End If
                           data_deuda.Recordset.MoveNext
                        Loop
                     End If
                     data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and mes >" & 0 & " and fecha_pago is null and origen <='" & "Refinan" & "' order by ano,mes"
                     data_deuda.Refresh
                     If data_deuda.Recordset.RecordCount > 0 Then
                        data_deuda.Recordset.MoveLast
                        If data_deuda.Recordset.RecordCount > 2 Then
                           Xop4 = data_deuda.Recordset("mes")
                           Xop5 = data_deuda.Recordset("ano")
                           Xq = 9
                           If Wxquepreg = 0 Then
                              Wxquepreg = 2 'es por cuota
                           End If
                        End If
                     End If
                     data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and fecha_pago is null and origen >='" & "Refinan" & "'"
                     data_deuda.Refresh
                     If data_deuda.Recordset.RecordCount > 0 Then
                        data_deuda.Recordset.MoveFirst
                        Do While Not data_deuda.Recordset.EOF
                           If IsNull(data_deuda.Recordset("nro_superv")) = False Then
                              Xladat = data_deuda.Recordset("fecha") + data_deuda.Recordset("nro_superv")
                           Else
                              Xladat = data_deuda.Recordset("fecha") + 30
                           End If
                           If Format(Xladat, "yyyy/mm/dd") < Format(Xhoy, "yyyy/mm/dd") Then
                              Xq = 9
                              Wxquepreg = 3 'es por refinanc
                           End If
                           data_deuda.Recordset.MoveNext
                        Loop
                     End If
                     
                     If Xq = 9 Then
                        If txt_mat.Text <> "" Then
                           XAlta = 599
                           Xtot = txt_mat.Text
                           Xhab = txt_mat.Text
                           frm_veodeuda.Show vbModal
                        Else
                           Xhab = 0
                        End If
                                      
                        Xdeb = 1
                        MensajeClave3 = MsgBox("PACIENTE CON DEUDA! ES UN LLAMADO DE URGENCIA?", vbExclamation + vbYesNo + vbDefaultButton2)
                        
                        If MensajeClave3 = vbYes Then
                           Xelcodigoaut = "URGENCIA"
                           Xq = 0
                           Xdeudasi = 0
                           data_aut.RecordSource = "select * from Codigos_aut"
                           data_aut.Refresh
                           data_aut.Recordset.AddNew
                           data_aut.Recordset("fecha") = Date
                           data_aut.Recordset("usuario") = Mid(txt_nomb.Text, 1, 50)
                           data_aut.Recordset("codaut") = "URGENCIA"
                           If txt_mat.Text <> "" Then
                              data_aut.Recordset("socio") = txt_mat.Text
                           Else
                              data_aut.Recordset("socio") = txt_mat.Text
                           End If
                           data_aut.Recordset("modulo") = "DESPACHO"
                           data_aut.Recordset("usuario_caja") = WElusuario
                           data_aut.Recordset.Update
                        Else
                            frm_autoriza.Show vbModal
                            '14063
                            '117670
                            '5112
                            Xelcodigoaut = InputBox("SOCIO CON CRÉDITOS PENDIENTES O CUOTAS, INGRESE CODIGO DE AUTORIZACIÓN SI ES CLAVE 3", "SOCIO CON CRÉDITOS PENDIENTES", Wopszond)
                            If Trim(Xelcodigoaut) <> "" Then
                               data_aut.RecordSource = "select * from Codigos_aut where codaut ='" & Trim(Xelcodigoaut) & "' and socio =" & txt_mat.Text
                               data_aut.Refresh
                               If data_aut.Recordset.RecordCount > 0 Then
                                  Xq = 0
                                  Xdeudasi = 0
                               Else
                                  MsgBox "ATENCION! No se encuentra código de autorización, realice nuevamente la autorización o comunique a Administración", vbCritical
                                  Xq = 9
                                  Xdeudasi = 9
                               End If
                            Else
                               MsgBox "Socio con créditos o cuotas(>=3) pendientes, NO SE PODRÁ GRABAR LLAMADO CLAVE 3.", vbCritical
                               Xq = 9
                               Xdeudasi = 9
                            End If
                        End If
                     Else
                        Xdeudasi = 0
                     End If
                     If XAlta = 599 Then
                        XAlta = Xaltaanterior
                     End If
                     If IsNull(data_clib.Recordset("saldo_chc2")) = False Then
                        If data_clib.Recordset("saldo_chc2") = 1 Then
                           Xq = 11
                        End If
                        If Xq = 11 Then
                           MsgBox "ATENCION!! Socio con servicios RESTRINGIDOS! Estimado Funcionario NO dar servicio." & chr(13) _
                           & "El hacerlo estará bajo su exclusiva responsabilidad." & chr(13) & "El sistema no permitirá la continuidad de dicho servicio.", vbCritical, "SOCIOS"
                           MsgBox "SI ES UN LLAMADO CLAVE 3, DEBERA SOLICITAR AUTORIZACION al 097215419 PARA PODER GRABAR DATOS", vbInformation, "LLAMADO"
                           Xdeudasi = 9
                        End If
                     End If
                   End If
              End If
          End If
      
      
          Wopspro = 99
          data_parsec.DatabaseName = App.path & "\mensa.mdb"
          data_parsec.RecordSource = "mensaje"
          data_parsec.Refresh
          
          If Check1.Value <> 1 Then
             If data_clib.Recordset.RecordCount > 0 Then
                If IsNull(data_clib.Recordset("cl_grupo")) = False Then
                   Xcodzoning = data_clib.Recordset("cl_grupo")
                Else
                   Xcodzoning = 0
                End If
             Else
                Xcodzoning = 0
             End If
             If (Xcodzoning = 400 Or Xcodzoning = 401 Or Xcodzoning = 402 Or Xcodzoning = 403 Or Xcodzoning = 670 Or Xcodzoning = 671) And _
                (txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM") Then
                 Wopscob = 0
             Else
                If Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Or Val(cbozona.Text) = 5 Or Val(cbozona.Text) = 6 Then
                   If txt_cat.Text <> "" Then
                      If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Or txt_cat.Text = "UNIVS" Or _
                         txt_cat.Text = "UNNSAM" Or txt_cat.Text = "HEVANO" Or txt_cat.Text = "EVNSAM" Or _
                         txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM" Or txt_cat.Text = "GANOS" Or _
                         txt_cat.Text = "CASANO" Or txt_cat.Text = "CASNSA" Then
                         ConectarBD
                         ConbdSapp.Open
                         Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806)"
                         With Xrecconve
                             .CursorLocation = adUseClient
                             .CursorType = adOpenKeyset
                             .LockType = adLockOptimistic
                             .Open Xsqlstr, ConbdSapp, , , adCmdText
                         End With
                         If Xrecconve.RecordCount > 0 Then
                            Wopspro = 0
                            Wopscob = 0
                            ConbdSapp.Close
                            data_parsec.DatabaseName = App.path & "\mensa.mdb"
                            data_parsec.RecordSource = "mensaje"
                            data_parsec.Refresh
                            data_parsec.Recordset.Edit
                            data_parsec.Recordset("text") = "RECUERDE! Confirmar socio con la mutualista."
                            data_parsec.Recordset.Update
                            frm_mensajesvar.Show vbModal
                         Else
                           ConbdSapp.Close
                           If IsNull(data_clib.Recordset("cl_decuota")) = False Then
                              If data_clib.Recordset("cl_decuota") = 0 Or _
                                 data_clib.Recordset("cl_decuota") = 1 Or _
                                 data_clib.Recordset("cl_decuota") = 3 Or _
                                 data_clib.Recordset("cl_decuota") = 4 Then
                                 data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                 data_parsec.RecordSource = "mensaje"
                                 data_parsec.Refresh
                                 data_parsec.Recordset.Edit
                                 If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                    If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual. Documentación a presentar:" _
                                       & "Fotocopia de CI vigente. Comunique al funcionario del móvil " _
                                       & "para realizar la misma." _
                                       & " RECUERDE! Confirmar socio con la mutualista."
                                    Else
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual. Requerimientos:" _
                                       & "Fotocopia de CI vigente. Comprobante domicilio:" _
                                       & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                       & " a nombre del cliente que sea del mes corriente o anterior. Comunique al funcionario del móvil" _
                                       & " para realizar la misma. RECUERDE! Confirmar socio con la mutualista."
                                    End If
                                 Else
                                    data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                    & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                    chr(13) & "para realizar la misma." _
                                    & " RECUERDE! Confirmar socio con la mutualista."
                                 End If
                                 data_parsec.Recordset.Update
                                 data_parsec.Refresh
                                 frm_mensajesvar.Show vbModal
                              Else
                                 ConectarBD
                                 ConbdSapp.Open
                                 Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806)"
                                 With Xrecconve
                                     .CursorLocation = adUseClient
                                     .CursorType = adOpenKeyset
                                     .LockType = adLockOptimistic
                                     .Open Xsqlstr, ConbdSapp, , , adCmdText
                                 End With
                                 If Xrecconve.RecordCount > 0 Then
                                 Else
                                    data_parsec.Recordset.Edit
                                    If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                       If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                          data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                          & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                          chr(13) & "para realizar la misma." _
                                          & " RECUERDE! Confirmar socio con la mutualista."
                                       Else
                                          data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" _
                                          & "Fotocopia de CI vigente. Comprobante domicilio:" _
                                          & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                          & " a nombre del cliente que sea del mes corriente o anterior.Comunique al funcionario del móvil " _
                                          & "para realizar la misma." _
                                          & "RECUERDE! Confirmar socio con la mutualista."
                                       End If
                                    Else
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                       & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                       chr(13) & "para realizar la misma." _
                                       & " RECUERDE! Confirmar socio con la mutualista."
                                    End If
                                    data_parsec.Recordset.Update
                                    data_parsec.Refresh
                                    frm_mensajesvar.Show vbModal
                                 End If
                                 ConbdSapp.Close
                              End If
                           Else
                              ConectarBD
                              ConbdSapp.Open
                              Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806)"
                              With Xrecconve
                                  .CursorLocation = adUseClient
                                  .CursorType = adOpenKeyset
                                  .LockType = adLockOptimistic
                                  .Open Xsqlstr, ConbdSapp, , , adCmdText
                              End With
                              If Xrecconve.RecordCount > 0 Then
                              Else
                                 data_parsec.Recordset.Edit
                                 If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                    If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                       & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                       chr(13) & "para realizar la misma." _
                                       & "RECUERDE! confirmar socio con la mutualista."
                                    Else
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" _
                                       & "Fotocopia de CI vigente. Comprobante domicilio:" _
                                       & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                       & " a nombre del cliente y del mes corriente o anterior.Comunique al funcionario del móvil" _
                                       & " para realizar la misma." _
                                       & "RECUERDE! Confirmar socio con la mutualista."
                                    End If
                                 Else
                                    data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                    & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                    chr(13) & "para realizar la misma." _
                                    & " RECUERDE! Confirmar socio con la mutualista."
                                 End If
                                 data_parsec.Recordset.Update
                                 data_parsec.Refresh
                                 frm_mensajesvar.Show vbModal
                              End If
                              ConbdSapp.Close
                           End If
                         End If
                      End If
                   End If
                End If
             End If
          
          End If
          Wopspro = 0
      
      End If
   End If
End If

End Sub

Private Sub Label7_DblClick()
If txt_cat.Text <> "" Then
   frm_verdesconv.Show vbModal
End If

End Sub

Private Sub Label8_DblClick()

If Label8.Caption = "Cédula:" Then
   Label8.Caption = "DNI:"
   t_codced.Visible = False
   t_codced.Text = 0
Else
   Label8.Caption = "Cédula:"
   t_codced.Visible = True
   t_codced.Text = ""
End If

'If txt_mat.Text <> "" Then
'   XAlta = 599
'   Xtot = txt_mat.Text
'   frm_veodeuda.Show vbModal
'Else
'   MsgBox "Debe ingresar matrícula"
'End If

End Sub

Private Sub mfecasig_GotFocus()
mfecasig.Text = Format(Date, "dd/mm/yyyy")
If txt_horasig.Enabled = True Then
   txt_horasig.Text = Format(Time, "HH:mm")
End If

End Sub

Private Sub mfecasig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_horasig.Enabled = True Then
      txt_horasig.SetFocus
   End If
End If

End Sub

Private Sub mfecasig_LostFocus()
If mfecasig.Text = "__/__/____" Then
Else
   If IsDate(mfecasig.Text) = False Then
      MsgBox "Error en fecha", vbCritical, "Mensaje"
      mfecasig.SetFocus
   End If
End If

End Sub

Private Sub mfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_hora.SetFocus
   txt_hora.Text = Format(Time, "HH:mm")
   txt_usua.Text = WElusuario
End If

End Sub

Private Sub mfecha_LostFocus()
If IsDate(mfecha.Text) = False Then
   MsgBox "Error en fecha", vbCritical, "Mensaje"
   mfecha.SetFocus
End If

End Sub

Private Sub mftrassol_GotFocus()
If mftrassol.Text = "__/__/____" Then
   mftrassol.Text = Format(Date, "dd/mm/yyyy")
End If

End Sub

Private Sub mhtrassol_GotFocus()
If mhtrassol.Text = "__:__" Then
   mhtrassol.Text = Format(Time, "HH:mm")
End If

End Sub

Private Sub mllegada_GotFocus()
mllegada.Text = Format(Date, "dd/mm/yyyy")
txt_horlle.Text = Format(Time, "HH:mm")

End Sub

Private Sub mllegada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_horlle.SetFocus
End If

End Sub

Private Sub mllegada_LostFocus()
If mllegada.Text = "__/__/____" Then
   txt_horlle.SetFocus
Else
   If IsDate(mllegada.Text) = False Then
      MsgBox "Error en fecha", vbCritical, "Mensaje"
      mllegada.SetFocus
   End If
End If

End Sub

Private Sub msalida_GotFocus()
msalida.Text = Format(Date, "dd/mm/yyyy")
If txt_horsal.Enabled = True Then
   txt_horsal.Text = Format(Time, "HH:mm")
End If

End Sub

Private Sub msalida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_horsal.Enabled = True Then
      txt_horsal.SetFocus
   End If
End If

End Sub

Private Sub msalida_LostFocus()
If msalida.Text = "__/__/____" Then
   txt_horsal.SetFocus
Else
   If IsDate(msalida.Text) = False Then
      MsgBox "Error en fecha", vbCritical, "Mensaje"
      msalida.SetFocus
   End If
End If

End Sub

Private Sub mtd_GotFocus()
mtd.Text = Format(Date, "dd/mm/yyyy")
txt_hortd.Text = Format(Time, "HH:mm")

End Sub

Private Sub mtd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_hortd.SetFocus
End If

End Sub

Private Sub mtd_LostFocus()
If mtd.Text = "__/__/____" Then
   txt_hortd.SetFocus
Else
   If IsDate(mtd.Text) = False Then
      MsgBox "Error en fecha", vbCritical, "Mensaje"
      mtd.SetFocus
   End If
End If

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub t_codced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomb.SetFocus
End If

End Sub

Private Sub txt_ante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_obs.SetFocus
End If

End Sub



Private Sub txt_cat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomcat.SetFocus
End If

End Sub

Private Sub txt_cat_LostFocus()
Dim Xrecconve As New ADODB.Recordset
Dim Xsqlstr As String
Dim Xcodzoning As Integer
Dim Xfechacartas As Date
Xfechacartas = Date - 150

If txt_cat.Text <> "" Then
   If txt_cat.Text = "911" Or UCase(txt_cat.Text) = "911B" Then
      Command1.Visible = True
      Command1.Enabled = True
      txt_locali.Visible = True
   Else
      Command1.Enabled = False
      Command1.Enabled = False
'      Label40.Visible = False
'      txt_locali.Visible = False
   End If
   data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & UCase(txt_cat.Text) & "' and cnv_umpago not in (1) and cnv_alta ='" & "SI" & "'"
   data_convbus.Refresh
   If data_convbus.Recordset.RecordCount > 0 Then
      txt_cat.Text = data_convbus.Recordset("cnv_codigo")
      txt_nomcat.Text = data_convbus.Recordset("cnv_desc")
      If IsNull(data_convbus.Recordset("cnv_fbaja")) = False Then
         MsgBox "ATENCION!! El convenio figura de BAJA, comuníquese con Administración al 097215419.", vbCritical
         MsgBox "Se ingresará cómo categoría PARTICULAR"
         txt_cat.Text = "PART"
         txt_nomcat.Text = "PARTICULARES"
      End If
   Else
      MsgBox "Convenio no encontrado.", vbCritical, "Mensaje"
      txt_cat.Text = "PART"
      txt_nomcat.Text = "PARTICULARES"
      frm_buscnvlla.Show vbModal
   End If
   If txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "CERDGI" Or _
      txt_cat.Text = "CERADU" Or txt_cat.Text = "CERHEV" Or txt_cat.Text = "CERMAT" Or txt_cat.Text = "CERSEV" Or txt_cat.Text = "CERVIS" Then
      If txt_costo.Enabled = True Then
         txt_costo.Text = 0
      End If
   End If

   If txt_ced.Text <> "" Then
      If txt_ced.Text <> 0 Then
         data_clib.RecordSource = "Select * from clientes where cl_cedula =" & txt_ced.Text
         data_clib.Refresh
         If data_clib.Recordset.RecordCount > 0 Then
            If XwYalomostro = 99 Then
            Else
                If IsNull(data_clib.Recordset("cl_grupo")) = False Then
                   Xcodzoning = data_clib.Recordset("cl_grupo")
                Else
                   Xcodzoning = 0
                End If
                If (Xcodzoning = 400 Or Xcodzoning = 401 Or Xcodzoning = 402 Or Xcodzoning = 403 Or Xcodzoning = 670 Or Xcodzoning = 671) And _
                   (txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM") Then
                    Wopscob = 0
                Else
                    Wopspro = 99
                    If Check1.Value <> 1 Then
                       If Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                          If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Or txt_cat.Text = "UNIVS" Or _
                             txt_cat.Text = "UNNSAM" Or txt_cat.Text = "HEVANO" Or txt_cat.Text = "EVNSAM" Or _
                             txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM" Or txt_cat.Text = "GANOS" Or _
                             txt_cat.Text = "CASANO" Or txt_cat.Text = "CASNSA" Then
                             If data_clib.Recordset.RecordCount > 0 Then
                                ConectarBD
                                ConbdSapp.Open
                                Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                                With Xrecconve
                                    .CursorLocation = adUseClient
                                    .CursorType = adOpenKeyset
                                    .LockType = adLockOptimistic
                                    .Open Xsqlstr, ConbdSapp, , , adCmdText
                                End With
                                If Xrecconve.RecordCount > 0 Then
                                   Wopspro = 0
                                   ConbdSapp.Close
                                   data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                   data_parsec.RecordSource = "mensaje"
                                   data_parsec.Refresh
                                   data_parsec.Recordset.Edit
                                   data_parsec.Recordset("text") = "RECUERDE! Confirmar socio con la mutualista."
                                   data_parsec.Recordset.Update
                                   frm_mensajesvar.Show vbModal
                                Else
                                   ConbdSapp.Close
                                   data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                   data_parsec.RecordSource = "mensaje"
                                   data_parsec.Refresh
                                   data_parsec.Recordset.Edit
                                   If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                      If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                         data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Documentación a presentar:" & chr(13) _
                                         & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil " & _
                                         "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                                      Else
                                         data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" & _
                                         "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                                         & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                         & " a nombre del cliente del mes corriente o anterior.) Comunique al funcionario del móvil " & _
                                         "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                                      End If
                                   Else
                                      data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                      & "Fotocopia de CI vigente.Comunique al funcionario del móvil " & _
                                      chr(13) & "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                                   End If
                                   data_parsec.Recordset.Update
                                   data_parsec.Refresh
                                   frm_mensajesvar.Show vbModal
                                End If
                             Else
                                   data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                   data_parsec.RecordSource = "mensaje"
                                   data_parsec.Refresh
                                   data_parsec.Recordset.Edit
                                   If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                      If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                         data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Documentación a presentar:" & chr(13) _
                                         & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil " & _
                                         "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                                      Else
                                         data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" & _
                                         "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                                         & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                         & " a nombre del cliente del mes corriente o anterior.) Comunique al funcionario del móvil " & _
                                         "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                                      End If
                                   Else
                                      data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                      & "Fotocopia de CI vigente.Comunique al funcionario del móvil " & _
                                      chr(13) & "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                                   End If
                                   data_parsec.Recordset.Update
                                   data_parsec.Refresh
                                   frm_mensajesvar.Show vbModal
                             End If
                          End If
                       End If
                    End If
                End If
            End If
         Else
            Wopspro = 99
            If Check1.Value <> 1 Then
               If Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                  If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Or txt_cat.Text = "UNIVS" Or _
                     txt_cat.Text = "UNNSAM" Or txt_cat.Text = "HEVANO" Or txt_cat.Text = "EVNSAM" Or _
                     txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM" Or txt_cat.Text = "GANOS" Or _
                     txt_cat.Text = "CASANO" Or txt_cat.Text = "CASNSA" Then
                     If data_clib.Recordset.RecordCount > 0 Then
                        ConectarBD
                        ConbdSapp.Open
                        Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                        With Xrecconve
                            .CursorLocation = adUseClient
                            .CursorType = adOpenKeyset
                            .LockType = adLockOptimistic
                            .Open Xsqlstr, ConbdSapp, , , adCmdText
                        End With
                        If Xrecconve.RecordCount > 0 Then
                           Wopspro = 0
                           ConbdSapp.Close
                           data_parsec.DatabaseName = App.path & "\mensa.mdb"
                           data_parsec.RecordSource = "mensaje"
                           data_parsec.Refresh
                           data_parsec.Recordset.Edit
                           data_parsec.Recordset("text") = "RECUERDE! Confirmar socio con la mutualista."
                           data_parsec.Recordset.Update
                           frm_mensajesvar.Show vbModal
                        Else
                           ConbdSapp.Close
                           data_parsec.DatabaseName = App.path & "\mensa.mdb"
                           data_parsec.RecordSource = "mensaje"
                           data_parsec.Refresh
                           data_parsec.Recordset.Edit
                           If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                              If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Documentación a presentar:" & chr(13) _
                                 & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil " & _
                                 "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                              Else
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" & _
                                 "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                                 & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                 & " a nombre del cliente del mes corriente o anterior.) Comunique al funcionario del móvil " & _
                                 "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                              End If
                           Else
                              data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                              & "Fotocopia de CI vigente.Comunique al funcionario del móvil " & _
                              chr(13) & "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                           End If
                           data_parsec.Recordset.Update
                           data_parsec.Refresh
                           frm_mensajesvar.Show vbModal
                        End If
                     Else
                           data_parsec.DatabaseName = App.path & "\mensa.mdb"
                           data_parsec.RecordSource = "mensaje"
                           data_parsec.Refresh
                           data_parsec.Recordset.Edit
                           If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                              If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Documentación a presentar:" & chr(13) _
                                 & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil " & _
                                 "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                              Else
                                 data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" & _
                                 "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                                 & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                 & " a nombre del cliente del mes corriente o anterior.) Comunique al funcionario del móvil " & _
                                 "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                              End If
                           Else
                              data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                              & "Fotocopia de CI vigente.Comunique al funcionario del móvil " & _
                              chr(13) & "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                           End If
                           data_parsec.Recordset.Update
                           data_parsec.Refresh
                           frm_mensajesvar.Show vbModal
                     End If
                  End If
               End If
            End If
         End If
      Else
         Wopspro = 99
         If Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
            If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Or txt_cat.Text = "UNIVS" Or _
               txt_cat.Text = "UNNSAM" Or txt_cat.Text = "HEVANO" Or txt_cat.Text = "EVNSAM" Or _
               txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM" Or txt_cat.Text = "GANOS" Or _
               txt_cat.Text = "CASANO" Or txt_cat.Text = "CASNSA" Then
               data_parsec.DatabaseName = App.path & "\mensa.mdb"
               data_parsec.RecordSource = "mensaje"
               data_parsec.Refresh
               data_parsec.Recordset.Edit
               If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                  If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                     data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Documentación a presentar:" & chr(13) _
                     & "Fotocopia de CI vigente. Comunique al funcionario del móvil " & _
                     "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                  Else
                     data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual. Requerimientos:" & _
                     "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                     & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                     & " a nombre del cliente del mes corriente o anterior.Comunique al funcionario del móvil " & _
                     "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
                  End If
               Else
                  data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Documentación a presentar:" & chr(13) _
                  & "Fotocopia de CI vigente. Comunique al funcionario del móvil " & _
                  "para realizar la misma. RECUERDE! Confirmar socio con la mutualista."
               End If
               data_parsec.Recordset.Update
               data_parsec.Refresh
               frm_mensajesvar.Show vbModal
            End If
         End If
      End If
   Else
      Wopspro = 99
      If Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
         If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Or txt_cat.Text = "UNIVS" Or _
            txt_cat.Text = "UNNSAM" Or txt_cat.Text = "HEVANO" Or txt_cat.Text = "EVNSAM" Or _
            txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM" Or txt_cat.Text = "GANOS" Or _
            txt_cat.Text = "CASANO" Or txt_cat.Text = "CASNSA" Then
            data_parsec.DatabaseName = App.path & "\mensa.mdb"
            data_parsec.RecordSource = "mensaje"
            data_parsec.Refresh
            data_parsec.Recordset.Edit
            If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
               If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                  data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual. Documentación a presentar:" & chr(13) _
                  & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil " & _
                  "para realizar la misma. RECUERDE! Confirmar socio con la mutualista."
               Else
                  data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" & _
                  "Fotocopia de CI vigente. Comprobante domicilio (puede ser:" _
                  & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                  & " a nombre del cliente del mes corriente o anterior.).Comunique al funcionario del móvil " & _
                  "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
               End If
            Else
               data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
               & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
               chr(13) & "para realizar la misma.RECUERDE! Confirmar socio con la mutualista."
            End If
            data_parsec.Recordset.Update
            data_parsec.Refresh
            frm_mensajesvar.Show vbModal
         End If
      End If
   End If
   Wopspro = 0
End If


End Sub

Private Sub txt_ced_Change()
If Not IsNumeric(txt_ced.Text) And _
 txt_ced.Text <> "" Then
 Beep
 MsgBox "Se debe ingresar solo números"
txt_ced.Text = ""
txt_ced.SetFocus
End If
 
End Sub

Private Sub txt_ced_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If t_codced.Visible = True Then
      t_codced.SetFocus
   Else
      txt_nomb.SetFocus
   End If
'   txt_direc.SetFocus
End If

End Sub

Private Sub txt_ced_LostFocus()
Dim Xdeseabuscar As String
Dim Xelcodigoaut, Xlapersona As String
Dim Xaltaanterior As Integer
Dim Xrecconve As New ADODB.Recordset
Dim Xsqlstr As String
Dim Xcodzoning As Integer
Dim MensajeClave3 As String
Dim Xfechacartas As Date
Xfechacartas = Date - 150

Wopscob = 0
XwYalomostro = 0
Xaltaanterior = XAlta
    If txt_ced.Text <> "" Then
       If txt_ced.Text <> 0 Then
          data_clib.RecordSource = "Select * from clientes where cl_cedula =" & txt_ced.Text
          data_clib.Refresh
          If data_clib.Recordset.RecordCount > 0 Then
             If IsNull(data_clib.Recordset("cl_codced")) = False Then
                t_codced.Text = Int(data_clib.Recordset("cl_codced"))
             Else
                t_codced.Text = 0
             End If
             ''If txt_direc.Text = "" Then
                If IsNull(data_clib.Recordset("estado")) = False Then
                   If data_clib.Recordset("estado") = 2 Or data_clib.Recordset("estado") = 3 Then
                      If IsNull(data_clib.Recordset("fecha_baja")) = False Then
                         MsgBox "ATENCION!! SOCIO FIGURA DE BAJA con FECHA: " & Format(data_clib.Recordset("fecha_baja") & " Comuníquese con administración al 097215419.", "dd/mm/yyyy"), vbCritical, "Mensaje"
                      Else
                         MsgBox "ATENCION!! SOCIO FIGURA DE BAJA. Comuníquese con administración al 097215419", vbCritical, "Mensaje"
                      End If
                   Else
                      If IsNull(data_clib.Recordset("cl_ruc")) = False Then
                         MsgBox "ATENCION! SOCIO TAMBIEN ACTIVO EN OTRO CONVENIO: " & data_clib.Recordset("cl_ruc"), vbInformation
                      End If
                   End If
                End If
                txt_nomb.SetFocus
                txt_nomb.Text = data_clib.Recordset("cl_apellid")
                txt_mat.Text = Int(data_clib.Recordset("cl_codigo"))
                If txt_ced.Text > 0 Then
                   If Check1.Value = 1 Then
                   Else
                      data_ref.RecordSource = "Select * from referen where mat =" & data_clib.Recordset("cl_codigo")
                      data_ref.Refresh
                      If data_ref.Recordset.RecordCount > 0 Then
                         If IsNull(data_ref.Recordset("refmat")) = False Then
                            txt_direc.Text = data_ref.Recordset("refmat")
                         Else
'                         txt_referen.Text = ""
                         End If
                      Else
'                      txt_referen.Text = ""
                         If IsNull(data_clib.Recordset("cl_direcci")) = False Then
                            If IsNull(data_clib.Recordset("cl_zona")) = False Then
                               If data_clib.Recordset("cl_zona") <> "*TODOS" Then
                                  txt_direc.Text = data_clib.Recordset("cl_direcci") & "--" & data_clib.Recordset("cl_zona")
                               Else
                                  txt_direc.Text = data_clib.Recordset("cl_direcci")
                               End If
                            Else
                               txt_direc.Text = data_clib.Recordset("cl_direcci")
                            End If
                         End If
                      End If
                      Dim Xlafecdecons As Date
                      Xlafecdecons = Date - 4
'                    data_cons.Recordset.FindFirst "mat =" & txt_mat.Text
                      data_cons.RecordSource = "Select * from consmas where mat =" & txt_ced.Text & " and fecha >=#" & Format(Xlafecdecons, "yyyy/mm/dd") & "#"
                      data_cons.Refresh
                      If data_cons.Recordset.RecordCount > 0 Then
                         data_cons.Recordset.MoveLast
                         MsgBox "CONSULTAS EN LAS 48HS -ULTIMA CONSULTA EL: " & Format(data_cons.Recordset("fecha"), "dd/mm/yyyy") & " POR: " & data_cons.Recordset("motivo"), vbInformation, "Mensaje"
                      End If
                      If IsNull(data_clib.Recordset("cl_edad")) = False Then
                         txt_edad.Text = data_clib.Recordset("cl_edad")
                         If IsNull(data_clib.Recordset("cl_uniedad")) = False Then
                            If data_clib.Recordset("cl_uniedad") = "A" Then
                               cboed.ListIndex = 0
                            Else
                               If data_clib.Recordset("cl_uniedad") = "M" Then
                                  cboed.ListIndex = 1
                               Else
                                  If data_clib.Recordset("cl_uniedad") = "D" Then
                                     cboed.ListIndex = 2
                                  Else
                                     cboed.ListIndex = 0
                                  End If
                               End If
                            End If
                         Else
                            cboed.ListIndex = 0
                         End If
                      Else
                         txt_edad.Text = 0
                         cboed.ListIndex = 0
                      End If
                   End If
                Else
'                    txt_direc.Text = ""
                End If
                If Check1.Value = 1 Then
                Else
                   If IsNull(data_clib.Recordset("cl_codconv")) = False Then
                      txt_cat.Text = data_clib.Recordset("cl_codconv")
                   End If
                   If IsNull(data_clib.Recordset("cl_nomconv")) = False Then
                      txt_nomcat.Text = data_clib.Recordset("cl_nomconv")
                   End If
                   If IsNull(data_clib.Recordset("cl_dpto")) = False Then
                      If IsNull(data_clib.Recordset("cl_telefon")) = False Then
                         txt_tel.Text = data_clib.Recordset("cl_dpto") & "//" & data_clib.Recordset("cl_telefon")
                      Else
                         txt_tel.Text = data_clib.Recordset("cl_dpto")
                      End If
                   Else
                      If IsNull(data_clib.Recordset("cl_telefon")) = False Then
                         txt_tel.Text = data_clib.Recordset("cl_telefon")
                      End If
                   End If
                   If IsNull(data_clib.Recordset("cl_zona")) = False Then
                      If data_clib.Recordset("cl_zona") <> "*TODOS" Then
                         txt_locali.Text = data_clib.Recordset("cl_zona")
                      Else
                         txt_locali.Text = ""
                      End If
                   Else
                      txt_locali.Text = ""
                   End If
                   If IsNull(data_clib.Recordset("cl_grupo")) = False Then
                      If data_clib.Recordset("cl_grupo") >= 100 And data_clib.Recordset("cl_grupo") <= 530 Then
                         cbozona.Text = "1"
                      Else
                         If data_clib.Recordset("cl_grupo") >= 600 And data_clib.Recordset("cl_grupo") <= 689 Then
                            cbozona.Text = "2"
                         Else
                            If data_clib.Recordset("cl_grupo") >= 700 And data_clib.Recordset("cl_grupo") <= 788 Then
                               cbozona.Text = "1"
                            Else
                               cbozona.Text = "2"
                            End If
                         End If
                      End If
                   Else
                      cbozona.Text = "1"
                   End If
                   If txt_cat.Text <> "" Then
                      data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & UCase(txt_cat.Text) & "' and cnv_umpago not in (1)"
                      data_convbus.Refresh
                      If data_convbus.Recordset.RecordCount > 0 Then
                         If IsNull(data_convbus.Recordset("cnv_fbaja")) = False Then
                            MsgBox "ATENCION!! El convenio figura de BAJA, Comuníquese con Administración al 097215419.", vbCritical
                            MsgBox "Se ingresará cómo categoría PARTICULAR"
                            txt_cat.Text = "PART"
                            txt_nomcat.Text = "PARTICULARES"
                         End If
                      Else
                         MsgBox "Convenio no encontrado.", vbCritical, "Mensaje"
                         txt_cat.Text = "PART"
                         txt_nomcat.Text = "PARTICULARES"
                         frm_buscnvlla.Show vbModal
                      End If
                   End If
                End If
                If IsNull(data_clib.Recordset("cl_sexo")) = False Then
                   If data_clib.Recordset("cl_sexo") = 1 Then
                      Combo3.ListIndex = 0
                   Else
                      Combo3.ListIndex = 1
                   End If
                Else
                   Combo3.ListIndex = 0
                End If
          Else
              If Check1.Value = 1 Then
              Else
                 data_ref.RecordSource = "Select * from referen where mat =" & txt_ced.Text
                 data_ref.Refresh
                 If data_ref.Recordset.RecordCount > 0 Then
                    If IsNull(data_ref.Recordset("refmat")) = False Then
                       txt_direc.Text = data_ref.Recordset("refmat")
                    End If
                 End If
                Xdeseabuscar = MsgBox("DOCUMENTO NO ENCONTRADO EN PADRÓN, DESEA BUSCAR DATOS ?", vbInformation + vbYesNo, "DESPACHO")
                If Xdeseabuscar = vbYes Then
                   frm_buslla.Show vbModal
                End If
             End If
          End If
          data_histant.RecordSource = "Select * from ante where ced =" & txt_ced.Text
          data_histant.Refresh
          If data_histant.Recordset.RecordCount > 0 Then
             txt_ante.Text = data_histant.Recordset("ante")
          Else
             txt_ante.Text = ""
          End If
          Wxquepreg = 0
          Wopszond = ""
          Xop4 = 0
          Xop5 = 0
          If txt_mat.Text <> "" Then
             Xhab = txt_mat.Text
          Else
             Xhab = 0
          End If
          Dim Xq As Integer
        If Check1.Value = 1 Then
           If data_clib.Recordset.RecordCount > 0 Then
              data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and mes >" & 0 & " and fecha_pago is null order by ano,mes"
              data_deuda.Refresh
              If data_deuda.Recordset.RecordCount > 0 Then
                 data_deuda.Recordset.MoveLast
                 If data_deuda.Recordset.RecordCount > 2 Then
                    Xop4 = data_deuda.Recordset("mes")
                    Xop5 = data_deuda.Recordset("ano")
                    Xq = 9
                    Wxquepreg = 2 'Deuda por cuota
                 End If
              End If
           End If
        Else
            If Trim(txt_cat.Text) = "" Then
               data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & "AABB" & "'"
               data_convbus.Refresh
            Else
               data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & UCase(txt_cat.Text) & "' and cnv_sindeuda in (1) and cnv_fbaja is null"
               data_convbus.Refresh
            End If
            If data_convbus.Recordset.RecordCount > 0 Then
               Xq = 0
            Else
                 If data_clib.Recordset.RecordCount > 0 Then
                    Dim Xladat, Xhoy As Date
                    Xhoy = Date
                    Xq = 0
                   data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanciacion" & "' and mes =" & 0
                   data_deuda.Refresh
                   If data_deuda.Recordset.RecordCount > 0 Then
                      data_deuda.Recordset.MoveFirst
                      Do While Not data_deuda.Recordset.EOF
                         If IsNull(data_deuda.Recordset("nro_superv")) = False Then
                            Xladat = data_deuda.Recordset("fecha") + data_deuda.Recordset("nro_superv")
                         Else
                            Xladat = data_deuda.Recordset("fecha") + 15
                         End If
                         If Format(Xladat, "yyyy/mm/dd") < Format(Xhoy, "yyyy/mm/dd") Then
                            Xq = 9
                            Wxquepreg = 1 'es deuda por servicio
                         End If
                         data_deuda.Recordset.MoveNext
                      Loop
                   End If
                   data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and mes >" & 0 & " and fecha_pago is null and origen <='" & "Refinan" & "' order by ano,mes"
                   data_deuda.Refresh
                   If data_deuda.Recordset.RecordCount > 0 Then
                      data_deuda.Recordset.MoveLast
                      If data_deuda.Recordset.RecordCount > 2 Then
                         Xop4 = data_deuda.Recordset("mes")
                         Xop5 = data_deuda.Recordset("ano")
                         Xq = 9
                         If Wxquepreg = 0 Then
                            Wxquepreg = 2 'es por cuota
                         End If
                      End If
                   End If
                   data_deuda.RecordSource = "Select * from deudas where cliente =" & data_clib.Recordset("cl_codigo") & " and fecha_pago is null and origen >='" & "Refinan" & "'"
                   data_deuda.Refresh
                   If data_deuda.Recordset.RecordCount > 0 Then
                      data_deuda.Recordset.MoveFirst
                      Do While Not data_deuda.Recordset.EOF
                         If IsNull(data_deuda.Recordset("nro_superv")) = False Then
                            Xladat = data_deuda.Recordset("fecha") + data_deuda.Recordset("nro_superv")
                         Else
                            Xladat = data_deuda.Recordset("fecha") + 30
                         End If
                         If Format(Xladat, "yyyy/mm/dd") < Format(Xhoy, "yyyy/mm/dd") Then
                            Xq = 9
                            Wxquepreg = 3 'es por refinanc
                         End If
                         data_deuda.Recordset.MoveNext
                      Loop
                   End If
                   
                   If Xq = 9 Then
                      If txt_mat.Text <> "" Then
                         XAlta = 599
                         Xtot = txt_mat.Text
                         Xhab = txt_mat.Text
                         frm_veodeuda.Show vbModal
                      Else
                         Xhab = 0
                      End If
                                    
                      Xdeb = 1
                      MensajeClave3 = MsgBox("PACIENTE CON DEUDA! ES UN LLAMADO DE URGENCIA?", vbExclamation + vbYesNo + vbDefaultButton2)
                      
                      If MensajeClave3 = vbYes Then
                         Xelcodigoaut = "URGENCIA"
                         Xq = 0
                         Xdeudasi = 0
                         data_aut.RecordSource = "select * from Codigos_aut"
                         data_aut.Refresh
                         data_aut.Recordset.AddNew
                         data_aut.Recordset("fecha") = Date
                         data_aut.Recordset("usuario") = Mid(txt_nomb.Text, 1, 50)
                         data_aut.Recordset("codaut") = "URGENCIA"
                         If txt_mat.Text <> "" Then
                            data_aut.Recordset("socio") = txt_mat.Text
                         Else
                            data_aut.Recordset("socio") = txt_mat.Text
                         End If
                         data_aut.Recordset("modulo") = "DESPACHO"
                         data_aut.Recordset("usuario_caja") = WElusuario
                         data_aut.Recordset.Update
                      Else
                          If txt_mat.Text <> "" Then
                             Xhab = txt_mat.Text
                          Else
                             Xhab = 0
                          End If
                          frm_autoriza.Show vbModal
                          '14063
                          '117670
                          '5112
                          Xelcodigoaut = InputBox("SOCIO CON CRÉDITOS PENDIENTES O CUOTAS, INGRESE CODIGO DE AUTORIZACIÓN SI ES CLAVE 3", "SOCIO CON CRÉDITOS PENDIENTES", Wopszond)
                          If Trim(Xelcodigoaut) <> "" Then
                             data_aut.RecordSource = "select * from Codigos_aut where codaut ='" & Trim(Xelcodigoaut) & "' and socio =" & txt_mat.Text
                             data_aut.Refresh
                             If data_aut.Recordset.RecordCount > 0 Then
                                Xq = 0
                                Xdeudasi = 0
                             Else
                                MsgBox "ATENCION! No se encuentra código de autorización, realice nuevamente la autorización o comunique a Administración", vbCritical
                                Xq = 9
                                Xdeudasi = 9
                             End If
                          Else
                             MsgBox "Socio con créditos o cuotas(>=3) pendientes, NO SE PODRÁ GRABAR LLAMADO CLAVE 3.", vbCritical
                             Xq = 9
                             Xdeudasi = 9
                          End If
                      End If
                   Else
                      Xdeudasi = 0
                   End If
                   If XAlta = 599 Then
                      XAlta = Xaltaanterior
                   End If
                   If IsNull(data_clib.Recordset("saldo_chc2")) = False Then
                      If data_clib.Recordset("saldo_chc2") = 1 Then
                         Xq = 11
                      End If
                      If Xq = 11 Then
                         MsgBox "ATENCION!! Socio con servicios RESTRINGIDOS! Estimado Funcionario NO dar servicio." & chr(13) _
                         & "El hacerlo estará bajo su exclusiva responsabilidad." & chr(13) & "El sistema no permitirá la continuidad de dicho servicio.", vbCritical, "SOCIOS"
                         MsgBox "SI ES UN LLAMADO CLAVE 3, DEBERA SOLICITAR AUTORIZACION al 097215419 PARA PODER GRABAR DATOS", vbInformation, "LLAMADO"
                         Xdeudasi = 9
                      End If
                   End If
                 End If
            End If
        End If
          
          
          Wopspro = 99
          data_parsec.DatabaseName = App.path & "\mensa.mdb"
          data_parsec.RecordSource = "mensaje"
          data_parsec.Refresh
          
          If Check1.Value <> 1 Then
             If data_clib.Recordset.RecordCount > 0 Then
                If IsNull(data_clib.Recordset("cl_grupo")) = False Then
                   Xcodzoning = data_clib.Recordset("cl_grupo")
                Else
                   Xcodzoning = 0
                End If
             Else
                Xcodzoning = 0
             End If
             If (Xcodzoning = 400 Or Xcodzoning = 401 Or Xcodzoning = 402 Or Xcodzoning = 403 Or Xcodzoning = 670 Or Xcodzoning = 671) And _
                (txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM") Then
                 Wopscob = 0
             Else
                If Val(cbozona.Text) = 1 Or Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Or Val(cbozona.Text) = 5 Or Val(cbozona.Text) = 6 Then
                   If txt_cat.Text <> "" Then
                      If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Or txt_cat.Text = "UNIVS" Or _
                         txt_cat.Text = "UNNSAM" Or txt_cat.Text = "HEVANO" Or txt_cat.Text = "EVNSAM" Or _
                         txt_cat.Text = "CCNOS" Or txt_cat.Text = "CCNSAM" Or txt_cat.Text = "GANOS" Or _
                         txt_cat.Text = "CASANO" Or txt_cat.Text = "CASNSA" Then
                         ConectarBD
                         ConbdSapp.Open
                         Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                         With Xrecconve
                             .CursorLocation = adUseClient
                             .CursorType = adOpenKeyset
                             .LockType = adLockOptimistic
                             .Open Xsqlstr, ConbdSapp, , , adCmdText
                         End With
                         If Xrecconve.RecordCount > 0 Then
                            Wopspro = 0
                            Wopscob = 0
                            ConbdSapp.Close
                            data_parsec.DatabaseName = App.path & "\mensa.mdb"
                            data_parsec.RecordSource = "mensaje"
                            data_parsec.Refresh
                            data_parsec.Recordset.Edit
                            data_parsec.Recordset("text") = "RECUERDE! Confirmar socio con la mutualista."
                            data_parsec.Recordset.Update
                            XwYalomostro = 99
                            frm_mensajesvar.Show vbModal
                         Else
                           ConbdSapp.Close
                           If IsNull(data_clib.Recordset("cl_decuota")) = False Then
                              If data_clib.Recordset("cl_decuota") = 0 Or _
                                 data_clib.Recordset("cl_decuota") = 1 Or _
                                 data_clib.Recordset("cl_decuota") = 3 Or _
                                 data_clib.Recordset("cl_decuota") = 4 Then
                                 data_parsec.DatabaseName = App.path & "\mensa.mdb"
                                 data_parsec.RecordSource = "mensaje"
                                 data_parsec.Refresh
                                 data_parsec.Recordset.Edit
                                 If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                    If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual. Documentación a presentar:" _
                                       & "Fotocopia de CI vigente. Comunique al funcionario del móvil " _
                                       & "para realizar la misma." _
                                       & " RECUERDE! Confirmar socio con la mutualista."
                                    Else
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual. Requerimientos:" _
                                       & "Fotocopia de CI vigente. Comprobante domicilio:" _
                                       & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                       & " a nombre del cliente que sea del mes corriente o anterior. Comunique al funcionario del móvil" _
                                       & " para realizar la misma. RECUERDE! Confirmar socio con la mutualista."
                                    End If
                                 Else
                                    data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                    & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                    chr(13) & "para realizar la misma." _
                                    & " RECUERDE! Confirmar socio con la mutualista."
                                 End If
                                 data_parsec.Recordset.Update
                                 data_parsec.Refresh
                                 XwYalomostro = 99
                                 frm_mensajesvar.Show vbModal
                              Else
                                 ConectarBD
                                 ConbdSapp.Open
                                 Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                                 With Xrecconve
                                     .CursorLocation = adUseClient
                                     .CursorType = adOpenKeyset
                                     .LockType = adLockOptimistic
                                     .Open Xsqlstr, ConbdSapp, , , adCmdText
                                 End With
                                 If Xrecconve.RecordCount > 0 Then
                                 Else
                                    data_parsec.Recordset.Edit
                                    If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                       If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                          data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                          & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                          chr(13) & "para realizar la misma." _
                                          & " RECUERDE! Confirmar socio con la mutualista."
                                       Else
                                          data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" _
                                          & "Fotocopia de CI vigente. Comprobante domicilio:" _
                                          & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                          & " a nombre del cliente que sea del mes corriente o anterior.Comunique al funcionario del móvil " _
                                          & "para realizar la misma." _
                                          & "RECUERDE! Confirmar socio con la mutualista."
                                       End If
                                    Else
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                       & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                       chr(13) & "para realizar la misma." _
                                       & " RECUERDE! Confirmar socio con la mutualista."
                                    End If
                                    data_parsec.Recordset.Update
                                    data_parsec.Refresh
                                    XwYalomostro = 99
                                    frm_mensajesvar.Show vbModal
                                 End If
                                 ConbdSapp.Close
                              End If
                           Else
                              ConectarBD
                              ConbdSapp.Open
                              Xsqlstr = "Select * from linmmdd where cod_cli =" & data_clib.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(Xfechacartas, "yyyy/mm/dd") & "'"
                              With Xrecconve
                                  .CursorLocation = adUseClient
                                  .CursorType = adOpenKeyset
                                  .LockType = adLockOptimistic
                                  .Open Xsqlstr, ConbdSapp, , , adCmdText
                              End With
                              If Xrecconve.RecordCount > 0 Then
                              Else
                                 data_parsec.Recordset.Edit
                                 If txt_cat.Text = "SMIN" Or txt_cat.Text = "SMINA" Then
                                    If Val(cbozona.Text) = 2 Or Val(cbozona.Text) = 3 Then
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                       & "Fotocopia de CI vigente. " & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                       chr(13) & "para realizar la misma." _
                                       & "RECUERDE! confirmar socio con la mutualista."
                                    Else
                                       data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual.Requerimientos:" _
                                       & "Fotocopia de CI vigente. Comprobante domicilio:" _
                                       & " Constancia policial(antiguedad 2meses), UTE, OSE, ANTEL, con 6 meses de consumo," _
                                       & " a nombre del cliente y del mes corriente o anterior.Comunique al funcionario del móvil" _
                                       & " para realizar la misma." _
                                       & "RECUERDE! Confirmar socio con la mutualista."
                                    End If
                                 Else
                                    data_parsec.Recordset("text") = "ATENCION!! Debe realizar carta mutual." & chr(13) & "Documentación a presentar:" & chr(13) _
                                    & "Fotocopia de Cédula de identidad vigente." & chr(13) & "Comunique al funcionario del móvil correspondiente" & _
                                    chr(13) & "para realizar la misma." _
                                    & " RECUERDE! Confirmar socio con la mutualista."
                                 End If
                                 data_parsec.Recordset.Update
                                 data_parsec.Refresh
                                 XwYalomostro = 99
                                 frm_mensajesvar.Show vbModal
                              End If
                              ConbdSapp.Close
                           End If
                         End If
                      End If
                   End If
                End If
             End If
          
          End If
          Wopspro = 0
       
       Else
          txt_nomb.SetFocus
       End If
    Else
    '   txt_nomb.SetFocus
       frm_buslla.Show vbModal
       
    End If

End Sub



Private Sub txt_codmed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   dbcbomed.SetFocus
End If

End Sub

Private Sub txt_codmed_LostFocus()
If txt_codmed.Text <> "" Then
   If txt_codmed.Text <> 0 Then
      data_med.Recordset.FindFirst "med_cod =" & txt_codmed.Text
      If Not data_med.Recordset.NoMatch Then
         If IsNull(data_med.Recordset("med_nombre")) = False Then
            txt_codmed.Text = data_med.Recordset("med_cod")
            dbcbomed.Text = data_med.Recordset("med_nombre")
         Else
            dbcbomed.Text = ""
         End If
      End If
   Else
      MsgBox "MEDICO NO ENCONTRADO", vbInformation, "Mensaje"
      dbcbomed.SetFocus
      txt_codmed.Text = 0
   End If
End If

End Sub

Private Sub txt_codmed2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   dbcbomed2.SetFocus
End If

End Sub

Private Sub txt_codmed2_LostFocus()
If txt_codmed2.Text <> "" Then
   If txt_codmed2.Text <> 0 Then
      data_med2.Recordset.FindFirst "med_cod =" & txt_codmed2.Text
      If Not data_med2.Recordset.NoMatch Then
         If IsNull(data_med2.Recordset("med_nombre")) = False Then
            txt_codmed2.Text = data_med2.Recordset("med_cod")
            dbcbomed2.Text = data_med2.Recordset("med_nombre")
         Else
            dbcbomed2.Text = ""
            txt_codmed2.Text = 0
         End If
      End If
   Else
      MsgBox "MEDICO NO ENCONTRADO", vbInformation, "Mensaje"
'      dbcbomed.SetFocus
      txt_codmed.Text = 0
   End If
End If

End Sub

Private Sub txt_costo_LostFocus()
If txt_costo.Text <> "" Then
   If Val(txt_costo.Text) > 29000 Then
      MsgBox "Verifique importe si es correcto!", vbCritical
   End If
End If

End Sub

Private Sub txt_diag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocolfin.SetFocus
End If

End Sub

Private Sub txt_direc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbozona.SetFocus
End If

End Sub

Private Sub txt_edad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboed.SetFocus
End If

End Sub

Private Sub txt_edad_LostFocus()
If txt_edad.Text = "" Then
'   MsgBox "Ingrese dato de años", vbInformation, "Mensaje"
   txt_edad.Text = 0
'   txt_edad.SetFocus
End If

End Sub

Private Sub txt_enca_Click()
txt_enca.Text = Format(Time, "HH:mm")

End Sub

Private Sub txt_enca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_enzona.SetFocus
End If

End Sub

Private Sub txt_enzona_Click()
txt_enzona.Text = Format(Time, "HH:mm")

End Sub

Private Sub txt_enzona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_movtra.SetFocus
End If

End Sub

Private Sub txt_hora_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   txt_usua.SetFocus
'End If

End Sub

Private Sub txt_horasig_GotFocus()
txt_horasig.Text = Format(Time, "HH:mm")

End Sub

Private Sub txt_horasig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   msalida.SetFocus
End If

End Sub

Private Sub txt_horasig_LostFocus()
If txt_horasig.Text <> "" Then
   If txt_horsal.Text <> "" Then
      If mfecasig.Text = msalida.Text Then
         If txt_horasig.Text > txt_horsal.Text Then
            MsgBox "Atención!! Tiempo de asignación es MAYOR al tiempo de SALIDA", vbCritical, "Despacho"
            txt_horasig.SetFocus
         End If
      
      End If
   End If
End If

End Sub

Private Sub txt_horlle_GotFocus()
txt_horlle.Text = Format(Time, "HH:mm")
End Sub

Private Sub txt_horlle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mtd.SetFocus
End If

End Sub

Private Sub txt_horlle_LostFocus()
If txt_horlle.Text <> "" Then
   If txt_horsal.Text <> "" Then
      If msalida.Text = mllegada.Text Then
         If txt_horsal.Text > txt_horlle.Text Then
            MsgBox "Atención!! Tiempo de LLEGADA es MENOR que tiempo de SALIDA", vbCritical, "Despacho"
            txt_horlle.SetFocus
         End If
      End If
   End If
End If

End Sub

Private Sub txt_horsal_GotFocus()
If txt_horasig.Text = "" Then
   MsgBox "No ingresó HORA de MOVIL"
Else
   txt_horsal.Text = Format(Time, "HH:mm")
End If

End Sub

Private Sub txt_horsal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mllegada.SetFocus
End If

End Sub

Private Sub txt_horsal_LostFocus()
If txt_horasig.Text <> "" Then
   If txt_horsal.Text <> "" Then
      If mfecasig.Text = msalida.Text Then
         If txt_horasig.Text > txt_horsal.Text Then
            MsgBox "Atención!! Tiempo de SALIDA es MENOR que tiempo de ASIGNADO", vbCritical, "Despacho"
            txt_horsal.SetFocus
         End If
      End If
   End If
End If

End Sub

Private Sub txt_hortd_GotFocus()
txt_hortd.Text = Format(Time, "HH:mm")

End Sub

Private Sub txt_hortd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_diag.SetFocus
End If

End Sub

Private Sub txt_hortd_LostFocus()
If txt_horlle.Text <> "" Then
   If txt_hortd.Text <> "" Then
      If mllegada.Text = mtd.Text Then
         If txt_horlle.Text > txt_hortd.Text Then
            MsgBox "Atención!! Tiempo de T/D es MENOR que tiempo de LLEGADA", vbCritical, "Despacho"
            txt_hortd.SetFocus
         End If
      End If
   End If
End If

End Sub

Private Sub txt_locali_LostFocus()
Dim Xregbusca As New ADODB.Recordset
Dim XsqlCons As String
Dim Confirmazona As String

On Error GoTo Allocal

If txt_locali.Text <> "" Then
   ConectarBD
   ConbdSapp.Open
    
   XsqlCons = "Select * from zonas where zo_nombre ='" & txt_locali.Text & "'"
    
   With Xregbusca
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open XsqlCons, ConbdSapp, , , adCmdText
    End With
    If Xregbusca.RecordCount > 0 Then
    Else
       Confirmazona = MsgBox("La zona ingresada no existe, confirma igual la zona ingresada?", vbYesNo + vbInformation)
       If Confirmazona = vbYes Then
       Else
          txt_locali.Text = ""
       End If
'       txt_locali.SetFocus
    End If
    ConbdSapp.Close
End If

Exit Sub

Allocal:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ALLOCAL ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ALLOCAL ERR:" & Err.Number
      End If

End Sub

Private Sub txt_lugar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_trassal.SetFocus
End If

End Sub

Private Sub txt_lugar_LostFocus()
Dim Xregbusca As New ADODB.Recordset
Dim XsqlCons As String
On Error GoTo Allugar

If txt_lugar.Text <> "" Then
   ConectarBD
   ConbdSapp.Open
    
   XsqlCons = "Select * from sociedad where soc_nombre ='" & txt_lugar.Text & "'"
    
   With Xregbusca
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open XsqlCons, ConbdSapp, , , adCmdText
    End With
    If Xregbusca.RecordCount > 0 Then
    Else
       MsgBox "No se encuentra dato de lugar ingresado. Verifique si es correcto.", vbCritical
    End If
    ConbdSapp.Close
End If

Exit Sub

Allugar:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ALLUGAR ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ALLUGAR ERR:" & Err.Number
      End If

End Sub

Private Sub txt_mat_Change()
If Not IsNumeric(txt_mat.Text) And _
 txt_mat.Text <> "" Then
 Beep
 MsgBox "Se debe ingresar solo números en matrícula o vacío"
  txt_mat.Text = ""
txt_mat.SetFocus
End If

End Sub

Private Sub txt_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_tel.SetFocus
End If

End Sub

Private Sub txt_mat_LostFocus()
'If XAlta <> 0 Then

End Sub

Private Sub txt_mot_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   cbocolor.SetFocus
   txt_ante.SetFocus
End If

End Sub

Private Sub txt_movil_GotFocus()
txt_movil.SelStart = 0
txt_movil.SelLength = Len(txt_movil.Text)

End Sub

Private Sub txt_movil_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfecasig.SetFocus
End If

End Sub

Private Sub txt_movil_LostFocus()
If txt_movil.Text <> "" Then
   If txt_movil.Text > 0 Then
      data_mov.Recordset.FindFirst "movil =" & txt_movil.Text
      If Not data_mov.Recordset.NoMatch Then
         txt_codmed.Text = data_mov.Recordset("codmed")
         dbcbomed.Text = data_mov.Recordset("nommed")
         If IsNull(data_mov.Recordset("codchof")) = False Then
            labcodchof.Caption = data_mov.Recordset("codchof")
         Else
            labcodchof.Caption = 0
         End If
         If IsNull(data_mov.Recordset("nomchof")) = False Then
            labnomchof.Caption = "Chof.:" & data_mov.Recordset("nomchof")
         Else
            labnomchof.Caption = ""
         End If
         If IsNull(data_mov.Recordset("ano")) = False Then
            txt_queb.Text = data_mov.Recordset("ano")
         Else
            txt_queb.Text = 0
         End If
      Else
         txt_codmed.Text = ""
         dbcbomed.Text = ""
         txt_queb.Text = 0
         labcodchof.Caption = 0
      End If
      If txt_movil.Text = 99 Or txt_movil.Text = 98 Then
      Else
         If cbotimbre.ListIndex = 1 Then
            MsgBox "Llamado con costo de timbre, comunique al MOVIL!", vbCritical
         End If
      End If
   Else
      txt_codmed.Text = ""
      dbcbomed.Text = ""
      txt_queb.Text = 0
      labcodchof.Caption = 0
   End If
Else
   txt_codmed.Text = ""
   dbcbomed.Text = ""
   txt_queb.Text = 0
   labcodchof.Caption = 0
End If

End Sub

Private Sub txt_movtra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_grabar.SetFocus
End If

End Sub

Private Sub txt_nomb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_edad.SetFocus
End If

End Sub

Private Sub txt_nomcat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mat.SetFocus
End If

End Sub

Private Sub txt_nro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfecha.SetFocus
End If

End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
On Error GoTo Alpasar

If KeyAscii = 13 Then
   If WDespa = 1 Then
      b_grabar.SetFocus
   Else
      txt_movil.SetFocus
   End If
End If

Exit Sub

Alpasar:
        If Err.Number = 5 Then
           MsgBox "Control no habilitado"
        Else
           MsgBox "Error:" & Err.Description
        End If
        
End Sub

Private Sub txt_otros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_direc.SetFocus
End If

End Sub


Private Sub txt_salca_Click()
txt_salca.Text = Format(Time, "HH:mm")

End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_direc.SetFocus
End If

End Sub

Private Sub txt_trassal_Click()
txt_trassal.Text = Format(Time, "HH:mm")

End Sub

Private Sub txt_trassal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_enca.SetFocus
End If

End Sub

Public Function igualar_lla()
'data_lla.Refresh
'data_lla.Recordset.MoveLast
On Error GoTo ALIgual

txt_nro.Text = data_lla.Recordset("nrolla")
mfecha.Text = Format(data_lla.Recordset("fecha"), "dd/mm/yyyy")
txt_hora.Text = Format(data_lla.Recordset("hora"), "HH:mm")
txt_usua.Text = data_lla.Recordset("usuario")
If IsNull(data_lla.Recordset("nomodif")) = False Then
   Check1.Value = data_lla.Recordset("nomodif")
Else
   Check1.Value = 0
End If
If IsNull(data_lla.Recordset("timbre")) = False Then
   cbotimbre.ListIndex = data_lla.Recordset("timbre")
Else
   cbotimbre.ListIndex = -1
End If
If IsNull(data_lla.Recordset("valor_timbre")) = False Then
   t_timbre.Text = data_lla.Recordset("valor_timbre")
Else
   t_timbre.Text = ""
End If
If IsNull(data_lla.Recordset("segui_covid")) = False Then
   chcovid.Value = data_lla.Recordset("segui_covid")
Else
   chcovid.Value = 0
End If
If IsNull(data_lla.Recordset("pend")) = False Then
    If data_lla.Recordset("pend") = 4 Then
       Command5.Visible = True
       Frame2.Visible = False
    Else
       Command5.Visible = False
       Frame2.Visible = True
    End If
Else
    Command5.Visible = False
    Frame2.Visible = True
End If
If IsNull(data_lla.Recordset("matric")) = False Then
   txt_mat.Text = data_lla.Recordset("matric")
Else
   txt_mat.Text = 0
End If
If IsNull(data_lla.Recordset("nombre")) = False Then
   txt_nomb.Text = data_lla.Recordset("nombre")
Else
   txt_nomb.Text = ""
End If
If IsNull(data_lla.Recordset("edad")) = False Then
   txt_edad.Text = data_lla.Recordset("edad")
Else
   txt_edad.Text = ""
End If
If IsNull(data_lla.Recordset("unied")) = False Then
   If data_lla.Recordset("unied") = 3 Then
      cboed.ListIndex = 0
   Else
      If data_lla.Recordset("unied") = 2 Then
         cboed.ListIndex = 1
      Else
         If data_lla.Recordset("unied") = 1 Then
            cboed.ListIndex = 2
         Else
            cboed.ListIndex = 0
         End If
      End If
   End If
Else
   cboed.ListIndex = 0
End If
If IsNull(data_lla.Recordset("categ")) = False Then
   txt_cat.Text = data_lla.Recordset("categ")
Else
   txt_cat.Text = ""
End If
If IsNull(data_lla.Recordset("nomcat")) = False Then
   txt_nomcat.Text = data_lla.Recordset("nomcat")
Else
   txt_nomcat.Text = ""
End If
If IsNull(data_lla.Recordset("hora_anterior")) = False Then
   labanthor.Caption = data_lla.Recordset("hora_anterior")
Else
   labanthor.Caption = ""
End If
If IsNull(data_lla.Recordset("ci")) = False Then
   txt_ced.Text = Int(data_lla.Recordset("ci"))
Else
   txt_ced.Text = 0
End If
If IsNull(data_lla.Recordset("telef")) = False Then
   txt_tel.Text = data_lla.Recordset("telef")
Else
   txt_tel.Text = ""
End If
If IsNull(data_lla.Recordset("codzon")) = False Then
   If data_lla.Recordset("codzon") = 2 Then
      cbozona.ListIndex = 1
   Else
      If data_lla.Recordset("codzon") = 3 Then
         cbozona.ListIndex = 2
      Else
         If data_lla.Recordset("codzon") = 4 Then
            cbozona.ListIndex = 3
         Else
            If data_lla.Recordset("codzon") = 5 Then
               cbozona.ListIndex = 4
            Else
               If data_lla.Recordset("codzon") = 6 Then
                  cbozona.ListIndex = 5
               Else
                  If data_lla.Recordset("codzon") = 7 Then
                     cbozona.ListIndex = 6
                  Else
                     cbozona.ListIndex = 0
                  End If
               End If
            End If
         End If
      End If
   End If
Else
   cbozona.ListIndex = 0
End If
If IsNull(data_lla.Recordset("base")) = False Then
   cbobase.Text = data_lla.Recordset("base")
Else
   cbobase.Text = 0
End If
If IsNull(data_lla.Recordset("referen")) = False Then
   txt_direc.Text = data_lla.Recordset("referen")
Else
   txt_direc.Text = ""
End If
If IsNull(data_lla.Recordset("obs")) = False Then
   txt_obs.Text = data_lla.Recordset("obs")
Else
   txt_obs.Text = ""
End If
If IsNull(data_lla.Recordset("motcon")) = False Then
   txt_ante.Text = data_lla.Recordset("motcon")
Else
   txt_ante.Text = ""
End If
If IsNull(data_lla.Recordset("obsmot")) = False Then
   txt_mot.Text = data_lla.Recordset("obsmot")
Else
   txt_mot.Text = ""
End If
If IsNull(data_lla.Recordset("codmot")) = False Then
   If data_lla.Recordset("codmot") = "R" Then
      cbocolor.ListIndex = 2
   Else
      If data_lla.Recordset("codmot") = "A" Then
         cbocolor.ListIndex = 1
      Else
         If data_lla.Recordset("codmot") = "C" Then
            cbocolor.ListIndex = 3
         Else
            If data_lla.Recordset("codmot") = "Z" Then
               cbocolor.ListIndex = 4
            Else
               If data_lla.Recordset("codmot") = "N" Then
                  cbocolor.ListIndex = 5
               Else
                  cbocolor.ListIndex = 0
               End If
            End If
         End If
      End If
   End If
Else
   cbocolor.ListIndex = 0
End If
If cbocolor.Text = "VERDE" Then
   cbocolor.BackColor = &HC000&
Else
   If cbocolor.Text = "ROJO" Then
      cbocolor.BackColor = &HFF&
   Else
      If cbocolor.Text = "AMARILLO" Then
         cbocolor.BackColor = &HFFFF&
      Else
         If cbocolor.Text = "CELESTE" Then
            cbocolor.BackColor = &HFFFF00
         Else
            If cbocolor.Text = "AZUL" Then
               cbocolor.BackColor = &HC00000
            Else
               If cbocolor.Text = "NEGRO" Then
                  cbocolor.BackColor = &H80000006
               Else
                  cbocolor.BackColor = &HFFFFFF
               End If
            End If
         End If
      End If
   End If
End If
If IsNull(data_lla.Recordset("movilpas")) = False Then
   txt_movil.Text = data_lla.Recordset("movilpas")
Else
   txt_movil.Text = ""
End If
If IsNull(data_lla.Recordset("fecpas")) = False Then
   mfecasig.Text = Format(data_lla.Recordset("fecpas"), "dd/mm/yyyy")
Else
   mfecasig.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("horpas")) = False Then
   txt_horasig.Text = Format(data_lla.Recordset("horpas"), "HH:mm")
Else
   txt_horasig.Text = ""
End If
If IsNull(data_lla.Recordset("fecsali")) = False Then
   msalida.Text = Format(data_lla.Recordset("fecsali"), "dd/mm/yyyy")
Else
   msalida.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("horsali")) = False Then
   txt_horsal.Text = Format(data_lla.Recordset("horsali"), "HH:mm")
Else
   txt_horsal.Text = ""
End If
If IsNull(data_lla.Recordset("fec_llega")) = False Then
   mllegada.Text = Format(data_lla.Recordset("fec_llega"), "dd/mm/yyyy")
Else
   mllegada.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("hor_llega")) = False Then
   txt_horlle.Text = Format(data_lla.Recordset("hor_llega"), "HH:mm")
Else
   txt_horlle.Text = ""
End If
If IsNull(data_lla.Recordset("fec_rea")) = False Then
   mtd.Text = Format(data_lla.Recordset("fec_rea"), "dd/mm/yyyy")
Else
   mtd.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("hor_rea")) = False Then
   If data_lla.Recordset("hor_rea") <> "" Then
      txt_hortd.Text = Format(data_lla.Recordset("hor_rea"), "HH:mm")
   Else
      txt_hortd.Text = "__:__"
   End If
Else
   txt_hortd.Text = "__:__"
End If
If IsNull(data_lla.Recordset("diag")) = False Then
   txt_diag.Text = data_lla.Recordset("diag")
Else
   txt_diag.Text = ""
End If
If IsNull(data_lla.Recordset("colormot")) = False Then
   If data_lla.Recordset("colormot") = "R" Then
      cbocolfin.ListIndex = 2
   Else
      If data_lla.Recordset("colormot") = "A" Then
         cbocolfin.ListIndex = 1
      Else
         If data_lla.Recordset("colormot") = "V" Then
            cbocolfin.ListIndex = 0
         Else
            If data_lla.Recordset("colormot") = "N" Then
               cbocolfin.ListIndex = 3
            Else
               cbocolfin.Text = ""
            End If
         End If
      End If
   End If
Else
   cbocolfin.Text = ""
End If

dbcbomed.ListField = ""
dbcbomed.BoundColumn = ""
If IsNull(data_lla.Recordset("nommed")) = False Then
   dbcbomed.Text = data_lla.Recordset("nommed")
Else
   dbcbomed.Text = ""
End If
dbcbomed.ListField = "med_nombre"
dbcbomed.BoundColumn = "med_nombre"

dbcbomed2.ListField = ""
dbcbomed2.BoundColumn = ""
dbcbomed2.Text = ""
dbcbomed2.ListField = "med_nombre"
dbcbomed2.BoundColumn = "med_nombre"

If IsNull(data_lla.Recordset("codmed")) = False Then
   txt_codmed.Text = data_lla.Recordset("codmed")
Else
   txt_codmed.Text = 0
End If
If IsNull(data_lla.Recordset("trasla")) = False Then
   If data_lla.Recordset("trasla") > 0 Then
      If data_lla.Recordset("trasla") > 10 Then
         If data_lla.Recordset("trasla") = 11 Then
            cbotras.ListIndex = 8
         Else
            cbotras.ListIndex = 9
         End If
      Else
         cbotras.ListIndex = data_lla.Recordset("trasla")
      End If
   Else
      cbotras.ListIndex = 0
   End If
Else
   cbotras.ListIndex = 0
End If
If IsNull(data_lla.Recordset("lugar")) = False Then
   txt_lugar.Text = data_lla.Recordset("lugar")
Else
   txt_lugar.Text = ""
End If
If IsNull(data_lla.Recordset("hsald")) = False Then
   txt_trassal.Text = Format(data_lla.Recordset("hsald"), "HH:mm")
Else
   txt_trassal.Text = ""
End If
If IsNull(data_lla.Recordset("hllega")) = False Then
   txt_enca.Text = Format(data_lla.Recordset("hllega"), "HH:mm")
Else
   txt_enca.Text = ""
End If
If IsNull(data_lla.Recordset("hzona")) = False Then
   txt_enzona.Text = Format(data_lla.Recordset("hzona"), "HH:mm")
Else
   txt_enzona.Text = ""
End If
If IsNull(data_lla.Recordset("movtras")) = False Then
   txt_movtra.Text = data_lla.Recordset("movtras")
Else
   txt_movtra.Text = ""
End If
If IsNull(data_lla.Recordset("dcobr")) = False Then
   Combo1.Text = data_lla.Recordset("dcobr")
Else
   Combo1.Text = ""
End If
If IsNull(data_lla.Recordset("activo")) = False Then
   Label3.Caption = Format(data_lla.Recordset("activo"), "HH:mm:ss")
Else
   Label3.Caption = "00:00:00"
End If
If IsNull(data_lla.Recordset("timdes")) = False Then
   Label39.Caption = data_lla.Recordset("timdes")
Else
   Label39.Caption = "Sin Largar"
End If
If IsNull(data_lla.Recordset("totdem")) = False Then
   txt_demora.Text = Format(data_lla.Recordset("totdem"), "HH:mm")
Else
   txt_demora.Text = ""
End If
If IsNull(data_lla.Recordset("motmov")) = True Then
   txt_locali.Text = ""
Else
   txt_locali.Text = data_lla.Recordset("motmov")
End If
If IsNull(data_lla.Recordset("mm")) = True Then
   Label41.Caption = 0
Else
   Label41.Caption = data_lla.Recordset("mm")
End If
If IsNull(data_lla.Recordset("thh")) = True Then
   Label42.Caption = 0
Else
   Label42.Caption = data_lla.Recordset("thh")
End If
If IsNull(data_lla.Recordset("tmm")) = True Then
   Label43.Caption = 0
Else
   Label43.Caption = data_lla.Recordset("tmm")
End If
If IsNull(data_lla.Recordset("pasado")) = True Then
   Label44.Caption = 0
Else
   Label44.Caption = data_lla.Recordset("pasado")
End If
If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
    If IsNull(data_lla.Recordset("ano")) = True Then
       Label45.Caption = 0
    Else
       Label45.Caption = data_lla.Recordset("ano")
    End If
    If IsNull(data_lla.Recordset("mes")) = True Then
       Label46.Caption = 0
    Else
       Label46.Caption = data_lla.Recordset("mes")
    End If
Else
    Label45.Caption = 0
    Label46.Caption = 0
End If
If IsNull(data_lla.Recordset("mes")) = False Then
   If data_lla.Recordset("mes") > 10 Then
      txt_costo.Text = data_lla.Recordset("mes")
   Else
      txt_costo.Text = 0
   End If
Else
   txt_costo.Text = 0
End If
If IsNull(data_lla.Recordset("realiza")) = False Then
   chtmut.Value = data_lla.Recordset("realiza")
Else
   chtmut.Value = 0
End If
If IsNull(data_lla.Recordset("ano")) = False Then
   If data_lla.Recordset("ano") > 10 Then
      txt_boleta.Text = data_lla.Recordset("ano")
   Else
      txt_boleta.Text = 0
   End If
Else
   txt_boleta.Text = 0
End If

If IsNull(data_lla.Recordset("timsi")) = True Then
   Label48.Caption = 0
Else
   Label48.Caption = data_lla.Recordset("timsi")
End If
If IsNull(data_lla.Recordset("aft")) = False Then
   Label40.Caption = "AFT:" & data_lla.Recordset("aft")
Else
   Label40.Caption = ""
End If

If IsNull(data_lla.Recordset("enfer")) = True Then
   Check2.Value = 0
Else
   Check2.Value = data_lla.Recordset("enfer")
End If
If IsNull(data_lla.Recordset("motcance")) = True Then
   txt_quien.Text = ""
Else
   txt_quien.Text = data_lla.Recordset("motcance")
End If
If IsNull(data_lla.Recordset("cancela")) = True Then
   If IsNull(data_lla.Recordset("hor_cance")) = False Then
      txt_salca.Text = data_lla.Recordset("hor_cance")
   Else
      txt_salca.Text = ""
   End If
End If
If IsNull(data_lla.Recordset("hh")) = True Then ' Sexo
   Combo3.ListIndex = -1
Else
   Combo3.ListIndex = data_lla.Recordset("hh")
End If
data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
data_llamod.Refresh
If data_llamod.Recordset.RecordCount > 0 Then
   If IsNull(data_llamod.Recordset("movil_rea")) = False Then
      If data_llamod.Recordset("movil_rea") > 0 Then
         data_chof.RecordSource = "Select * from movil where nromov =" & data_llamod.Recordset("movil_rea")
         data_chof.Refresh
         If data_chof.Recordset.RecordCount > 0 Then
            labcodchof.Caption = data_llamod.Recordset("movil_rea")
            labnomchof.Caption = "Chof.:" & data_chof.Recordset("chofer")
         Else
            labcodchof.Caption = 0
            labnomchof.Caption = ""
         End If
      Else
         labcodchof.Caption = 0
         labnomchof.Caption = ""
      End If
   Else
      labcodchof.Caption = 0
      labnomchof.Caption = ""
   End If
   
   If IsNull(data_llamod.Recordset("mes")) = False Then
      t_codced.Text = Int(data_llamod.Recordset("mes"))
   Else
      t_codced.Text = 0
   End If
   If IsNull(data_llamod.Recordset("pasado")) = False Then
      Check4.Value = data_llamod.Recordset("pasado")
   Else
      Check4.Value = 0
   End If
   If IsNull(data_llamod.Recordset("telef")) = False Then
      If data_llamod.Recordset("telef") = "RECIBO" Then
         Combo2.ListIndex = 0
      Else
         If data_llamod.Recordset("telef") = "CONFORME" Then
            Combo2.ListIndex = 1
         Else
            Combo2.ListIndex = -1
         End If
      End If
   Else
      Combo2.ListIndex = -1
   End If
   If IsNull(data_llamod.Recordset("movilpas")) = False Then
      data_med2.Recordset.FindFirst "med_cod =" & data_llamod.Recordset("movilpas")
'      data_med2.RecordSource = "Select * from medicos where med_cod =" & data_llamod.Recordset("movilpas")
'      data_med2.Refresh
      If Not data_med2.Recordset.NoMatch Then
         dbcbomed2.Text = data_med2.Recordset("med_nombre")
      Else
         dbcbomed2.Text = ""
      End If
      txt_codmed2.Text = data_llamod.Recordset("movilpas")
   Else
      txt_codmed2.Text = 0
   End If
   If IsNull(data_llamod.Recordset("fec_llega")) = False Then
      mftrassol.Text = Format(data_llamod.Recordset("fec_llega"), "dd/mm/yyyy")
   Else
      mftrassol.Text = "__/__/____"
   End If
   If IsNull(data_llamod.Recordset("hor_llega")) = False Then
      mhtrassol.Text = Format(data_llamod.Recordset("hor_llega"), "HH:mm")
   Else
      mhtrassol.Text = "__:__"
   End If
   If IsNull(data_llamod.Recordset("hzona")) = False Then
      labcmt.Visible = True
      labcmt.Caption = "PASADO A CMT HORA:" & Format(data_llamod.Recordset("hzona"), "HH:mm")
      If IsNull(data_llamod.Recordset("mm")) = False Then
         If data_llamod.Recordset("mm") = 1 Then
            labcmt.Caption = labcmt.Caption & " NO RESUELTO H."
            If IsNull(data_llamod.Recordset("hsald")) = False Then
               labcmt.Caption = labcmt.Caption & data_llamod.Recordset("hsald")
            End If
            If IsNull(data_llamod.Recordset("totend")) = False Then
               If data_llamod.Recordset("totend") = "R" Then
                  labcmt.Caption = labcmt.Caption & " RECLASIFICA A ROJO"
               End If
               If data_llamod.Recordset("totend") = "A" Then
                  labcmt.Caption = labcmt.Caption & " RECLASIFICA A AMARILLO"
               End If
            End If
         End If
         If data_llamod.Recordset("mm") = 2 Then
            labcmt.Caption = labcmt.Caption & " RESUELTO HORA:"
            If IsNull(data_llamod.Recordset("hor_rea")) = False Then
               labcmt.Caption = labcmt.Caption & data_llamod.Recordset("hor_rea")
            End If
         End If
         If data_llamod.Recordset("mm") = 2 Or data_llamod.Recordset("mm") = 3 Then
            Command5.Visible = True
            Frame2.Visible = False
         Else
            If data_llamod.Recordset("mm") = 1 Then
               Command5.Visible = False
               If WDespa = 1 Then
                  Frame2.Visible = False
               Else
                  Frame2.Visible = True
               End If
            Else
               Command5.Visible = True
               Frame2.Visible = False
            End If
         End If
      Else
         Command5.Visible = True
         Frame2.Visible = False
      End If
   Else
      labcmt.Caption = ""
      labcmt.Visible = False
      Command5.Visible = False
      If WDespa = 1 Then
         Frame2.Visible = False
      Else
         Frame2.Visible = True
      End If
   End If
Else
   txt_codmed2.Text = 0
   Check4.Value = 0
   dbcbomed2.Text = ""
   mftrassol.Text = "__/__/____"
   mhtrassol.Text = "__:__"
   t_codced.Text = 0
   labcmt.Caption = ""
   labcmt.Visible = False
   Command5.Visible = False
   Combo2.ListIndex = -1
End If

Exit Function

ALIgual:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ALIGUAL ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ALIGUAL ERR:" & Err.Number
      End If

End Function

Public Function borra_ya()

txt_nro.Text = ""
mfecha.Text = "__/__/____"
Check1.Value = 0
txt_hora.Text = ""
txt_usua.Text = ""
txt_mat.Text = ""
txt_nomb.Text = ""
txt_edad.Text = ""
cboed.Text = ""
txt_cat.Text = ""
txt_nomcat.Text = ""
txt_ced.Text = ""
txt_direc.Text = ""
txt_tel.Text = ""
cbozona.Text = ""
cbobase.Text = ""
t_codced.Text = ""
txt_obs.Text = ""
txt_mot.Text = ""
txt_ante.Text = ""
cbocolor.Text = ""
txt_movil.Text = ""
mfecasig.Text = "__/__/____"
txt_horasig.Text = ""
msalida.Text = "__/__/____"
txt_horsal.Text = ""
mllegada.Text = "__/__/____"
txt_horlle.Text = ""
mtd.Text = "__/__/____"
txt_hortd.Text = "__:__"
txt_demora.Text = ""
txt_diag.Text = ""
cbocolfin.Text = ""
labanthor.Caption = ""
dbcbomed.ListField = ""
dbcbomed.BoundColumn = ""
dbcbomed.Text = ""
dbcbomed.ListField = "med_nombre"
dbcbomed.BoundColumn = "med_nombre"
dbcbomed2.ListField = ""
dbcbomed2.BoundColumn = ""
dbcbomed2.Text = ""
dbcbomed2.ListField = "med_nombre"
dbcbomed2.BoundColumn = "med_nombre"

txt_codmed.Text = ""
cbotras.Text = ""
txt_lugar.Text = ""
txt_trassal.Text = ""
txt_enca.Text = ""
txt_enzona.Text = ""
txt_movtra.Text = ""
Combo1.Text = ""
Label3.Caption = ""
Label39.Caption = ""
txt_locali.Text = ""
labnomchof.Caption = ""
Check2.Value = 0
Label41.Caption = 0
Label42.Caption = 0
Label43.Caption = 0
Label44.Caption = 0
Label45.Caption = 0
Label45.Caption = -1
Label48.Caption = 0
txt_salca.Text = ""
'Check3.Value = 0
Combo3.ListIndex = -1
dbcbomed2.ListField = ""
dbcbomed2.BoundColumn = ""
dbcbomed2.Text = ""
dbcbomed2.ListField = "med_nombre"
dbcbomed2.BoundColumn = "med_nombre"
Check2.Value = 0
Check4.Value = 0
txt_codmed2.Text = ""
mftrassol.Text = "__/__/____"
mhtrassol.Text = "__:__"
Label26.Caption = "Traslado:"
labcmt.Caption = ""
chtmut.Value = 0
txt_boleta.Text = 0
txt_costo.Text = 0
Combo2.ListIndex = -1
Label40.Caption = ""
chcovid.Value = 0
cbotimbre.ListIndex = -1
t_timbre.Text = ""

End Function

Public Function igualar_sin()
On Error GoTo ALIgualsin

txt_nro.Text = data_lla.Recordset("nrolla")
mfecha.Text = data_lla.Recordset("fecha")
txt_hora.Text = Format(data_lla.Recordset("hora"), "HH:mm")
txt_usua.Text = data_lla.Recordset("usuario")
If IsNull(data_lla.Recordset("matric")) = False Then
   txt_mat.Text = data_lla.Recordset("matric")
Else
   txt_mat.Text = 0
End If
If IsNull(data_lla.Recordset("timbre")) = False Then
   cbotimbre.ListIndex = data_lla.Recordset("timbre")
Else
   cbotimbre.ListIndex = -1
End If
If IsNull(data_lla.Recordset("valor_timbre")) = False Then
   t_timbre.Text = data_lla.Recordset("valor_timbre")
Else
   t_timbre.Text = ""
End If

If IsNull(data_lla.Recordset("segui_covid")) = False Then
   chcovid.Value = data_lla.Recordset("segui_covid")
Else
   chcovid.Value = 0
End If
If IsNull(data_lla.Recordset("nombre")) = False Then
   txt_nomb.Text = data_lla.Recordset("nombre")
Else
   txt_nomb.Text = ""
End If
If IsNull(data_lla.Recordset("edad")) = False Then
   txt_edad.Text = data_lla.Recordset("edad")
Else
   txt_edad.Text = ""
End If
If IsNull(data_lla.Recordset("nomodif")) = False Then
   Check1.Value = data_lla.Recordset("nomodif")
Else
   Check1.Value = 0
End If
If IsNull(data_lla.Recordset("unied")) = False Then
   If data_lla.Recordset("unied") = 3 Then
      cboed.ListIndex = 0
   Else
      If data_lla.Recordset("unied") = 2 Then
         cboed.ListIndex = 1
      Else
         If data_lla.Recordset("unied") = 1 Then
            cboed.ListIndex = 2
         Else
            cboed.ListIndex = 0
         End If
      End If
   End If
Else
   cboed.ListIndex = 0
End If
If IsNull(data_lla.Recordset("pend")) = False Then
   If data_lla.Recordset("pend") = 4 Then
      Command5.Visible = True
      Frame2.Visible = False
   Else
      Command5.Visible = False
      Frame2.Visible = True
   End If
Else
   Command5.Visible = False
   Frame2.Visible = True
End If
If IsNull(data_lla.Recordset("categ")) = False Then
   txt_cat.Text = data_lla.Recordset("categ")
Else
   txt_cat.Text = ""
End If
If IsNull(data_lla.Recordset("nomcat")) = False Then
   txt_nomcat.Text = data_lla.Recordset("nomcat")
Else
   txt_nomcat.Text = ""
End If
If IsNull(data_lla.Recordset("hora_anterior")) = False Then
   labanthor.Caption = data_lla.Recordset("hora_anterior")
Else
   labanthor.Caption = ""
End If
If IsNull(data_lla.Recordset("aft")) = False Then
   Label40.Caption = "AFT:" & data_lla.Recordset("aft")
Else
   Label40.Caption = ""
End If

If IsNull(data_lla.Recordset("ci")) = False Then
   txt_ced.Text = Int(data_lla.Recordset("ci"))
Else
   txt_ced.Text = 0
End If
If IsNull(data_lla.Recordset("telef")) = False Then
   txt_tel.Text = data_lla.Recordset("telef")
Else
   txt_tel.Text = ""
End If
If IsNull(data_lla.Recordset("codzon")) = False Then
   If data_lla.Recordset("codzon") = 2 Then
      cbozona.ListIndex = 1
   Else
      If data_lla.Recordset("codzon") = 3 Then
         cbozona.ListIndex = 2
      Else
         If data_lla.Recordset("codzon") = 4 Then
            cbozona.ListIndex = 3
         Else
            If data_lla.Recordset("codzon") = 5 Then
               cbozona.ListIndex = 4
            Else
               If data_lla.Recordset("codzon") = 6 Then
                  cbozona.ListIndex = 5
               Else
                  If data_lla.Recordset("codzon") = 7 Then
                     cbozona.ListIndex = 6
                  Else
                     cbozona.ListIndex = 0
                  End If
               End If
            End If
         End If
      End If
   End If
Else
   cbozona.ListIndex = 0
End If
If IsNull(data_lla.Recordset("base")) = False Then
   cbobase.Text = data_lla.Recordset("base")
Else
   cbobase.Text = 0
End If
If IsNull(data_lla.Recordset("referen")) = False Then
   txt_direc.Text = data_lla.Recordset("referen")
Else
   txt_direc.Text = ""
End If
If IsNull(data_lla.Recordset("motcon")) = False Then
   txt_ante.Text = data_lla.Recordset("motcon")
Else
   txt_ante.Text = ""
End If
If IsNull(data_lla.Recordset("obs")) = False Then
   txt_obs.Text = data_lla.Recordset("obs")
Else
   txt_obs.Text = ""
End If
If IsNull(data_lla.Recordset("obsmot")) = False Then
   txt_mot.Text = data_lla.Recordset("obsmot")
Else
   txt_mot.Text = ""
End If
If IsNull(data_lla.Recordset("codmot")) = False Then
   If data_lla.Recordset("codmot") = "R" Then
      cbocolor.ListIndex = 2
   Else
      If data_lla.Recordset("codmot") = "A" Then
         cbocolor.ListIndex = 1
      Else
         If data_lla.Recordset("codmot") = "C" Then
            cbocolor.ListIndex = 3
         Else
            If data_lla.Recordset("codmot") = "Z" Then
               cbocolor.ListIndex = 4
            Else
               If data_lla.Recordset("codmot") = "N" Then
                  cbocolor.ListIndex = 5
               Else
                  cbocolor.ListIndex = 0
               End If
            End If
         End If
      End If
   End If
Else
   cbocolor.ListIndex = 0
End If
If cbocolor.Text = "VERDE" Then
   cbocolor.BackColor = &HC000&
Else
   If cbocolor.Text = "ROJO" Then
      cbocolor.BackColor = &HFF&
   Else
      If cbocolor.Text = "AMARILLO" Then
         cbocolor.BackColor = &HFFFF&
      Else
         If cbocolor.Text = "CELESTE" Then
            cbocolor.BackColor = &HFFFF00
         Else
            If cbocolor.Text = "AZUL" Then
               cbocolor.BackColor = &HC00000
            Else
               If cbocolor.Text = "NEGRO" Then
                  cbocolor.BackColor = &H80000006
               Else
                  cbocolor.BackColor = &HFFFFFF
               End If
            End If
         End If
      End If
   End If
End If
If IsNull(data_lla.Recordset("movilpas")) = False Then
   txt_movil.Text = data_lla.Recordset("movilpas")
Else
   txt_movil.Text = ""
End If
If IsNull(data_lla.Recordset("fecpas")) = False Then
   mfecasig.Text = Format(data_lla.Recordset("fecpas"), "dd/mm/yyyy")
Else
   mfecasig.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("horpas")) = False Then
   txt_horasig.Text = Format(data_lla.Recordset("horpas"), "HH:mm")
Else
   txt_horasig.Text = ""
End If
If IsNull(data_lla.Recordset("fecsali")) = False Then
   msalida.Text = Format(data_lla.Recordset("fecsali"), "dd/mm/yyyy")
Else
   msalida.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("horsali")) = False Then
   txt_horsal.Text = Format(data_lla.Recordset("horsali"), "HH:mm")
Else
   txt_horsal.Text = ""
End If
If IsNull(data_lla.Recordset("fec_llega")) = False Then
   mllegada.Text = Format(data_lla.Recordset("fec_llega"), "dd/mm/yyyy")
Else
   mllegada.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("hor_llega")) = False Then
   txt_horlle.Text = Format(data_lla.Recordset("hor_llega"), "HH:mm")
Else
   txt_horlle.Text = ""
End If
If IsNull(data_lla.Recordset("fec_rea")) = False Then
   mtd.Text = Format(data_lla.Recordset("fec_rea"), "dd/mm/yyyy")
Else
   mtd.Text = "__/__/____"
End If
If IsNull(data_lla.Recordset("hor_rea")) = False Then
   If data_lla.Recordset("hor_rea") <> "" Then
      txt_hortd.Text = Format(data_lla.Recordset("hor_rea"), "HH:mm")
   Else
      txt_hortd.Text = "__:__"
   End If
Else
   txt_hortd.Text = "__:__"
End If
If IsNull(data_lla.Recordset("diag")) = False Then
   txt_diag.Text = data_lla.Recordset("diag")
Else
   txt_diag.Text = ""
End If
If IsNull(data_lla.Recordset("colormot")) = False Then
   If data_lla.Recordset("colormot") = "R" Then
      cbocolfin.ListIndex = 2
   Else
      If data_lla.Recordset("colormot") = "A" Then
         cbocolfin.ListIndex = 1
      Else
         If data_lla.Recordset("colormot") = "V" Then
            cbocolfin.ListIndex = 0
         Else
            If data_lla.Recordset("colormot") = "N" Then
               cbocolfin.ListIndex = 3
            Else
               cbocolfin.Text = ""
            End If
         End If
      End If
   End If
Else
   cbocolfin.Text = ""
End If
dbcbomed.ListField = ""
dbcbomed.BoundColumn = ""
If IsNull(data_lla.Recordset("nommed")) = False Then
   dbcbomed.Text = data_lla.Recordset("nommed")
Else
   dbcbomed.Text = ""
End If
dbcbomed.ListField = "med_nombre"
dbcbomed.BoundColumn = "med_nombre"

dbcbomed2.ListField = ""
dbcbomed2.BoundColumn = ""
dbcbomed2.Text = ""
dbcbomed2.ListField = "med_nombre"
dbcbomed2.BoundColumn = "med_nombre"

If IsNull(data_lla.Recordset("codmed")) = False Then
   txt_codmed.Text = data_lla.Recordset("codmed")
Else
   txt_codmed.Text = 0
End If
If IsNull(data_lla.Recordset("trasla")) = False Then
   If data_lla.Recordset("trasla") > 0 Then
      If data_lla.Recordset("trasla") > 10 Then
         If data_lla.Recordset("trasla") = 11 Then
            cbotras.ListIndex = 8
         Else
            cbotras.ListIndex = 9
         End If
      Else
         cbotras.ListIndex = data_lla.Recordset("trasla")
      End If
   Else
      cbotras.ListIndex = 0
   End If
Else
   cbotras.ListIndex = 0
End If
If IsNull(data_lla.Recordset("lugar")) = False Then
   txt_lugar.Text = data_lla.Recordset("lugar")
Else
   txt_lugar.Text = ""
End If
If IsNull(data_lla.Recordset("hsald")) = False Then
   txt_trassal.Text = Format(data_lla.Recordset("hsald"), "HH:mm")
Else
   txt_trassal.Text = ""
End If
If IsNull(data_lla.Recordset("hllega")) = False Then
   txt_enca.Text = Format(data_lla.Recordset("hllega"), "HH:mm")
Else
   txt_enca.Text = ""
End If
If IsNull(data_lla.Recordset("hzona")) = False Then
   txt_enzona.Text = Format(data_lla.Recordset("hzona"), "HH:mm")
Else
   txt_enzona.Text = ""
End If
If IsNull(data_lla.Recordset("movtras")) = False Then
   txt_movtra.Text = data_lla.Recordset("movtras")
Else
   txt_movtra.Text = ""
End If
If IsNull(data_lla.Recordset("dcobr")) = False Then
   Combo1.Text = data_lla.Recordset("dcobr")
Else
   Combo1.Text = ""
End If
If IsNull(data_lla.Recordset("totdem")) = False Then
   txt_demora.Text = Format(data_lla.Recordset("totdem"), "HH:mm")
Else
   txt_demora.Text = ""
End If
If IsNull(data_lla.Recordset("activo")) = False Then
   Label3.Caption = Format(data_lla.Recordset("activo"), "HH:mm:ss")
Else
   Label3.Caption = "00:00:00"
End If
If IsNull(data_lla.Recordset("timdes")) = False Then
   Label39.Caption = data_lla.Recordset("timdes")
Else
   Label39.Caption = "Sin Largar"
End If
If IsNull(data_lla.Recordset("motmov")) = True Then
   txt_locali.Text = ""
Else
   txt_locali.Text = data_lla.Recordset("motmov")
End If
If IsNull(data_lla.Recordset("mm")) = True Then
   Label41.Caption = 0
Else
   Label41.Caption = data_lla.Recordset("mm")
End If
If IsNull(data_lla.Recordset("thh")) = True Then
   Label42.Caption = 0
Else
   Label42.Caption = data_lla.Recordset("thh")
End If
If IsNull(data_lla.Recordset("tmm")) = True Then
   Label43.Caption = 0
Else
   Label43.Caption = data_lla.Recordset("tmm")
End If
If IsNull(data_lla.Recordset("pasado")) = True Then
   Label44.Caption = 0
Else
   Label44.Caption = data_lla.Recordset("pasado")
End If
If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
    If IsNull(data_lla.Recordset("ano")) = True Then
       Label45.Caption = 0
    Else
       Label45.Caption = data_lla.Recordset("ano")
    End If
    If IsNull(data_lla.Recordset("mes")) = True Then
       Label46.Caption = 0
    Else
       Label46.Caption = data_lla.Recordset("mes")
    End If
Else
    Label45.Caption = 0
    Label46.Caption = 0
End If
If IsNull(data_lla.Recordset("mes")) = False Then
   If data_lla.Recordset("mes") > 10 Then
      txt_costo.Text = data_lla.Recordset("mes")
   Else
      txt_costo.Text = 0
   End If
Else
   txt_costo.Text = 0
End If
If IsNull(data_lla.Recordset("ano")) = False Then
   If data_lla.Recordset("ano") > 10 Then
      txt_boleta.Text = data_lla.Recordset("ano")
   Else
      txt_boleta.Text = 0
   End If
Else
   txt_boleta.Text = 0
End If
If IsNull(data_lla.Recordset("realiza")) = False Then
   chtmut.Value = data_lla.Recordset("realiza")
Else
   chtmut.Value = 0
End If
If IsNull(data_lla.Recordset("timsi")) = True Then
   Label48.Caption = 0
Else
   Label48.Caption = data_lla.Recordset("timsi")
End If
If IsNull(data_lla.Recordset("enfer")) = True Then
   Check2.Value = 0
Else
   Check2.Value = data_lla.Recordset("enfer")
End If
If IsNull(data_lla.Recordset("motcance")) = True Then
   txt_quien.Text = ""
Else
   txt_quien.Text = data_lla.Recordset("motcance")
End If
If IsNull(data_lla.Recordset("cancela")) = True Then
   If IsNull(data_lla.Recordset("hor_cance")) = False Then
      txt_salca.Text = data_lla.Recordset("hor_cance")
   Else
      txt_salca.Text = ""
   End If
End If
If IsNull(data_lla.Recordset("hh")) = True Then ' Sexo
   Combo3.ListIndex = -1
Else
   Combo3.ListIndex = data_lla.Recordset("hh")
End If
data_llamod.RecordSource = "Select * from resplla where nro =" & txt_nro.Text
data_llamod.Refresh
If data_llamod.Recordset.RecordCount > 0 Then
   If IsNull(data_llamod.Recordset("movil_rea")) = False Then
      If data_llamod.Recordset("movil_rea") > 0 Then
         data_chof.RecordSource = "Select * from movil where nromov =" & data_llamod.Recordset("movil_rea")
         data_chof.Refresh
         If data_chof.Recordset.RecordCount > 0 Then
            labcodchof.Caption = data_llamod.Recordset("movil_rea")
            labnomchof.Caption = "Chof.:" & data_chof.Recordset("chofer")
         Else
            labcodchof.Caption = 0
            labnomchof.Caption = ""
         End If
      Else
         labcodchof.Caption = 0
         labnomchof.Caption = ""
      End If
   Else
      labcodchof.Caption = 0
      labnomchof.Caption = ""
   End If
   If IsNull(data_llamod.Recordset("pasado")) = False Then
      Check4.Value = data_llamod.Recordset("pasado")
   Else
      Check4.Value = 0
   End If
   If IsNull(data_llamod.Recordset("mes")) = False Then
      t_codced.Text = Int(data_llamod.Recordset("mes"))
   Else
      t_codced.Text = 0
   End If
   If IsNull(data_llamod.Recordset("telef")) = False Then
      If data_llamod.Recordset("telef") = "RECIBO" Then
         Combo2.ListIndex = 0
      Else
         If data_llamod.Recordset("telef") = "CONFORME" Then
            Combo2.ListIndex = 1
         Else
            Combo2.ListIndex = -1
         End If
      End If
   Else
      Combo2.ListIndex = -1
   End If
   If IsNull(data_llamod.Recordset("hzona")) = False Then
      labcmt.Visible = True
      labcmt.Caption = "PASADO A CMT HORA:" & Format(data_llamod.Recordset("hzona"), "HH:mm")
      If IsNull(data_llamod.Recordset("mm")) = False Then
         If data_llamod.Recordset("mm") = 1 Then
            labcmt.Caption = labcmt.Caption & " NO RESUELTO H."
            If IsNull(data_llamod.Recordset("hsald")) = False Then
               labcmt.Caption = labcmt.Caption & data_llamod.Recordset("hsald")
            End If
            If IsNull(data_llamod.Recordset("totend")) = False Then
               If data_llamod.Recordset("totend") = "R" Then
                  labcmt.Caption = labcmt.Caption & " RECLASIFICA A ROJO"
               End If
               If data_llamod.Recordset("totend") = "A" Then
                  labcmt.Caption = labcmt.Caption & " RECLASIFICA A AMARILLO"
               End If
            End If
         End If
         If data_llamod.Recordset("mm") = 2 Then
            labcmt.Caption = labcmt.Caption & " RESUELTO HORA:"
            If IsNull(data_llamod.Recordset("hor_rea")) = False Then
               labcmt.Caption = labcmt.Caption & data_llamod.Recordset("hor_rea")
            End If
         End If
         If data_llamod.Recordset("mm") = 2 Or data_llamod.Recordset("mm") = 3 Then
            Command5.Visible = True
            Frame2.Visible = False
         Else
            If data_llamod.Recordset("mm") = 1 Then
               Command5.Visible = False
               If WDespa = 1 Then
                  Frame2.Visible = False
               Else
                  Frame2.Visible = True
               End If
            Else
               Command5.Visible = True
               Frame2.Visible = False
            End If
         End If
      Else
         Command5.Visible = True
         Frame2.Visible = False
      End If
   Else
      labcmt.Caption = ""
      labcmt.Visible = False
      Command5.Visible = False
      If WDespa = 1 Then
         Frame2.Visible = False
      Else
         Frame2.Visible = True
      End If
   End If
   If IsNull(data_llamod.Recordset("movilpas")) = False Then
      data_med2.Recordset.FindFirst "med_cod =" & data_llamod.Recordset("movilpas")
      If Not data_med2.Recordset.NoMatch Then
         dbcbomed2.Text = data_med2.Recordset("med_nombre")
      Else
         dbcbomed2.Text = ""
      End If
      txt_codmed2.Text = data_llamod.Recordset("movilpas")
   Else
      txt_codmed2.Text = 0
   End If
   If IsNull(data_llamod.Recordset("fec_llega")) = False Then
      mftrassol.Text = Format(data_llamod.Recordset("fec_llega"), "dd/mm/yyyy")
   Else
      mftrassol.Text = "__/__/____"
   End If
   If IsNull(data_llamod.Recordset("hor_llega")) = False Then
      mhtrassol.Text = Format(data_llamod.Recordset("hor_llega"), "HH:mm")
   Else
      mhtrassol.Text = "__:__"
   End If
Else
   txt_codmed2.Text = 0
   Check4.Value = 0
   dbcbomed2.Text = ""
   t_codced.Text = 0
   labcmt.Caption = ""
   labcmt.Visible = False
   Combo2.ListIndex = -1
   labcodchof.Caption = 0
   labnomchof.Caption = ""
End If

Exit Function

ALIgualsin:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ALIGUALSIN ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ALIGUALSIN ERR:" & Err.Number
      End If


End Function

Private Function get_Usuario() As String

    Dim Nombre As String, ret As Long
    ' Buffer
    Nombre = Space$(250)
    ' Tamaño
    ret = Len(Nombre)
    If GetUserName(Nombre, ret) = 0 Then
        get_Usuario = vbNullString
    Else
        ' Extrae solo los caracteres
        get_Usuario = Left$(Nombre, ret - 1)
    End If

End Function


Public Sub controlasiesta()
Dim Sqlconssiesta As String
Dim Regsiesta As New ADODB.Recordset
On Error GoTo ControlSI

ConectarBD
ConbdSapp.Open
If txt_nro.Text <> "" Then
   If Val(txt_nro.Text) > 0 Then
        Sqlconssiesta = "Select * from llamado where nrolla =" & txt_nro.Text
        With Regsiesta
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlconssiesta, ConbdSapp, , , adCmdText
        End With
        If Regsiesta.RecordCount > 0 Then
           Xhayregistros = 9
           b_imp.Enabled = True
        Else
           Xhayregistros = 0
        End If
    Else
       Xhayregistros = 0
    End If
Else
    Xhayregistros = 0
End If
Regsiesta.Close
ConbdSapp.Close

Exit Sub

ControlSI:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática CONTROLSI ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática CONTROLSI ERR:" & Err.Number
      End If



End Sub

Public Sub historial()
On Error GoTo Histor

If Check1.Value = 1 Then
Else
    If txt_direc.Text <> "" Then
       If txt_mat.Text <> "" Then
          If txt_mat.Text <> 0 Then
             If txt_mat.Text <= 99999998 Then
                data_ref.RecordSource = "Select * from referen where mat =" & txt_mat.Text
                data_ref.Refresh
                If data_ref.Recordset.RecordCount > 0 Then
                   data_ref.Recordset.Edit
                   data_ref.Recordset("refmat") = txt_direc.Text
                   data_ref.Recordset.Update
                Else
                   data_ref.Recordset.AddNew
                   data_ref.Recordset("mat") = txt_mat.Text
                   data_ref.Recordset("refmat") = txt_direc.Text
                   data_ref.Recordset.Update
                End If
             End If
          Else
             If txt_ced.Text <> "" Then
                data_ref.RecordSource = "Select * from referen where mat =" & txt_ced.Text
                data_ref.Refresh
                If data_ref.Recordset.RecordCount > 0 Then
                   data_ref.Recordset.Edit
                   data_ref.Recordset("refmat") = txt_direc.Text
                   data_ref.Recordset.Update
                Else
                   data_ref.Recordset.AddNew
                   data_ref.Recordset("mat") = txt_ced.Text
                   data_ref.Recordset("refmat") = txt_direc.Text
                   data_ref.Recordset.Update
                End If
             End If
          End If
       Else
          If txt_ced.Text <> "" Then
             data_ref.RecordSource = "Select * from referen where mat =" & txt_ced.Text
             data_ref.Refresh
             If data_ref.Recordset.RecordCount > 0 Then
                data_ref.Recordset.Edit
                data_ref.Recordset("refmat") = txt_direc.Text
                data_ref.Recordset.Update
             Else
                data_ref.Recordset.AddNew
                data_ref.Recordset("mat") = txt_ced.Text
                data_ref.Recordset("refmat") = txt_direc.Text
                data_ref.Recordset.Update
             End If
          End If
       End If
    End If
End If

If txt_ante.Text <> "" Then
   If txt_ced.Text <> "" Then
         data_histant.RecordSource = "Select * from ante where ced =" & txt_ced.Text
         data_histant.Refresh
         If data_histant.Recordset.RecordCount > 0 Then
            data_histant.Recordset.Edit
            data_histant.Recordset("fecha") = Date
            data_histant.Recordset("ante") = txt_ante.Text
            If txt_mat.Text <> "" Then
               If txt_mat.Text > 0 Then
                  data_histant.Recordset("matric") = txt_mat.Text
               End If
            End If
            data_histant.Recordset.Update
         Else
            data_histant.Recordset.AddNew
            data_histant.Recordset("fecha") = Date
            data_histant.Recordset("ced") = txt_ced.Text
            data_histant.Recordset("ante") = txt_ante.Text
            If txt_mat.Text <> "" Then
               If txt_mat.Text > 0 Then
                  data_histant.Recordset("matric") = txt_mat.Text
               End If
            End If
            data_histant.Recordset.Update
         End If
   End If
End If

Exit Sub

Histor:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática HISTOR ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática HISTOR ERR:" & Err.Number
      End If



End Sub

Public Sub despuesdegraba()

Frame1.Enabled = False
Frame2.Enabled = False
b_nuevo.Enabled = True
b_modif.Enabled = True
b_imp.Enabled = True
b_hist.Enabled = True
b_buscar.Enabled = True
b_grabar.Enabled = False
b_cancel.Enabled = True
b_cancela.Enabled = False
b_pend.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
b_covid.Enabled = True

End Sub

'Public Function ConectarBD()
'ConbdSapp.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xconexrmt & ";PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=sapp1987;OPTION=3;"

'End Function
Public Sub ControlCosto()
Dim Xcodestudio As Long
Dim XImp, Ximpest, Ximppart, Xdescuento As Double
XImp = 0
Ximpest = 0
On Error GoTo Veralcosto

If mtd.Text = "__/__/____" Then
   If txt_hortd.Text = "__:__" Then
      If txt_cat.Text = "SAMCB" Or cbobase.Text > 0 Or cbozona.ListIndex >= 4 Then
         XImp = 0
         txt_costo.Text = 0
      Else
         If cbocolor.Text = "" Then
         Else
            If cbocolor.Text = "VERDE" Or cbocolor.Text = "CELESTE" Then
               If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                  Xcodestudio = 10014
               Else
                  Xcodestudio = 10002
               End If
            Else
               If cbocolor.Text = "AMARILLO" Then
                  If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                     Xcodestudio = 10013
                  Else
                     Xcodestudio = 10004
                  End If
               Else
                  If cbocolor.Text = "ROJO" Then
                     If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                        Xcodestudio = 10012
                     Else
                        Xcodestudio = 10006
                     End If
                  Else
                     If cbocolor.Text = "AZUL" Then
                        Xcodestudio = 14004
                     Else
                        If cbocolor.Text = "NEGRO" Then
                           If txt_cat.Text = "911" Or txt_cat.Text = "911B" Then
                              Xcodestudio = 10012
                           Else
                              Xcodestudio = 10016
                           End If
                        Else
                           Xcodestudio = 10016
                        End If
                     End If
                  End If
               End If
            End If
            If txt_cat.Text = "MSP" Then
               If cbocolor.Text = "ROJO" Then
                  Xcodestudio = 90017
               Else
                  If cbocolor.Text = "AMARILLO" Then
                     Xcodestudio = 90018
                  Else
                     Xcodestudio = 90019
                  End If
               End If
            End If
             
            If txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "CERDGI" Or _
               txt_cat.Text = "CERADU" Or txt_cat.Text = "CERHEV" Or txt_cat.Text = "CERMAT" Or txt_cat.Text = "CERSEV" Or txt_cat.Text = "CERVIS" Then
        '''''''UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and base =" & 0 & " and movilpas <>" & 99
               Xcodestudio = 10008
            End If
              
'Desde acá recuperar
            If txt_cat.Text <> "" Then
               data_aran.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
               data_aran.Refresh
               If data_aran.Recordset.RecordCount > 0 Then
                  If IsNull(data_aran.Recordset("cnv_aran")) = False Then
                     Xop1 = data_aran.Recordset("cnv_aran")
                  Else
                     Xop1 = 0
                  End If
               Else
                  Xop1 = 0
               End If
            Else
               Xop1 = 0
            End If
            data_aran.RecordSource = "Select * from estudios where codest =" & Xcodestudio
            data_aran.Refresh
            If data_aran.Recordset.RecordCount > 0 Then
               If IsNull(data_aran.Recordset("cons")) = False Then
                  Ximpest = data_aran.Recordset("cons")
                  Ximppart = data_aran.Recordset("part")
               Else
                  Ximpest = 0
                  Ximppart = 0
               End If
               data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & Xop1 & " and id_serv =" & Xcodestudio
               data_aran.Refresh
               If data_aran.Recordset.RecordCount > 0 Then
                  If data_aran.Recordset("prec_serv") > 0 Then
                     XImp = data_aran.Recordset("prec_serv")
                  Else
                     If data_aran.Recordset("por_serv") = 100 Then
                        XImp = 0
                     Else
                        If data_aran.Recordset("por_serv") = 0 Then
                           XImp = Ximpest
                        Else
                           Xdescuento = data_aran.Recordset("por_serv") * Ximpest / 100
                           XImp = Ximpest - Xdescuento
                        End If
                     End If
                  End If
               Else
                  XImp = Ximppart
               End If
               If txt_cat.Text = "PART" Then
                  XImp = Ximppart
               End If
                            
               If txt_cat.Text <> "PART" Then
                  data_convbus.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
                  data_convbus.Refresh
                  If data_convbus.Recordset.RecordCount > 0 Then
                     If IsNull(data_convbus.Recordset("cnv_colrec")) = False Then
                        If txt_cat.Text = "SA" Or txt_cat.Text = "SAF" Or txt_cat.Text = "CCOMS" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "SEGAM" Or _
                           txt_cat.Text = "EMERN" Or txt_cat.Text = "EMERG" Or txt_cat.Text = "EMERNE" Or txt_cat.Text = "CAAM" Or _
                           txt_cat.Text = "SAP" Or txt_cat.Text = "VIV19" Or txt_cat.Text = "VIV20" Or txt_cat.Text = "CAAMEP" Or _
                           txt_cat.Text = "MSP" Or txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "APNORE" Or _
                           txt_cat.Text = "SJ01" Or data_convbus.Recordset("cnv_colrec") = "M" Or data_convbus.Recordset("cnv_colrec") = "R" Or _
                           Mid(txt_cat.Text, 1, 4) = "TALA" Or cbozona.Text = 3 Or txt_cat.Text = "SEMM" Or txt_cat.Text = "SEMM1" Then
                           XImp = 0
                        End If
                     Else
                        If txt_cat.Text = "SA" Or txt_cat.Text = "SAF" Or txt_cat.Text = "CCOMS" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "SEGAM" Or _
                           txt_cat.Text = "EMERN" Or txt_cat.Text = "EMERG" Or txt_cat.Text = "EMERNE" Or txt_cat.Text = "CAAM" Or txt_cat.Text = "911" Or _
                           txt_cat.Text = "SAP" Or txt_cat.Text = "VIV19" Or txt_cat.Text = "VIV20" Or txt_cat.Text = "CAAMEP" Or txt_cat.Text = "911B" Or _
                           txt_cat.Text = "MSP" Or txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "APNORE" Or txt_cat.Text = "CASH" Or _
                           txt_cat.Text = "SJ01" Or Mid(txt_cat.Text, 1, 4) = "TALA" Or cbozona.Text = 3 Then
                           XImp = 0
                        End If
                     End If
                  Else
                     If txt_cat.Text = "SA" Or txt_cat.Text = "SAF" Or txt_cat.Text = "CCOMS" Or txt_cat.Text = "CERCAS" Or txt_cat.Text = "SEGAM" Or _
                        txt_cat.Text = "EMERN" Or txt_cat.Text = "EMERG" Or txt_cat.Text = "EMERNE" Or txt_cat.Text = "CAAM" Or txt_cat.Text = "911" Or _
                        txt_cat.Text = "SAP" Or txt_cat.Text = "VIV19" Or txt_cat.Text = "VIV20" Or txt_cat.Text = "CAAMEP" Or txt_cat.Text = "911B" Or _
                        txt_cat.Text = "MSP" Or txt_cat.Text = "UDEMM" Or txt_cat.Text = "CERSEM" Or txt_cat.Text = "APNORE" Or txt_cat.Text = "CASH" Or _
                        txt_cat.Text = "SJ01" Or Mid(txt_cat.Text, 1, 4) = "TALA" Or cbozona.Text = 3 Then
                        XImp = 0
                     End If
                  End If
                  If txt_cat.Text = "MUCAFL" Or txt_cat.Text = "MUCAMA" Or txt_cat.Text = "MUCAMI" Or txt_cat.Text = "MUCAMM" Or txt_cat.Text = "MUCAMP" Or _
                     txt_cat.Text = "MUCAMS" Or txt_cat.Text = "MUCAMT" Or txt_cat.Text = "MUCATA" Or txt_cat.Text = "SOLEME" Or txt_cat.Text = "UNIPA" Or _
                     txt_cat.Text = "CAAMEP" Or txt_cat.Text = "SOLAF" Or txt_cat.Text = "SOLAMB" Or txt_cat.Text = "SOC" Or txt_cat.Text = "CPS" Then
                     XImp = 0
                  End If
               End If
            End If
         End If
         If txt_costo.Text <> "" Then
            If txt_costo.Text > 0 Then
            Else
               txt_costo.Text = Format(XImp, "Standard")
            End If
         Else
            txt_costo.Text = Format(XImp, "Standard")
         End If
      End If
   End If
End If

Exit Sub

Veralcosto:
           If Err.Number = 3155 Then
              MsgBox "ERROR al grabar el nuevo costo del llamado, avise a informática", vbInformation
           Else
              MsgBox "ERROR al modificar al costo de llamado, avise a informática", vbInformation
           End If
           
End Sub


Public Sub GrabaNosapp()
             
data_parsec.DatabaseName = ""
data_parsec.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_parsec.RecordSource = "select * from cartasnosapp where nrolla =" & txt_nro.Text
data_parsec.Refresh

If data_parsec.Recordset.RecordCount > 0 Then
   data_parsec.Recordset.Edit
   If cbocolor.Text <> "" Then
      data_parsec.Recordset("codigo") = cbocolor.Text
   Else
      data_parsec.Recordset("codigo") = "S/D"
   End If
   If txt_mot.Text <> "" Then
      data_parsec.Recordset("motivo") = Mid(txt_mot.Text, 1, 100)
   End If
   If cbozona.Text <> "" Then
      data_parsec.Recordset("zona") = cbozona.Text
   End If
   data_parsec.Recordset("hora_lla") = txt_hora.Text
   data_parsec.Recordset.Update
   data_parsec.Refresh
Else
   MsgBox "No se encontró registro de carta para el llamado. Avise a informática.", vbInformation
End If
data_parsec.Connect = ""
data_parsec.DatabaseName = App.path & "\mensa.mdb"
data_parsec.RecordSource = "mensaje"
data_parsec.Refresh


End Sub

Public Sub carga_trasl()
Dim Sqlconstr As String
Dim Regtr As New ADODB.Recordset
On Error GoTo ControlTrasl

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Sqlconstr = "Select * from traslados"
With Regtr
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Sqlconstr, ConbdSapp, , , adCmdText
End With
cbotras.Clear
If Regtr.RecordCount > 0 Then
   Do While Not Regtr.EOF
      cbotras.AddItem Regtr("descrip")
      Regtr.MoveNext
   Loop
End If
Regtr.Close
ConbdSapp.Close

Exit Sub

ControlTrasl:
      If Err.Number = 444 Then
         MsgBox "No se pudo cargar, comunique a informática TRASLADOS ERR:" & Err.Description
      Else
         MsgBox "Error al cargar, comunique a informática TRASLADOS ERR:" & Err.Number
      End If

End Sub

Public Sub Consulta_cedcovid()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from llamado where ci =" & txt_ced.Text & " and segui_covid in (1) and cierre_hora is null"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   MsgBox "ATENCION!!! Este paciente ya figura ingresado para seguimiento. Verifique!", vbCritical
End If

Xrecclii.Close
ConbdSapp.Close


End Sub

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
Xsqlpromo = "Select * from clientes_history where cl_codigo =" & Val(txt_mat.Text) & " and fecha_modif >='" & Format(Fecha_Datos, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount <= 0 Then
   XorigenDatos = 2 'Despacho
Else
   DatosVerificadosOk = 0
End If

Xrecclii.Close
ConbdSapp.Close

If DatosVerificadosOk = 1 Then
   frm_valida_datos_socio.Show vbModal
End If

End Sub
Public Sub altaValidacionDatosabm()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from abmsocio where cl_codigo =" & Val(txt_mat.Text)
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
Xrecclii("cl_codigo") = Val(txt_mat.Text)
Xrecclii("desc") = "MODIF"
Xrecclii("cl_motivo") = "CANCELA VALIDACION"
Xrecclii("convenio") = txt_cat.Text
Xrecclii("base") = data_par.Recordset("base")
Xrecclii.Update


Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub guarda_Alcancelar()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

On Error GoTo Nopudograbar

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

Xsqlpromo = "Select * from llamado where nrolla =" & Val(txt_nro.Text)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xrecclii("editando") = 1
Xrecclii.Update


Xrecclii.Close
ConbdSapp.Close

Exit Sub

Nopudograbar:
            If Err.Number = 3197 Then
               MsgBox "Ya está modificado."
            Else
               MsgBox "No se puede grabar."
            End If
            Xrecclii.Close
            ConbdSapp.Close


End Sub



Public Sub Llamado_Ap()
Dim Xsqlpromo, XsqlCons, XsqlCantCons As String
Dim Xrecclii As New ADODB.Recordset
Dim XrecCons As New ADODB.Recordset
Dim XCantidadCons As New ADODB.Recordset

Dim XfechaDesde As String
Dim XfecD, XfecDAnual As Date
Dim MensualAnual, XcantConsultas As Integer
Dim Xelprecio As Double
Xelprecio = 0
MensualAnual = 0
XcantConsultas = 0
XfecDAnual = Date - 365
XllevacostoAp = 0
If Month(Date) > 9 Then
   XfechaDesde = "01/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
Else
   XfechaDesde = "01/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
End If
XfecD = CDate(XfechaDesde)

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

Xsqlpromo = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "' and cnv_cant_r in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cnv_menanio")) = False Then
      MensualAnual = Xrecclii("cnv_menanio")
   End If
   If IsNull(Xrecclii("cnv_grupoap")) = False Then
      If IsNull(Xrecclii("cnv_cantcons")) = False Then
         XcantConsultas = Xrecclii("cnv_cantcons")
         If IsNull(Xrecclii("cnv_preccons")) = False Then
            XsqlCons = "Select * from convenio_llam where convenio ='" & Xrecclii("cnv_codigo") & "' and fecha >='" & Format(XfecD, "yyyy-mm-dd") & "' and fecha <='" & Format(Date, "yyyy-mm-dd") & "'"
            With XrecCons
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open XsqlCons, ConbdSapp, , , adCmdText
            End With
            If XrecCons.RecordCount > 0 Then
               XrecCons.MoveLast
               If XrecCons.RecordCount >= XcantConsultas Then
                  MsgBox "ATENCION! Llamado con costo, comunique al solicitante! Valor: " & Val(Xrecclii("cnv_preccons")) & " Pesos, que será incluido en la factura mensual.", vbCritical, "Despacho"
                  XllevacostoAp = 9
               Else
                  XllevacostoAp = 8
               End If
            Else
               XllevacostoAp = 8
            End If
            XrecCons.Close
         Else
            XllevacostoAp = 8
         End If
      Else
         XsqlCantCons = "Select * from convenio where cnv_codigo ='" & Xrecclii("cnv_grupoap") & "' and cnv_cant_r in (1)"
         With XCantidadCons
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open XsqlCantCons, ConbdSapp, , , adCmdText
         End With
         If XCantidadCons.RecordCount > 0 Then
            If IsNull(XCantidadCons("cnv_menanio")) = False Then
               MensualAnual = XCantidadCons("cnv_menanio")
            End If
            If IsNull(XCantidadCons("cnv_cantcons")) = False Then
               XcantConsultas = XCantidadCons("cnv_cantcons")
               Xelprecio = XCantidadCons("cnv_preccons")
            Else
               XcantConsultas = 0
            End If
         End If
         XCantidadCons.Close
         
         If MensualAnual = 1 Then
            XsqlCons = "Select * from convenio_llam where convenio ='" & Xrecclii("cnv_grupoap") & "' and fecha >='" & Format(XfecDAnual, "yyyy-mm-dd") & "' and fecha <='" & Format(Date, "yyyy-mm-dd") & "'"
         Else
            XsqlCons = "Select * from convenio_llam where convenio ='" & Xrecclii("cnv_grupoap") & "' and fecha >='" & Format(XfecD, "yyyy-mm-dd") & "' and fecha <='" & Format(Date, "yyyy-mm-dd") & "'"
         End If
         With XrecCons
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open XsqlCons, ConbdSapp, , , adCmdText
         End With
         If XrecCons.RecordCount > 0 Then
            XrecCons.MoveLast
            If XrecCons.RecordCount >= XcantConsultas Then
               MsgBox "ATENCION! Llamado con costo, comunique al solicitante! Valor: " & Val(Xelprecio) & " Pesos, que será incluido en la factura mensual.", vbCritical, "Despacho"
               XllevacostoAp = 9
            Else
               XllevacostoAp = 8
            End If
         Else
            XllevacostoAp = 8
         End If
         XrecCons.Close
      End If
   Else
      If IsNull(Xrecclii("cnv_cantcons")) = False Then
         XcantConsultas = Xrecclii("cnv_cantcons")
         If IsNull(Xrecclii("cnv_preccons")) = False Then
            XsqlCons = "Select * from convenio_llam where convenio ='" & Xrecclii("cnv_codigo") & "' and fecha >='" & Format(XfecD, "yyyy-mm-dd") & "' and fecha <='" & Format(Date, "yyyy-mm-dd") & "'"
            With XrecCons
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open XsqlCons, ConbdSapp, , , adCmdText
            End With
            If XrecCons.RecordCount > 0 Then
               XrecCons.MoveLast
               If XrecCons.RecordCount >= XcantConsultas Then
                  MsgBox "ATENCION! Llamado con costo, comunique al solicitante! Valor: " & Val(Xrecclii("cnv_preccons")) & " Pesos, que será incluido en la factura mensual.", vbCritical, "Despacho"
                  XllevacostoAp = 9
               Else
                  XllevacostoAp = 8
               End If
            Else
               XllevacostoAp = 8
            End If
            XrecCons.Close
         Else
            XllevacostoAp = 8
         End If
      End If
   End If
End If

Xrecclii.Close
ConbdSapp.Close


End Sub

Public Sub Graba_CostoAp()
Dim Xsqlpromo, XsqlCons, XelGrupo As String
Dim Xrecclii As New ADODB.Recordset
Dim XrecCons As New ADODB.Recordset
Dim XrecGraba As New ADODB.Recordset

Dim Xvalor As Double
Xvalor = 0
XelGrupo = ""

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

Xsqlpromo = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cnv_grupoap")) = False Then
      XelGrupo = Xrecclii("cnv_grupoap")
   Else
      If IsNull(Xrecclii("cnv_preccons")) = False Then
         Xvalor = Xrecclii("cnv_preccons")
         XelGrupo = Xrecclii("cnv_codigo")
      End If
   End If
End If

If Trim(XelGrupo) <> "" Then
   Xsqlpromo = "Select * from convenio where cnv_codigo ='" & XelGrupo & "'"
   With XrecCons
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   If XrecCons.RecordCount > 0 Then
      If IsNull(XrecCons("cnv_preccons")) = False Then
         Xvalor = XrecCons("cnv_preccons")
      End If
   End If
   XrecCons.Close
End If


Xsqlpromo = "Select * from convenio_tiquets where fecha ='" & Format(Date, "yyyy-mm-dd") & "'"
With XrecGraba
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

XrecGraba.AddNew
XrecGraba("id_convenio") = txt_cat.Text
XrecGraba("fecha") = Date
XrecGraba("hora") = Format(Time, "HH:mm")
XrecGraba("importe") = Xvalor
XrecGraba("nom_grupo") = XelGrupo
If Trim(txt_nomb.Text) <> "" Then
   XrecGraba("nombre") = Mid(txt_nomb.Text, 1, 80)
End If
If Trim(txt_ced.Text) <> "" Then
   If Trim(t_codced.Text) <> "" Then
      XrecGraba("cedula") = txt_ced.Text & t_codced.Text
   Else
      XrecGraba("cedula") = txt_ced.Text
   End If
End If

XrecGraba.Update

XrecGraba.Close
Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Grabar_llamadoAp()
Dim Xsqlpromo, XsqlCons, XelGrupo As String
Dim Xrecclii As New ADODB.Recordset
Dim XrecGraba As New ADODB.Recordset

XelGrupo = ""

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

Xsqlpromo = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cnv_grupoap")) = False Then
      XelGrupo = Xrecclii("cnv_grupoap")
   Else
      If IsNull(Xrecclii("cnv_cantcons")) = False Then
         XelGrupo = Xrecclii("cnv_codigo")
      End If
   End If
End If
Xrecclii.Close

If Trim(XelGrupo) <> "" Then
   Xsqlpromo = "Select * from convenio_llam where convenio ='" & XelGrupo & "'"
   With XrecGraba
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   XrecGraba.AddNew
   XrecGraba("fecha") = Date
   XrecGraba("convenio") = XelGrupo
   XrecGraba.Update
   XrecGraba.Close
End If

ConbdSapp.Close

End Sub

Public Function Verificar_digito() As Integer
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot, Xlacedu As Long
Dim Xcedtex, Xtottex, Xcodced As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo As Long

If Trim(txt_ced.Text) <> "" Then
   If txt_ced.Text <> 0 Then
      Xn1 = 2
      Xn2 = 9
      Xn3 = 8
      Xn4 = 7
      Xn5 = 6
      Xn6 = 3
      Xn7 = 4
      Xpond = 10
      
      If Len(txt_ced.Text) = 6 Or Len(txt_ced.Text) = 7 Then
         Xcedtex = Trim(txt_ced.Text)
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
         Xced7 = Val(Mid(Xcedtex, 7, 1))
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
         If t_codced.Text = Val(Xtot) Then
            Verificar_digito = Xtot
         Else
            Verificar_digito = 99
         End If
      Else
         Verificar_digito = 99
      End If
   Else
      Verificar_digito = 999
   End If
Else
   Verificar_digito = 999
End If


End Function
Public Sub Lleva_timbre()
Dim Xsqlpromo, XsqlCons As String
Dim Xrecclii As New ADODB.Recordset
Dim XrecCosto As New ADODB.Recordset
Dim XsiTimbre As Integer
XsiTimbre = 0

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

If Trim(txt_cat.Text) <> "" Then
   Xsqlpromo = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
   With Xrecclii
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   If Xrecclii.RecordCount > 0 Then
      If IsNull(Xrecclii("cnv_sald")) = False Then
         XsiTimbre = Xrecclii("cnv_sald")
      End If
   End If
   If XsiTimbre = 1 Then
      XsqlCons = "Select * from estudios where codest =" & 995
      With XrecCosto
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open XsqlCons, ConbdSapp, , , adCmdText
      End With
      If XrecCosto.RecordCount > 0 Then
         MsgBox "ATENCION!! El llamado lleva costo de Timbre Profesional. Pesos:" & Trim(str(Val(XrecCosto("cons")))) & ". Comunique al solicitante y al móvil para el cobro.", vbCritical
         cbotimbre.ListIndex = 1
         t_timbre.Text = Val(XrecCosto("cons"))
      Else
         cbotimbre.ListIndex = -1
         t_timbre.Text = ""
      End If
      XrecCosto.Close
   Else
      cbotimbre.ListIndex = -1
      t_timbre.Text = ""
   End If
   Xrecclii.Close
Else
   cbotimbre.ListIndex = -1
   t_timbre.Text = ""
End If

ConbdSapp.Close

End Sub

Public Sub Grabar_RepeticionMed()
Dim Xsqlpromo As String
Dim Xrep As Integer
Xrep = 0
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

If Trim(txt_nro.Text) <> "" Then
   Xsqlpromo = "Select * from llamado where nrolla =" & Val(txt_nro.Text)
   With Xrecclii
      .CursorLocation = adUseClient
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
Else
   MsgBox "No hay llamado seleccionado para cargar repetición.", vbCritical
End If

If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("repite")) = False Then
      Xrep = Xrecclii("repite")
      If Xrep <> 1 Then
         Xrecclii("repite") = 1
         Xrecclii.Update
      End If
   Else
      Xrecclii("repite") = 1
      Xrecclii.Update
   End If
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Public Sub Grabar_CmtAt()

Dim Xsqlpromo As String
Dim XsqlSegundo As String

Dim Xrecclii As New ADODB.Recordset
Dim XrecSegundo As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

If Trim(txt_nro.Text) <> "" Then
   Xsqlpromo = "Select * from llamado where nrolla =" & Val(txt_nro.Text)
   With Xrecclii
      .CursorLocation = adUseClient
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
Else
   MsgBox "No hay llamado seleccionado para cargar repetición.", vbCritical
End If

If Xrecclii.RecordCount > 0 Then
   If Xrecclii("pend") = 4 Or Xrecclii("pend") = 1 Or Xrecclii("pend") = 2 Then
      If IsNull(Xrecclii("codmedcmt")) = False Then
         If Xrecclii("codmedcmt") <> 9999 Then
            Xrecclii("codmedcmt") = 9999
            Xrecclii("editando") = 1
            Xrecclii.Update
         End If
      Else
         Xrecclii("pend") = 4
         Xrecclii("codmedcmt") = 9999
         Xrecclii("editando") = 1
         Xrecclii.Update
      End If
   Else
      Xrecclii("pend") = 4
      Xrecclii("codmedcmt") = 9999
      Xrecclii("editando") = 1
      Xrecclii.Update
      
      XsqlSegundo = "Select * from resplla where nro =" & Xrecclii("nrolla")
      With XrecSegundo
         .CursorLocation = adUseClient
         .CursorType = adOpenKeyset
         .LockType = adLockOptimistic
         .Open XsqlSegundo, ConbdSapp, , , adCmdText
      End With
      
      If XrecSegundo.RecordCount > 0 Then
         XrecSegundo("hzona") = Format(Time, "HH:mm")
         XrecSegundo("pend") = 1
         XrecSegundo.Update
      Else
         XrecSegundo.AddNew
         XrecSegundo("nro") = txt_nro.Text
         XrecSegundo("fecha") = mfecha.Text
         XrecSegundo("hzona") = Format(Time, "HH:mm")
         XrecSegundo("pend") = 1
         XrecSegundo.Update
      End If
      labcmt.Visible = True
      labcmt.Caption = "PASADO A CMT HORA:" & Format(Time, "HH:mm")
      b_cmt.Enabled = False
      XrecSegundo.Close
   End If
   
Else
   MsgBox "No hay llamado seleccionado para cargar repetición.", vbCritical
End If

Xrecclii.Close
ConbdSapp.Close
                          

End Sub
Public Function Verificar_siTieneL() As Integer
Dim Xverifica As Integer
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open

If Trim(txt_ced.Text) <> "" Then
   If Val(txt_ced.Text) > 0 Then
       Xsqlpromo = "Select * from llamado where ci =" & Val(txt_ced.Text) & " and fecha ='" & Format(Date, "yyyy-mm-dd") & "' and pend in (0,4) and cancela is null"
       With Xrecclii
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Xsqlpromo, ConbdSapp, , , adCmdText
       End With
    
       If Xrecclii.RecordCount > 0 Then
          Verificar_siTieneL = 1
       Else
          Verificar_siTieneL = 99
       End If
   Else
       Verificar_siTieneL = 99
   End If
Else
   Verificar_siTieneL = 99
End If

Xrecclii.Close
ConbdSapp.Close

End Function

