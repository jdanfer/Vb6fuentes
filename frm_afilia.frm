VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_afilia 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de afiliación"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   Icon            =   "frm_afilia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Documento Extranjero"
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
      Left            =   6960
      TabIndex        =   116
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Data data_nrosoc 
      Caption         =   "data_nrosoc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   4560
   End
   Begin Crystal.CrystalReport cr1print 
      Left            =   5160
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton b_imp 
      Height          =   495
      Left            =   120
      Picture         =   "frm_afilia.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   111
      ToolTipText     =   "Realizar impresión del contrato"
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton b_correo 
      Height          =   495
      Left            =   1320
      Picture         =   "frm_afilia.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   110
      ToolTipText     =   "Enviar correo con copia del contrato al titular"
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport cr2pant 
      Left            =   6360
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowControls  =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
   End
   Begin VB.Data data_hist 
      Caption         =   "data_hist"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      Picture         =   "frm_afilia.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "Consultar datos en padrón SAPP"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Data data_abm 
      Caption         =   "data_abm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_afilcons 
      Caption         =   "data_afilcons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_cob 
      Caption         =   "data_cob"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11040
      Picture         =   "frm_afilia.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Consultar afiliaciones ingresadas"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Data data_afilia 
      Caption         =   "data_afilia"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_nro 
      Caption         =   "data_nro"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   10200
      Picture         =   "frm_afilia.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Cancelar afiliación"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      Picture         =   "frm_afilia.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Guardar los datos de la afiliación"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8520
      Picture         =   "frm_afilia.frx":26C6
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Crear nueva afiliación"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la afiliación"
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
      ForeColor       =   &H00FF0000&
      Height          =   6975
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   11655
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         Picture         =   "frm_afilia.frx":2C50
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Grabar modificaciones en la afiliación. Sólo nombres, nacimiento, sexo, teléfonos, correo, dirección."
         Top             =   6480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cbovende 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         Left            =   4920
         TabIndex        =   49
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos para el cobro sin tarjeta"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   47
         Top             =   4440
         Visible         =   0   'False
         Width           =   11415
         Begin MSMask.MaskEdBox mph 
            Height          =   495
            Left            =   2880
            TabIndex        =   107
            Top             =   840
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox mpd 
            Height          =   495
            Left            =   240
            TabIndex        =   106
            Top             =   840
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.ComboBox cbodiacob 
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
            Left            =   10320
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox t_casacobro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8880
            MaxLength       =   40
            TabIndex        =   82
            Top             =   840
            Width           =   2295
         End
         Begin VB.ComboBox cbozonacobro 
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
            Left            =   2280
            TabIndex        =   80
            Top             =   1320
            Width           =   3975
         End
         Begin VB.TextBox t_dircobro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   59
            Top             =   840
            Width           =   5055
         End
         Begin VB.ComboBox cbocobnom 
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
            Left            =   3000
            TabIndex        =   56
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox t_cobnro 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   55
            Top             =   360
            Width           =   735
         End
         Begin VB.Label labplaz 
            BackColor       =   &H00FF0000&
            Caption         =   "Plazos de la afiliación:"
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
            Left            =   240
            TabIndex        =   105
            Top             =   480
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.Label labcodzoncobr 
            Height          =   375
            Left            =   5520
            TabIndex        =   83
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label35 
            BackColor       =   &H00C00000&
            Caption         =   "Localidad:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   81
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label34 
            BackColor       =   &H00C00000&
            Caption         =   "Nombre de la casa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   7440
            TabIndex        =   79
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C00000&
            Caption         =   "Dirección de cobro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   58
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C00000&
            Caption         =   "Día probable de cobro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   7440
            TabIndex        =   57
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C00000&
            Caption         =   "Cobrador sugerido:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos para el cobro con tarjeta (Autorización para débito automático)"
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
         Height          =   2055
         Left            =   120
         TabIndex        =   46
         Top             =   4320
         Visible         =   0   'False
         Width           =   11415
         Begin VB.CheckBox ch_debbrou 
            BackColor       =   &H00800000&
            Caption         =   "Verificado que Titular ya realizó formulario en la web del BROU (www.brou.com.uy)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   3840
            TabIndex        =   95
            Top             =   360
            Visible         =   0   'False
            Width           =   7455
         End
         Begin VB.TextBox t_teleftarj 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7680
            MaxLength       =   50
            TabIndex        =   78
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox t_domitarj 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   76
            Top             =   1320
            Width           =   4695
         End
         Begin VB.TextBox t_codcedtit 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8400
            MaxLength       =   1
            TabIndex        =   74
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox t_cedtit 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            MaxLength       =   9
            TabIndex        =   73
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboaniov 
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   360
            Left            =   10200
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cbomesv 
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   360
            Left            =   9480
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox t_nomtit 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   68
            Top             =   840
            Width           =   3975
         End
         Begin VB.ComboBox cbosello 
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   360
            ItemData        =   "frm_afilia.frx":31DA
            Left            =   1320
            List            =   "frm_afilia.frx":31ED
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox t_nrotarj 
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
            Height          =   405
            Left            =   5040
            MaxLength       =   4
            TabIndex        =   117
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox t_nrotarj2 
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
            Height          =   405
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   118
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox t_nrotarj3 
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
            Height          =   405
            Left            =   6480
            MaxLength       =   4
            TabIndex        =   119
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox t_nrotarj4 
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
            Height          =   405
            Left            =   7200
            MaxLength       =   4
            TabIndex        =   120
            Top             =   360
            Width           =   735
         End
         Begin VB.Label labcodsello 
            Height          =   255
            Left            =   480
            TabIndex        =   84
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C00000&
            Caption         =   "Teléfono/s:"
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
            Left            =   6360
            TabIndex        =   77
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C00000&
            Caption         =   "Domicilio:"
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
            Left            =   240
            TabIndex        =   75
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label31 
            BackColor       =   &H00C00000&
            Caption         =   "CI Titular:"
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
            Left            =   5640
            TabIndex        =   72
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C00000&
            Caption         =   "Vence:"
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
            Left            =   8520
            TabIndex        =   69
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C00000&
            Caption         =   "Nombre Titular:"
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
            Left            =   240
            TabIndex        =   67
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C00000&
            Caption         =   "Nro.Tarjeta"
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
            Left            =   3840
            TabIndex        =   66
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C00000&
            Caption         =   "Tarjeta:"
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
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox t_valor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   45
         Top             =   3480
         Width           =   1695
      End
      Begin VB.ComboBox cbomut 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         Left            =   8160
         TabIndex        =   43
         Top             =   2880
         Width           =   3255
      End
      Begin VB.ComboBox cbozona 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         Left            =   1800
         TabIndex        =   41
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox t_casa 
         Height          =   375
         Left            =   8640
         MaxLength       =   45
         TabIndex        =   39
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox t_sol 
         Height          =   375
         Left            =   10800
         MaxLength       =   10
         TabIndex        =   37
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox t_manz 
         Height          =   375
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   36
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox t_entre 
         Height          =   405
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   33
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox t_calle 
         Height          =   375
         Left            =   1800
         MaxLength       =   70
         TabIndex        =   32
         Top             =   2040
         Width           =   4815
      End
      Begin VB.TextBox t_correo 
         Height          =   375
         Left            =   7800
         MaxLength       =   250
         TabIndex        =   30
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox t_celu 
         Height          =   375
         Left            =   4680
         MaxLength       =   9
         TabIndex        =   28
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox t_telef 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   26
         ToolTipText     =   "Si no tiene, ingrese: NO APLICA"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox cbosexo 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         ItemData        =   "frm_afilia.frx":3222
         Left            =   9240
         List            =   "frm_afilia.frx":322C
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1080
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mfnac 
         Height          =   375
         Left            =   9960
         TabIndex        =   22
         Top             =   720
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
      Begin VB.ComboBox cbopromo 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         ItemData        =   "frm_afilia.frx":3245
         Left            =   8040
         List            =   "frm_afilia.frx":3247
         TabIndex        =   20
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox cbocat 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         ItemData        =   "frm_afilia.frx":3249
         Left            =   2520
         List            =   "frm_afilia.frx":324B
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox t_ape2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         MaxLength       =   45
         TabIndex        =   16
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox t_ape1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MaxLength       =   45
         TabIndex        =   15
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox t_nom2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         MaxLength       =   45
         TabIndex        =   14
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox t_nom1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   45
         TabIndex        =   13
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Haga doble click aquí para cambiar a pago con tarjeta"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   2040
         MouseIcon       =   "frm_afilia.frx":324D
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   6480
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C00000&
         Caption         =   "Con descuento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label labantcat 
         Height          =   255
         Left            =   480
         TabIndex        =   102
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label labtotafil 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1800
         TabIndex        =   101
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label labcedtit 
         Height          =   255
         Left            =   9960
         TabIndex        =   100
         Top             =   6480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label labnompromo 
         Height          =   255
         Left            =   4920
         TabIndex        =   99
         Top             =   3360
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label labintegra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Height          =   375
         Left            =   10200
         TabIndex        =   97
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Integrante:"
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
         Height          =   375
         Left            =   8880
         TabIndex        =   96
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label labapagar 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   9120
         TabIndex        =   93
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label labdapagar 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7920
         TabIndex        =   92
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label labdescpromo 
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
         Left            =   3600
         TabIndex        =   91
         Top             =   3840
         Width           =   3975
      End
      Begin VB.Label labcodvende 
         Height          =   255
         Left            =   4080
         TabIndex        =   63
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label labcodmut 
         Height          =   375
         Left            =   7320
         TabIndex        =   62
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label labcodzon 
         Height          =   375
         Left            =   720
         TabIndex        =   61
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label labcodpromo 
         Height          =   255
         Left            =   5520
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label23 
         BackColor       =   &H00800000&
         Caption         =   "Promotor:"
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
         Left            =   3600
         TabIndex        =   48
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   11640
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label22 
         BackColor       =   &H00800000&
         Caption         =   "Valor de la cuota:"
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
         TabIndex        =   44
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H00800000&
         Caption         =   "Mutualista:"
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
         Left            =   6360
         TabIndex        =   42
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackColor       =   &H00800000&
         Caption         =   "Zona:"
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
         TabIndex        =   40
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H00800000&
         Caption         =   "Nombre de la casa:"
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
         Left            =   7080
         TabIndex        =   38
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00800000&
         Caption         =   "Manzana y Solar:"
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
         Left            =   7080
         TabIndex        =   35
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label17 
         BackColor       =   &H00800000&
         Caption         =   "Dirección - Entre:"
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
         TabIndex        =   34
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00800000&
         Caption         =   "Dirección - Calle:"
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackColor       =   &H00800000&
         Caption         =   "Correo electrónico:"
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
         Left            =   6600
         TabIndex        =   29
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00800000&
         Caption         =   "Celular:"
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
         Left            =   3720
         TabIndex        =   27
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H00800000&
         Caption         =   "Teléfono de línea:"
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
         TabIndex        =   25
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00800000&
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
         Height          =   375
         Left            =   7920
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00800000&
         Caption         =   "Fecha nacimiento:"
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
         Left            =   7920
         TabIndex        =   21
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00800000&
         Caption         =   "Seleccione promoción:"
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
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         Caption         =   "Seleccione Categoría:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800000&
         Caption         =   "Apellidos:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         Caption         =   "Nombres:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consultar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox t_codced 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox t_ced 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label labesmutual 
      Height          =   255
      Left            =   1680
      TabIndex        =   121
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labcatrealdes 
      Height          =   255
      Left            =   4560
      TabIndex        =   115
      Top             =   1800
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label labcatreal 
      Height          =   255
      Left            =   2520
      TabIndex        =   114
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label labpromocion 
      Height          =   255
      Left            =   2040
      TabIndex        =   109
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label labfact 
      Height          =   255
      Left            =   1800
      TabIndex        =   108
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label labcl_codigo 
      Height          =   255
      Left            =   6840
      TabIndex        =   98
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labcatnomsol 
      Height          =   255
      Left            =   5040
      TabIndex        =   90
      Top             =   1680
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label labcatcodsol 
      Height          =   255
      Left            =   4560
      TabIndex        =   89
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labauto 
      Height          =   255
      Left            =   6480
      TabIndex        =   88
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labnomconv 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   87
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label labcodconv 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   86
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C00000&
      Caption         =   "Convenio que figura en padrón:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   85
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label labnro 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Número de Afiliación:"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "FP09 R006 Ed_03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   0
      X2              =   11880
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingrese cédula:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   0
      X2              =   11880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
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
      Left            =   10080
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORMULARIO AFILIACIÓN SAPP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Fecha actual:"
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
      Left            =   8520
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frm_afilia.frx":37D7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frm_afilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()

cbopromo.Enabled = True
Borrar_campos
Command1.Enabled = True
Command2.Enabled = True
b_alta.Enabled = False
b_cance.Enabled = True
b_graba.Enabled = True
b_busca.Enabled = False
b_imp.Enabled = False
b_correo.Enabled = False
Command3.Visible = False
t_ced.SetFocus
labnro.Caption = data_nro.Recordset("p_afilia") + 1
data_nro.Recordset.Edit
data_nro.Recordset("p_afilia") = Val(labnro.Caption)
data_nro.Recordset.Update
labintegra.Caption = 1
Label3.Caption = Format(Date, "dd/mm/yyyy")
Frame3.Visible = True
cbocat.ListIndex = -1
labpromocion.Caption = "SI"
Verifica_matricula_existe
Check1.Value = 0
labesmutual.Caption = ""

End Sub

Private Sub b_busca_Click()
frm_afilbusca.Show vbModal

End Sub

Private Sub b_cance_Click()
Dim XCancela As String
On Error GoTo AfilCancela

XCancela = MsgBox("Desea cancelar toda la afiliación?", vbInformation + vbYesNo)
If XCancela = vbYes Then
   data_afilia.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(labnro.Caption) & " and cant_impre is null"
   data_afilia.Refresh
   If data_afilia.Recordset.RecordCount > 0 Then
      data_afilia.Recordset.MoveFirst
      Do While Not data_afilia.Recordset.EOF
         data_afilia.Recordset.Delete
         data_afilia.Recordset.MoveNext
      Loop
   End If
   Label39.Visible = False
   Borrar_campos
   b_alta.Enabled = True
   b_cance.Enabled = False
   b_graba.Enabled = False
   b_busca.Enabled = True
   b_imp.Enabled = True
'   b_correo.Enabled = True
   Frame1.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False
   labintegra.Caption = ""
   labnro.Caption = ""
   MsgBox "Afiliación cancelada.", vbInformation
Else
    Borrar_campos
    Label39.Visible = False
    b_alta.Enabled = True
    b_cance.Enabled = False
    b_graba.Enabled = False
    b_busca.Enabled = True
    b_imp.Enabled = True
'    b_correo.Enabled = True
    Frame1.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Frame2.Visible = False
    Frame3.Visible = False
    MsgBox "Afiliación cancelada.", vbInformation
End If

Exit Sub

AfilCancela:
            If Err.Number = 3081 Then
               MsgBox "ERROR:" & Err.Description
            Else
               MsgBox "ERROR:" & Err.Description
            End If
            
End Sub

Private Sub b_correo_Click()
Dim Xarchtex As String
Dim Xsiimpafil, ImprimeContra As String
Dim Correo As String

If labnro.Caption <> "" Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(labnro.Caption) & " and integra_nro in (1)"
   data_afilcons.Refresh
   If IsNull(data_afilcons.Recordset("correo")) = False Then
      Correo = data_afilcons.Recordset("correo")
        data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(labnro.Caption) & " order by integra_nro"
        data_afilcons.Refresh
        If data_afilcons.Recordset.RecordCount > 0 Then
           If IsNull(data_afilcons.Recordset("sifact")) = False Then
              data_afilcons.Recordset.MoveFirst
              Do While Not data_afilcons.Recordset.EOF
                 Genera_contrato
                 data_afilcons.Recordset.MoveNext
              Loop
              data_afilcons.Recordset.MovePrevious
              data_hist.RecordSource = "select * from afiliaciones_impre"
              data_hist.Refresh
              data_hist.Recordset.AddNew
              data_hist.Recordset("fecha") = Date
              data_hist.Recordset("hora") = Format(Time, "HH:mm")
              data_hist.Recordset("usuario") = WElusuario
              data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
              data_hist.Recordset("nro_aflia") = data_afilcons.Recordset("afilia_nro")
              data_hist.Recordset("accion") = "ENVIO POR CORREO"
              data_hist.Recordset.Update
              data_afilcons.Recordset.MoveNext
              data_inf.RecordSource = "select * from infcli order by cl_cantpag"
              data_inf.Refresh
              If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
                 If data_inf.Recordset("cl_nom_sup") = "Cobro por tarjeta" Then
                    cr1print.ReportFileName = App.path & "\contrato_deb.rpt"
                 Else
                    cr1print.ReportFileName = App.path & "\contrato_afil.rpt"
                 End If
              Else
                 cr1print.ReportFileName = App.path & "\contrato_afil.rpt"
              End If
              cr1print.Action = 1
              MsgBox "Se enviará el correo con el contrato adjunto", vbInformation
              If Dir("C:\planillas\contrato.pdf") <> "" Then
                 Name "C:\planillas\contrato.pdf" As "C:\planillas\contrato_Nro_" & labnro.Caption & ".pdf"
                 Xarchtex = "C:\planillas\contrato_Nro_" & labnro.Caption & ".pdf"
              
                 Dim MenCorreo As String
                 Dim oMail As Class1
                 Set oMail = New Class1
                     With oMail
                         .servidor = "smtp.gmail.com"
                         .puerto = 465
                         .UseAuntentificacion = True
                         .ssl = True
                         .Usuario = "jdanfer@gmail.com"
                         .PassWord = "PpasJfsh8719"
                         .Asunto = "Contrato SAPP"
                         .de = "jdanfer@gmail.com"
                         .para = Correo
                '         .para = "sappjorge@hotmail.com; despachosapp@hotmail.com; sappsusanadominguez@hotmail.com; sappdirecciontecnica@hotmail.com; sappenrique@hotmail.com"
                         .Adjunto = Xarchtex
                         .Mensaje = "Se adjunta contrato SAPP Nro. " & labnro.Caption
                         .Enviar_Backup ' manda el mail
                     End With
                     Set oMail = Nothing
                     MsgBox "Correo enviado.", vbInformation
                     
              Else
                 MsgBox "No se pudo generar el archivo correctamente, reintente nuevamente.", vbCritical
              End If
           
           Else
               MsgBox "No se puede enviar porque falta la facturación.", vbCritical
           End If
        Else
           MsgBox "No se encuentra afiliación."
        End If
   Else
        MsgBox "La afiliación no tiene correo ingresado el titular.", vbCritical
       
   End If
Else
   MsgBox "No seleccionó afiliación", vbExclamation
End If



End Sub

Private Sub b_graba_Click()
Dim VerificaDatos, XX, VerificaCorreo, Xsiestalac, Xnroafantes, VerificaTarjeta, VerificaCobrador, VerificaDigit, Sipuedegrabar, Haynoauto As Integer
Dim textocorreo, MasIntegrante, FacturaSi, EnviarCorreo, ImprimeContra As String
Dim TotalAfil As Double
Dim Nrodetarjeta As String
Nrodetarjeta = ""

On Error GoTo AfilGraba

labfact.Caption = ""

TotalAfil = 0
Sipuedegrabar = 0
Haynoauto = 0

VerificaDatos = 0
VerificaCorreo = 0
VerificaTarjeta = 0
VerificaCobrador = 0
VerificaDigit = 0

textocorreo = ""
FacturaSi = ""
EnviarCorreo = ""
ImprimeContra = ""

data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If


data_afilia.RecordSource = "afiliaciones_new"
data_afilia.Refresh

If Frame3.Visible = True Then
   If t_cobnro.Text = "" Then
      t_cobnro.Text = 0
      cbocobnom.Text = "*TODOS"
   End If
   If cbodiacob.ListIndex < 0 Then
      If cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "EXTENSION SUAT" Then
         cbodiacob.ListIndex = 0
      Else
         VerificaCobrador = 2
      End If
   End If
Else
   If Frame2.Visible = True Then
      If Trim(labcodsello.Caption) = "" Then
         VerificaTarjeta = 2
      End If
      If cbosello.ListIndex < 0 Then
         VerificaTarjeta = 2
      Else
         If Trim(t_nrotarj.Text) <> "" Then
            If Trim(t_nrotarj2.Text) <> "" Then
               If Trim(t_nrotarj3.Text) <> "" Then
                  If Trim(t_nrotarj4.Text) <> "" Then
                     Nrodetarjeta = Trim(t_nrotarj.Text) & Trim(t_nrotarj2.Text) & Trim(t_nrotarj3.Text) & Trim(t_nrotarj4.Text)
                  End If
               End If
            End If
         End If
         If cbosello.Text = "OCA CARD" Then
'VISA / MASTER CARD /CABAL /OCA CARD
            If Len(Trim(Nrodetarjeta)) <> 16 Then
               VerificaTarjeta = 2
               MsgBox "Verifique el número de la tarjeta OCA", vbCritical
            End If
         End If
         If cbosello.Text = "VISA" Then
            If Len(Trim(Nrodetarjeta)) <> 16 Then
               VerificaTarjeta = 2
               MsgBox "Verifique el número de la tarjeta VISA", vbCritical
            End If
         End If
         If cbosello.Text = "MASTER CARD" Then
            If Len(Trim(Nrodetarjeta)) <> 16 Then
               VerificaTarjeta = 2
               MsgBox "Verifique el número de la tarjeta MASTER CARD", vbCritical
            End If
         End If
         If cbosello.Text = "CABAL" Then
            If Len(Trim(Nrodetarjeta)) <> 16 Then
               VerificaTarjeta = 2
               MsgBox "Verifique el número de la tarjeta CABAL", vbCritical
            End If
         End If
         If cbosello.Text = "DEBITO BROU" Then
            If ch_debbrou.Value <> 1 Then
               VerificaTarjeta = 2
               MsgBox "Debe realizar formulario de débito online", vbCritical
            End If
         End If
      End If
      If Trim(t_nrotarj.Text) = "" Then
         If cbosello.Text <> "DEBITO BROU" Then
            VerificaTarjeta = 2
         End If
      End If
      If cbomesv.ListIndex < 0 Then
         If cbosello.Text <> "DEBITO BROU" Then
            VerificaTarjeta = 2
         End If
      End If
      If cboaniov.ListIndex < 0 Then
         If cbosello.Text <> "DEBITO BROU" Then
            VerificaTarjeta = 2
         End If
      End If
      If Trim(t_nomtit.Text) = "" Then
         VerificaTarjeta = 2
      End If
      If t_cedtit.Text <> "" And t_codcedtit.Text <> "" Then
         VerificaDigit = Verifica_CedTarj
         If VerificaDigit = 2 Then
            VerificaTarjeta = 2
            MsgBox "Hay un error en la cédula del titular de tarjeta, verifique!", vbCritical
         End If
      End If
      If Trim(t_cedtit.Text) = "" Then
         VerificaTarjeta = 2
      Else
         If IsNumeric(t_cedtit.Text) = False Then
            VerificaTarjeta = 2
         End If
         If t_cedtit.Text <= 99999 Then
            VerificaTarjeta = 2
         End If
      End If
      If Trim(t_codcedtit.Text) = "" Then
         VerificaTarjeta = 2
      End If
      If Trim(t_domitarj.Text) = "" Then
         VerificaTarjeta = 2
      End If
      If Trim(t_teleftarj.Text) = "" Then
         VerificaTarjeta = 2
      End If
   Else
      VerificaCobrador = 2
   End If
End If

If labnro.Caption = "" Then
   VerificaDatos = 1
   MsgBox "Falta dato de número de contrato", vbInformation
Else
   Xnroafantes = Val(labnro.Caption)
End If
If t_ced.Text = "" Then
   VerificaDatos = 1
   MsgBox "Falta dato de cédula"
End If
If t_codced.Text = "" Then
   VerificaDatos = 1
   MsgBox "Falta dato de DV"
End If

If cbocat.ListIndex < 0 Then
   VerificaDatos = 1
   MsgBox "Falta categoría"
End If
If t_nom1.Text = "" Then
   VerificaDatos = 1
   MsgBox "Falta ingresar nombre"
End If
If t_ape1.Text = "" Then
   VerificaDatos = 1
   MsgBox "Falta ingresar apellido"
End If
If mfnac.Text = "__/__/____" Then
   VerificaDatos = 1
   MsgBox "Falta ingresar fecha nacimiento."
End If
If labcatcodsol.Caption = "" Then
   VerificaDatos = 1
   MsgBox "Falta categoría de solicitud."
End If
If labcodmut.Caption = "" Then
   VerificaDatos = 1
   MsgBox "Falta ingresar mutualista."
End If

If labcodzon.Caption = "" Then
   VerificaDatos = 1
   MsgBox "Falta ingresar Zona"
End If

If cbosexo.ListIndex = -1 Then
   VerificaDatos = 1
   MsgBox "Falta ingresar sexo."
End If
If labintegra.Caption = "" Then
   VerificaDatos = 1
   MsgBox "Faltan integrantes."
End If
If t_telef.Text = "" Then
   t_telef.Text = "NO APLICA"
End If
If t_telef.Text <> "NO APLICA" Then
   If Len(t_telef.Text) <= 6 Then
      VerificaDatos = 1
      MsgBox "Falta teléfono."
   End If
   If IsNumeric(t_telef.Text) = False Then
      VerificaDatos = 1
      MsgBox "Falta teléfono."
   End If
End If
If t_celu.Text = "" Then
   t_celu.Text = "NO APLICA"
End If
If t_celu.Text <> "NO APLICA" Then
   If Len(t_celu.Text) <> 9 Then
      VerificaDatos = 1
      MsgBox "Falta celular."
   End If
   If IsNumeric(t_celu.Text) = False Then
      VerificaDatos = 1
      MsgBox "Falta celular."
   End If
End If
If t_correo.Text = "" Or t_correo.Text = "no aplica" Then
   t_correo.Text = "NO APLICA"
End If
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
      VerificaCorreo = 2
      MsgBox "Error en correo electrónico."
   End If
End If
If VerificaCorreo = 2 Then
   VerificaDatos = 1
End If
If t_calle.Text = "" Then
   VerificaDatos = 1
   MsgBox "Falta dirección."
End If
If Trim(cbozona.Text) = "" Then
   VerificaDatos = 1
   MsgBox "Falta zona."
End If
If Trim(cbomut.Text) = "" Then
   VerificaDatos = 1
   MsgBox "Falta mutualista."
End If
If t_valor.Text = "" Then
   VerificaDatos = 1
   MsgBox "Error en el importe."
End If
If labcodvende.Caption = "" Then
   VerificaDatos = 1
   MsgBox "Falta promotor."
End If
If Trim(cbovende.Text) = "" Then
   VerificaDatos = 1
   MsgBox "Falta promotor."
End If
If VerificaTarjeta = 2 Then
   VerificaDatos = 1
End If
If VerificaCobrador = 2 Then
   VerificaDatos = 1
End If

If labcodconv.Caption = "CCNOS" Or labcodconv.Caption = "SMIN" Or _
   labcodconv.Caption = "UNIVS" Or labcodconv.Caption = "HEVANO" Or _
   labcodconv.Caption = "GANOS" Or labcodconv.Caption = "CASANO" Then
   If cbocat.Text = "COMPLEMENTO" Or cbocat.Text = "COMPLEMENTO C.GALICIA" Then
      If labauto.Caption = "NOC" Then
         MsgBox "Recuerde: debe realizar carta para autorizar la afiliación.", vbCritical
      End If
   Else
      labauto.Caption = "SI"
   End If
End If
If cbocat.Text = "COMPLEMENTO" Or cbocat.Text = "COMPLEMENTO C.GALICIA" Then
   If Trim(labesmutual.Caption) <> "SI" Then
      VerificaDatos = 1
   End If
End If

Xsiestalac = Existe_afiliado()

If Xsiestalac = 1 Then
   VerificaDatos = 1
   MsgBox "La cédula ya figura con una afiliación realizada", vbCritical
End If

If VerificaDatos = 1 Then
   If VerificaCorreo = 2 Then
      MsgBox "Hay un error en los datos del correo electrónico", vbInformation
   Else
      MsgBox "Hay error en el ingreso de los datos, verifique!", vbCritical
      t_nom1.SetFocus
   End If
Else
   data_afilia.Recordset.AddNew
   data_afilia.Recordset("wusuario") = WElusuario
   data_afilia.Recordset("wbase") = frm_menu.data_parse.Recordset("base")
   data_afilia.Recordset("afilia_nro") = Val(labnro.Caption)
   data_afilia.Recordset("fecha") = Format(Label3.Caption, "yyyy-mm-dd")
   data_afilia.Recordset("integra_nro") = Val(labintegra.Caption)
   data_afilia.Recordset("cedula") = t_ced.Text & t_codced.Text
   data_afilia.Recordset("convenio") = cbocat.Text
   data_afilia.Recordset("catcontrato") = Devuelve_catContrato()
   data_afilia.Recordset("categ") = labcatcodsol.Caption
   data_afilia.Recordset("nomcateg") = labcatnomsol.Caption
   data_afilia.Recordset("tipoced") = Check1.Value
   If Trim(labcatreal.Caption) <> "" Then
      data_afilia.Recordset("catreal") = labcatreal.Caption
      If Trim(labcatrealdes.Caption) <> "" Then
         data_afilia.Recordset("catrealdes") = labcatrealdes.Caption
      End If
   End If
   
   If labcodpromo.Caption <> "" Then
      data_afilia.Recordset("codpromo") = Val(labcodpromo.Caption)
   End If
   data_afilia.Recordset("nom1") = t_nom1.Text
   If t_nom2.Text <> "" Then
      data_afilia.Recordset("nom2") = t_nom2.Text
   End If
   data_afilia.Recordset("ape1") = t_ape1.Text
   If t_ape2.Text <> "" Then
      data_afilia.Recordset("ape2") = t_ape2.Text
   End If
   data_afilia.Recordset("fnac") = mfnac.Text
   If cbosexo.Text = "MASCULINO" Then
      data_afilia.Recordset("sexo") = 1
   Else
      data_afilia.Recordset("sexo") = 2
   End If
   data_afilia.Recordset("telef") = t_telef.Text
   data_afilia.Recordset("celular") = t_celu.Text
   data_afilia.Recordset("correo") = t_correo.Text
   data_afilia.Recordset("direc1") = t_calle.Text
   If t_entre.Text <> "" Then
      data_afilia.Recordset("direc2") = t_entre.Text
   End If
   If t_manz.Text <> "" Then
      data_afilia.Recordset("manz") = t_manz.Text
   End If
   If t_sol.Text <> "" Then
      data_afilia.Recordset("solar") = t_sol.Text
   End If
   If t_casa.Text <> "" Then
      data_afilia.Recordset("casa") = t_casa.Text
   End If
   data_afilia.Recordset("codzon") = Val(labcodzon.Caption)
   data_afilia.Recordset("nomzona") = cbozona.Text
   data_afilia.Recordset("codmut") = Val(labcodmut.Caption)
   data_afilia.Recordset("valorcuota") = Val(t_valor.Text)
   If Trim(labtotafil.Caption) <> "" Then
      data_afilia.Recordset("importe_fin") = Val(labtotafil.Caption)
      data_afilia.Recordset("desc_porce") = Val(labdapagar.Caption)
      data_afilia.Recordset("desc_imp") = Val(t_valor.Text) - Int(labtotafil.Caption)
   Else
      data_afilia.Recordset("importe_fin") = t_valor.Text
      data_afilia.Recordset("desc_porce") = 0
      data_afilia.Recordset("desc_imp") = 0
   End If
   data_afilia.Recordset("codvende") = Val(labcodvende.Caption)
   If Frame3.Visible = True Then
      If Trim(t_cobnro.Text) = "" Then
         t_cobnro.Text = 0
      End If
      data_afilia.Recordset("codcob") = Val(t_cobnro.Text)
      data_afilia.Recordset("dia_cobro") = Val(cbodiacob.Text)
      data_afilia.Recordset("direc_cobro") = t_dircobro.Text
      If t_casacobro.Text <> "" Then
         data_afilia.Recordset("casanomcob") = t_casacobro.Text
      End If
      data_afilia.Recordset("codzoncob") = Val(labcodzoncobr.Caption)
      data_afilia.Recordset("zonacobro") = cbozonacobro.Text
      If mpd.Text <> "__/__/____" Then
         data_afilia.Recordset("fec_desde") = mpd.Text
      End If
      If mph.Text <> "__/__/____" Then
         data_afilia.Recordset("fec_hasta") = mph.Text
      End If
   Else
      If Frame2.Visible = True Then
         If cbosello.Text = "OCA CARD" Then
            data_afilia.Recordset("codcob") = 690
         End If
         If cbosello.Text = "VISA" Then
            data_afilia.Recordset("codcob") = 514
         End If
         If cbosello.Text = "MASTER CARD" Then
            data_afilia.Recordset("codcob") = 683
         End If
         If cbosello.Text = "CABAL" Then
            data_afilia.Recordset("codcob") = 673
         End If
         If cbosello.Text = "DEBITO BROU" Then
            data_afilia.Recordset("codcob") = 607
            data_afilia.Recordset("debito_brou") = ch_debbrou.Value
         End If
         data_afilia.Recordset("tarj_sello") = cbosello.Text
         data_afilia.Recordset("tarj_codsello") = Val(labcodsello.Caption)
         data_afilia.Recordset("tarj_nro") = Trim(Nrodetarjeta)
         data_afilia.Recordset("tarj_vencmes") = Val(cbomesv.Text)
         data_afilia.Recordset("tarj_vencanio") = Val(cboaniov.Text)
         data_afilia.Recordset("tarj_titular") = UCase(t_nomtit.Text)
         data_afilia.Recordset("tarj_cedtit") = t_cedtit.Text
         data_afilia.Recordset("tarj_codced") = t_codcedtit.Text
         data_afilia.Recordset("tarj_domic") = t_domitarj.Text
         data_afilia.Recordset("tarj_telef") = t_teleftarj.Text
      Else
         data_afilia.Recordset("codcob") = 0
      End If
   End If
   If labauto.Caption = "SI" Then
      If cbocat.Text = "COMPLEMENTO" Or cbocat.Text = "COMPLEMENTO C.GALICIA" Then
         data_afilia.Recordset("pendiente") = 3
         If labcl_codigo.Caption <> "" Then
            data_afilia.Recordset("matricula") = Val(labcl_codigo.Caption)
         End If
      Else
         data_afilia.Recordset("pendiente") = 0
      End If
   Else
      If labauto.Caption = "NO" Then
         MsgBox "ATENCION! Este registro no se cargará al padrón hasta que sea autorizado por administración.", vbCritical
         MsgBox "La afiliación quedará en suspenso hasta su autorización.", vbExclamation
         data_afilia.Recordset("pendiente") = 2
         If labcl_codigo.Caption <> "" Then
            data_afilia.Recordset("matricula") = Val(labcl_codigo.Caption)
            data_afilia.Recordset("obs_noaut") = "NO AUTORIZADO DEUDA/NOSAPP"
         End If
      Else
         If labauto.Caption = "NOC" Then
            MsgBox "ATENCION! Este registro no se cargará al padrón hasta que sea realizada la carta mutual.", vbCritical
            MsgBox "La afiliación quedará en suspenso hasta su autorización por padrón social.", vbExclamation
            data_afilia.Recordset("pendiente") = 2
            If labcl_codigo.Caption <> "" Then
               data_afilia.Recordset("matricula") = Val(labcl_codigo.Caption)
               data_afilia.Recordset("obs_noaut") = "NO AUTORIZADO-FALTA CARTA"
            End If
         Else
            data_afilia.Recordset("pendiente") = 0
         End If
      End If
   End If
   If cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "CONVENIO DE VERANO" Then
   Else
      If cbomut.Text = "ASSE" Then
         If Frame2.Visible = True And cbosello.ListIndex >= 0 Then
         Else
            MsgBox "Socio con mutual ASSE sin débito automático de tarjeta, pasará para autorización.", vbCritical
            data_afilia.Recordset("pendiente") = 2
         End If
      End If
   End If
   If labcodconv.Caption = "INCUMP" Then
      MsgBox "Socio con categoría no habilitada, quedará pendiente para autorización.", vbCritical
      data_afilia.Recordset("pendiente") = 2
   End If
   
   data_afilia.Recordset("promosi") = labpromocion.Caption
   If labcl_codigo.Caption <> "" Then
      data_afilia.Recordset("matricula") = Val(labcl_codigo.Caption)
   End If
   data_afilia.Recordset.Update
   If cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "EXTENSION SUAT" Then
      MsgBox "Un integrante por afiliación, se grabará este integrante.", vbInformation
      MasIntegrante = vbNo
      labantcat.Caption = ""
   Else
      MasIntegrante = MsgBox("Desea agregar un nuevo integrante a esta afiliación?", vbInformation + vbYesNo, "Afiliaciones")
   End If
   If MasIntegrante = vbYes Then
      labantcat.Caption = cbocat.Text
      cbopromo.Enabled = False
      Borrar_camposDos
      labnro.Caption = Xnroafantes
      If Frame3.Visible = True Then
         Frame3.Enabled = False
      End If
      If Frame2.Visible = True Then
         Frame2.Enabled = False
      End If
      labintegra.Caption = Val(labintegra.Caption) + 1
      t_ced.SetFocus
      data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(labnro.Caption)
      data_afilcons.Refresh
      If data_afilcons.Recordset.RecordCount > 0 Then
         data_afilcons.Recordset.MoveFirst
         Do While Not data_afilcons.Recordset.EOF
            TotalAfil = TotalAfil + data_afilcons.Recordset("importe_fin")
            labapagar.Caption = "TOTAL a Pagar:" & vbCrLf & Format(TotalAfil, "Standard")
            data_afilcons.Recordset.MoveNext
         Loop
      End If
      labcl_codigo.Caption = ""
   Else
      If cbocat.Text = "Grupo de 3 o más" Then
         If Val(labintegra.Caption) < 3 Then
            Sipuedegrabar = 1
         Else
            Sipuedegrabar = 0
         End If
      Else
         Sipuedegrabar = 0
      End If
     If Sipuedegrabar = 0 Then
        labantcat.Caption = ""
        data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(labnro.Caption) & " order by integra_nro"
        data_afilcons.Refresh
        If data_afilcons.Recordset.RecordCount > 0 Then
           data_afilcons.Recordset.MoveFirst
           Do While Not data_afilcons.Recordset.EOF
              If IsNull(data_afilcons.Recordset("pendiente")) = False Then
                 If data_afilcons.Recordset("pendiente") = 2 Then
                    Haynoauto = 1
                 End If
              End If
              TotalAfil = TotalAfil + data_afilcons.Recordset("importe_fin")
              labapagar.Caption = "TOTAL a Pagar:" & vbCrLf & Format(TotalAfil, "Standard")
              data_afilcons.Recordset.MoveNext
           Loop
           If Haynoauto = 1 Then
              data_afilcons.Recordset.MoveFirst
              Do While Not data_afilcons.Recordset.EOF
                 If IsNull(data_afilcons.Recordset("pendiente")) = False Then
                    If data_afilcons.Recordset("pendiente") <> 2 Then
                       data_afilcons.Recordset.Edit
                       data_afilcons.Recordset("pendiente") = 2
                       data_afilcons.Recordset.Update
                    End If
                 End If
                 data_afilcons.Recordset.MoveNext
              Loop
           End If
           data_afilcons.Recordset.MoveFirst
           If Haynoauto <> 1 And TotalAfil > 0 Then
              MsgBox "Verifique POS y si está preparada la impresora de e-ticket", vbInformation
              FacturaSi = MsgBox("Confirma realizar el e-ticket por $." & Format(TotalAfil, "Standard") & " ?", vbInformation + vbYesNo, "Afiliación")
              If FacturaSi = vbYes Then
                 Do While Not data_afilcons.Recordset.EOF
                    If data_afilcons.Recordset("pendiente") = 3 Or data_afilcons.Recordset("pendiente") = 2 Then
                    Else
                       Alta_Modif
                    End If
                    data_afilcons.Recordset.MoveNext
                 Loop
                 data_afilcons.Recordset.MovePrevious
                 frm_afilfactura.Show vbModal
                 
                 data_hist.RecordSource = "select * from afiliaciones_impre"
                 data_hist.Refresh
                 data_hist.Recordset.AddNew
                 data_hist.Recordset("fecha") = Date
                 data_hist.Recordset("hora") = Format(Time, "HH:mm")
                 data_hist.Recordset("usuario") = WElusuario
                 data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
                 data_hist.Recordset("nro_afilia") = data_afilcons.Recordset("afilia_nro")
                 data_hist.Recordset("accion") = "CREA AFILIACION"
                 data_hist.Recordset.Update
                 data_afilcons.Recordset.MoveNext
                 ImprimeContra = MsgBox("Desea imprimir contrato?", vbInformation + vbYesNo, "Afiliaciones SAPP")
                 If ImprimeContra = vbYes Then
                    data_afilcons.Recordset.MoveFirst
                    If IsNull(data_afilcons.Recordset("sifact")) = False Then
                        Do While Not data_afilcons.Recordset.EOF
                           Genera_contrato
                           data_afilcons.Recordset.Edit
                           If IsNull(data_afilcons.Recordset("cant_impre")) = False Then
                              data_afilcons.Recordset("cant_impre") = data_afilcons.Recordset("cant_impre") + 1
                           Else
                              data_afilcons.Recordset("cant_impre") = 1
                           End If
                           data_afilcons.Recordset.Update
                           data_afilcons.Recordset.MoveNext
                        Loop
                        data_afilcons.Recordset.MovePrevious
                        data_hist.RecordSource = "select * from afiliaciones_impre"
                        data_hist.Refresh
                        data_hist.Recordset.AddNew
                        data_hist.Recordset("fecha") = Date
                        data_hist.Recordset("hora") = Format(Time, "HH:mm")
                        data_hist.Recordset("usuario") = WElusuario
                        data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
                        data_hist.Recordset("nro_afilia") = data_afilcons.Recordset("afilia_nro")
                        data_hist.Recordset("accion") = "IMPRESION"
                        data_hist.Recordset.Update
                        data_afilcons.Recordset.MoveNext
                        data_inf.RecordSource = "select * from infcli order by cl_cantpag"
                        data_inf.Refresh
                        If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
                           If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
                              cr1print.ReportFileName = App.path & "\contrato_debprint.rpt"
                           Else
                              cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
                           End If
                        Else
                           cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
                        End If
                        cr1print.Action = 1
                        b_imp.Enabled = False
                    Else
                        MsgBox "No está realizada la factura, no se puede imprimir el contrato", vbCritical
                        b_imp.Enabled = False
                    End If
                 Else
                    data_afilcons.Recordset.MoveFirst
                    Do While Not data_afilcons.Recordset.EOF
                       Genera_contrato
                       data_afilcons.Recordset.MoveNext
                    Loop
                    data_inf.RecordSource = "select * from infcli order by cl_cantpag"
                    data_inf.Refresh
                    If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
                       If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
                          cr2pant.ReportFileName = App.path & "\contrato_debprint.rpt"
                       Else
                          cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
                       End If
                    Else
                       cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
                    End If
                    cr2pant.Action = 1
                    b_imp.Enabled = False
                 End If
              Else
                 MsgBox "La afiliación no será cargada al padrón hasta ser facturada.", vbCritical
              End If
           Else
              If Val(TotalAfil) = 0 And cbocat.Text = "CAMBIO DE CATEGORIA" Then
                 Do While Not data_afilcons.Recordset.EOF
                    data_afilcons.Recordset.Edit
                    data_afilcons.Recordset("sifact") = 1
                    data_afilcons.Recordset.Update
                    Alta_Modif
                    data_afilcons.Recordset.MoveNext
                 Loop
                 data_afilcons.Recordset.MoveFirst
              Else
                 MsgBox "Cuando la Afiliación sea autorizada por administración, deberá realizar la facturación.", vbCritical
                 b_imp.Enabled = False
              End If
           End If
           Label39.Visible = False
           b_alta.Enabled = True
           b_cance.Enabled = False
           b_graba.Enabled = False
           b_busca.Enabled = True
'           b_imp.Enabled = True
           b_correo.Enabled = True
           Check1.Value = 0
           labtotafil.Caption = ""
           Label38.Visible = False
           labcl_codigo.Caption = ""
           Frame2.Enabled = True
           Frame3.Enabled = True
           Frame2.Visible = False
           Frame3.Visible = False
           Frame1.Enabled = False
           Command1.Enabled = False
           Command2.Enabled = False
           labintegra.Caption = ""
           labnro.Caption = ""
           labauto.Caption = "SI"
           labdescpromo.Caption = ""
           labdapagar.Caption = ""
           labapagar.Caption = ""
           labpromocion.Caption = "SI"
        Else
           MsgBox "Si no realiza la factura, no se cargará la afiliación al padrón.", vbCritical
           b_imp.Enabled = False
        End If
     Else
        MsgBox "La promoción seleccionada debe tener cómo mínimo 3 integrantes en la afiliación.", vbCritical
        b_imp.Enabled = False
     End If
   End If
End If

Exit Sub

AfilGraba:
          If Err.Number = 3081 Then
             MsgBox "ERROR: " & Err.Description
          Else
             MsgBox "ERROR: " & Err.Description
          End If


End Sub

Private Sub b_imp_Click()
Dim Xsiimpafil, ImprimeContra As String
data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

If labnro.Caption <> "" Then
   Xsiimpafil = MsgBox("Desea imprimir el contrato seleccionado?", vbInformation + vbYesNo, "Afiliaciones")
   If Xsiimpafil = vbYes Then
        data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(labnro.Caption) & " order by integra_nro"
        data_afilcons.Refresh
        If data_afilcons.Recordset.RecordCount > 0 Then
           ImprimeContra = MsgBox("Desea imprimir contrato?", vbInformation + vbYesNo, "Afiliaciones SAPP")
           If ImprimeContra = vbYes Then
              data_afilcons.Recordset.MoveFirst
              Do While Not data_afilcons.Recordset.EOF
                 Genera_contrato
                 data_afilcons.Recordset.Edit
                 If IsNull(data_afilcons.Recordset("cant_impre")) = False Then
                    data_afilcons.Recordset("cant_impre") = data_afilcons.Recordset("cant_impre") + 1
                 Else
                    data_afilcons.Recordset("cant_impre") = 1
                 End If
                 data_afilcons.Recordset.Update
                 data_afilcons.Recordset.MoveNext
              Loop
              data_afilcons.Recordset.MovePrevious
              data_hist.RecordSource = "select * from afiliaciones_impre"
              data_hist.Refresh
              data_hist.Recordset.AddNew
              data_hist.Recordset("fecha") = Date
              data_hist.Recordset("hora") = Format(Time, "HH:mm")
              data_hist.Recordset("usuario") = WElusuario
              data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
              data_hist.Recordset("nro_afilia") = data_afilcons.Recordset("afilia_nro")
              data_hist.Recordset("accion") = "IMPRESION"
              data_hist.Recordset.Update
              data_afilcons.Recordset.MoveNext
              data_inf.RecordSource = "select * from infcli order by cl_cantpag"
              data_inf.Refresh
              If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
                 If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
                    cr1print.ReportFileName = App.path & "\contrato_debprint.rpt"
                 Else
                    cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
                 End If
              Else
                 cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
              End If
              cr1print.Action = 1
           Else
              data_afilcons.Recordset.MoveFirst
              Do While Not data_afilcons.Recordset.EOF
                 Genera_contrato
                 data_afilcons.Recordset.MoveNext
              Loop
              data_inf.RecordSource = "select * from infcli order by cl_cantpag"
              data_inf.Refresh
              If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
                 If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
                    cr2pant.ReportFileName = App.path & "\contrato_debprint.rpt"
                 Else
                    cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
                 End If
              Else
                 cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
              End If
              cr2pant.Action = 1
           End If
        Else
           MsgBox "No se encuentran datos para este contrato"
        End If

   End If
Else
   MsgBox "No seleccionó afiliación", vbExclamation
   
End If

End Sub

Private Sub cboaniov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nomtit.SetFocus
End If

End Sub

Private Sub cbocat_Click()
If cbocat.ListIndex >= 0 Then
   If cbocat.Text = "CAMBIO DE CATEGORIA" Then
      afilia_cambio.Show vbModal
   Else
      Consulta_precio
      If Trim(cbopromo.Text) <> "" Then
         Consulta_promos
      End If
   End If
Else
   t_valor.Text = ""

End If
cbopromo.Enabled = True

If Val(labintegra.Caption) > 1 Then
   If labantcat.Caption = "AMBULATORIO" Or labantcat.Caption = "COMPLEMENTO" Or labantcat.Caption = "C.GALICIA" Then
      If cbocat.Text = "EMERGENCIA" Or cbocat.Text = "PARCIAL" Or cbocat.Text = "TALA" Or cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "EXTENSION SUAT" Then
         MsgBox "La categoría debe ser AMBULATORIO", vbCritical
         cbocat.SetFocus
      End If
   Else
      If labantcat.Caption = "EMERGENCIA" Then
         If cbocat.Text = "AMBULATORIO" Or cbocat.Text = "COMPLEMENTO" Or cbocat.Text = "COMPLEMENTO C.GALICIA" Or cbocat.Text = "PARCIAL" Or cbocat.Text = "TALA" Or cbocat.Text = "TALA" Or cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "EXTENSION SUAT" Then
            MsgBox "La categoría debe ser EMERGENCIA", vbCritical
            cbocat.SetFocus
         End If
      Else
         If labantcat.Caption = "PARCIAL" Then
            If cbocat.Text = "AMBULATORIO" Or cbocat.Text = "COMPLEMENTO" Or cbocat.Text = "COMPLEMENTO C.GALICIA" Or cbocat.Text = "EMERGENCIA" Or cbocat.Text = "TALA" Or cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "EXTENSION SUAT" Then
               MsgBox "La categoría debe ser PARCIAL", vbCritical
               cbocat.SetFocus
            End If
         Else
            If labantcat.Caption = "TALA" Then
               If cbocat.Text = "AMBULATORIO" Or cbocat.Text = "COMPLEMENTO" Or cbocat.Text = "COMPLEMENTO C.GALICIA" Or cbocat.Text = "EMERGENCIA" Or cbocat.Text = "PARCIAL" Or cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "EXTENSION SUAT" Then
                  MsgBox "La categoría debe ser TALA", vbCritical
                  cbocat.SetFocus
               End If
            End If
         End If
      End If
   End If
      
End If
If cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "CONVENIO DE VERANO" Then
   Frame2.Visible = False
   Frame3.Visible = True
   Label39.Caption = "Haga doble click aquí para cambiar a pago con tarjeta"
   Veo_plazos
   mpd.Text = Format(Date, "dd/mm/yyyy")
   mph.Text = Format(Date + 30, "dd/mm/yyyy")
   cbopromo.ListIndex = -1
   cbopromo.Enabled = False
Else
   If Frame2.Visible = True Then
   Else
      Frame3.Visible = True
   End If
   Label39.Caption = "Haga doble click aquí para cambiar a pago con tarjeta"
   NoVeo_plazos
End If


End Sub

Private Sub cbocat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cbopromo.Enabled = True Then
      cbopromo.SetFocus
   End If
End If

End Sub

Private Sub cbocat_LostFocus()
If cbocat.ListIndex < 0 Then
   MsgBox "Seleccione Categoría para la afiliación.", vbCritical
   cbocat.SetFocus
End If

End Sub

Private Sub cbocobnom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbodiacob.SetFocus
End If

End Sub

Private Sub cbocobnom_LostFocus()
If cbocobnom.Text <> "" Then
   Buscar_cobrador_nombre
End If

End Sub

Private Sub cbomesv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboaniov.SetFocus
End If

End Sub

Private Sub cbomut_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   cbovende.SetFocus
End If

End Sub

Private Sub cbomut_LostFocus()
If cbomut.Text <> "" Then
   Consulta_mutual
Else
   labcodmut.Caption = ""
End If

End Sub

Private Sub cbopromo_Click()
If Trim(cbopromo.Text) <> "" Then
   Consulta_promos
Else
   labcodpromo.Caption = ""
   labdapagar.Caption = ""
   labdescpromo.Caption = ""
   cbopromo.Text = ""
   labapagar.Caption = ""
   Label38.Visible = False
   labtotafil.Caption = ""
End If

If cbopromo.Text = "Tarjeta de crédito" Then
   Frame2.Visible = True
   Frame3.Visible = False
   If cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SUAT" Then
      MsgBox "Este tipo de contrato no puede llevar la opción de tarjeta. VERIFIQUE!", vbCritical
   End If
   Label39.Caption = "Haga doble click aquí para cambiar a cobrar a domicilio."
Else
   Frame2.Visible = False
   Frame3.Visible = True
   If cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "CONVENIO DE VERANO" Or cbocat.Text = "EXTENSION SUAT" Then
      Veo_plazos
   Else
      NoVeo_plazos
   End If
   Label39.Caption = "Haga doble click aquí para cambiar a pago con tarjeta"

End If

End Sub

Private Sub cbopromo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom1.SetFocus
End If

End Sub

Private Sub cbopromo_LostFocus()
If Trim(cbopromo.Text) <> "" Then
Else
   labcodpromo.Caption = ""
   labdapagar.Caption = ""
   labdescpromo.Caption = ""
   cbopromo.Text = ""
   labapagar.Caption = ""
   Label38.Visible = False
   labtotafil.Caption = ""
   If Frame3.Visible = False Then
      Frame3.Visible = True
      Frame2.Visible = False
   End If
End If


End Sub

Private Sub cbosello_Click()
If cbosello.ListIndex >= 0 Then
   If cbosello.Text = "DEBITO BROU" Then
      Debito_brou
      labcodsello.Caption = 5
   Else
      NoDebito_brou
      If cbosello.Text = "OCA CARD" Then
         labcodsello.Caption = 9
      End If
      If cbosello.Text = "VISA" Then
         labcodsello.Caption = 1
      End If
      If cbosello.Text = "MASTER CARD" Then
         labcodsello.Caption = 7
      End If
      If cbosello.Text = "CABAL" Then
         labcodsello.Caption = 4
      End If
   End If
Else
   labcodsello.Caption = ""
End If

End Sub

Private Sub cbosello_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nrotarj.SetFocus
End If

End Sub

Private Sub cbosexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_telef.SetFocus
End If

End Sub

Private Sub cbovende_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomut.SetFocus
End If

End Sub

Private Sub cbovende_LostFocus()
If cbovende.Text <> "" Then
   Consulta_vendedor
Else
   labcodvende.Caption = ""
End If

End Sub

Private Sub cbozona_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   cbomut.SetFocus
End If

End Sub

Private Sub cbozona_LostFocus()

If cbozona.Text <> "" Then
   Consulta_zonas (cbozona.Text)
Else
   labcodzon.Caption = ""
End If
If cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "CONVENIO DE VERANO" Then
Else
    If labcodzon.Caption <> "" Then
       If Val(labcodzon.Caption) <> 999 Then
          If Frame3.Visible = True Then
             If t_cobnro.Visible = True Then
                If t_cobnro.Text <> "" Then
                   If Val(t_cobnro.Text) > 0 Then
                      Buscar_cobrador
                   End If
                End If
             End If
          End If
       End If
    End If
End If

End Sub

Private Sub cbozonacobro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

End Sub

Private Sub cbozonacobro_LostFocus()
If cbozonacobro.Text <> "" Then
   Consulta_zonasCob (cbozonacobro.Text)
Else
   labcodzoncobr.Caption = ""
End If

End Sub

Private Sub Check1_Click()
Dim Tieneautoriza As String

If Check1.Value = 1 Then
   Tieneautoriza = MsgBox("Para ingresar un documento extranjero deberá seleccionar el PAIS del documento. Desea continuar?", vbCritical + vbYesNo)
   If Tieneautoriza = vbYes Then
'      t_codced.Visible = False
'      t_codced.Text = 0
      Check1.Value = 0
      t_codced.Text = ""
      t_codced.Visible = True
   Else
      Check1.Value = 0
      t_codced.Text = ""
      t_codced.Visible = True
   End If
Else
   t_codced.Text = ""
   t_codced.Visible = True
End If

End Sub

Private Sub Command1_Click()
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo, Xantnro As Long


Xn1 = 2
Xn2 = 9
Xn3 = 8
Xn4 = 7
Xn5 = 6
Xn6 = 3
Xn7 = 4
Xpond = 10
If Trim(labnro.Caption) <> "" Then
   Xantnro = Val(labnro.Caption)
Else
   Xantnro = 0
End If
If Check1.Value = 1 Then
   t_codced.Text = 0
End If

If t_ced.Text <> "" And t_codced.Text <> "" Then
   If IsNumeric(t_ced.Text) = False Then
      MsgBox "La cédula debe contener solo números", vbInformation
      t_ced.Text = ""
   Else
      Xcedtex = Trim(str(t_ced.Text))
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
      If Check1.Value = 1 Then
         Frame1.Enabled = True
         cbocat.SetFocus
         t_codced.Text = 0
      Else
        If Xtot <> t_codced.Text Then
           MsgBox "Hay un error en la cédula, verifique!", vbCritical
           t_nom1.Text = ""
           t_nom2.Text = ""
           t_ape1.Text = ""
           t_ape2.Text = ""
           If Xantnro <> 0 Then
              labnro.Caption = Xantnro
           End If
           Frame1.Enabled = False
        Else
           Consulta_Cli
'''           Consultar_Yaexiste
           Label39.Visible = True
           If Trim(labcodconv.Caption) <> "" Then
              If Trim(labcodconv.Caption) = "APS" Then
                 MsgBox "ATENCION!! Socio en una categoría no habilitada para afiliar! Consulte con Administración.", vbCritical
                 b_cance_Click
              End If
           End If
           
        End If
      End If
   End If
Else
'    MsgBox "Debe ingresar cédula para poder grabar la afiliación", vbInformation
   labcl_codigo.Caption = ""
   
End If
End Sub


Private Sub Command2_Click()
frm_afilbuscli.Show vbModal

End Sub




Private Sub Command3_Click()
'On Error GoTo Nograbanada
data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(labnro.Caption) & " and integra_nro =" & Val(labintegra.Caption)
data_afilcons.Refresh
If data_afilcons.Recordset.RecordCount > 0 Then
   data_afilcons.Recordset.Edit
   data_afilcons.Recordset("nom1") = t_nom1.Text
   If t_nom2.Text <> "" Then
      data_afilcons.Recordset("nom2") = t_nom2.Text
   End If
   data_afilcons.Recordset("ape1") = t_ape1.Text
   If t_ape2.Text <> "" Then
      data_afilcons.Recordset("ape2") = t_ape2.Text
   End If
   data_afilcons.Recordset("fnac") = mfnac.Text
   If cbosexo.Text = "MASCULINO" Then
      data_afilcons.Recordset("sexo") = 1
   Else
      data_afilcons.Recordset("sexo") = 2
   End If
   data_afilcons.Recordset("telef") = t_telef.Text
   data_afilcons.Recordset("celular") = t_celu.Text
   data_afilcons.Recordset("correo") = t_correo.Text
   data_afilcons.Recordset("direc1") = t_calle.Text
   If t_entre.Text <> "" Then
      data_afilcons.Recordset("direc2") = t_entre.Text
   End If
   If t_manz.Text <> "" Then
      data_afilcons.Recordset("manz") = t_manz.Text
   End If
   If t_sol.Text <> "" Then
      data_afilcons.Recordset("solar") = t_sol.Text
   End If
   If t_casa.Text <> "" Then
      data_afilcons.Recordset("casa") = t_casa.Text
   End If
   If cbomut.Text <> "" Then
      data_afilcons.Recordset("codmut") = Val(labcodmut.Caption)
   End If
   data_afilcons.Recordset("codzon") = Val(labcodzon.Caption)
   data_afilcons.Recordset("nomzona") = cbozona.Text
   
   data_afilcons.Recordset.Update
   Modif_SocioAfil
   MsgBox "Modificado."
   Borrar_campos
End If

Frame1.Enabled = False
Command3.Visible = False

'Exit Sub
'Nograbanada:
'            If Err.Number = 3155 Then
'               MsgBox "No hay datos para grabar"
'               Unload Me
'            Else
'               MsgBox "No hay datos para grabar"
'               Unload Me
'            End If

End Sub

Private Sub Form_Load()
Dim XXa, XXaa As Integer

data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_afilia.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_afilcons.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hist.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_abm.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_nro.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_nro.RecordSource = "Select * from param_gral"
data_nro.Refresh

data_nrosoc.DatabaseName = App.path & "\parse.mdb"
data_nrosoc.RecordSource = "parsec0"
data_nrosoc.Refresh

'data_nro.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_nro.RecordSource = "param_gral"
'data_nro.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"

Carga_catego
Carga_promos
Carga_zonas
Carga_mutua
Carga_vendedores
Carga_cobradors
Carga_zonasCob

cbomesv.Clear
For XXa = 1 To 12
    cbomesv.AddItem Trim(str(XXa))
Next
For XXa = 1 To 31
    cbodiacob.AddItem Trim(str(XXa))
Next

XXa = Year(Date) + 10

For XXaa = Year(Date) To XXa
    cboaniov.AddItem Trim(str(XXaa))
Next

End Sub

Private Sub mfeccobro_Change()

End Sub

Private Sub Frame2_DblClick()
Frame3.Visible = True
Frame2.Visible = False

End Sub

Private Sub Frame3_DblClick()
Frame2.Visible = True
Frame3.Visible = False

End Sub

Private Sub Label39_DblClick()
If Frame2.Visible = True Then
   Frame2.Visible = False
   cbomesv.ListIndex = -1
   cboaniov.ListIndex = -1
   cbosello.ListIndex = -1
   t_nomtit.Text = ""
   t_domitarj.Text = ""
   t_teleftarj.Text = ""
   t_cedtit.Text = ""
   t_codcedtit.Text = ""
   ch_debbrou.Value = 0
   t_nrotarj.Text = ""
   Frame3.Visible = True
   Label39.Caption = "Haga doble click aquí para cambiar a pago con tarjeta"

Else
   Frame2.Visible = True
   Frame3.Visible = False
   t_cobnro.Text = ""
   cbocobnom.Text = ""
   cbodiacob.ListIndex = -1
   t_dircobro.Text = ""
   t_casacobro.Text = ""
   cbozonacobro.Text = ""
   mpd.Text = "__/__/____"
   mph.Text = "__/__/____"
   Label39.Caption = "Haga doble click aquí para cambiar a cobrar a domicilio."
   
End If
End Sub

Private Sub mfnac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   cbosexo.SetFocus
End If

End Sub

Private Sub t_ape1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   KeyAscii = 0
   t_ape1.Text = Trim(t_ape1.Text)
   t_ape2.SetFocus
End If

End Sub

Private Sub t_ape1_LostFocus()
   t_ape1.Text = Trim(t_ape1.Text)

End Sub

Private Sub t_ape2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   KeyAscii = 0
   t_ape2.Text = Trim(t_ape2.Text)
   mfnac.SetFocus
End If

End Sub

Private Sub t_ape2_LostFocus()
   If Trim(t_ape2.Text) = "NO APLICA" Then
      t_ape2.Text = ""
   End If
   t_ape2.Text = Trim(t_ape2.Text)

End Sub

Private Sub t_calle_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_entre.SetFocus
End If

End Sub

Private Sub t_calle_LostFocus()
MsgBox "RECUERDE! Consultar si ya tienen cobrador en domicilio para asignarlo.", vbInformation

End Sub

Private Sub t_casa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   cbozona.SetFocus
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   t_codced.SetFocus
End If

End Sub

Public Sub Borrar_campos()
labauto.Caption = ""
labcl_codigo.Caption = ""
t_ced.Text = ""
t_codced.Text = ""
cbocat.ListIndex = -1
cbopromo.Text = ""
mpd.Text = "__/__/____"
mph.Text = "__/__/____"
t_nom1.Text = ""
t_nom2.Text = ""
t_ape1.Text = ""
t_ape2.Text = ""
mfnac.Text = "__/__/____"
cbosexo.ListIndex = -1
t_telef.Text = ""
t_celu.Text = ""
t_correo.Text = ""
t_calle.Text = ""
t_entre.Text = ""
t_manz.Text = ""
t_sol.Text = ""
t_casa.Text = ""
cbozona.Text = ""
cbomut.Text = ""
t_valor.Text = ""
cbovende.Text = ""
labnro.Caption = ""
labcodpromo.Caption = ""
labcodzon.Caption = ""
labcodmut.Caption = ""
labcodvende.Caption = ""
labcodsello.Caption = ""
t_cobnro.Text = ""
cbocobnom.Text = ""
cbodiacob.ListIndex = -1
t_dircobro.Text = ""
t_casacobro.Text = ""
cbozonacobro.Text = ""
cbosello.ListIndex = -1
t_nrotarj.Text = ""
t_nrotarj2.Text = ""
t_nrotarj3.Text = ""
t_nrotarj4.Text = ""
cbomesv.ListIndex = -1
cboaniov.ListIndex = -1
t_nomtit.Text = ""
t_cedtit.Text = ""
t_codcedtit.Text = ""
t_domitarj.Text = ""
t_teleftarj.Text = ""
ch_debbrou.Value = 0
labcatcodsol.Caption = ""
labcatnomsol.Caption = ""
labcodconv.Caption = ""
labnomconv.Caption = ""
labcatreal.Caption = ""
labcatrealdes.Caption = ""
Check1.Value = 0

End Sub

Public Sub Consulta_Cli()
Dim Xsqlconsulta, Xladeuda, Xtienecarta, XconveMut As String
Dim Xfecha_baja As Date
Dim Xbajasi As Integer
Dim Xrecclii As New ADODB.Recordset
Dim Xrecdeuda As New ADODB.Recordset
Dim Xreccarta As New ADODB.Recordset
Dim Xrecconvem As New ADODB.Recordset
Dim VerificarCod As Integer
Dim XSinservicios As Integer
Dim FechaCartas As Date
FechaCartas = Date - 150
XSinservicios = 0
VerificarCod = 0
Xbajasi = 0
Xfecha_baja = Date - 180

ConectarBD
ConbdSapp.Open
Xsqlconsulta = "Select * from clientes where cl_cedula =" & t_ced.Text

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlconsulta, ConbdSapp, , , adCmdText
End With
labauto.Caption = "SI"
If Xrecclii.RecordCount > 0 Then
   labcl_codigo.Caption = Xrecclii("cl_codigo")
   If IsNull(Xrecclii("fecha_baja")) = False Then
      If Format(Xrecclii("fecha_baja"), "yyyy/mm/dd") > Format(Xfecha_baja, "yyyy/mm/dd") Then
         MsgBox "Socio de baja con fecha menor a 6 meses, NO INCLUYE PAGO DE PROMOCION AL VENDEDOR.", vbCritical
         Xbajasi = 1
         labpromocion.Caption = "NO"
      Else
         Xbajasi = 0
      End If
      MsgBox "El socio se encuentra de baja desde el " & Format(Xrecclii("fecha_baja"), "dd/mm/yyyy"), vbInformation
      
   Else
      Xbajasi = 0
   End If
   Xladeuda = "Select * from deudas where cliente =" & Xrecclii("cl_codigo") & " and fecha_pago is null and total >" & 0
   With Xrecdeuda
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xladeuda, ConbdSapp, , , adCmdText
   End With
   If Xrecdeuda.RecordCount > 0 Then
      Xrecdeuda.MoveLast
      If Xrecdeuda.RecordCount >= 1 Then
         If Xbajasi = 1 Then
            MsgBox "Socio con deuda pendiente de pago y fecha de baja menor a seis meses.", vbCritical
            MsgBox "La afiliación no podrá ser ingresada al padrón hasta que sea autorizada por administración.", vbCritical
            labauto.Caption = "NO"
         Else
            If Xrecdeuda.RecordCount >= 2 Then
               MsgBox "Socio con deuda pendiente de pago.", vbCritical
               MsgBox "La afiliación no podrá ser ingresada al padrón hasta que sea autorizada por administración.", vbCritical
               labauto.Caption = "NO"
            End If
         End If
      Else
         labauto.Caption = "SI"
      End If
   End If
   Xrecdeuda.Close
   
   If labauto.Caption = "SI" And IsNull(Xrecclii("fecha_baja")) = True Then
'   If labauto.Caption = "SI" Then
      If Xrecclii("cl_codconv") = "CCNOS" Or Xrecclii("cl_codconv") = "SMIN" Or _
         Xrecclii("cl_codconv") = "UNIVS" Or Xrecclii("cl_codconv") = "CCNRE" Or Xrecclii("cl_codconv") = "UNIVNR" Or _
         Xrecclii("cl_codconv") = "HEVANO" Or Xrecclii("cl_codconv") = "SMINR" Or Xrecclii("cl_codconv") = "CASANR" Or _
         Xrecclii("cl_codconv") = "GANOS" Or Xrecclii("cl_codconv") = "HEVAN" Or _
         Xrecclii("cl_codconv") = "CASANO" Then
         Xtienecarta = "Select * from linmmdd where cod_cli =" & Xrecclii("cl_codigo") & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format(FechaCartas, "yyyy/mm/dd") & "'"
         With Xreccarta
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Xtienecarta, ConbdSapp, , , adCmdText
         End With
         If Xreccarta.RecordCount > 0 Then
            labauto.Caption = "NO"
         Else
            MsgBox "No tiene realizada la carta mutual, la afiliación deberá ser autorizada por administración.", vbCritical
            labauto.Caption = "NOC"
         End If
         Xreccarta.Close
      End If
   End If
      
   If IsNull(Xrecclii("estado")) = False Then
      If Xrecclii("estado") = 1 Then
         XconveMut = "Select * from convenio where cnv_codigo ='" & Xrecclii("cl_codconv") & "' and cnv_cant_r not in (2) and cnv_grupo in ('CCOU','UNIVERSAL','SMI','H.EVANGELICO','CASA DE GALICIA')"
         With Xrecconvem
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open XconveMut, ConbdSapp, , , adCmdText
         End With
         If Xrecconvem.RecordCount > 0 Then
            MsgBox "ATENCION!!! Socio MUTUAL? Categoría sugerida: COMPLEMENTO", vbInformation
            labesmutual.Caption = "SI"
         Else
            labesmutual.Caption = ""
         End If
         Xrecconvem.Close
         XconveMut = "Select * from convenio where cnv_codigo ='" & Xrecclii("cl_codconv") & "' and cnv_cant_r in (2) and cnv_grupo in ('CCOU','UNIVERSAL','SMI','H.EVANGELICO','CASA DE GALICIA')"
         With Xrecconvem
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open XconveMut, ConbdSapp, , , adCmdText
         End With
         If Xrecconvem.RecordCount > 0 Then
            MsgBox "ATENCION!!! Socio MUTUAL ACTIVO CON COMPLEMENTO. Verifique con Padrón social.", vbCritical
            b_graba.Enabled = False
         End If
         Xrecconvem.Close
         If Xrecclii("cl_codconv") <> "PART" Then
             XconveMut = "Select * from convenio where cnv_codigo ='" & Xrecclii("cl_codconv") & "' and cnv_cant_r in (2) and cnv_grupo not in ('CCOU','UNIVERSAL','SMI','H.EVANGELICO','CASA DE GALICIA')"
             With Xrecconvem
                 .CursorLocation = adUseClient
                 .CursorType = adOpenKeyset
                 .LockType = adLockOptimistic
                 .Open XconveMut, ConbdSapp, , , adCmdText
             End With
             If Xrecconvem.RecordCount > 0 Then
                MsgBox "ATENCION!!! Socio SAPP ACTIVO CON EMISION. Verifique con Padrón social.", vbCritical
    '            b_graba.Enabled = False
             End If
             
             Xrecconvem.Close
         End If
      Else
         labesmutual.Caption = ""
      End If
   End If
   
   If IsNull(Xrecclii("saldo_chc2")) = False Then
      If Xrecclii("saldo_chc2") = 1 Then
         labauto.Caption = "NO"
         MsgBox "ATENCION!! Socio con servicios restringidos, no se puede afiliar. Consulte con Administración.", vbCritical
         XSinservicios = 1
         t_ced.Text = ""
         t_codced.Text = ""
      End If
   End If
   If XSinservicios <> 1 Then
    If IsNull(Xrecclii("cl_codconv")) = False Then
       labcodconv.Caption = Xrecclii("cl_codconv")
       labnomconv.Caption = Xrecclii("cl_nomconv")
    Else
       labcodconv.Caption = ""
       labnomconv.Caption = ""
    End If
    If IsNull(Xrecclii("cl_apellid")) = False Then
       t_nom1.Text = Xrecclii("cl_apellid")
    Else
       t_nom1.Text = ""
    End If
    If IsNull(Xrecclii("cl_apellid")) = False Then
       t_ape1.Text = Xrecclii("cl_apellid")
    Else
       t_ape1.Text = ""
    End If
    If IsNull(Xrecclii("cl_fnac")) = False Then
       mfnac.Text = Xrecclii("cl_fnac")
    Else
       mfnac.Text = "__/__/____"
    End If
    If IsNull(Xrecclii("cl_sexo")) = False Then
       If Xrecclii("cl_sexo") = 1 Then
          cbosexo.ListIndex = 1
       Else
          If Xrecclii("cl_sexo") = 2 Then
             cbosexo.ListIndex = 0
          Else
             cbosexo.ListIndex = -1
          End If
       End If
    End If
    If IsNull(Xrecclii("cl_telefon")) = False Then
       If Xrecclii("cl_telefon") = "NO APLICA" Then
          If Trim(t_telef.Text) <> "" Then
          Else
             t_telef.Text = ""
          End If
       Else
          If Trim(t_telef.Text) <> "" Then
          Else
             t_telef.Text = Xrecclii("cl_telefon")
          End If
       End If
    Else
       If Trim(t_telef.Text) <> "" Then
       Else
          t_telef.Text = ""
       End If
    End If
    If IsNull(Xrecclii("cl_dpto")) = False Then
       If Xrecclii("cl_dpto") = "NO APLICA" Then
          t_celu.Text = ""
       Else
          t_celu.Text = Xrecclii("cl_dpto")
       End If
    Else
       t_celu.Text = ""
    End If
    If IsNull(Xrecclii("cl_referen")) = False Then
       If Xrecclii("cl_referen") = "NO APLICA" Then
          t_correo.Text = ""
       Else
          t_correo.Text = Xrecclii("cl_referen")
       End If
    Else
       t_correo.Text = ""
    End If
    If IsNull(Xrecclii("cl_direcci")) = False Then
       If Trim(t_calle.Text) <> "" Then
       Else
          t_calle.Text = Xrecclii("cl_direcci")
       End If
    End If
    Frame1.Enabled = True
    cbocat.SetFocus
  End If
Else
   labcl_codigo.Caption = ""
   labauto.Caption = "SI"
   t_nom1.Text = ""
   t_ape1.Text = ""
   mfnac.Text = "__/__/____"
   cbosexo.ListIndex = -1
   t_telef.Text = ""
   t_celu.Text = ""
   t_correo.Text = ""
   Frame1.Enabled = True
   t_nom1.SetFocus
   labcodconv.Caption = ""
   labnomconv.Caption = ""
   cbocat.SetFocus
   labesmutual.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Private Sub t_cedtit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codcedtit.SetFocus
End If

End Sub

Private Sub t_celu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_correo.SetFocus
End If

End Sub

Private Sub t_cobnro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocobnom.SetFocus
End If

End Sub

Private Sub t_cobnro_LostFocus()
If t_cobnro.Text <> "" Then
   Buscar_cobrador
End If

End Sub

Private Sub t_codced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If Command1.Enabled = True Then
      Command1.SetFocus
      'Command1_Click
   End If
End If

End Sub

Public Sub Busca_codigo(texto As String)
Dim Sqlconstr As String
Dim Regtr As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
'cambiar codigos_aut por Codigos_aut
Sqlconstr = "Select * from codigos_aut where codaut ='" & Trim(texto) & "' and modulo ='" & "Afiliaciones" & "'"
With Regtr
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Sqlconstr, ConbdSapp, , , adCmdText
End With

If Regtr.RecordCount > 0 Then
   Regtr.Close
   ConbdSapp.Close
Else
   MsgBox "No se encuentra código de autorización, verifique!", vbCritical
   Regtr.Close
   ConbdSapp.Close
   b_cance_Click
End If


End Sub

Public Sub Carga_promos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
If WElusuario = "MPEREZ" Or WElusuario = "FFONTES" Or WElusuario = "VVIVAS" Then
   Xsqlpromo = "Select * from promocion_gpo"
Else
   Xsqlpromo = "Select * from promocion_gpo where id not in (4)"
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
      If IsNull(Xrecclii("descrip")) = False Then
         cbopromo.AddItem Xrecclii("descrip")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Carga_zonas()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from zonas order by zo_nombre"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("zo_nombre")) = False Then
         cbozona.AddItem Xrecclii("zo_nombre")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Carga_mutua()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm order by ca_nom"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("ca_nom")) = False Then
         cbomut.AddItem Xrecclii("ca_nom")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_zonas(zona As String)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pregunta As String

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from zonas where zo_nombre ='" & Trim(zona) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcodzon.Caption = Xrecclii("zo_grupo")
   If cbocat.Text = "EXTENSION SEMM" Or cbocat.Text = "EXTENSION UCM" Or cbocat.Text = "CONVENIO DE VERANO" Then
   Else
      If Frame3.Visible = True Then
         If Val(labcodzon.Caption) <> 999 Then
            If IsNull(Xrecclii("zo_cob")) = False Then
               Pregunta = MsgBox("Cobrador sugerido " & Xrecclii("zo_cob") & " DESEA ASIGNARLO?", vbInformation + vbYesNo, "Afiliaciones SAPP")
               If Pregunta = vbYes Then
                  t_cobnro.Text = Xrecclii("zo_cob")
               End If
            End If
         End If
      End If
   End If
Else
   labcodzon.Caption = ""
   cbozona.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Private Sub t_codced_LostFocus()
   If Command1.Enabled = True Then
      'Command1.SetFocus
      Command1_Click
   End If

End Sub

Private Sub t_codcedtit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_domitarj.SetFocus
End If

End Sub

Private Sub t_correo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(chr(KeyAscii)))
If KeyAscii = 13 Then
   t_calle.SetFocus
End If

End Sub

Private Sub t_dircobro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

End Sub

Private Sub t_domitarj_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_teleftarj.SetFocus
End If

End Sub

Private Sub t_entre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_manz.SetFocus
End If

End Sub

Private Sub t_manz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_sol.SetFocus
End If

End Sub

Private Sub t_nom1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   KeyAscii = 0
   t_nom1.Text = Trim(t_nom1.Text)
   t_nom2.SetFocus
End If

End Sub

Private Sub t_nom1_LostFocus()

t_nom1.Text = Trim(t_nom1.Text)

End Sub

Private Sub t_nom2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   KeyAscii = 0
   t_nom2.Text = Trim(t_nom2.Text)
   t_ape1.SetFocus
End If

End Sub

Private Sub t_nom2_LostFocus()
   If Trim(t_nom2.Text) = "NO APLICA" Then
      t_nom2.Text = ""
   End If
   t_nom2.Text = Trim(t_nom2.Text)

End Sub

Private Sub t_nomtit_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_cedtit.SetFocus
End If

End Sub

Private Sub t_nrotarj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nrotarj2.SetFocus
End If

End Sub

Private Sub t_nrotarj2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nrotarj3.SetFocus
End If

End Sub

Private Sub t_nrotarj3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nrotarj4.SetFocus
End If

End Sub

Private Sub t_nrotarj4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomesv.SetFocus
End If

End Sub

Private Sub t_sol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_casa.SetFocus
End If

End Sub

Private Sub t_telef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_celu.SetFocus
End If

End Sub

Private Sub t_valor_Change()
'If IsNumeric(t_valor.Text) = False Then
'   MsgBox "Se admiten solo números", vbCritical
'   t_valor.Text = ""
'End If

End Sub


Public Sub Consulta_promos()
Dim Xsqlpromo, Xporcent As String
Dim Xrecclii As New ADODB.Recordset
Dim ValorDesc As Double

ConectarBD
ConbdSapp.Open
ValorDesc = 0
Xporcent = ""
Xsqlpromo = "Select * from promocion_gpo where descrip ='" & cbopromo.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcodpromo.Caption = Xrecclii("id")
   Xporcent = "0." & Xrecclii("descu_por")
   If t_valor.Text <> "" Then
      ValorDesc = CDbl(t_valor.Text) * CDbl(Xporcent)
      labtotafil.Caption = Val(t_valor.Text) - Int(ValorDesc)
      Label38.Visible = True
      labdescpromo.Caption = Trim(str(Xrecclii("descu_por"))) & " % por " & cbopromo.Text
   Else
      ValorDesc = 0
      labtotafil.Caption = ""
      Label38.Visible = False
      labdescpromo.Caption = ""
   End If
   labdapagar.Caption = Xrecclii("descu_por")
Else
   labcodpromo.Caption = ""
   labdapagar.Caption = ""
   labdescpromo.Caption = ""
   cbopromo.Text = ""
   ValorDesc = 0
   Label38.Visible = False
   labtotafil.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Public Sub Consulta_vendedor()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
If IsNumeric(cbovende.Text) = True Then
   Xsqlpromo = "Select * from vende_func where idfunc =" & Val(cbovende.Text)
Else
   Xsqlpromo = "Select * from vende_func where nombre ='" & cbovende.Text & "'"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcodvende.Caption = Xrecclii("idfunc")
   cbovende.Text = Xrecclii("nombre")
Else
   labcodvende.Caption = ""
   cbovende.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Carga_vendedores()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from vende_func order by nombre"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("nombre")) = False Then
         cbovende.AddItem Xrecclii("nombre")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Public Function Existe_afiliado() As Integer
Dim Xsqlpromo, XcedulaAfi As String
Dim Xrecclii As New ADODB.Recordset
Dim XsionoAfi As Integer
Dim FechaDe As Date

XcedulaAfi = t_ced.Text & t_codced.Text
XsionoAfi = 0
ConectarBD
ConbdSapp.Open

FechaDe = Date - 120
Xsqlpromo = "Select * from afiliaciones_new where cedula ='" & XcedulaAfi & "' and pendiente not in (20) and fecha >='" & Format(FechaDe, "yyyy-mm-dd") & "'"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   XsionoAfi = 1
Else
   XsionoAfi = 0
End If
Existe_afiliado = XsionoAfi

Xrecclii.Close
ConbdSapp.Close

End Function



Public Sub Carga_cobradors()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from cobrador where cb_recatra not in (2) order by cb_nombre"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("cb_nombre")) = False Then
         cbocobnom.AddItem Xrecclii("cb_nombre")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Buscar_cobrador()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from cobrador where cb_numero =" & t_cobnro.Text
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   t_cobnro.Text = Xrecclii("cb_numero")
   cbocobnom.Text = Xrecclii("cb_nombre")
Else
   t_cobnro.Text = ""
   cbocobnom.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Carga_zonasCob()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from zonas order by zo_nombre"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("zo_nombre")) = False Then
         cbozonacobro.AddItem Xrecclii("zo_nombre")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_zonasCob(zonac As String)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from zonas where zo_nombre ='" & Trim(zonac) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcodzoncobr.Caption = Xrecclii("zo_grupo")
Else
   labcodzoncobr.Caption = ""
   cbozonacobro.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Function Verifica_CedTarj() As Integer

Dim Xpond, Xn1v, Xn2v, Xn3v, Xn4v, Xn5v, Xn6v, Xn7v, Xtotv As Long
Dim Xcedtexv, Xtottexv As String
Dim Xced1v, Xced2v, Xced3v, Xced4v, Xced5v, Xced6v, Xced7v, Xlargov As Long

Xn1v = 2
Xn2v = 9
Xn3v = 8
Xn4v = 7
Xn5v = 6
Xn6v = 3
Xn7v = 4
Xpondv = 10
If IsNumeric(t_cedtit.Text) = False Then
   MsgBox "La cédula debe contener solo números", vbInformation
Else
   Xcedtexv = Trim(str(t_cedtit.Text))
   Xlargov = Len(Xcedtexv)
   If Xlargov = 6 Then
      Xcedtexv = "0" & Trim(Xcedtexv)
   End If
   Xced1v = Val(Mid(Trim(Xcedtexv), 1, 1))
   Xced2v = Val(Mid(Xcedtexv, 2, 1))
   Xced3v = Val(Mid(Xcedtexv, 3, 1))
   Xced4v = Val(Mid(Xcedtexv, 4, 1))
   Xced5v = Val(Mid(Xcedtexv, 5, 1))
   Xced6v = Val(Mid(Xcedtexv, 6, 1))
   Xced7v = Val(Mid(Xcedtexv, 7, 1))
   Xced1v = Xced1v * Xn1v
   Xced2v = Xced2v * Xn2v
   Xced3v = Xced3v * Xn3v
   Xced4v = Xced4v * Xn4v
   Xced5v = Xced5v * Xn5v
   Xced6v = Xced6v * Xn6v
   Xced7v = Xced7v * Xn7v
   Xtotv = Xced1v + Xced2v + Xced3v + Xced4v + Xced5v + Xced6v + Xced7v
   If Len(Trim(str(Xtotv))) = 1 Then
      Xtottexv = "0000" & Trim(str(Xtotv))
   End If
   If Len(Trim(str(Xtotv))) = 2 Then
      Xtottexv = "000" & Trim(str(Xtotv))
   End If
   If Len(Trim(str(Xtotv))) = 3 Then
      Xtottexv = "00" & Trim(str(Xtotv))
   End If
   If Len(Trim(str(Xtotv))) = 4 Then
      Xtottexv = "0" & Trim(str(Xtotv))
   End If
   Xtotv = Val(Mid(Xtottexv, 5, 1))
   If Xtotv <> 0 Then
      Xtotv = Xpondv - Xtotv
   Else
      Xtotv = 0
   End If
   If Xtotv <> t_codcedtit.Text Then
      Verifica_CedTarj = 2
   Else
      Verifica_CedTarj = 0
   End If
End If

End Function

Public Sub Carga_catego()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_categ where id not in (12)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("descrip")) = False Then
         cbocat.AddItem Xrecclii("descrip")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_precio()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_categ where descrip ='" & cbocat.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   t_valor.Text = Val(Xrecclii("valor"))
   labcatnomsol.Caption = Xrecclii("catnom")
   labcatcodsol.Caption = Xrecclii("catsapp")
   If IsNull(Xrecclii("catrealcod")) = False Then
      labcatreal.Caption = Xrecclii("catrealcod")
      labcatrealdes.Caption = Xrecclii("catrealdes")
   Else
      labcatreal.Caption = ""
      labcatrealdes.Caption = ""
   End If
Else
   t_valor.Text = ""
   labcatnomsol.Caption = ""
   labcatcodsol.Caption = ""
   labcatreal.Caption = ""
   labcatrealdes.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close
End Sub

Public Sub Debito_brou()
ch_debbrou.Visible = True
Label28.Visible = False
t_nrotarj.Text = ""
t_nrotarj2.Text = ""
t_nrotarj3.Text = ""
t_nrotarj4.Text = ""
t_nrotarj.Visible = False
t_nrotarj2.Visible = False
t_nrotarj3.Visible = False
t_nrotarj4.Visible = False
Label30.Visible = False
cbomesv.Visible = False
cboaniov.Visible = False

End Sub

Public Sub NoDebito_brou()
ch_debbrou.Value = 0
ch_debbrou.Visible = False
Label28.Visible = True
t_nrotarj.Visible = True
t_nrotarj2.Visible = True
t_nrotarj3.Visible = True
t_nrotarj4.Visible = True
Label30.Visible = True
cbomesv.Visible = True
cboaniov.Visible = True

End Sub

Public Sub Borrar_camposDos()
labauto.Caption = "SI"
labcl_codigo.Caption = ""
t_ced.Text = ""
t_codced.Text = ""
mpd.Text = "__/__/____"
mph.Text = "__/__/____"
'cbocat.ListIndex = -1
t_nom1.Text = ""
t_nom2.Text = ""
t_ape1.Text = ""
t_ape2.Text = ""
mfnac.Text = "__/__/____"
cbosexo.ListIndex = -1
't_telef.Text = ""
t_celu.Text = ""
t_correo.Text = ""
't_calle.Text = ""
't_entre.Text = ""
't_manz.Text = ""
't_sol.Text = ""
't_casa.Text = ""
'cbozona.Text = ""
cbomut.Text = ""
't_valor.Text = ""
'cbovende.Text = ""
'labcodpromo.Caption = ""
'labcodzon.Caption = ""
labcodmut.Caption = ""
'labcodvende.Caption = ""
labcodconv.Caption = ""
labnomconv.Caption = ""

End Sub

Public Sub Consultar_Yaexiste()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Lacedulatexto As String
Dim Fechahasta As Date

ConectarBD
ConbdSapp.Open

Fechahasta = Date - 90
Lacedulatexto = t_ced.Text & t_codced.Text
If Lacedulatexto <> "" Then
   Xsqlconsulta = "Select * from afiliaciones_new where cedula =" & Lacedulatexto & " and fecha >='" & Format(Fechahasta, "yyyy/mm/dd") & "' and pendiente not in (20)"
Else
   Lacedulatexto = "0"
   Xsqlconsulta = "Select * from afiliaciones_new where cedula =" & Lacedulatexto & " and fecha >='" & Format(Fechahasta, "yyyy/mm/dd") & "' and pendiente not in (20)"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlconsulta, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   
   MsgBox "La cédula ya figura registrada recientemente en una afiliación, Verifique!", vbCritical
   t_ced.Text = ""
   t_codced.Text = ""
   b_cance_Click
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Alta_Modif()
Dim Cl_apellid, Cl_entre, Cl_dir, VenceT, MutAfil, VendeAfil As String
Dim Xmatnew As Long
Xmatnew = 0
Cl_apellid = ""
Cl_entre = ""
VenceT = ""
Cl_dir = ""
MutAfil = ""
VendeAfil = ""

If IsNull(data_afilcons.Recordset("matricula")) = False Then
   data_cli.RecordSource = "select * from clientes where cl_codigo =" & data_afilcons.Recordset("matricula")
Else
   data_cli.RecordSource = "select * from clientes where cl_codigo =" & 0
End If
data_cli.Refresh

If data_cli.Recordset.RecordCount > 0 And IsNull(data_afilcons.Recordset("matricula")) = False Then
   If IsNull(data_cli.Recordset("fecha_baja")) = False Then
      data_cli.Recordset.Edit
      data_cli.Recordset("fecha_baja") = Null
      data_cli.Recordset.Update
   End If
   If IsNull(data_cli.Recordset("estado")) = False Then
      If data_cli.Recordset("estado") <> 1 Then
         data_cli.Recordset.Edit
         data_cli.Recordset("estado") = 1
         data_cli.Recordset.Update
      End If
   Else
      data_cli.Recordset.Edit
      data_cli.Recordset("estado") = 1
      data_cli.Recordset.Update
   End If
   data_cli.Recordset.Edit
   If IsNull(data_afilcons.Recordset("catreal")) = False Then
      data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("catreal")
      data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("catrealdes")
   Else
      data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("categ")
      data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("nomcateg")
   End If
   If IsNull(data_afilcons.Recordset("nom2")) = False Then
      If IsNull(data_afilcons.Recordset("ape2")) = False Then
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
      Else
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
      End If
   Else
      If IsNull(data_afilcons.Recordset("ape2")) = False Then
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1")
      Else
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
      End If
   End If
   data_cli.Recordset("cl_apellid") = Mid(Cl_apellid, 1, 60)
   If IsNull(data_afilcons.Recordset("manz")) = False Then
      If IsNull(data_afilcons.Recordset("solar")) = False Then
         Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz") & " SL." & data_afilcons.Recordset("solar")
      Else
         Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz")
      End If
   Else
      Cl_dir = data_afilcons.Recordset("direc1")
   End If
   data_cli.Recordset("cl_direcci") = Mid(Cl_dir, 1, 80)
   If IsNull(data_afilcons.Recordset("casa")) = False Then
      If IsNull(data_afilcons.Recordset("direc2")) = False Then
         Cl_entre = data_afilcons.Recordset("casa") & " " & data_afilcons.Recordset("direc2")
      Else
         Cl_entre = data_afilcons.Recordset("casa")
      End If
   Else
      If IsNull(data_afilcons.Recordset("direc2")) = False Then
         Cl_entre = data_afilcons.Recordset("direc2")
      End If
   End If
   If Trim(Cl_entre) <> "" Then
      data_cli.Recordset("cl_entre") = Mid(Cl_entre, 1, 80)
   End If
   data_cli.Recordset("cl_dpto") = data_afilcons.Recordset("celular")
   data_cli.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
   data_cli.Recordset("cl_cedula_t") = Trim(data_afilcons.Recordset("cedula"))
   If IsNull(data_afilcons.Recordset("celular")) = False Then
      If data_afilcons.Recordset("celular") = "NO APLICA" Then
      Else
         data_cli.Recordset("cl_celular_n") = Trim(data_afilcons.Recordset("celular"))
      End If
   End If
   If Check1.Value = 1 Then
      data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
      data_cli.Recordset("cl_codced") = 0
   Else
      If Len(data_afilcons.Recordset("cedula")) = 7 Then
         data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
         data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
      Else
         If Len(data_afilcons.Recordset("cedula")) = 8 Then
            data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
            data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
         Else
            data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
            data_cli.Recordset("cl_codced") = 0
         End If
      End If
   End If
   If Check1.Value = 1 Then
      data_cli.Recordset("cl_tipoced") = 1
   Else
      data_cli.Recordset("cl_tipoced") = 0
   End If
   data_cli.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
   data_cli.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
   If Frame2.Visible = True Then
      data_cli.Recordset("cl_forpago") = 2
      data_cli.Recordset("cl_descpag") = "Debito Automatico"
   Else
      data_cli.Recordset("cl_forpago") = 1
      data_cli.Recordset("cl_descpag") = "Abono Mensual"
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") > 0 Then
         data_cli.Recordset("idpromos") = data_afilcons.Recordset("codpromo")
      End If
   End If
   data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
   data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
   If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
      data_cli.Recordset("cl_dircobr") = data_afilcons.Recordset("direc_cobro")
   End If
   If Day(Date) >= 26 Then
      If Month(Date) = 12 Then
         data_cli.Recordset("cl_ultmesp") = 1
         data_cli.Recordset("cl_ultanop") = Year(Date) + 1
      Else
         data_cli.Recordset("cl_ultmesp") = Month(Date) + 1
         data_cli.Recordset("cl_ultanop") = Year(Date)
      End If
   Else
      data_cli.Recordset("cl_ultmesp") = Month(Date)
      data_cli.Recordset("cl_ultanop") = Year(Date)
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") = 2 Then
         If Month(Date) = 12 Then
            data_cli.Recordset("cl_ultmesp") = 11
            data_cli.Recordset("cl_ultanop") = Year(Date) + 1
         Else
            data_cli.Recordset("cl_ultmesp") = Month(Date) - 1
            data_cli.Recordset("cl_ultanop") = Year(Date) + 1
         End If
      End If
   End If
   If data_afilcons.Recordset("codcob") > 0 Then
      If Frame3.Visible = True Then
         If cbocobnom.Text <> "" Then
            data_cli.Recordset("cl_nrocobr") = data_afilcons.Recordset("codcob")
            data_cli.Recordset("cl_nomcobr") = Mid(cbocobnom.Text, 1, 25)
         Else
            data_cli.Recordset("cl_nrocobr") = 0
            data_cli.Recordset("cl_nomcobr") = "*TODOS"
         End If
      Else
         If Frame2.Visible = True Then
            If data_afilcons.Recordset("tarj_sello") = "OCA CARD" Then
               data_cli.Recordset("cl_nrocobr") = 690
               data_cli.Recordset("cl_nomcobr") = "OCA DEBITO"
            End If
            If data_afilcons.Recordset("tarj_sello") = "VISA" Then
               data_cli.Recordset("cl_nrocobr") = 514
               data_cli.Recordset("cl_nomcobr") = "DEBITO AUTOMATICO VISA"
            End If
            If data_afilcons.Recordset("tarj_sello") = "MASTER CARD" Then
               data_cli.Recordset("cl_nrocobr") = 683
               data_cli.Recordset("cl_nomcobr") = "DEBITO MASTERCARD"
            End If
            If data_afilcons.Recordset("tarj_sello") = "CABAL" Then
               data_cli.Recordset("cl_nrocobr") = 673
               data_cli.Recordset("cl_nomcobr") = "DEBITO CABAL"
            End If
            If data_afilcons.Recordset("tarj_sello") = "DEBITO BROU" Then
               data_cli.Recordset("cl_nrocobr") = 607
               data_cli.Recordset("cl_nomcobr") = "DEBITO BROU"
            End If
            data_cli.Recordset("tit_tarj") = data_afilcons.Recordset("tarj_titular")
            If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
               data_cli.Recordset("cl_nrotarj") = data_afilcons.Recordset("tarj_nro")
            End If
            data_cli.Recordset("ci_tarj") = data_afilcons.Recordset("tarj_cedtit")
            data_cli.Recordset("codcitarj") = data_afilcons.Recordset("tarj_codced")
            data_cli.Recordset("cl_tjemi_c") = data_afilcons.Recordset("tarj_codsello")
            data_cli.Recordset("cl_tjemi_n") = data_afilcons.Recordset("tarj_sello")
            If data_afilcons.Recordset("tarj_vencmes") <> 0 Then
               If data_afilcons.Recordset("tarj_vencmes") > 9 Then
                  VenceT = "01/" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
               Else
                  VenceT = "01/0" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
               End If
               data_cli.Recordset("cl_tj_venc") = Format(CDate(VenceT), "dd/mm/yyyy")
            End If
            data_cli.Recordset("tarj_domi") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 60)
            data_cli.Recordset("tarj_telef") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 45)
         End If
      End If
   Else
      data_cli.Recordset("cl_nrocobr") = 0
      data_cli.Recordset("cl_nomcobr") = "*TODOS"
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") = 1 Then
         If data_afilcons.Recordset("integra_nro") <> 1 Then
            Consulta_AfilRuta
            If labcedtit.Caption <> "" Then
               data_cli.Recordset("cl_codruta") = Val(labcedtit.Caption)
            End If
         End If
      End If
   End If
   If Day(Date) >= 26 Then
      If Month(Date) = 11 Or Month(Date) = 12 Then
         If Month(Date) = 11 Then
            data_cli.Recordset("mesproxemi") = 1
            data_cli.Recordset("anoproxemi") = Year(Date) + 1
         Else
            data_cli.Recordset("mesproxemi") = 2
            data_cli.Recordset("anoproxemi") = Year(Date) + 1
         End If
      Else
         data_cli.Recordset("mesproxemi") = Month(Date) + 2
         data_cli.Recordset("anoproxemi") = Year(Date)
      End If
   Else
      If Month(Date) = 12 Then
         data_cli.Recordset("mesproxemi") = 1
         data_cli.Recordset("anoproxemi") = Year(Date) + 1
      Else
         data_cli.Recordset("mesproxemi") = Month(Date) + 1
         data_cli.Recordset("anoproxemi") = Year(Date)
      End If
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") = 2 Then
         If Day(Date) >= 26 Then
            If Month(Date) = 12 Then
               data_cli.Recordset("mesproxemi") = 1
               data_cli.Recordset("anoproxemi") = Year(Date) + 2
            Else
               data_cli.Recordset("mesproxemi") = Month(Date) + 1
               data_cli.Recordset("anoproxemi") = Year(Date) + 1
            End If
         Else
            If Month(Date) = 12 Then
               data_cli.Recordset("mesproxemi") = 1
               data_cli.Recordset("anoproxemi") = Year(Date) + 2
            Else
               data_cli.Recordset("mesproxemi") = Month(Date)
               data_cli.Recordset("anoproxemi") = Year(Date) + 1
            End If
         End If
      End If
   End If
   data_cli.Recordset("cl_referen") = Mid(data_afilcons.Recordset("correo"), 1, 74)
   If IsNull(data_afilcons.Recordset("codzon")) = False Then
      data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
      data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
   Else
      data_cli.Recordset("cl_grupo") = 0
      data_cli.Recordset("cl_zona") = "*TODOS"
   End If
   data_cli.Recordset("cl_sexo") = data_afilcons.Recordset("sexo")
   MutAfil = Devuelve_mut()
   data_cli.Recordset("cl_socmnom") = Mid(MutAfil, 1, 25)
   If IsNull(data_cli.Recordset("fecha_baja")) = False Then
      data_cli.Recordset("fecha_baja") = Null
   End If
   VendeAfil = Devuelve_vende()
   data_cli.Recordset("cl_nrovend") = data_afilcons.Recordset("codvende")
   data_cli.Recordset("cl_nomvend") = Mid(VendeAfil, 1, 35)
   If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
      data_cli.Recordset("cl_diacobr") = Trim(str(data_afilcons.Recordset("dia_cobro"))) & " C/MES"
   End If
   data_cli.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
   data_cli.Recordset.Update
   data_abm.RecordSource = "select * from abmsocio where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
   data_abm.Refresh
   data_abm.Recordset.AddNew
   data_abm.Recordset("usuario") = WElusuario
   data_abm.Recordset("fecha") = Date
   data_abm.Recordset("hora") = Format(Time, "HH:mm")
   data_abm.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
   data_abm.Recordset("desc") = "MODIF"
   data_abm.Recordset("cl_motivo") = "MODIFICACION ESPECIAL"
   data_abm.Recordset("convenio") = data_afilcons.Recordset("categ")
   data_abm.Recordset("base") = frm_menu.data_parse.Recordset("base")
   data_abm.Recordset.Update
Else
'   Xmatnew = data_nro.Recordset("p_newsocio") + 1
   Xmatnew = data_nrosoc.Recordset("ultimo_soc") + 1
   
   data_nrosoc.Recordset.Edit
   data_nrosoc.Recordset("ultimo_soc") = Xmatnew
   data_nrosoc.Recordset.Update
   
   data_afilcons.Recordset.Edit
   data_afilcons.Recordset("matricula") = Xmatnew
   data_afilcons.Recordset.Update
   data_cli.Recordset.AddNew
   data_cli.Recordset("cl_codigo") = Xmatnew
   data_cli.Recordset("estado") = 1
   If IsNull(data_afilcons.Recordset("catreal")) = False Then
      data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("catreal")
      data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("catrealdes")
   Else
      data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("categ")
      data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("nomcateg")
   End If
   If IsNull(data_afilcons.Recordset("nom2")) = False Then
      If IsNull(data_afilcons.Recordset("ape2")) = False Then
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
      Else
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
      End If
   Else
      If IsNull(data_afilcons.Recordset("ape2")) = False Then
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1")
      Else
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
      End If
   End If
   data_cli.Recordset("cl_apellid") = Mid(Cl_apellid, 1, 60)
   If IsNull(data_afilcons.Recordset("manz")) = False Then
      If IsNull(data_afilcons.Recordset("solar")) = False Then
         Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz") & " SL." & data_afilcons.Recordset("solar")
      Else
         Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz")
      End If
   Else
      Cl_dir = data_afilcons.Recordset("direc1")
   End If
   data_cli.Recordset("cl_direcci") = Mid(Cl_dir, 1, 80)
   If IsNull(data_afilcons.Recordset("casa")) = False Then
      If IsNull(data_afilcons.Recordset("direc2")) = False Then
         Cl_entre = data_afilcons.Recordset("casa") & " " & data_afilcons.Recordset("direc2")
      Else
         Cl_entre = data_afilcons.Recordset("casa")
      End If
   Else
      If IsNull(data_afilcons.Recordset("direc2")) = False Then
         Cl_entre = data_afilcons.Recordset("direc2")
      End If
   End If
   If Trim(Cl_entre) <> "" Then
      data_cli.Recordset("cl_entre") = Mid(Cl_entre, 1, 80)
   End If
   data_cli.Recordset("cl_dpto") = data_afilcons.Recordset("celular")
   data_cli.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
   
   data_cli.Recordset("cl_cedula_t") = Trim(data_afilcons.Recordset("cedula"))
   If IsNull(data_afilcons.Recordset("celular")) = False Then
      If data_afilcons.Recordset("celular") = "NO APLICA" Then
      Else
         data_cli.Recordset("cl_celular_n") = Trim(data_afilcons.Recordset("celular"))
      End If
   End If
   
   If Check1.Value = 1 Then
      data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
      data_cli.Recordset("cl_codced") = 0
   Else
      If Len(data_afilcons.Recordset("cedula")) = 7 Then
         data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
         data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
      Else
         If Len(data_afilcons.Recordset("cedula")) = 8 Then
            data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
            data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
         Else
            data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
            data_cli.Recordset("cl_codced") = 0
         End If
      End If
   End If
   If Check1.Value = 1 Then
      data_cli.Recordset("cl_tipoced") = 1
   Else
      data_cli.Recordset("cl_tipoced") = 0
   End If
   data_cli.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
   data_cli.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
   If Frame2.Visible = True Then
      data_cli.Recordset("cl_forpago") = 2
      data_cli.Recordset("cl_descpag") = "Debito Automatico"
   Else
      data_cli.Recordset("cl_forpago") = 1
      data_cli.Recordset("cl_descpag") = "Abono Mensual"
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") > 0 Then
         data_cli.Recordset("idpromos") = data_afilcons.Recordset("codpromo")
      End If
   End If
   If Day(Date) >= 26 Then
      If Month(Date) = 12 Then
         data_cli.Recordset("cl_ultmesp") = 1
         data_cli.Recordset("cl_ultanop") = Year(Date) + 1
      Else
         data_cli.Recordset("cl_ultmesp") = Month(Date) + 1
         data_cli.Recordset("cl_ultanop") = Year(Date)
      End If
   Else
      data_cli.Recordset("cl_ultmesp") = Month(Date)
      data_cli.Recordset("cl_ultanop") = Year(Date)
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") = 2 Then
         If Month(Date) = 12 Then
            data_cli.Recordset("cl_ultmesp") = 11
            data_cli.Recordset("cl_ultanop") = Year(Date) + 1
         Else
            data_cli.Recordset("cl_ultmesp") = Month(Date) - 1
            data_cli.Recordset("cl_ultanop") = Year(Date) + 1
         End If
      End If
   End If
   data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
   data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
   If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
      data_cli.Recordset("cl_dircobr") = data_afilcons.Recordset("direc_cobro")
   End If
   If data_afilcons.Recordset("codcob") > 0 Then
      If Frame3.Visible = True Then
         If cbocobnom.Text <> "" Then
            data_cli.Recordset("cl_nrocobr") = data_afilcons.Recordset("codcob")
            data_cli.Recordset("cl_nomcobr") = Mid(cbocobnom.Text, 1, 25)
         Else
            data_cli.Recordset("cl_nrocobr") = 0
            data_cli.Recordset("cl_nomcobr") = "*TODOS"
         End If
      Else
         If Frame2.Visible = True Then
            If data_afilcons.Recordset("tarj_sello") = "OCA CARD" Then
               data_cli.Recordset("cl_nrocobr") = 690
               data_cli.Recordset("cl_nomcobr") = "OCA DEBITO"
            End If
            If data_afilcons.Recordset("tarj_sello") = "VISA" Then
               data_cli.Recordset("cl_nrocobr") = 514
               data_cli.Recordset("cl_nomcobr") = "DEBITO AUTOMATICO VISA"
            End If
            If data_afilcons.Recordset("tarj_sello") = "MASTER CARD" Then
               data_cli.Recordset("cl_nrocobr") = 683
               data_cli.Recordset("cl_nomcobr") = "DEBITO MASTERCARD"
            End If
            If data_afilcons.Recordset("tarj_sello") = "CABAL" Then
               data_cli.Recordset("cl_nrocobr") = 673
               data_cli.Recordset("cl_nomcobr") = "DEBITO CABAL"
            End If
            If data_afilcons.Recordset("tarj_sello") = "DEBITO BROU" Then
               data_cli.Recordset("cl_nrocobr") = 607
               data_cli.Recordset("cl_nomcobr") = "DEBITO BROU"
            End If
            data_cli.Recordset("tit_tarj") = data_afilcons.Recordset("tarj_titular")
            If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
               data_cli.Recordset("cl_nrotarj") = data_afilcons.Recordset("tarj_nro")
            End If
            data_cli.Recordset("ci_tarj") = data_afilcons.Recordset("tarj_cedtit")
            data_cli.Recordset("codcitarj") = data_afilcons.Recordset("tarj_codced")
            data_cli.Recordset("cl_tjemi_c") = data_afilcons.Recordset("tarj_codsello")
            data_cli.Recordset("cl_tjemi_n") = data_afilcons.Recordset("tarj_sello")
            If data_afilcons.Recordset("tarj_vencmes") <> 0 Then
               If data_afilcons.Recordset("tarj_vencmes") > 9 Then
                  VenceT = "01/" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
               Else
                  VenceT = "01/0" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
               End If
               data_cli.Recordset("cl_tj_venc") = Format(CDate(VenceT), "dd/mm/yyyy")
            End If
            data_cli.Recordset("tarj_domi") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 60)
            data_cli.Recordset("tarj_telef") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 45)
         End If
      End If
   Else
      data_cli.Recordset("cl_nrocobr") = 0
      data_cli.Recordset("cl_nomcobr") = "*TODOS"
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") = 1 Then
         If data_afilcons.Recordset("integra_nro") <> 1 Then
            Consulta_AfilRuta
            If labcedtit.Caption <> "" Then
               data_cli.Recordset("cl_codruta") = Val(labcedtit.Caption)
            End If
         End If
      End If
   End If
   If Day(Date) >= 26 Then
      If Month(Date) = 11 Or Month(Date) = 12 Then
         If Month(Date) = 11 Then
            data_cli.Recordset("mesproxemi") = 1
            data_cli.Recordset("anoproxemi") = Year(Date) + 1
         Else
            data_cli.Recordset("mesproxemi") = 2
            data_cli.Recordset("anoproxemi") = Year(Date) + 1
         End If
      Else
         data_cli.Recordset("mesproxemi") = Month(Date) + 2
         data_cli.Recordset("anoproxemi") = Year(Date)
      End If
   Else
      If Month(Date) = 12 Then
         data_cli.Recordset("mesproxemi") = 1
         data_cli.Recordset("anoproxemi") = Year(Date) + 1
      Else
         data_cli.Recordset("mesproxemi") = Month(Date) + 1
         data_cli.Recordset("anoproxemi") = Year(Date)
      End If
   End If
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      If data_afilcons.Recordset("codpromo") = 2 Then
         If Day(Date) >= 26 Then
            If Month(Date) = 12 Then
               data_cli.Recordset("mesproxemi") = 1
               data_cli.Recordset("anoproxemi") = Year(Date) + 2
            Else
               data_cli.Recordset("mesproxemi") = Month(Date) + 1
               data_cli.Recordset("anoproxemi") = Year(Date) + 1
            End If
         Else
            If Month(Date) = 12 Then
               data_cli.Recordset("mesproxemi") = 1
               data_cli.Recordset("anoproxemi") = Year(Date) + 2
            Else
               data_cli.Recordset("mesproxemi") = Month(Date)
               data_cli.Recordset("anoproxemi") = Year(Date) + 1
            End If
         End If
      End If
   End If
   data_cli.Recordset("cl_referen") = Mid(data_afilcons.Recordset("correo"), 1, 74)
   If IsNull(data_afilcons.Recordset("codzon")) = False Then
      data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
      data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
   Else
      data_cli.Recordset("cl_grupo") = 0
      data_cli.Recordset("cl_zona") = "*TODOS"
   End If
   data_cli.Recordset("cl_sexo") = data_afilcons.Recordset("sexo")
   MutAfil = Devuelve_mut()
   data_cli.Recordset("cl_socmnom") = Mid(MutAfil, 1, 25)
   If IsNull(data_cli.Recordset("fecha_baja")) = False Then
      data_cli.Recordset("fecha_baja") = Null
   End If
   VendeAfil = Devuelve_vende()
   data_cli.Recordset("cl_nrovend") = data_afilcons.Recordset("codvende")
   data_cli.Recordset("cl_nomvend") = Mid(VendeAfil, 1, 35)
   If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
      data_cli.Recordset("cl_diacobr") = Trim(str(data_afilcons.Recordset("dia_cobro"))) & " C/MES"
   End If
   data_cli.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
   data_cli.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
   data_cli.Recordset.Update
   data_abm.RecordSource = "select * from abmsocio where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
   data_abm.Refresh
   data_abm.Recordset.AddNew
   data_abm.Recordset("usuario") = WElusuario
   data_abm.Recordset("fecha") = Date
   data_abm.Recordset("hora") = Format(Time, "HH:mm")
   data_abm.Recordset("cl_codigo") = Xmatnew
   data_abm.Recordset("desc") = "ALTA"
   data_abm.Recordset("cl_motivo") = "ALTA DE FICHA"
   data_abm.Recordset("convenio") = data_afilcons.Recordset("categ")
   data_abm.Recordset("base") = frm_menu.data_parse.Recordset("base")
   data_abm.Recordset.Update
   
End If

End Sub

Public Sub Consulta_vendeAfil()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from vende_func where idfunc =" & data_afilcons.Recordset("codvende")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labnompromo.Caption = Xrecclii("nombre")
Else
   labnompromo.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Public Sub Consulta_AfilRuta()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcedtit.Caption = Xrecclii("cedula")
Else
   labcedtit.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Function Devuelve_mut() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm where id =" & data_afilcons.Recordset("codmut")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_mut = Xrecclii("ca_nom")
Else
   Devuelve_mut = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

Public Function Devuelve_vende() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from vende_func where idfunc =" & data_afilcons.Recordset("codvende")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_vende = Xrecclii("nombre")
Else
   Devuelve_vende = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

Public Sub Veo_plazos()
labplaz.Visible = True
mpd.Visible = True
mph.Visible = True
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label34.Visible = False
Label35.Visible = False
t_cobnro.Visible = False
cbocobnom.Visible = False
cbodiacob.Visible = False
t_dircobro.Visible = False
t_casacobro.Visible = False
cbozonacobro.Visible = False

End Sub

Public Sub NoVeo_plazos()
labplaz.Visible = False
mpd.Visible = False
mph.Visible = False
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
Label34.Visible = True
Label35.Visible = True
t_cobnro.Visible = True
cbocobnom.Visible = True
cbodiacob.Visible = True
t_dircobro.Visible = True
t_casacobro.Visible = True
cbozonacobro.Visible = True

End Sub

Public Sub Genera_contrato()
Dim Direc, Xcontrato As String
Direc = ""

If data_afilcons.Recordset("convenio") = "COMPLEMENTO" Or data_afilcons.Recordset("convenio") = "COMPLEMENTO C.GALICIA" Or data_afilcons.Recordset("convenio") = "AMBULATORIO" Then
'   data_inf.DatabaseName = App.path & "\contrato.mdb"
'   data_inf.RecordSource = "contrato"
'   data_inf.Refresh
'   If IsNull(data_inf.Recordset("contrato")) = False Then
'      Xcontrato = data_inf.Recordset("contrato")
'   Else
'      Xcontrato = "Sin datos"
'   End If
   data_inf.DatabaseName = ""
   data_inf.Connect = "odbc;dsn=sappnew;"
   data_inf.RecordSource = "afiliaciones_contratoa"
   data_inf.Refresh
   If IsNull(data_inf.Recordset("descrip")) = False Then
      Xcontrato = data_inf.Recordset("descrip")
   Else
      Xcontrato = "Sin datos"
   End If
   data_inf.Connect = ""

Else
   data_inf.DatabaseName = ""
   data_inf.Connect = "odbc;dsn=sappnew;"
   data_inf.RecordSource = "afiliaciones_contratoe"
   data_inf.Refresh
   If IsNull(data_inf.Recordset("descrip")) = False Then
      Xcontrato = data_inf.Recordset("descrip")
   Else
      Xcontrato = "Sin datos"
   End If
   data_inf.Connect = ""
End If

data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh

If data_afilcons.Recordset.RecordCount > 0 Then
   data_inf.Recordset.AddNew
   data_inf.Recordset("cl_codigo") = data_afilcons.Recordset("afilia_nro")
   If IsNull(data_afilcons.Recordset("catcontrato")) = False Then
      data_inf.Recordset("cl_descpag") = data_afilcons.Recordset("catcontrato")
   Else
      If data_afilcons.Recordset("convenio") = "COMPLEMENTO" Or data_afilcons.Recordset("convenio") = "COMPLEMENTO C.GALICIA" Then
         data_inf.Recordset("cl_descpag") = "AMBULATORIO"
      Else
         data_inf.Recordset("cl_descpag") = data_afilcons.Recordset("convenio")
      End If
   End If
   data_inf.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
   data_inf.Recordset("cl_cantpag") = data_afilcons.Recordset("integra_nro")
   data_inf.Recordset("cl_apellid") = data_afilcons.Recordset("ape1")
   data_inf.Recordset("cl_medflia") = Mid(Devuelve_titular(), 1, 30)
   data_inf.Recordset("tit_tarj") = Mid(Devuelve_titularApe(), 1, 30)
   If IsNull(data_afilcons.Recordset("ape2")) = False Then
      data_inf.Recordset("cl_localid") = Mid(data_afilcons.Recordset("ape2"), 1, 35)
   End If
   data_inf.Recordset("cl_nomvend") = Mid(data_afilcons.Recordset("nom1"), 1, 35)
   If IsNull(data_afilcons.Recordset("nom2")) = False Then
      data_inf.Recordset("cl_nombre") = Mid(data_afilcons.Recordset("nom2"), 1, 30)
   End If
   data_inf.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
   If Len(data_afilcons.Recordset("cedula")) = 7 Then
      data_inf.Recordset("cl_fax") = Mid(data_afilcons.Recordset("cedula"), 1, 6) & "-" & Mid(data_afilcons.Recordset("cedula"), 7, 1)
   Else
      data_inf.Recordset("cl_fax") = Mid(data_afilcons.Recordset("cedula"), 1, 7) & "-" & Mid(data_afilcons.Recordset("cedula"), 8, 1)
   End If
   data_inf.Recordset("cl_tjemi_n") = Devuelve_titularCed()
   If IsNull(data_afilcons.Recordset("telef")) = False Then
      data_inf.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
   End If
   data_inf.Recordset("cl_celular") = data_afilcons.Recordset("celular")
   If IsNull(data_afilcons.Recordset("correo")) = False Then
      data_inf.Recordset("cl_dircobr") = data_afilcons.Recordset("correo")
   End If
   If IsNull(data_afilcons.Recordset("codmut")) = False Then
      data_inf.Recordset("cl_socmnom") = Devuelve_mut()
   End If
   If IsNull(data_afilcons.Recordset("direc2")) = False Then
      Direc = data_afilcons.Recordset("direc1") & " E/" & data_afilcons.Recordset("direc2")
   Else
      Direc = data_afilcons.Recordset("direc1")
   End If
   data_inf.Recordset("cl_direcci") = Mid(Direc, 1, 80)
   If IsNull(data_afilcons.Recordset("manz")) = False Then
      data_inf.Recordset("cl_estadoc") = data_afilcons.Recordset("manz")
   End If
   If IsNull(data_afilcons.Recordset("solar")) = False Then
      data_inf.Recordset("cl_tipcli") = Mid(data_afilcons.Recordset("solar"), 1, 3)
   End If
   If IsNull(data_afilcons.Recordset("nomzona")) = False Then
      data_inf.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
   End If
   data_inf.Recordset("cl_atrasoa") = data_afilcons.Recordset("valorcuota")
   If IsNull(data_afilcons.Recordset("desc_imp")) = False Then
      data_inf.Recordset("cl_seg_vto") = data_afilcons.Recordset("desc_imp")
   Else
      data_inf.Recordset("cl_seg_vto") = 0
   End If
   If IsNull(data_afilcons.Recordset("importe_fin")) = False Then
      data_inf.Recordset("cl_ter_vto") = data_afilcons.Recordset("importe_fin")
   Else
      data_inf.Recordset("cl_ter_vto") = 0
   End If
   If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
      data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta"
      data_inf.Recordset("info_debit") = "COBRO POR DÉBITO AUTOMÁTICO:" & chr(13)
      data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "Se adjunta autorización débito automático al final del contrato."
      If IsNull(data_afilcons.Recordset("codvende")) = False Then
         data_inf.Recordset("cl_entre") = Devuelve_vende()
      Else
         data_inf.Recordset("cl_entre") = "Sin promotor"
      End If
      If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
         data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
      End If
      data_inf.Recordset("cl_tipclin") = data_afilcons.Recordset("tarj_sello")
      data_inf.Recordset("cl_email") = Mid(data_afilcons.Recordset("tarj_titular"), 1, 30)
      data_inf.Recordset("cl_nrovend") = data_afilcons.Recordset("tarj_cedtit")
      data_inf.Recordset("cl_forpago") = data_afilcons.Recordset("tarj_codced")
      data_inf.Recordset("cl_nomconv") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 30)
      data_inf.Recordset("cl_nomcobr") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 25)
      data_inf.Recordset("cl_nrotarj") = Mid(data_afilcons.Recordset("tarj_nro"), 1, 20)
      data_inf.Recordset("cl_ultmesp") = data_afilcons.Recordset("tarj_vencmes")
      data_inf.Recordset("cl_ultanop") = data_afilcons.Recordset("tarj_vencanio")
   Else
      If IsNull(data_afilcons.Recordset("debito_brou")) = False Then
         data_inf.Recordset("cl_nom_sup") = "Débito BROU"
         data_inf.Recordset("info_debit") = "CONFIRMA QUE REALIZÓ FORMULARIO PARA DÉBITO BROU?:--->SI" & chr(13)
         data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "NOMBRE DE TITULAR DE LA CUENTA:" & data_afilcons.Recordset("tarj_titular")
         If IsNull(data_afilcons.Recordset("codvende")) = False Then
            data_inf.Recordset("cl_entre") = Devuelve_vende()
         Else
            data_inf.Recordset("cl_entre") = "Sin promotor"
         End If
         If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
            data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
         End If
      Else
         data_inf.Recordset("cl_nom_sup") = "Cobrador a domicilio"
         data_inf.Recordset("info_debit") = "DOMICILIO DE COBRO:" & chr(13)
         If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
            data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & data_afilcons.Recordset("direc_cobro") & chr(13)
            If IsNull(data_afilcons.Recordset("zonacobro")) = False Then
               data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "ZONA: " & data_afilcons.Recordset("zonacobro")
               If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: " & data_afilcons.Recordset("dia_cobro") & " c/mes" & chr(13)
               Else
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: sin datos." & chr(13)
               End If
            Else
               If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: " & data_afilcons.Recordset("dia_cobro") & " c/mes" & chr(13)
               Else
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: sin datos." & chr(13)
               End If
            End If
         Else
            data_inf.Recordset("info_debit") = "Misma dirección." & chr(13)
         End If
         If IsNull(data_afilcons.Recordset("codvende")) = False Then
            data_inf.Recordset("cl_entre") = Devuelve_vende()
         Else
            data_inf.Recordset("cl_entre") = "Sin promotor"
         End If
         If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
            data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
         End If
         
      End If
   End If
   data_inf.Recordset("obsp") = Xcontrato
   data_inf.Recordset.Update
Else
   MsgBox "No hay datos de afiliación para imprimir. Verifique!", vbCritical
   
End If

End Sub

Public Function Devuelve_titular() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_titular = Xrecclii("nom1")
Else
   Devuelve_titular = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function
Public Function Devuelve_titularApe() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_titularApe = Xrecclii("ape1")
Else
   Devuelve_titularApe = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function

Public Sub Consulta_mutual()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from ca_adm where ca_nom ='" & cbomut.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcodmut.Caption = Xrecclii("id")
Else
   labcodmut.Caption = ""
   cbomut.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
If Label39.ForeColor = &H80000012 Then
   Label39.ForeColor = &HFF&
   If Label39.BackColor = &HFFFF& Then
      Label39.BackColor = &HFFFFFF
   Else
      Label39.BackColor = &HFFFF&
   End If
Else
   If Label39.ForeColor = &HFF& Then
      Label39.ForeColor = &H808000
      If Label39.BackColor = &HFFFF& Then
         Label39.BackColor = &HFFFFFF
      Else
         Label39.BackColor = &HFFFF&
      End If
   Else
      If Label39.ForeColor = &H808000 Then
         Label39.ForeColor = &H80000012
         If Label39.BackColor = &HFFFF& Then
            Label39.BackColor = &HFFFFFF
         Else
            Label39.BackColor = &HFFFF&
         End If
      End If
   End If
End If

End Sub

Public Sub Modif_SocioAfil()
Dim Cl_apellid, Cl_entre, Cl_dir, VenceT, MutAfil, VendeAfil As String
Cl_apellid = ""
Cl_entre = ""
VenceT = ""
Cl_dir = ""
MutAfil = ""
VendeAfil = ""

If IsNull(data_afilcons.Recordset("matricula")) = False Then
   data_cli.RecordSource = "select * from clientes where cl_codigo =" & data_afilcons.Recordset("matricula")
Else
   data_cli.RecordSource = "select * from clientes where cl_codigo =" & 0
End If
data_cli.Refresh

If data_cli.Recordset.RecordCount > 0 Then
   If IsNull(data_afilcons.Recordset("nom2")) = False Then
      If IsNull(data_afilcons.Recordset("ape2")) = False Then
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
      Else
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
      End If
   Else
      If IsNull(data_afilcons.Recordset("ape2")) = False Then
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1")
      Else
         Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
      End If
   End If
   If data_cli.Recordset("cl_apellid") <> Mid(Cl_apellid, 1, 60) Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_apellid") = Mid(Cl_apellid, 1, 60)
      data_cli.Recordset.Update
   End If
   If IsNull(data_afilcons.Recordset("manz")) = False Then
      If IsNull(data_afilcons.Recordset("solar")) = False Then
         Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz") & " SL." & data_afilcons.Recordset("solar")
      Else
         Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz")
      End If
   Else
      Cl_dir = data_afilcons.Recordset("direc1")
   End If
   If data_cli.Recordset("cl_direcci") <> Mid(Cl_dir, 1, 80) Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_direcci") = Mid(Cl_dir, 1, 80)
      data_cli.Recordset.Update
   End If
   If IsNull(data_afilcons.Recordset("casa")) = False Then
      If IsNull(data_afilcons.Recordset("direc2")) = False Then
         Cl_entre = data_afilcons.Recordset("casa") & " " & data_afilcons.Recordset("direc2")
      Else
         Cl_entre = data_afilcons.Recordset("casa")
      End If
   Else
      If IsNull(data_afilcons.Recordset("direc2")) = False Then
         Cl_entre = data_afilcons.Recordset("direc2")
      End If
   End If
   If Trim(Cl_entre) <> "" Then
      If data_cli.Recordset("cl_entre") <> Mid(Cl_entre, 1, 80) Then
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_entre") = Mid(Cl_entre, 1, 80)
         data_cli.Recordset.Update
      End If
   End If
   If data_cli.Recordset("cl_dpto") <> data_afilcons.Recordset("celular") Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_dpto") = data_afilcons.Recordset("celular")
      data_cli.Recordset.Update
   End If
   If data_cli.Recordset("cl_telefon") <> data_afilcons.Recordset("telef") Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
      data_cli.Recordset.Update
   End If
   If data_cli.Recordset("cl_fnac") <> data_afilcons.Recordset("fnac") Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
      data_cli.Recordset.Update
   End If
   If data_cli.Recordset("cl_grupo") <> data_afilcons.Recordset("codzon") Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
      data_cli.Recordset.Update
   End If
   If data_cli.Recordset("cl_zona") <> Mid(data_afilcons.Recordset("nomzona"), 1, 25) Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
      data_cli.Recordset.Update
   End If
   
   If data_cli.Recordset("cl_referen") <> Mid(data_afilcons.Recordset("correo"), 1, 74) Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_referen") = Mid(data_afilcons.Recordset("correo"), 1, 74)
      data_cli.Recordset.Update
   End If
   
   If IsNull(data_afilcons.Recordset("codzon")) = False Then
      If data_cli.Recordset("cl_grupo") <> data_afilcons.Recordset("codzon") Then
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
         data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
         data_cli.Recordset.Update
      End If
   Else
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
      data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
      data_cli.Recordset.Update
   End If
   If data_cli.Recordset("cl_sexo") <> data_afilcons.Recordset("sexo") Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_sexo") = data_afilcons.Recordset("sexo")
      data_cli.Recordset.Update
   End If
   MutAfil = Devuelve_mut()
   If data_cli.Recordset("cl_socmnom") <> Mid(MutAfil, 1, 25) Then
      data_cli.Recordset.Edit
      data_cli.Recordset("cl_socmnom") = Mid(MutAfil, 1, 25)
      data_cli.Recordset.Update
   End If
   
'   data_abm.RecordSource = "select * from abmsocio where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
'   data_abm.Refresh
'   data_abm.Recordset.AddNew
'   data_abm.Recordset("usuario") = WElusuario
'   data_abm.Recordset("fecha") = Date
'   data_abm.Recordset("hora") = Format(Time, "HH:mm")
'   data_abm.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
'   data_abm.Recordset("desc") = "MODIF"
'   data_abm.Recordset("cl_motivo") = "MODIFICACION ESPECIAL"
'   data_abm.Recordset("convenio") = data_afilcons.Recordset("categ")
'   data_abm.Recordset("base") = frm_menu.data_parse.Recordset("base")
'   data_abm.Recordset.Update
Else
   MsgBox "No se pudo modificar, socio no encontrado.", vbExclamation
End If


End Sub
Public Function Devuelve_titularCed() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If Len(Xrecclii("cedula")) = 7 Then
      Devuelve_titularCed = Mid(Trim(str(Xrecclii("cedula"))), 1, 6) & "-" & Mid(Trim(str(Xrecclii("cedula"))), 7, 1)
   Else
      Devuelve_titularCed = Mid(Trim(str(Xrecclii("cedula"))), 1, 7) & "-" & Mid(Trim(str(Xrecclii("cedula"))), 8, 1)
   End If
Else
   Devuelve_titularCed = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function

Public Sub Buscar_cobrador_nombre()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from cobrador where cb_nombre ='" & cbocobnom.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   t_cobnro.Text = Xrecclii("cb_numero")
   cbocobnom.Text = Xrecclii("cb_nombre")
Else
   t_cobnro.Text = ""
   cbocobnom.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Verifica_matricula_existe()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim XmatVerifica As Long
XmatVerifica = 0
ConectarBD
ConbdSapp.Open
             
XmatVerifica = data_nrosoc.Recordset("ultimo_soc") + 1
             
Xsqlpromo = "Select * from clientes where cl_codigo =" & XmatVerifica
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   MsgBox "ATENCION!!, hay error en los parámetros de numeración, comunique a informática.", vbCritical
   Unload Me
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Function Devuelve_catContrato() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
If cbocat.Text = "CAMBIO DE CATEGORIA" Then
   If labcatcodsol.Caption = "SOLEME" Then
      Devuelve_catContrato = "EMERGENCIA"
   Else
      If labcatcodsol.Caption = "SOLAMB" Then
         Devuelve_catContrato = "AMBULATORIO"
      Else
         If labcatcodsol.Caption = "SOLPAR" Then
            Devuelve_catContrato = "PARCIAL"
         Else
            Devuelve_catContrato = "AMBULATORIO"
         End If
      End If
   End If
Else
    Xsqlpromo = "Select * from afiliaciones_categ where descrip ='" & Trim(cbocat.Text) & "'"
    With Xrecclii
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open Xsqlpromo, ConbdSapp, , , adCmdText
    End With
    If Xrecclii.RecordCount > 0 Then
       Devuelve_catContrato = Xrecclii("catcontrato")
    Else
       Devuelve_catContrato = "Sin Dato"
    End If
    Xrecclii.Close
End If

ConbdSapp.Close

End Function

