VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_srvenferm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicios de enfermeria a domicilio"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11940
   Icon            =   "frm_srvenferm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      Picture         =   "frm_srvenferm.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Ingreso de observaciones administrador"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9840
      TabIndex        =   59
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   4560
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_parainf 
      Caption         =   "data_parainf"
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
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_verord 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11280
      Picture         =   "frm_srvenferm.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Ver la orden escaneada del registro seleccionado"
      Top             =   5160
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cm1 
      Left            =   1200
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
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
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_buscasrv 
      Caption         =   "data_buscasrv"
      Connect         =   "ODBC;DSN=sappespecial;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "serv_enferm"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_serv 
      Caption         =   "data_serv"
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
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox t_busced 
      BackColor       =   &H00FF8080&
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
      Height          =   375
      Left            =   6600
      TabIndex        =   43
      ToolTipText     =   "Ingresar solo los números antes del guión"
      Top             =   5280
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_srvenferm.frx":109E
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frm_srvenferm.frx":10BA
      TabIndex        =   41
      Top             =   5640
      Width           =   11655
   End
   Begin VB.CommandButton b_infos 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      Picture         =   "frm_srvenferm.frx":31F9
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      Picture         =   "frm_srvenferm.frx":3783
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      Picture         =   "frm_srvenferm.frx":3D0D
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      Picture         =   "frm_srvenferm.frx":4297
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_srvenferm.frx":4821
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5040
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para el servicio de enfermería"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000FF00&
         Caption         =   "Cuidados Paliativos"
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
         Left            =   7080
         TabIndex        =   64
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox t_obss 
         Height          =   855
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Top             =   3720
         Width           =   4935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Internación a domicilio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   60
         Top             =   4680
         Width           =   3015
      End
      Begin MSMask.MaskEdBox mfsusp 
         Height          =   375
         Left            =   8040
         TabIndex        =   58
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.TextBox t_realiza 
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
         Left            =   7800
         MaxLength       =   60
         TabIndex        =   56
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
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
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data data_desc 
         Caption         =   "data_desc"
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
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_abre 
         Caption         =   "data_abre"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   9720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data data_arch 
         Caption         =   "data_arch"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox t_cant 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   5520
         TabIndex        =   52
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox t_arch 
         Height          =   285
         Left            =   720
         TabIndex        =   50
         Top             =   4680
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton b_sube 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         Picture         =   "frm_srvenferm.frx":4DAB
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Cargar orden ESCANEADA en PDF"
         Top             =   4440
         Width           =   495
      End
      Begin VB.ComboBox combo1 
         BackColor       =   &H00FF8080&
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
         ItemData        =   "frm_srvenferm.frx":5335
         Left            =   7800
         List            =   "frm_srvenferm.frx":534B
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox t_bol 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   8760
         MaxLength       =   20
         TabIndex        =   45
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF00&
         Caption         =   "ACTO VERIFICADO"
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
         Left            =   6720
         TabIndex        =   35
         Top             =   4200
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   16777215
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   1680
         TabIndex        =   33
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   16777215
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
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   10680
         TabIndex        =   31
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   16777215
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   9480
         TabIndex        =   30
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   16777215
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
      Begin VB.TextBox t_acto 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   1680
         MaxLength       =   130
         TabIndex        =   28
         Top             =   2760
         Width           =   5295
      End
      Begin VB.TextBox t_telef 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   8760
         MaxLength       =   80
         TabIndex        =   26
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox t_local 
         BackColor       =   &H00FF8080&
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
         Height          =   405
         Left            =   8760
         MaxLength       =   70
         TabIndex        =   24
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox t_ref 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   1680
         MaxLength       =   130
         TabIndex        =   22
         Top             =   2280
         Width           =   5535
      End
      Begin VB.TextBox t_direc 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   1680
         MaxLength       =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox t_convd 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   3000
         MaxLength       =   70
         TabIndex        =   16
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox t_conv 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox t_nom 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   4680
         MaxLength       =   100
         TabIndex        =   13
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox t_codc 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox t_b 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   6840
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   16777215
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   16777215
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
      Begin VB.Label Label24 
         BackColor       =   &H00C00000&
         Caption         =   "Observaciones:"
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
         TabIndex        =   61
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C00000&
         Caption         =   "FEC.SUSPEND:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   57
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C00000&
         Caption         =   "REALIZA:"
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
         Left            =   6720
         TabIndex        =   55
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "CANTIDAD:"
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
         Left            =   4200
         TabIndex        =   51
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C00000&
         Caption         =   "ESTADO:"
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
         Left            =   6720
         TabIndex        =   47
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   9840
         TabIndex        =   46
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "BOLETA NRO:"
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
         Left            =   7320
         TabIndex        =   44
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C00000&
         Caption         =   "FECHA HASTA:"
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
         TabIndex        =   32
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C00000&
         Caption         =   "F.DESDE:"
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
         TabIndex        =   29
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C00000&
         Caption         =   "ACTO ENF:"
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
         TabIndex        =   27
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C00000&
         Caption         =   "TELEFONOS:"
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
         Left            =   7320
         TabIndex        =   25
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C00000&
         Caption         =   "LOCALIDAD:"
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
         Left            =   7320
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C00000&
         Caption         =   "REFERENCIAS:"
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
         TabIndex        =   21
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "DIRECCION:"
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
         TabIndex        =   19
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label labedad 
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
         Height          =   375
         Left            =   10200
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "EDAD:"
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
         Left            =   9480
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "CONVENIO:"
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
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
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
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "CEDULA:"
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
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label labu 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "USUARIO:"
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
         Left            =   7800
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "BASE:"
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
         Left            =   6000
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "HORA:"
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
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "FECHA:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solo los nros antes del guión"
      Height          =   255
      Left            =   8160
      TabIndex        =   54
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por cédula"
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
      Left            =   4920
      TabIndex        =   42
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   9000
      Picture         =   "frm_srvenferm.frx":53C3
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1575
   End
End
Attribute VB_Name = "frm_srvenferm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdfpath, pdfpath1 As String
Public pdffile As ADODB.Stream
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub b_cance_Click()
borrarc
b_nuevo.Enabled = True
b_edita.Enabled = True
b_cance.Enabled = False
b_infos.Enabled = True
DBGrid1.Enabled = True
t_busced.Enabled = True
b_graba.Enabled = False
b_verord.Enabled = True
Command1.Enabled = True
Frame1.Enabled = False
XAlta = 0


End Sub

Private Sub b_edita_Click()
Frame1.Enabled = True
b_nuevo.Enabled = False
b_edita.Enabled = False
b_cance.Enabled = True
b_infos.Enabled = False
DBGrid1.Enabled = False
t_busced.Enabled = False
b_graba.Enabled = True
Command1.Enabled = False
b_verord.Enabled = False
XAlta = 2
t_ced.SetFocus

End Sub

Private Sub b_graba_Click()
On Error GoTo Xelerrsrvenf

If XAlta = 1 Then
   If t_ced.Text = "" Or t_nom.Text = "" Or t_conv.Text = "" Or t_direc.Text = "" Or t_codc.Text = "" Or _
      t_local.Text = "" Or t_telef.Text = "" Or t_acto.Text = "" Or t_ref.Text = "" Or t_cant.Text = "" Or _
      t_bol.Text = "" Or mfd.Text = "__/__/____" Or mfh.Text = "__/__/____" Or Combo1.ListIndex < 0 Or t_arch.Text = "" Then
      MsgBox "Faltan ingresar datos para poder grabar, VERIFIQUE!"
   Else
      frm_srvenferm.MousePointer = 11
      data_serv.RecordSource = "Select * from serv_enferm where id <=" & 0
      data_serv.Refresh
      data_serv.Recordset.AddNew
      data_serv.Recordset("srv_fec") = mf.Text
      data_serv.Recordset("srv_hora") = mh.Text
      data_serv.Recordset("srv_base") = t_b.Text
      data_serv.Recordset("srv_usua") = labu.Caption
      data_serv.Recordset("srv_ced") = t_ced.Text
      data_serv.Recordset("srv_cced") = t_codc.Text
      data_serv.Recordset("srv_nom") = t_nom.Text
      If labedad.Caption = "" Then
         data_serv.Recordset("srv_edad") = "Sin F.NAC"
      Else
         data_serv.Recordset("srv_edad") = labedad.Caption
      End If
      data_serv.Recordset("srv_palia") = Check3.Value
      data_serv.Recordset("srv_local") = t_local.Text
      data_serv.Recordset("srv_direc") = t_direc.Text
      data_serv.Recordset("srv_cconv") = t_conv.Text
      data_serv.Recordset("srv_conv") = t_convd.Text
      data_serv.Recordset("srv_telef") = t_telef.Text
      data_serv.Recordset("srv_ref") = t_ref.Text
      data_serv.Recordset("srv_boleta") = t_bol.Text
      data_serv.Recordset("srv_acto") = t_acto.Text
      data_serv.Recordset("srv_fecd") = mfd.Text
      data_serv.Recordset("srv_hord") = mhd.Text
      data_serv.Recordset("srv_fech") = mfh.Text
      data_serv.Recordset("srv_horh") = mhh.Text
      data_serv.Recordset("srv_estad") = Combo1.Text
      data_serv.Recordset("srv_estan") = Combo1.ListIndex
      data_serv.Recordset("srv_verif") = Check1.Value
      data_serv.Recordset("srv_cant") = t_cant.Text
      data_serv.Recordset("srv_usmod") = WElusuario
      data_serv.Recordset("srv_usfmod") = Date
      data_serv.Recordset("srv_ushmod") = Format(Time, "HH:mm")
      If t_obss.Text <> "" Then
         data_serv.Recordset("srv_obss") = t_obss.Text
      End If
      If mfsusp.Text <> "__/__/____" Then
         data_serv.Recordset("srv_fecsusp") = Format(mfsusp.Text, "dd/mm/yyyy")
      End If
      If t_realiza.Text <> "" Then
         data_serv.Recordset("srv_realiza") = t_realiza.Text
      End If
      data_serv.Recordset("srv_intdom") = Check2.Value
      data_serv.Recordset.Update
      data_serv.Refresh
      data_buscasrv.Refresh
      Text1.Text = data_buscasrv.Recordset("id")
      
      If data_serv.Recordset.RecordCount > 0 Then
         data_serv.Recordset.MoveLast
      End If
      If t_arch.Text <> "" Then
         data_arch.RecordSource = "Select * from arch_orden where id <=" & 1
         data_arch.Refresh
   
         data_arch.Recordset.AddNew
         Set pdffile = New ADODB.Stream
         pdffile.Type = adTypeBinary
         pdffile.Open
         pdffile.LoadFromFile pdfpath
         data_arch.Recordset.Fields("archivo") = pdffile.Read
         data_arch.Recordset("idsrv") = Text1.Text
         data_arch.Recordset("fecha") = Date
         data_arch.Recordset("hora") = Format(Time, "HH:mm")
         data_arch.Recordset.Update
         pdffile.Close
         Set pdffile = Nothing
         frm_srvenferm.MousePointer = 0
         MsgBox "Guardado"
         data_arch.Refresh
      Else
         frm_srvenferm.MousePointer = 0
         MsgBox "Guardado pero sin archivo adjunto de ORDEN"
      End If
      data_buscasrv.Refresh
      frm_srvenferm.MousePointer = 0
      borrarc
      b_nuevo.Enabled = True
      b_edita.Enabled = True
      b_cance.Enabled = False
      b_infos.Enabled = True
      DBGrid1.Enabled = True
      t_busced.Enabled = True
      b_graba.Enabled = False
      Command1.Enabled = True
      b_verord.Enabled = True
      Frame1.Enabled = False
      XAlta = 0
   End If
End If
If XAlta = 2 Then
   If t_ced.Text = "" Or t_nom.Text = "" Or t_conv.Text = "" Or t_direc.Text = "" Or t_codc.Text = "" Or _
      t_local.Text = "" Or t_telef.Text = "" Or t_acto.Text = "" Or t_ref.Text = "" Or t_cant.Text = "" Or _
      t_bol.Text = "" Or mfd.Text = "__/__/____" Or mfh.Text = "__/__/____" Or Combo1.ListIndex < 0 Or Label19.Caption = "" Then
      MsgBox "Faltan ingresar datos para poder grabar, VERIFIQUE!"
   Else
      frm_srvenferm.MousePointer = 11
      data_serv.RecordSource = "Select * from serv_enferm where id =" & Label19.Caption
      data_serv.Refresh
      If data_serv.Recordset.RecordCount > 0 Then
         data_serv.Recordset.Edit
         If IsNull(data_serv.Recordset("srv_fec")) = False Then
            If Format(data_serv.Recordset("srv_fec"), "yyyy/mm/dd") <> Format(mf.Text, "yyyy/mm/dd") Then
               data_serv.Recordset("srv_fec") = mf.Text
            End If
         End If
         data_serv.Recordset("srv_hora") = mh.Text
         data_serv.Recordset("srv_base") = t_b.Text
         data_serv.Recordset("srv_usua") = labu.Caption
         data_serv.Recordset("srv_ced") = t_ced.Text
         data_serv.Recordset("srv_cced") = t_codc.Text
         data_serv.Recordset("srv_nom") = t_nom.Text
         data_serv.Recordset("srv_palia") = Check3.Value
         If labedad.Caption = "" Then
            data_serv.Recordset("srv_edad") = "Sin F.NAC"
         Else
            data_serv.Recordset("srv_edad") = labedad.Caption
         End If
         If t_obss.Text <> "" Then
            data_serv.Recordset("srv_obss") = t_obss.Text
         Else
            If IsNull(data_serv.Recordset("srv_obss")) = False Then
               data_serv.Recordset("srv_obss") = Null
            End If
         End If
         data_serv.Recordset("srv_local") = t_local.Text
         data_serv.Recordset("srv_direc") = t_direc.Text
         data_serv.Recordset("srv_cconv") = t_conv.Text
         data_serv.Recordset("srv_conv") = t_convd.Text
         data_serv.Recordset("srv_telef") = t_telef.Text
         data_serv.Recordset("srv_ref") = t_ref.Text
         data_serv.Recordset("srv_boleta") = t_bol.Text
         data_serv.Recordset("srv_acto") = t_acto.Text
         data_serv.Recordset("srv_fecd") = mfd.Text
         data_serv.Recordset("srv_hord") = mhd.Text
         data_serv.Recordset("srv_fech") = mfh.Text
         data_serv.Recordset("srv_horh") = mhh.Text
         data_serv.Recordset("srv_estad") = Combo1.Text
         data_serv.Recordset("srv_estan") = Combo1.ListIndex
         data_serv.Recordset("srv_verif") = Check1.Value
         data_serv.Recordset("srv_cant") = t_cant.Text
         data_serv.Recordset("srv_usmod") = WElusuario
         data_serv.Recordset("srv_usfmod") = Date
         data_serv.Recordset("srv_ushmod") = Format(Time, "HH:mm")
         If mfsusp.Text <> "__/__/____" Then
            data_serv.Recordset("srv_fecsusp") = Format(mfsusp.Text, "dd/mm/yyyy")
         Else
            If IsNull(data_serv.Recordset("srv_fecsusp")) = False Then
               data_serv.Recordset("srv_fecsusp") = Null
            End If
         End If
         If t_realiza.Text <> "" Then
            data_serv.Recordset("srv_realiza") = t_realiza.Text
         Else
            If IsNull(data_serv.Recordset("srv_realiza")) = False Then
               data_serv.Recordset("srv_realiza") = Null
            End If
         End If
         data_serv.Recordset("srv_intdom") = Check2.Value
         data_serv.Recordset.Update
         data_serv.Refresh
         If t_arch.Text <> "" Then
            data_arch.RecordSource = "Select * from arch_orden where idsrv =" & Label19.Caption
            data_arch.Refresh
            If data_arch.Recordset.RecordCount > 0 Then
               data_arch.Recordset.Delete
               data_arch.Refresh
            End If
            data_arch.Recordset.AddNew
            Set pdffile = New ADODB.Stream
            pdffile.Type = adTypeBinary
            pdffile.Open
            pdffile.LoadFromFile pdfpath
            data_arch.Recordset.Fields("archivo") = pdffile.Read
            data_arch.Recordset("idsrv") = Label19.Caption
            data_arch.Recordset("fecha") = Date
            data_arch.Recordset("hora") = Format(Time, "HH:mm")
            data_arch.Recordset.Update
            pdffile.Close
            Set pdffile = Nothing
            frm_srvenferm.MousePointer = 0
            data_arch.Refresh
            MsgBox "Guardado"
         End If
         data_buscasrv.Refresh
         frm_srvenferm.MousePointer = 0
         borrarc
         b_nuevo.Enabled = True
         b_edita.Enabled = True
         b_cance.Enabled = False
         b_infos.Enabled = True
         DBGrid1.Enabled = True
         t_busced.Enabled = True
         b_graba.Enabled = False
         Command1.Enabled = True
         b_verord.Enabled = True
         Frame1.Enabled = False
         XAlta = 0
      Else
         MsgBox "No se puede grabar, cancele y vuelva a intentar"
         
      End If
   End If
End If

Exit Sub

Xelerrsrvenf:
             If Err.Number = 3155 Then
                MsgBox "No hay modificaciones para grabar, presione el botón de cancelar.", vbInformation
             Else
                MsgBox "Error al grabar, verifique datos " & Err.Number & " " & Err.Description, vbInformation
             End If
             
End Sub

Private Sub b_infos_Click()
Dim Xdesde, Xhasta As String
Dim Ximpelform, CreaPlani As String
CreaPlani = ""

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija, Xcantsrv, Xcanttot As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

XCol = 1
Xlin = 1
Xnrocan = 1
Xcanttot = 0

Ximpelform = MsgBox("Desea imprimir el formulario?", vbYesNo)
If Ximpelform = vbYes Then
   frm_srvenferm.PrintForm
Else
    Xdesde = InputBox("Ingrese desde que fecha")
    Xhasta = InputBox("Ingrese hasta que fecha")
    
    If data_inf.Recordset.RecordCount > 0 Then
       data_inf.Recordset.MoveFirst
       Do While Not data_inf.Recordset.EOF
          data_inf.Recordset.Delete
          data_inf.Recordset.MoveNext
       Loop
    End If
    
    If Xdesde = "" And Xhasta = "" Then
       MsgBox "Debe ingresar fechas para listar"
    Else
       data_parainf.RecordSource = "Select * from serv_enferm where srv_fec >=#" & Format(Xdesde, "yyyy/mm/dd") & "# and srv_fec <=#" & Format(Xhasta, "yyyy/mm/dd") & "# order by id"
       data_parainf.Refresh
       If data_parainf.Recordset.RecordCount > 0 Then
          data_parainf.Recordset.MoveFirst
          CreaPlani = MsgBox("Desea crear planilla excel con actos?", vbInformation + vbYesNo, "Enfermería")
          If CreaPlani = vbYes Then
             Set Xobjexel = New Excel.Application
             Set Xlibexel = Xobjexel.Workbooks.Add
             Set Xarchexel = Xlibexel.Worksheets.Add
             Xarchexel.Name = "SERVICIOS ENF"
             Xlibexel.SaveAs ("C:\planillas\" & "ServiciosEnf" & ".xls")
             Xarchtex = "C:\planillas\" & "ServiciosEnf" & ".xls"
             frm_srvenferm.MousePointer = 11
             Xarchexel.Cells(Xlin, XCol) = "SAPP - CÓMPUTOS"
             Xlin = Xlin + 1
             XCol = XCol + 1
             Xarchexel.Range("A1", "C3").Font.Size = 16
             Xarchexel.Cells(Xlin, XCol) = "INFORME DE SERVICIOS ENFERMERÍA DESDE: " & Xdesde & " HASTA: " & Xhasta
             Xarchexel.Range("B" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
             XCol = 1
             Xlin = Xlin + 2
             Xnrocan = Xnrocan + Xlin
             Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
             Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
             Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
             Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
             Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
             Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
             Xarchexel.Range("A" & Trim(str(Xlin)), "AD" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
             Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
             Xarchexel.Cells(Xlin, XCol) = "ACTO NRO"
             XCol = XCol + 1
             Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 30
             Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
             XCol = XCol + 1
             Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
             Xarchexel.Cells(Xlin, XCol) = "FEC.INICIO"
             XCol = XCol + 1
             Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
             Xarchexel.Cells(Xlin, XCol) = "FECHA FIN"
             XCol = XCol + 1
             Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
             Xarchexel.Cells(Xlin, XCol) = "C/CUANTO?"
             XCol = XCol + 1
             Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 25
             Xarchexel.Cells(Xlin, XCol) = "PROCEDIMIENTO"
             XCol = XCol + 1
             Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
             Xarchexel.Cells(Xlin, XCol) = "ENFERMERA/O"
             XCol = XCol + 1
             Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 20
             Xarchexel.Cells(Xlin, XCol) = "ZONA"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "SAPP/MUT."
             XCol = XCol + 1
             Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 12
             Xarchexel.Cells(Xlin, XCol) = "ACTOS"
             XCol = XCol + 1
             Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
             Xarchexel.Cells(Xlin, XCol) = "INT.DOMI"
             
             Xlin = Xlin + 1
             XCol = 1
          
             Do While Not data_parainf.Recordset.EOF
                Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("id")
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_nom")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("srv_nom")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_fecd")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = "'" & Format(data_parainf.Recordset("srv_fecd"), "dd/mm/yyyy")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_fech")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = "'" & Format(data_parainf.Recordset("srv_fech"), "dd/mm/yyyy")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_obss")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("srv_obss")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_acto")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("srv_acto")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_realiza")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("srv_realiza")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_local")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("srv_local")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_cconv")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("srv_cconv")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_cant")) = False Then
                   Xarchexel.Cells(Xlin, XCol) = data_parainf.Recordset("srv_cant")
                End If
                XCol = XCol + 1
                If IsNull(data_parainf.Recordset("srv_intdom")) = False Then
                   If data_parainf.Recordset("srv_intdom") = 1 Then
                      Xarchexel.Cells(Xlin, XCol) = "SI"
                   Else
                      Xarchexel.Cells(Xlin, XCol) = "NO"
                   End If
                Else
                   Xarchexel.Cells(Xlin, XCol) = "NO"
                End If
                
                data_parainf.Recordset.MoveNext
                Xlin = Xlin + 1
                Xcanttot = Xcanttot + 1
                XCol = 1
             Loop
             Xlin = Xlin + 1
             Xarchexel.Cells(Xlin, XCol) = "TOTAL GENERAL: " & Xcanttot
             Xlibexel.Save
             Xlibexel.Close
             Xobjexel.Quit
             Xlabrir.Workbooks.Open Xarchtex, , False
             Xlabrir.Visible = True
             Xlabrir.WindowState = xlMaximized
             frm_srvenferm.MousePointer = 0
             MsgBox "Proceso terminado"
          Else
             Do While Not data_parainf.Recordset.EOF
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_fnac") = data_parainf.Recordset("srv_fec")
                data_inf.Recordset("cl_nrocobr") = data_parainf.Recordset("srv_base")
                data_inf.Recordset("cl_cedula") = data_parainf.Recordset("srv_ced")
                data_inf.Recordset("cl_codced") = data_parainf.Recordset("srv_cced")
                data_inf.Recordset("cl_apellid") = Mid(data_parainf.Recordset("srv_nom"), 1, 50)
                data_inf.Recordset("cl_codconv") = data_parainf.Recordset("srv_cconv")
                data_inf.Recordset("cl_direcci") = Mid(data_parainf.Recordset("srv_direc"), 1, 50)
                data_inf.Recordset("cl_telefon") = Mid(data_parainf.Recordset("srv_telef"), 1, 20)
                If IsNull(data_parainf.Recordset("srv_palia")) = False Then
                   If data_parainf.Recordset("srv_palia") = 1 Then
                      data_inf.Recordset("info_debit") = data_parainf.Recordset("srv_acto") & "--Servs.Paliativos"
                   Else
                      data_inf.Recordset("info_debit") = data_parainf.Recordset("srv_acto")
                   End If
                Else
                   data_inf.Recordset("info_debit") = data_parainf.Recordset("srv_acto")
                End If
                data_inf.Recordset("cl_nrovend") = data_parainf.Recordset("srv_cant")
                data_inf.Recordset("cl_nomcobr") = Mid(data_parainf.Recordset("srv_estad"), 1, 20)
                data_inf.Recordset("cl_email") = Mid(data_parainf.Recordset("srv_realiza"), 1, 30)
                data_inf.Recordset("cl_nom_sup") = Mid(data_parainf.Recordset("srv_edad"), 1, 20)
                data_inf.Recordset.Update
                data_parainf.Recordset.MoveNext
             Loop
             MsgBox "Proceso terminado"
             cr1.ReportFileName = App.path & "\infsrvenf.rpt"
             cr1.ReportTitle = "Informe de Actos de enfermería desde: " & Xdesde & " Hasta: " & Xhasta
             cr1.Action = 1
          End If
       Else
          MsgBox "No existen registros"
       End If
    
    End If
End If

End Sub

Private Sub b_nuevo_Click()
Frame1.Enabled = True
b_nuevo.Enabled = False
b_edita.Enabled = False
b_cance.Enabled = True
b_infos.Enabled = False
DBGrid1.Enabled = False
t_busced.Enabled = False
b_graba.Enabled = True
Command1.Enabled = False
b_verord.Enabled = False
XAlta = 1
mf.Text = Date
mh.Text = Format(Time, "HH:mm")
t_b.Text = frm_menu.data_parse.Recordset("base")
labu.Caption = WElusuario
t_ced.SetFocus


End Sub

Private Sub b_sube_Click()
With cm1
     .FileName = ""
     .Filter = "PDF (*.pdf;) | *.pdf;"
     .ShowOpen
     If Len(.FileName) <> 0 Then
        pdfpath = .FileName
        pdfpath1 = .FileTitle
        t_arch.Text = .FileTitle
     Else
        t_arch.Text = ""
     End If
End With


End Sub

Private Sub b_verord_Click()
Dim x, Xbandlab As Integer
Dim Xlac As String

Xlac = ""
Xbandlab = 0
On Error GoTo Noestaarch

frm_srvenferm.MousePointer = 11

b_verord.Enabled = False

If Dir(App.path & "\laboratorio\temporal.pdf") <> "" Then

   Kill App.path & "\laboratorio\temporal.pdf"
End If

data_abre.Recordset.Edit
data_abre.Recordset("numero") = data_buscasrv.Recordset("id")
data_abre.Recordset.Update
data_abre.Refresh
data_arch.RecordSource = "Select * from arch_orden where idsrv =" & data_abre.Recordset("numero")
data_arch.Refresh
If data_arch.Recordset.RecordCount > 0 Then
   Shell App.path & "\archenf.exe", vbMinimizedFocus
   ShellExecute Me.hwnd, "open", App.path & "\laboratorio\temporal.pdf", "", "", 4

'   Shell data_desc.Recordset("desc") & " " & App.Path & "\laboratorio\temporal" & ".pdf", vbMaximizedFocus
Else
   frm_srvenferm.MousePointer = 0
   MsgBox "No tiene archivo escaneado"
End If
frm_srvenferm.MousePointer = 0

b_verord.Enabled = True

Exit Sub

Noestaarch:
           If Err.Number = 53 Then
              MsgBox "No se ecuentra el archivo, verifique"
           Else
              MsgBox "Error al cargar el archivo"
           End If
           

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_sube.SetFocus
End If

End Sub

Private Sub Command1_Click()
If Label19.Caption <> "" Then
    If WElusuario = "MCURBELO" Or WElusuario = "SORTEGA" Or WElusuario = "JFERNAN" Or WElusuario = "SDOMINGUEZ" Then
       frm_obsenfadm.Show vbModal
    Else
       MsgBox "Usuario no autorizado"
    End If
End If
End Sub

Private Sub DBGrid1_DblClick()
borrarc
mf.Text = data_buscasrv.Recordset("srv_fec")
Label19.Caption = data_buscasrv.Recordset("id")
mh.Text = data_buscasrv.Recordset("srv_hora")
t_b.Text = data_buscasrv.Recordset("srv_base")
labu.Caption = data_buscasrv.Recordset("srv_usua")
t_ced.Text = data_buscasrv.Recordset("srv_ced")
t_codc.Text = data_buscasrv.Recordset("srv_cced")
t_nom.Text = data_buscasrv.Recordset("srv_nom")
If IsNull(data_buscasrv.Recordset("srv_cconv")) = False Then
   t_conv.Text = data_buscasrv.Recordset("srv_cconv")
Else
   t_conv.Text = "AA"
End If
If IsNull(data_buscasrv.Recordset("srv_conv")) = False Then
   t_convd.Text = data_buscasrv.Recordset("srv_conv")
Else
   t_convd.Text = "BB"
End If
t_local.Text = data_buscasrv.Recordset("srv_local")
t_direc.Text = data_buscasrv.Recordset("srv_direc")
t_telef.Text = data_buscasrv.Recordset("srv_telef")
t_ref.Text = data_buscasrv.Recordset("srv_ref")
t_bol.Text = data_buscasrv.Recordset("srv_boleta")
t_acto.Text = data_buscasrv.Recordset("srv_acto")
If IsNull(data_buscasrv.Recordset("srv_realiza")) = False Then
   t_realiza.Text = data_buscasrv.Recordset("srv_realiza")
Else
   t_realiza.Text = ""
End If
If IsNull(data_buscasrv.Recordset("srv_palia")) = False Then
   Check3.Value = data_buscasrv.Recordset("srv_palia")
Else
   Check3.Value = 0
End If
If IsNull(data_buscasrv.Recordset("srv_obss")) = False Then
   t_obss.Text = data_buscasrv.Recordset("srv_obss")
Else
   t_obss.Text = ""
End If

If IsNull(data_buscasrv.Recordset("srv_fecsusp")) = False Then
   mfsusp.Text = Format(data_buscasrv.Recordset("srv_fecsusp"), "dd/mm/yyyy")
Else
   mfsusp.Text = "__/__/____"
End If

If IsNull(data_buscasrv.Recordset("srv_fecd")) = False Then
   mfd.Text = data_buscasrv.Recordset("srv_fecd")
Else
   mfd.Text = "__/__/____"
End If
If IsNull(data_buscasrv.Recordset("srv_hord")) = False Then
   mhd.Text = data_buscasrv.Recordset("srv_hord")
Else
   mhd.Text = "__:__"
End If
If IsNull(data_buscasrv.Recordset("srv_intdom")) = False Then
   Check2.Value = data_buscasrv.Recordset("srv_intdom")
Else
   Check2.Value = 0
End If
mfh.Text = data_buscasrv.Recordset("srv_fech")
If IsNull(data_buscasrv.Recordset("srv_horh")) = False Then
   mhh.Text = data_buscasrv.Recordset("srv_horh")
Else
   mhh.Text = "__:__"
End If
Combo1.ListIndex = data_buscasrv.Recordset("srv_estan")
Check1.Value = data_buscasrv.Recordset("srv_verif")
t_arch.Text = ""
t_cant.Text = data_buscasrv.Recordset("srv_cant")
If IsNull(data_buscasrv.Recordset("srv_edad")) = False Then
   labedad.Caption = data_buscasrv.Recordset("srv_edad")
Else
   labedad.Caption = "S/D"
End If


End Sub

Private Sub Form_Load()
data_serv.Connect = "ODBC;DSN=sappespecial;"
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_buscasrv.Connect = "ODBC;DSN=sappespecial;"
data_buscasrv.RecordSource = "Select * from serv_enferm order by id DESC"
data_buscasrv.Refresh

data_arch.Connect = "ODBC;DSN=sappespecial;"

data_abre.DatabaseName = App.path & "\abrir.mdb"
data_abre.RecordSource = "abrir"
data_abre.Refresh

data_desc.DatabaseName = App.path & "\desc.mdb"
data_desc.RecordSource = "desc"
data_desc.Refresh

data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh

data_parainf.Connect = "ODBC;DSN=sappespecial;"

'item internación domiciliaria
'nombre de enfermer asignado
'observaciones general
'cant. act inicial y final
'busqueda por fecha
'asignado a móvil que lo vea despacho
'filtro de actos terminados


End Sub

Public Sub borrarc()
mf.Text = "__/__/____"
mh.Text = "__:__"
t_b.Text = ""
labu.Caption = ""
t_ced.Text = ""
t_codc.Text = ""
t_nom.Text = ""
mfsusp.Text = "__/__/____"
t_realiza.Text = ""
t_conv.Text = ""
t_convd.Text = ""
t_local.Text = ""
t_direc.Text = ""
t_telef.Text = ""
t_ref.Text = ""
t_bol.Text = ""
t_acto.Text = ""
Check3.Value = 0
mfd.Text = "__/__/____"
mhd.Text = "__:__"
mfh.Text = "__/__/____"
mhh.Text = "__:__"
Combo1.ListIndex = -1
Check1.Value = 0
t_arch.Text = ""
t_cant.Text = ""
labedad.Caption = ""
Check2.Value = 0
t_obss.Text = ""

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhd.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhh.SetFocus
End If

End Sub

Private Sub mhd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mhh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cant.SetFocus
End If

End Sub

Private Sub t_acto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfd.SetFocus
End If

End Sub

Private Sub t_bol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_acto.SetFocus
End If

End Sub

Private Sub t_busced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_busced.Text <> "" Then
      data_buscasrv.RecordSource = "Select * from serv_enferm where srv_ced =" & t_busced.Text & " order by id DESC"
      data_buscasrv.Refresh
   Else
      data_buscasrv.RecordSource = "Select * from serv_enferm order by id DESC"
      data_buscasrv.Refresh
   End If
End If

End Sub

Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_ced_LostFocus()
If t_ced.Text <> "" Then
   data_cli.RecordSource = "Select * from clientes where cl_cedula =" & t_ced.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      t_codc.Text = data_cli.Recordset("cl_codced")
      t_nom.Text = data_cli.Recordset("cl_apellid")
      t_conv.Text = data_cli.Recordset("cl_codconv")
      If IsNull(data_cli.Recordset("cl_fnac")) = False Then
         CalculaEdad (data_cli.Recordset("cl_fnac"))
      Else
         labedad.Caption = "S/D"
      End If
      t_convd.Text = data_cli.Recordset("cl_nomconv")
      If IsNull(data_cli.Recordset("cl_zona")) = False Then
         t_local.Text = data_cli.Recordset("cl_zona")
      End If
      If IsNull(data_cli.Recordset("cl_direcci")) = False Then
         t_direc.Text = data_cli.Recordset("cl_direcci")
      End If
      If IsNull(data_cli.Recordset("cl_telefon")) = False Then
         If IsNull(data_cli.Recordset("cl_dpto")) = False Then
            t_telef.Text = data_cli.Recordset("cl_telefon") & "//" & data_cli.Recordset("cl_dpto")
         Else
            t_telef.Text = data_cli.Recordset("cl_telefon")
         End If
      Else
         If IsNull(data_cli.Recordset("cl_dpto")) = False Then
            t_telef.Text = data_cli.Recordset("cl_dpto")
         End If
      End If
      t_nom.SetFocus
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
   labedad.Caption = Anios & " AÑOS"
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
'   labunie.Caption = Meses
'   labdias.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   labedad.Caption = "S/D"
End If

End Sub


Private Sub t_conv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_local.SetFocus
End If

End Sub

Private Sub t_direc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_telef.SetFocus
End If

End Sub

Private Sub t_local_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_direc.SetFocus
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_conv.SetFocus
End If

End Sub

Private Sub t_ref_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_bol.SetFocus
End If

End Sub

Private Sub t_telef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_ref.SetFocus
End If

End Sub
