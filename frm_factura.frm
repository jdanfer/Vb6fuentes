VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_factura 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Facturación"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_lindbgri 
      Caption         =   "data_lindbgri"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_aut 
      Caption         =   "data_aut"
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_estudiobus 
      Caption         =   "data_estudiobus"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_facafil 
      Caption         =   "data_facafil"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_func 
      Caption         =   "data_func"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBCtls.DBCombo dbcboprom 
      Bindings        =   "frm_factura.frx":0000
      Height          =   570
      Left            =   360
      TabIndex        =   65
      Top             =   2400
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1005
      _Version        =   393216
      Style           =   1
      ForeColor       =   12582912
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data data_cablocal 
      Caption         =   "data_cablocal"
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
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_errfact 
      Caption         =   "data_errfact"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Deudas"
      Height          =   375
      Left            =   9120
      TabIndex        =   60
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tim emi"
      Height          =   375
      Left            =   9360
      TabIndex        =   58
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data data_eror 
      Caption         =   "data_eror"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_imagen 
      Caption         =   "data_imagen"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      DataField       =   "qr"
      DataSource      =   "data_imagen"
      Height          =   1695
      Left            =   3120
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   48
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5280
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowControlBox=   0   'False
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      DiscardSavedData=   -1  'True
      WindowState     =   1
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowZoomCtl=   0   'False
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.CommandButton b_ncefct 
      Caption         =   "b_ncefct"
      Height          =   495
      Left            =   2520
      TabIndex        =   47
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton b_ncetck 
      Caption         =   "b_ncetck"
      Height          =   495
      Left            =   1680
      TabIndex        =   46
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data data_lincance 
      Caption         =   "data_lincance"
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
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton b_verfaccance 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5640
      Picture         =   "frm_factura.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Buscar factura para cancelar"
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "e-fct"
      Height          =   495
      Left            =   960
      TabIndex        =   40
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Data data_lincab 
      Caption         =   "data_lincab"
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
      Top             =   6000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox mf 
      Height          =   375
      Left            =   8280
      TabIndex        =   39
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   8454143
      ForeColor       =   16711680
      Enabled         =   0   'False
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
   Begin VB.Data data_param 
      Caption         =   "data_param"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_cabeza2 
      Caption         =   "data_cabeza2"
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
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "rut"
      Height          =   615
      Left            =   5760
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIN"
      Height          =   495
      Left            =   4080
      TabIndex        =   36
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton b_etck 
      Caption         =   "e-tck"
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data data_ui 
      Caption         =   "data_ui"
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data data_ctr 
      Caption         =   "data_ctr"
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
      Top             =   2760
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
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox t_codaut 
      Height          =   285
      Left            =   3600
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_u 
      Caption         =   "data_u"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_consdeu 
      Caption         =   "data_consdeu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_er 
      Caption         =   "data_er"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo dbcbomedo 
      Bindings        =   "frm_factura.frx":05A2
      Height          =   600
      Left            =   6480
      TabIndex        =   30
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1058
      _Version        =   393216
      Style           =   1
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txt_rut 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2640
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FACTURA CON RUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancelar Factura"
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
      Left            =   8280
      MouseIcon       =   "frm_factura.frx":05BD
      MousePointer    =   99  'Custom
      Picture         =   "frm_factura.frx":08C7
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_codcaja 
      Caption         =   "data_codcaja"
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
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_deudas 
      Caption         =   "data_deudas"
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
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_arancel 
      Caption         =   "data_arancel"
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_linmmdd 
      Caption         =   "data_linmmdd"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton btn_fin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Terminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Picture         =   "frm_factura.frx":0D09
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5880
      Width           =   975
   End
   Begin VB.Data data_medicos 
      Caption         =   "data_medicos"
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
      Top             =   5880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBCtls.DBCombo dbcbomed 
      Bindings        =   "frm_factura.frx":1293
      Height          =   600
      Left            =   6480
      TabIndex        =   16
      Top             =   2040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1058
      _Version        =   393216
      Style           =   1
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txt_ano 
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
      Left            =   8880
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txt_mes 
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
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox cbotim 
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
      ItemData        =   "frm_factura.frx":12AE
      Left            =   8040
      List            =   "frm_factura.frx":12B8
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txt_precio 
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
      Left            =   6480
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Data data_estudio 
      Caption         =   "data_estudio"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data data_lineas 
      Caption         =   "data_lineas"
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
      Top             =   6120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_factura.frx":12C4
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frm_factura.frx":12E0
      TabIndex        =   4
      Top             =   3840
      Width           =   9975
   End
   Begin VB.CommandButton btn_graba 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00808080&
      Picture         =   "frm_factura.frx":2513
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Data data_caja 
      Caption         =   "data_caja"
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
      Top             =   6360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "frm_factura.frx":2A9D
      Height          =   660
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1164
      _Version        =   393216
      Style           =   1
      ListField       =   ""
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
   End
   Begin VB.TextBox t_cant 
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
      Height          =   405
      Left            =   1680
      TabIndex        =   62
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label labmedicacion 
      Height          =   375
      Left            =   4800
      TabIndex        =   69
      Top             =   6360
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label labcorreo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   0
      TabIndex        =   68
      Top             =   1560
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.Label labnomprom 
      Height          =   255
      Left            =   360
      TabIndex        =   67
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label labcodpro 
      BackColor       =   &H00C00000&
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
      Left            =   2040
      TabIndex        =   66
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C00000&
      Caption         =   "PROMOTOR:"
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
      Left            =   360
      TabIndex        =   64
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labdevol 
      Height          =   255
      Left            =   6600
      TabIndex        =   63
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label13 
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
      Left            =   360
      TabIndex        =   61
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label labmotivo 
      Height          =   255
      Left            =   4800
      TabIndex        =   59
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label labidemi 
      Height          =   375
      Left            =   9720
      TabIndex        =   57
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labdeudaemi 
      Height          =   375
      Left            =   6960
      TabIndex        =   56
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label labtimemi 
      Height          =   255
      Left            =   8520
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label labfeccance 
      Height          =   375
      Left            =   2880
      TabIndex        =   54
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label labvenceok 
      Height          =   135
      Left            =   2280
      TabIndex        =   53
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labvence 
      Height          =   375
      Left            =   3600
      TabIndex        =   52
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labhasta 
      Height          =   255
      Left            =   3600
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label labdesde 
      Height          =   255
      Left            =   3600
      TabIndex        =   50
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label labautoriza 
      Height          =   255
      Left            =   3600
      TabIndex        =   49
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label labmonecance 
      Height          =   255
      Left            =   3600
      TabIndex        =   44
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labseriecance 
      Height          =   255
      Left            =   4560
      TabIndex        =   43
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labfaccance 
      Height          =   255
      Left            =   2880
      TabIndex        =   42
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lablinea 
      Height          =   255
      Left            =   3000
      TabIndex        =   41
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labserie 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   37
      Top             =   0
      Width           =   375
   End
   Begin VB.Label labfpago 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6360
      TabIndex        =   34
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Forma de pago:"
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
      Left            =   4440
      TabIndex        =   33
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label labtimme 
      Height          =   255
      Left            =   2760
      TabIndex        =   32
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labmedo 
      BackColor       =   &H00C00000&
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
      Left            =   8880
      TabIndex        =   29
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C00000&
      Caption         =   "MEDICO QUE ORDENA"
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
      Left            =   6480
      TabIndex        =   28
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "TOTAL:"
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
      Height          =   375
      Left            =   7080
      TabIndex        =   24
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "IVA 10%:"
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
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   5040
      TabIndex        =   22
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label labfac 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Factura:"
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
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label labtot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   8160
      TabIndex        =   18
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   10320
      X2              =   0
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label labmed 
      BackColor       =   &H00FFC0C0&
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
      Left            =   8880
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   6360
      X2              =   6360
      Y1              =   600
      Y2              =   3720
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "MEDICO QUE REALIZA"
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
      Left            =   6480
      TabIndex        =   15
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label labtim 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
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
      Left            =   8880
      TabIndex        =   14
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
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
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "MES/AÑO PAGO:"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "TIMBRE"
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
      Left            =   8040
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "PRECIO"
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
      Left            =   6480
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "SERVICIO:"
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
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label labnomb 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label labmatri 
      BackColor       =   &H00FFFFC0&
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
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1200
      Picture         =   "frm_factura.frx":2AB8
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
End
Attribute VB_Name = "frm_factura"
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

Public Xquehag, Xcoddeu As Integer


Private Sub b_cance_Click()
On Error GoTo Alcancelar

frmabm.btn_fact.Enabled = True

If data_lineas.Recordset.RecordCount > 0 Then
   data_lineas.Recordset.MoveFirst
   Do While Not data_lineas.Recordset.EOF
      data_lineas.Recordset.Delete
      data_lineas.Recordset.MoveNext
   Loop
   data_lineas.Refresh
   
End If
If data_cabeza2.Recordset.RecordCount > 0 Then
   data_cabeza2.Recordset.MoveFirst
   Do While Not data_cabeza2.Recordset.EOF
      data_cabeza2.Recordset.Delete
      data_cabeza2.Recordset.MoveNext
   Loop
   data_cabeza2.Refresh
End If

data_eror.DatabaseName = App.path & "\selec.mdb"
data_eror.RecordSource = "selec"
data_eror.Refresh
If data_eror.Recordset.RecordCount > 0 Then
   data_eror.Recordset.MoveFirst
   Do While Not data_eror.Recordset.EOF
      data_eror.Recordset.Delete
      data_eror.Recordset.MoveNext
   Loop
End If
data_eror.DatabaseName = App.path & "\erores.mdb"
data_eror.RecordSource = "erores"
data_eror.Refresh

Xquehag = 0
Xcoddeu = 0
t_codaut.Text = ""

labtim.Caption = ""
labmed.Caption = ""
txt_precio.Text = 0
cbotim.ListIndex = 0
dbcbomed.Text = ""
DBCombo1.Text = ""
txt_mes.Text = ""
txt_ano.Text = ""
If DBCombo1.Enabled = True Then
   DBCombo1.SetFocus
End If
XQuefac = 0
Unload Me
Exit Sub

Alcancelar:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cancelar"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cancelar"
            data_errfact.Recordset.Update
            Unload Me
         End If


End Sub

Private Sub b_etck_Click()
Dim strIdTransac As String
Dim Xn As Integer
Xn = 0
On Error GoTo Xquepasaalenvio

Set objPosCfe = New PosCfe
    
Dim objresultado As Resultado
If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
   If frm_menu.data_parse.Recordset("base") = 78 Then 'Notebook JF
      Set objresultado = objPosCfe.Inicializar("SAPP-105", frm_menu.data_parse.Recordset("contacto"), vbNullString)
   Else
      Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
   End If
Else
    If frm_menu.data_parse.Recordset("base") = 1 Then
       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-301", vbNullString)
    Else
       If frm_menu.data_parse.Recordset("base") = 2 Then
          Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-002", vbNullString)
       Else
          If frm_menu.data_parse.Recordset("base") = 3 Then
             Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-001", vbNullString)
          Else
             If frm_menu.data_parse.Recordset("base") = 4 Then
                Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-004", vbNullString)
             Else
                If frm_menu.data_parse.Recordset("base") = 6 Then
                   Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-006", vbNullString)
                Else
                   If frm_menu.data_parse.Recordset("base") = 8 Then
                      Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-008", vbNullString)
                   Else
                      If frm_menu.data_parse.Recordset("base") = 10 Then
                         Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-010", vbNullString)
                      Else
                         If frm_menu.data_parse.Recordset("base") = 13 Then
                            Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-013", vbNullString)
                         Else
                            If frm_menu.data_parse.Recordset("base") = 16 Then
                               Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-116", vbNullString)
                            Else
                               If frm_menu.data_parse.Recordset("base") = 91 Then
                                  Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-216", vbNullString)
                               Else
                                  If frm_menu.data_parse.Recordset("base") = 17 Then
                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-017", vbNullString)
                                  Else
                                     If frm_menu.data_parse.Recordset("base") = 18 Then
                                        Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-018", vbNullString)
                                     Else
                                        If frm_menu.data_parse.Recordset("base") = 96 Then
                                           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-218", vbNullString)
                                        Else
                                           If frm_menu.data_parse.Recordset("base") = 38 Then
                                              Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
                                           Else
                                              If frm_menu.data_parse.Recordset("base") = 11 Then
                                                 Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-111", vbNullString)
                                              Else
                                                 If frm_menu.data_parse.Recordset("base") = 93 Then
                                                    Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-217", vbNullString)
                                                 Else
                                                     If frm_menu.data_parse.Recordset("base") = 33 Or frm_menu.data_parse.Recordset("base") = 34 Then  ' B3 adm
                                                        If frm_menu.data_parse.Recordset("base") = 33 Then
                                                           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-333", vbNullString)
                                                        Else
                                                           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-334", vbNullString)
                                                        End If
                                                     Else
                                                        If frm_menu.data_parse.Recordset("base") = 92 Then
                                                           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-318", vbNullString)
                                                        Else
                                                           Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
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
                      End If
                   End If
                End If
             End If
          End If
       End If
    End If
End If

'data_temp.Recordset.MoveFirst
Dim strMensaje As String
strMensaje = "No se pudo inicializar el POS"
If objresultado Is Nothing Then
   If strMensaje = "No se pudo inicializar el POS" Then
      MsgBox strMensaje
      Exit Sub
      
   End If
End If
Xn = 1
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
Xn = 2
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
Xn = 3
'MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'       "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
       
    'Enviando
If Not EstaInicializado() Then Exit Sub
    
Dim objCfe As CFE
Set objCfe = New CFE

Dim objCf As ClassFactory

Set objCf = New ClassFactory
       
Set objCfe.ETck = New ETck
With objCfe.ETck.Encabezado.IdDoc
     .TipoCFE = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
     .FchEmis.SetDate Year(data_cabeza2.Recordset("cl_fnac")), Month(data_cabeza2.Recordset("cl_fnac")), Day(data_cabeza2.Recordset("cl_fnac"))
     .IsValidMntBruto = True
     .MntBruto = IdDoc_Tck_MntBruto_1
     If data_cabeza2.Recordset("cl_forpago") = 1 Then
        .FmaPago = IdDoc_Tck_FmaPago_1
     Else
        .FmaPago = IdDoc_Tck_FmaPago_2
     End If
End With
Xn = 4
With objCfe.ETck.Encabezado.Emisor
     .RUCEmisor = data_param.Recordset("ruc")
     .RznSoc = data_param.Recordset("nomc")
     .CdgDGISucur.FromString Trim(str(data_param.Recordset("codsuc")))
     .DomFiscal = data_param.Recordset("domic")
     .Ciudad = data_param.Recordset("ciudad")
     .Departamento = data_param.Recordset("dpto")
End With
Xn = 5
Set objCfe.ETck.Encabezado.Receptor = New Receptor_Tck
Set objCfe.ETck.Encabezado.Receptor.Receptor_Tck_Choice = New Receptor_Tck_Choice
    With objCfe.ETck.Encabezado.Receptor
        If data_cabeza2.Recordset("cl_nro_sup") = 4 Then
           .TipoDocRecep = DocType_4
        Else
           If data_cabeza2.Recordset("cl_nro_sup") = 3 Then
              .TipoDocRecep = DocType_3
           Else
              If data_cabeza2.Recordset("cl_nro_sup") = 2 Then
                 .TipoDocRecep = DocType_2
              Else
                 .TipoDocRecep = DocType_4
              End If
           End If
        End If
        .CodPaisRecep = CodPaisType_UY
        .Receptor_Tck_Choice.DocRecepExt = data_cabeza2.Recordset("cl_nom_sup")
'        .CodPaisRecep = CodPaisType_UY
        If data_cabeza2.Recordset("cl_nro_sup") = 3 Or data_cabeza2.Recordset("cl_nro_sup") = 2 Then
           If IsNull(data_cabeza2.Recordset("cl_nom_sup")) = False Then
              .Receptor_Tck_Choice.DocRecep = data_cabeza2.Recordset("cl_nom_sup")
           Else
              .Receptor_Tck_Choice.DocRecep = "0"
           End If
        End If
'        .Receptor_Tck_Choice.DocRecepExt = data_cabezal.Recordset("cl_nom_sup")
        .RznSocRecep = data_cabeza2.Recordset("info_debit")
        .DirRecep = data_cabeza2.Recordset("cl_direcci")
        .CiudadRecep = data_cabeza2.Recordset("cl_zona")
    End With
    Xn = 6
    With objCfe.ETck.Encabezado.Totales
        .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
        .IsValidTpoCambio = True
        .TpoCambio.FromString "1"
        .IsValidMntNetoIvaTasaMin = True
        .IsValidMntIVATasaMin = True
        .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
        .IVATasaMin = TasaIVAType_10FullStop000
        .MntIVATasaMin.FromString Format(data_cabeza2.Recordset("saldo_cc"), "0.00")
        If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
           If data_cabeza2.Recordset("cl_cedula") > 0 Then
              .IsValidMntNoGrv = True
              .MntNoGrv.FromString Format(data_cabeza2.Recordset("cl_cedula"), "0.00")
           End If
        End If
        .CantLinDet.FromString Trim(str(data_cabeza2.Recordset("cl_grupo")))
        .MntTotal.FromString Format(data_cabeza2.Recordset("saldo_doc"), "0.00")
        .MntPagar.FromString Format(data_cabeza2.Recordset("saldo_chc"), "0.00")
    End With
    Xn = 8
    Do While Not data_lineas.Recordset.EOF
       With objCfe.ETck.Detalle.Item.AddNew
          .NroLinDet.FromString Trim(str(data_lineas.Recordset("linea")))
          .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_lineas.Recordset("tipo_mov"))))
          .NomItem = data_lineas.Recordset("nom_prod")
          If data_lineas.Recordset("cod_prod") = 60103 Or data_lineas.Recordset("cod_prod") = 60107 Then
             .IsValidDscItem = True
             .DscItem = data_lineas.Recordset("nom_medic")
          End If
          .cantidad.FromString Trim(str(data_lineas.Recordset("cantidad")))
          .UniMed = "N/A"
          .PrecioUnitario.FromString Format(data_lineas.Recordset("arancel"), "0.00")
          .MontoItem.FromString Format(data_lineas.Recordset("tot_lin"), "0.00")
       End With
       data_lineas.Recordset.MoveNext
    Loop
    Xn = 7
    Dim s As String
    s = objCfe.ToXml(True, XmlFormatting_Indented)
    Dim texto As String
    Dim strGuid As String
    strGuid = objPosCfe.CrearGuid()
    Dim objResultadoCfe As ResultadoCfe
'    Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    
    If IsNull(data_cabeza2.Recordset("obsp")) = False Then
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))
    Else
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    End If

    Xn = 9
    If Val(objResultadoCfe.EstadoCfe.CodigoRespuesta) = 0 Or _
       Val(objResultadoCfe.EstadoCfe.CodigoRespuesta) = 11 Then

       labserie.Caption = objResultadoCfe.EstadoCfe.SerieNumeroCfe.serie
       labfac.Caption = objResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
       data_cabeza2.Recordset.Edit
       labvence.Caption = CStr(objResultadoCfe.EstadoCfe.DatosCae.Vencimiento)
       labautoriza.Caption = CStr(objResultadoCfe.EstadoCfe.DatosCae.Autorizacion)
       labdesde.Caption = labserie.Caption & " " & CStr(objResultadoCfe.EstadoCfe.DatosCae.NumeroDesde) & " - " & labserie.Caption & " " & CStr(objResultadoCfe.EstadoCfe.DatosCae.NumeroHasta)
       labhasta.Caption = CStr(objResultadoCfe.EstadoCfe.CodigoSeguridad)
       If Len(labvence.Caption) = 8 Then
          labvenceok.Caption = Mid(labvence.Caption, 7, 2) & "/" & Mid(labvence.Caption, 5, 2) & "/" & Mid(labvence.Caption, 1, 4)
       Else
          labvenceok.Caption = "31/12/2016"
       End If
       If labvenceok.Caption <> "" Then
          data_cabeza2.Recordset("cl_fultpag") = CDate(labvenceok.Caption)
       Else
          data_cabeza2.Recordset("cl_fultpag") = CDate("01/01/2018")
       End If
       If labautoriza.Caption <> "" Then
          data_cabeza2.Recordset("cl_nrocobr") = Val(labautoriza.Caption)
       Else
          data_cabeza2.Recordset("cl_nrocobr") = 0
       End If
       data_cabeza2.Recordset("cl_medflia") = Trim(labdesde.Caption)
       data_cabeza2.Recordset("cl_fax") = Trim(labhasta.Caption)
       data_cabeza2.Recordset.Update
       Xn = 10
       Dim objResultado44 As ResultadoObtenerQr
       Set objResultado44 = objPosCfe.ObtenerQr(objResultadoCfe.EstadoCfe.DatosQr, 100)
       
       Dim strFile As String
       strFile = App.path & "\qr.bmp"
       Dim f As Long
       f = FreeFile()
       Open strFile For Binary As #f
       Put #f, , objResultado44.ImagenQr
       Close #f

 '''   imgQr.Picture = LoadPicture(strFile)

        Set objUltimaSerieNumero = Nothing
 
''''''''''    DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
        If Not objUltimaSerieNumero Is Nothing Then _
        ' cmdFirmarNc.Enabled = True
'       MsgBox "firmar NC"
        End If
    
        Command1_Click
    Else
       Xn = 11
        MsgBox "Factura rechazada, VERIFIQUE!!", vbInformation
       data_eror.Recordset.AddNew
       data_eror.Recordset("nro") = 11
       data_eror.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
       data_eror.Recordset("hora") = Format(Time, "HH:mm")
       data_eror.Recordset("obs") = "FACT CANCE"
       data_eror.Recordset.Update
       MsgBox "Comprobante RECHAZADO, NO FUE ACEPTADO, debe realizarlo nuevamente, verifique datos!", vbInformation
       End
    End If
        
Exit Sub

Xquepasaalenvio:
              If Err.Number = 3155 Then
                 MsgBox "Error al terminar la factura: " & Err.Number & " " & Err.Description, vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = Err.Number
                 data_errfact.Recordset("desc") = Trim(str(Xn)) & Mid(Err.Description, 1, 110)
                 data_errfact.Recordset.Update
              Else
                 MsgBox "Error al terminar la factura:" & Err.Number & " " & Err.Description, vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = Err.Number
                 data_errfact.Recordset("desc") = Trim(str(Xn)) & Mid(texto, 1, 110)
                 data_errfact.Recordset.Update
              End If
              Unload Me

End Sub

Private Sub b_ncefct_Click()
Dim strIdTransac As String

Set objPosCfe = New PosCfe
    
Dim objresultado As Resultado
Dim tipo As Integer
tipo = 0
If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
   Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
Else
    If frm_menu.data_parse.Recordset("base") = 1 Then
       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-301", vbNullString)
    Else
       If frm_menu.data_parse.Recordset("base") = 2 Then
          Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-002", vbNullString)
       Else
          If frm_menu.data_parse.Recordset("base") = 3 Then
             Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-001", vbNullString)
          Else
             If frm_menu.data_parse.Recordset("base") = 4 Then
                Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-004", vbNullString)
             Else
                If frm_menu.data_parse.Recordset("base") = 6 Then
                   Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-006", vbNullString)
                Else
                   If frm_menu.data_parse.Recordset("base") = 8 Then
                      Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-008", vbNullString)
                   Else
                      If frm_menu.data_parse.Recordset("base") = 10 Then
                         Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-010", vbNullString)
                      Else
                         If frm_menu.data_parse.Recordset("base") = 13 Then
                            Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-013", vbNullString)
                         Else
                            If frm_menu.data_parse.Recordset("base") = 16 Then
                               Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-116", vbNullString)
                            Else
                               If frm_menu.data_parse.Recordset("base") = 91 Then
                                  Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-216", vbNullString)
                               Else
                                  If frm_menu.data_parse.Recordset("base") = 17 Then
                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-017", vbNullString)
                                  Else
                                     If frm_menu.data_parse.Recordset("base") = 18 Then
                                        Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-018", vbNullString)
                                     Else
                                        If frm_menu.data_parse.Recordset("base") = 96 Then
                                           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-218", vbNullString)
                                        Else
                                           If frm_menu.data_parse.Recordset("base") = 38 Then
                                              Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
                                           Else
                                              If frm_menu.data_parse.Recordset("base") = 11 Then
                                                 Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-111", vbNullString)
                                              Else
                                                 If frm_menu.data_parse.Recordset("base") = 33 Or frm_menu.data_parse.Recordset("base") = 34 Then  ' B3 adm
                                                    If frm_menu.data_parse.Recordset("base") = 33 Then
                                                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-333", vbNullString)
                                                    Else
                                                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-334", vbNullString)
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
                   End If
                End If
             End If
          End If
       End If
    End If
End If

'Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
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
    
'MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'      "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
        
    'Enviando
If Not EstaInicializado() Then Exit Sub
    
Dim objCfe As CFE
Set objCfe = New CFE

Dim objCf As ClassFactory

Set objCf = New ClassFactory
       
Set objCfe.EFact = New EFact
With objCfe.EFact.Encabezado.IdDoc
    .TipoCFE = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
    .FchEmis.SetDate Year(data_cabeza2.Recordset("cl_fnac")), Month(data_cabeza2.Recordset("cl_fnac")), Day(data_cabeza2.Recordset("cl_fnac"))
    .IsValidMntBruto = True
    .MntBruto = IdDoc_Tck_MntBruto_1
     If data_cabeza2.Recordset("cl_forpago") = 1 Then
       .FmaPago = IdDoc_Fact_FmaPago_1
     Else
       .FmaPago = IdDoc_Fact_FmaPago_2
     End If
End With
    
With objCfe.EFact.Encabezado.Emisor
    .RUCEmisor = data_param.Recordset("ruc")
    .RznSoc = data_param.Recordset("nomc")
    .CdgDGISucur.FromString Trim(str(data_param.Recordset("codsuc")))
    .DomFiscal = data_param.Recordset("domic")
    .Ciudad = data_param.Recordset("ciudad")
    .Departamento = data_param.Recordset("dpto")
End With
    
With objCfe.EFact.Encabezado.Receptor
    .TipoDocRecep = DocType_2
    .CodPaisRecep = CodPaisType_UY
    .DocRecep = data_cabeza2.Recordset("cl_nom_sup")
    .RznSocRecep = data_cabeza2.Recordset("info_debit")
    .DirRecep = data_cabeza2.Recordset("cl_direcci")
    .CiudadRecep = data_cabeza2.Recordset("cl_zona")
End With
With objCfe.EFact.Encabezado.Totales
    .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
    .IsValidTpoCambio = True
    .TpoCambio.FromString "1"
    .IsValidMntNetoIvaTasaMin = True
    .IsValidMntIVATasaMin = True
    .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
    If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
       If data_cabeza2.Recordset("cl_cedula") > 0 Then
          .IsValidMntNoGrv = True
          .MntNoGrv.FromString Format(data_cabeza2.Recordset("cl_cedula"), "0.00")
       End If
    End If
    .IVATasaMin = TasaIVAType_10FullStop000
    .MntIVATasaMin.FromString Format(data_cabeza2.Recordset("saldo_cc"), "0.00")
    .CantLinDet.FromString data_cabeza2.Recordset("cl_grupo")
    .MntTotal.FromString Format(data_cabeza2.Recordset("saldo_doc"), "0.00")
    .MntPagar.FromString Format(data_cabeza2.Recordset("saldo_chc"), "0.00")
End With
    
Do While Not data_lineas.Recordset.EOF
   With objCfe.EFact.Detalle.Item.AddNew
        .NroLinDet.FromString Trim(str(data_lineas.Recordset("linea")))
        .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_lineas.Recordset("tipo_mov"))))
        .NomItem = data_lineas.Recordset("nom_prod")
        .cantidad.FromString Trim(str(data_lineas.Recordset("cantidad")))
        .UniMed = "N/A"
        .PrecioUnitario.FromString Format(data_lineas.Recordset("arancel"), "0.00")
        .MontoItem.FromString Format(data_lineas.Recordset("tot_lin"), "0.00")
    End With
    data_lineas.Recordset.MoveNext
Loop
data_lineas.Recordset.MoveFirst

Set objCfe.EFact.Referencia = New Referencia
Do While Not data_lineas.Recordset.EOF
   With objCfe.EFact.Referencia.ReferenciaA.AddNew
       .NroLinRef.FromString Trim(str(data_lineas.Recordset("linearef")))
       .IsValidIndGlobal = False
       .IsValidTpoDocRef = True
       .TpoDocRef = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(data_lineas.Recordset("tipodocref"))))
       .IsValidSerie = True
       .serie = Trim(data_lineas.Recordset("serieref"))
       .IsValidNroCFERef = True
       .NroCFERef.FromLong data_lineas.Recordset("nrofactref")
       .IsValidFechaCFEref = True
       .FechaCFEref.SetDate Year(data_lineas.Recordset("fechafact")), Month(data_lineas.Recordset("fechafact")), Day(data_lineas.Recordset("fechafact"))
   End With
   If Label7.Caption = "NC E-FACTURA" Then
      If Val(labidemi.Caption) = 5 Then
      Else
         data_lincance.RecordSource = "Select * from linmmdd where factura =" & data_lineas.Recordset("nrofactref") & " and moneda ='" & data_lineas.Recordset("serieref") & "' and linea =" & data_lineas.Recordset("linearef")
         data_lincance.Refresh
         If data_lincance.Recordset.RecordCount > 0 Then
            If IsNull(data_lincance.Recordset("descuento")) = True Then
               data_lincance.Recordset.Edit
               data_lincance.Recordset("descuento") = 1
               data_lincance.Recordset.Update
            End If
         End If
      End If
   End If
   data_lineas.Recordset.MoveNext
Loop
    
Dim s As String
s = objCfe.ToXml(True, XmlFormatting_Indented)
Dim strGuid As String
strGuid = objPosCfe.CrearGuid()
Dim objResultadoCfe As ResultadoCfe
Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    
Set objUltimaSerieNumero = Nothing
DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
If Not objUltimaSerieNumero Is Nothing Then _
   ' cmdFirmarNc.Enabled = True
    '   MsgBox "firmar NC"
End If
If frm_menu.data_parse.Recordset("base") = 38 Then
   MsgBox "Terminado"
   Unload Me
Else
   Command1_Click
End If

End Sub

Private Sub b_ncetck_Click()
Dim strIdTransac As String

On Error GoTo Alenvncet

Set objPosCfe = New PosCfe
    
''''MsgBox "Llegó bien al envío"

Dim objresultado As Resultado

If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
   If frm_menu.data_parse.Recordset("base") = 78 Then 'Notebook JF
      Set objresultado = objPosCfe.Inicializar("SAPP-105", frm_menu.data_parse.Recordset("contacto"), vbNullString)
   Else
      Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
   End If
Else
    If frm_menu.data_parse.Recordset("base") = 1 Then
       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-301", vbNullString)
    Else
       If frm_menu.data_parse.Recordset("base") = 2 Then
          Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-002", vbNullString)
       Else
          If frm_menu.data_parse.Recordset("base") = 3 Then
             Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-001", vbNullString)
          Else
             If frm_menu.data_parse.Recordset("base") = 4 Then
                Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-004", vbNullString)
             Else
                If frm_menu.data_parse.Recordset("base") = 6 Then
                   Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-006", vbNullString)
                Else
                   If frm_menu.data_parse.Recordset("base") = 8 Then
                      Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-008", vbNullString)
                   Else
                      If frm_menu.data_parse.Recordset("base") = 10 Then
                         Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-010", vbNullString)
                      Else
                         If frm_menu.data_parse.Recordset("base") = 13 Then
                            Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-013", vbNullString)
                         Else
                            If frm_menu.data_parse.Recordset("base") = 16 Then
                               Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-116", vbNullString)
                            Else
                               If frm_menu.data_parse.Recordset("base") = 91 Then
                                  Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-216", vbNullString)
                               Else
                                  If frm_menu.data_parse.Recordset("base") = 17 Then
                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-017", vbNullString)
                                  Else
                                     If frm_menu.data_parse.Recordset("base") = 18 Then
                                        Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-018", vbNullString)
                                     Else
                                        If frm_menu.data_parse.Recordset("base") = 96 Then
                                           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-218", vbNullString)
                                        Else
                                           If frm_menu.data_parse.Recordset("base") = 38 Then
                                              Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
                                           Else
                                              If frm_menu.data_parse.Recordset("base") = 11 Then
                                                 Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-111", vbNullString)
                                              Else
                                                 If frm_menu.data_parse.Recordset("base") = 93 Then
                                                    Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-217", vbNullString)
                                                 Else
                                                    If frm_menu.data_parse.Recordset("base") = 92 Then
                                                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-318", vbNullString)
                                                    Else
                                                       If frm_menu.data_parse.Recordset("base") = 33 Or frm_menu.data_parse.Recordset("base") = 34 Then  ' B3 adm
                                                          If frm_menu.data_parse.Recordset("base") = 33 Then
                                                             Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-333", vbNullString)
                                                          Else
                                                             Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-334", vbNullString)
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
                         End If
                      End If
                   End If
                End If
             End If
          End If
       End If
    End If
End If

'Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
Dim strMensaje As String
Dim tipo As Integer
tipo = 0
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
'MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'   "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
        
    'Enviando
If Not EstaInicializado() Then Exit Sub
    
Dim objCfe As CFE
Set objCfe = New CFE

Dim objCf As ClassFactory
Set objCf = New ClassFactory

Set objCfe.ETck = New ETck
With objCfe.ETck.Encabezado.IdDoc
    .TipoCFE = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
    .FchEmis.SetDate Year(data_cabeza2.Recordset("cl_fnac")), Month(data_cabeza2.Recordset("cl_fnac")), Day(data_cabeza2.Recordset("cl_fnac"))
    .IsValidMntBruto = True
    .MntBruto = IdDoc_Tck_MntBruto_1
    If data_cabeza2.Recordset("cl_forpago") = 1 Then
       .FmaPago = IdDoc_Tck_FmaPago_1
    Else
       .FmaPago = IdDoc_Tck_FmaPago_2
    End If
End With
With objCfe.ETck.Encabezado.Emisor
    .RUCEmisor = data_param.Recordset("ruc")
    .RznSoc = data_param.Recordset("nomc")
    .CdgDGISucur.FromString Trim(str(data_param.Recordset("codsuc")))
    .DomFiscal = data_param.Recordset("domic")
    .Ciudad = data_param.Recordset("ciudad")
    .Departamento = data_param.Recordset("dpto")
End With
Set objCfe.ETck.Encabezado.Receptor = New Receptor_Tck
Set objCfe.ETck.Encabezado.Receptor.Receptor_Tck_Choice = New Receptor_Tck_Choice
With objCfe.ETck.Encabezado.Receptor
    If data_cabeza2.Recordset("cl_nro_sup") = 2 Then
       .TipoDocRecep = DocType_2
       tipo = 2
    Else
       If data_cabeza2.Recordset("cl_nro_sup") = 3 Then
          .TipoDocRecep = DocType_3
          tipo = 3
       Else
          If data_cabeza2.Recordset("cl_nro_sup") = 5 Then
             .TipoDocRecep = DocType_5
             tipo = 5
          Else
             If data_cabeza2.Recordset("cl_nro_sup") = 6 Then
                .TipoDocRecep = DocType_6
                tipo = 6
             Else
                .TipoDocRecep = DocType_4
                tipo = 4
             End If
          End If
       End If
    End If
    .CodPaisRecep = CodPaisType_UY
    If tipo = 4 Then
       .Receptor_Tck_Choice.DocRecepExt = data_cabeza2.Recordset("cl_nom_sup")
    Else
       .Receptor_Tck_Choice.DocRecep = data_cabeza2.Recordset("cl_nom_sup")
    End If
    .RznSocRecep = data_cabeza2.Recordset("info_debit")
    .DirRecep = data_cabeza2.Recordset("cl_direcci")
    .CiudadRecep = data_cabeza2.Recordset("cl_zona")
End With
With objCfe.ETck.Encabezado.Totales
     .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
     .IsValidTpoCambio = True
     .TpoCambio.FromString "1"
     .IsValidMntNetoIvaTasaMin = True
     .IsValidMntIVATasaMin = True
     .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
     If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
        If data_cabeza2.Recordset("cl_cedula") > 0 Then
           .IsValidMntNoGrv = True
           .MntNoGrv.FromString Format(data_cabeza2.Recordset("cl_cedula"), "0.00")
        End If
     End If
     .IVATasaMin = TasaIVAType_10FullStop000
     .MntIVATasaMin.FromString Format(data_cabeza2.Recordset("saldo_cc"), "0.00")
     .CantLinDet.FromString data_cabeza2.Recordset("cl_grupo")
     .MntTotal.FromString Format(data_cabeza2.Recordset("saldo_doc"), "0.00")
     .MntPagar.FromString Format(data_cabeza2.Recordset("saldo_chc"), "0.00")
End With
Do While Not data_lineas.Recordset.EOF
   With objCfe.ETck.Detalle.Item.AddNew
       .NroLinDet.FromString Trim(str(data_lineas.Recordset("linea")))
       .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_lineas.Recordset("tipo_mov"))))
       .NomItem = data_lineas.Recordset("nom_prod")
       .cantidad.FromString Trim(str(data_lineas.Recordset("cantidad")))
       .UniMed = "N/A"
       .PrecioUnitario.FromString Format(data_lineas.Recordset("arancel"), "0.00")
       .MontoItem.FromString Format(data_lineas.Recordset("tot_lin"), "0.00")
   End With
   data_lineas.Recordset.MoveNext
Loop
data_lineas.Recordset.MoveFirst
Set objCfe.ETck.Referencia = New Referencia
Do While Not data_lineas.Recordset.EOF
   With objCfe.ETck.Referencia.ReferenciaA.AddNew
       .NroLinRef.FromString Trim(str(data_lineas.Recordset("linearef")))
       .IsValidIndGlobal = False
       .IsValidTpoDocRef = True
       .TpoDocRef = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_lineas.Recordset("tipodocref"))))
       .IsValidSerie = True
       .serie = Trim(data_lineas.Recordset("serieref"))
       .IsValidNroCFERef = True
       .NroCFERef.FromLong data_lineas.Recordset("nrofactref")
       .IsValidFechaCFEref = True
       .FechaCFEref.SetDate Year(data_lineas.Recordset("fechafact")), Month(data_lineas.Recordset("fechafact")), Day(data_lineas.Recordset("fechafact"))
   End With
   If Label7.Caption = "NC E-TICKET" Then
      If Val(labidemi.Caption) = 5 Then
      Else
         data_lincance.RecordSource = "Select * from linmmdd where factura =" & data_lineas.Recordset("nrofactref") & " and moneda ='" & data_lineas.Recordset("serieref") & "' and linea =" & data_lineas.Recordset("linearef")
         data_lincance.Refresh
         If data_lincance.Recordset.RecordCount > 0 Then
            If IsNull(data_lincance.Recordset("descuento")) = True Then
               data_lincance.Recordset.Edit
               data_lincance.Recordset("descuento") = 1
               data_lincance.Recordset.Update
            End If
         End If
      End If
   End If
   data_lineas.Recordset.MoveNext
Loop
Dim s As String
s = objCfe.ToXml(True, XmlFormatting_Indented)

Dim strGuid As String
strGuid = objPosCfe.CrearGuid()
Dim objResultadoCfe As ResultadoCfe
'Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))
    
If Val(objResultadoCfe.EstadoCfe.CodigoRespuesta) = 0 Or _
   Val(objResultadoCfe.EstadoCfe.CodigoRespuesta) = 11 Then
   labserie.Caption = objResultadoCfe.EstadoCfe.SerieNumeroCfe.serie
   labfac.Caption = objResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
   data_cabeza2.Recordset.Edit
   labvence.Caption = CStr(objResultadoCfe.EstadoCfe.DatosCae.Vencimiento)
   labautoriza.Caption = CStr(objResultadoCfe.EstadoCfe.DatosCae.Autorizacion)
   labdesde.Caption = labserie.Caption & " " & CStr(objResultadoCfe.EstadoCfe.DatosCae.NumeroDesde) & " - " & labserie.Caption & " " & CStr(objResultadoCfe.EstadoCfe.DatosCae.NumeroHasta)
   labhasta.Caption = CStr(objResultadoCfe.EstadoCfe.CodigoSeguridad)
   If Len(labvence.Caption) = 8 Then
      labvenceok.Caption = Mid(labvence.Caption, 7, 2) & "/" & Mid(labvence.Caption, 5, 2) & "/" & Mid(labvence.Caption, 1, 4)
   Else
      labvenceok.Caption = "31/12/2016"
   End If
   If labvenceok.Caption <> "" Then
      data_cabeza2.Recordset("cl_fultpag") = CDate(labvenceok.Caption)
   Else
      data_cabeza2.Recordset("cl_fultpag") = CDate("01/01/2018")
   End If
   If labautoriza.Caption <> "" Then
      data_cabeza2.Recordset("cl_nrocobr") = Val(labautoriza.Caption)
   Else
      data_cabeza2.Recordset("cl_nrocobr") = 0
   End If
   data_cabeza2.Recordset("cl_medflia") = Trim(labdesde.Caption)
   data_cabeza2.Recordset("cl_fax") = Trim(labhasta.Caption)
   data_cabeza2.Recordset.Update
    
   Dim objResultado44 As ResultadoObtenerQr
   Set objResultado44 = objPosCfe.ObtenerQr(objResultadoCfe.EstadoCfe.DatosQr, 100)
       
   Dim strFile As String
   strFile = App.path & "\qr.bmp"
   Dim f As Long
   f = FreeFile()
   Open strFile For Binary As #f
   Put #f, , objResultado44.ImagenQr
   Close #f
    
   Set objUltimaSerieNumero = Nothing
 
''''''''''    DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
   If Not objUltimaSerieNumero Is Nothing Then _
        ' cmdFirmarNc.Enabled = True
'       MsgBox "firmar NC"
   End If
   Command1_Click
Else
    MsgBox "Factura rechazada, VERIFIQUE!!", vbInformation
    data_eror.Recordset.AddNew
    data_eror.Recordset("nro") = 11
    data_eror.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
    data_eror.Recordset("hora") = Format(Time, "HH:mm")
    data_eror.Recordset("obs") = "FACT CANCE"
    data_eror.Recordset.Update
    MsgBox "Comprobante RECHAZADO, NO FUE ACEPTADO, debe realizarlo nuevamente, verifique datos!", vbInformation
    End

End If

Exit Sub

Alenvncet:
             MsgBox "Error al terminar la factura. AVISE A INFORMATICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al envio de nc etck"
             data_errfact.Recordset.Update
             End


End Sub

Private Sub b_verfaccance_Click()
On Error GoTo Almostrarcance

frm_factcancela.Show vbModal

Exit Sub

Almostrarcance:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al mostrar fact"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al mostrar fact"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub btn_fin_Click()
Dim Xsaldocaj As Long
Dim Xquetot As Long
Dim Xquees As Integer
Dim Xsaldofac As Double
Dim Xlincan As Long
Dim Xnromedic As Long
Dim Xtotverif, Xivaverif, Ximpivaverif As Double
Dim EnfermeroCovid As String
EnfermeroCovid = ""
Xtotverif = 0

On Error GoTo Elerrdefact
'''''' terminado
Dim Xlatasa, Xlatasa22 As Double
Dim Xlaui, Xmasdiezui As Double
Dim XivaVer As Double
XivaVer = 0
Xlaui = 0
Xmasdiezui = 0

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")

If Val(Label5.Caption) = 997 And data_lineas.Recordset.RecordCount > 0 Then
   MsgBox "Solo se puede facturar una línea para cobranza de deudas.", vbCritical
   b_cance_Click
Else
    btn_fin.Enabled = False
    'b_borr.Enabled = False
    b_cance.Enabled = False
    btn_graba.Enabled = False
    If labtot.Caption > 0 Then
       If data_ui.Recordset.RecordCount > 0 Then
          Xlaui = CDbl(data_ui.Recordset("descrip"))
          Xmasdiezui = CDbl(labtot.Caption) / Xlaui
       End If
    End If
    
    frm_factura.Enabled = False
    
    If txt_rut.Text <> "" Then
       Xtipodedocumento = 2
    Else
       If Xmasdiezui > 10000 Then
          Xtipodedocumento = 3
       Else
          Xtipodedocumento = 4
       End If
    End If
    If labtim.Caption = "" Then
       labtim.Caption = 0
    End If
    If labtot.Caption > 0 Then
    Else
       Label7.Caption = "REG."
    End If
    Dim Xnuevaslin As Integer
    Xnuevaslin = 1
'    Dim Xporquedevol As String
'    If XQuefac = 21 Then
'       Xporquedevol = InputBox("INGRESE MOTIVO DE LA DEVOLUCIÓN")
'       labdevol.Caption = Xporquedevol
'    End If
    
    If data_lineas.Recordset.RecordCount > 0 Then
       data_lineas.Refresh
       data_lineas.Recordset.MoveFirst
       If data_lineas.Recordset("cod_prod") = 30081 Then
          EnfermeroCovid = InputBox("Ingrese nombre del enfermero que realizará el domicilio:", "Facturación")
          If Trim(EnfermeroCovid) <> "" Then
             data_deudas.RecordSource = "select * from sol_hisopos where cedula =" & data_lineas.Recordset("ced_socio") & " and fecha_fact is null"
             data_deudas.Refresh
             If data_deudas.Recordset.RecordCount > 0 Then
                data_deudas.Recordset.Edit
                data_deudas.Recordset("enf_realiza") = Trim(EnfermeroCovid)
                data_deudas.Recordset.Update
             End If
          End If
       End If
       
       Do While Not data_lineas.Recordset.EOF
          XivaVer = XivaVer + data_lineas.Recordset("imp_iva")
          If data_lineas.Recordset("cod_prod") = 60107 Or data_lineas.Recordset("cod_prod") = 60103 Then
             If data_lineas.Recordset("tot_lin") = 0 Then
                data_lineas.Recordset.Edit
                data_lineas.Recordset("tipo_mov") = 5
                data_lineas.Recordset.Update
             End If
          End If
          data_lineas.Recordset.Edit
          data_lineas.Recordset("linea") = Xnuevaslin
          data_lineas.Recordset("unidad") = "N"
          data_lineas.Recordset.Update
          Xnuevaslin = Xnuevaslin + 1
          data_lineas.Recordset.MoveNext
       Loop
       data_lineas.Recordset.MoveFirst
       data_cabeza2.Recordset.AddNew
       data_cabeza2.Recordset("cl_tipcli") = "1.0"
       data_cabeza2.Recordset("cl_telefon") = Label7.Caption
       If Label7.Caption = "E-FACTURA" Then
          data_cabeza2.Recordset("cl_tipocli") = 111
          data_cabeza2.Recordset("cl_nombre") = "RUT COMPRADOR"
       Else
          If Label7.Caption = "E-TICKET" Then
             data_cabeza2.Recordset("cl_tipocli") = 101
             data_cabeza2.Recordset("cl_nombre") = "CONSUMO FINAL"
          Else
             If Label7.Caption = "NC E-FACTURA" Then
                data_cabeza2.Recordset("cl_tipocli") = 112
                data_cabeza2.Recordset("cl_nombre") = "RUT COMPRADOR"
             Else
                If Label7.Caption = "NC E-TICKET" Then
                   data_cabeza2.Recordset("cl_tipocli") = 102
                   data_cabeza2.Recordset("cl_nombre") = "CONSUMO FINAL"
                Else
                   If Label7.Caption = "ND E-FACTURA" Then
                      data_cabeza2.Recordset("cl_tipocli") = 113
                      data_cabeza2.Recordset("cl_nombre") = "RUT COMPRADOR"
                   Else
                      If Label7.Caption = "ND E-TICKET" Then
                         data_cabeza2.Recordset("cl_tipocli") = 103
                         data_cabeza2.Recordset("cl_nombre") = "CONSUMO FINAL"
                      Else
                         If Label7.Caption = "REG." Then
                            data_cabeza2.Recordset("cl_tipocli") = 887
                            data_cabeza2.Recordset("cl_nombre") = "REG."
                            labfpago.Caption = "REG."
                         Else
                            If Label7.Caption = "DEV.RECIBO" Then
                               data_cabeza2.Recordset("cl_tipocli") = 889
                               data_cabeza2.Recordset("cl_nombre") = "DEV.RECIBO"
                            Else
                               data_cabeza2.Recordset("cl_tipocli") = 888
                               data_cabeza2.Recordset("cl_nombre") = "RECIBO"
                            End If
                         End If
                      End If
                   End If
                End If
             End If
          End If
       End If
       data_cabeza2.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
       data_cabeza2.Recordset("cl_nrovend") = 1 'lineas de detalle con iva inc
       If labfpago.Caption = "CONTADO" Then
          data_cabeza2.Recordset("cl_forpago") = 1
       Else
          If labfpago.Caption = "CREDITO" Then
             data_cabeza2.Recordset("cl_forpago") = 2
          Else
             data_cabeza2.Recordset("cl_forpago") = 0
         End If
       End If
       data_cabeza2.Recordset("cl_celular") = labfpago.Caption
    '   data_cabeza2.Recordset("fecha_modi") = Format(labvence.Caption, "dd/mm/yyyy")
       data_cabeza2.Recordset("cl_diacobr") = Trim(str(data_param.Recordset("ruc")))
       data_cabeza2.Recordset("cl_nrotarj") = data_param.Recordset("nombre")
       data_cabeza2.Recordset("cl_tjemi_n") = data_param.Recordset("nombre")
       data_cabeza2.Recordset("cl_tjemi_c") = data_param.Recordset("codsuc")
       data_cabeza2.Recordset("cl_referen") = data_param.Recordset("domic")
       data_cabeza2.Recordset("tit_tarj") = data_param.Recordset("ciudad")
       data_cabeza2.Recordset("cl_nomconv") = data_param.Recordset("dpto")
        'receptor
       data_cabeza2.Recordset("cl_nro_sup") = Xtipodedocumento
       data_cabeza2.Recordset("hora_baja") = "UY"
       If Xtipodedocumento = 3 Then
          data_cabeza2.Recordset("cl_nom_sup") = Trim(str(data_lineas.Recordset("ced_socio"))) & Trim(str(data_lineas.Recordset("fact")))
          data_cabeza2.Recordset("cl_nomcobr") = "CI"
       Else
          If Xtipodedocumento = 2 Then
             data_cabeza2.Recordset("cl_nom_sup") = txt_rut.Text
             data_cabeza2.Recordset("cl_nomcobr") = "RUT"
          Else
             data_cabeza2.Recordset("cl_nom_sup") = labmatri.Caption
             data_cabeza2.Recordset("cl_nomcobr") = "Otro"
          End If
       End If
       If Trim(frmabm.t_rs.Text) <> "" Then
          If Check1.Value = 1 Then
             data_cabeza2.Recordset("info_debit") = frmabm.t_rs.Text
          Else
             data_cabeza2.Recordset("info_debit") = labnomb.Caption
          End If
       Else
          data_cabeza2.Recordset("info_debit") = labnomb.Caption
       End If
       If Xtipodedocumento = 4 Then
          data_cabeza2.Recordset("cl_direcci") = "S/D"
          data_cabeza2.Recordset("cl_zona") = "S/D"
          data_cabeza2.Recordset("cl_localid") = "URUGUAY" 'opcional
       Else
          data_cabeza2.Recordset("cl_direcci") = frmabm.txt_direcc1.Text
          data_cabeza2.Recordset("cl_zona") = frmabm.cbolocalid.Text
          data_cabeza2.Recordset("cl_localid") = "URUGUAY" 'opcional
       End If
       data_cabeza2.Recordset("cl_codigo") = Val(labmatri.Caption)
       data_cabeza2.Recordset("usu_baja") = "UYU"
       If labtim.Caption <> "" Then
          If Val(labtim.Caption) > 0 Then
             data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labtim.Caption, "Standard")
             data_cabeza2.Recordset("cl_cedula") = Format(labtim.Caption, "0.00")
          Else
             If labtimemi.Caption <> "" Then
                If Val(labtimemi.Caption) > 0 Then
                   If labdeudaemi.Caption <> "" Then
                      If Val(labdeudaemi.Caption) > 0 Then
                         data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labdeudaemi.Caption, "Standard") - Format(labtimemi.Caption, "Standard")
                      Else
                         data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labtimemi.Caption, "Standard")
                      End If
                   Else
                      data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labtimemi.Caption, "Standard")
                   End If
                Else
                   If labdeudaemi.Caption <> "" Then
                      If Val(labdeudaemi.Caption) > 0 Then
                         data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labdeudaemi.Caption, "Standard")
                      Else
                         data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
                      End If
                   Else
                      data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
                   End If
                End If
             Else
                If labdeudaemi.Caption <> "" Then
                   If Val(labdeudaemi.Caption) > 0 Then
                      data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labdeudaemi.Caption, "Standard")
                   Else
                      data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
                   End If
                Else
                   data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
                End If
             End If
          End If
       Else
          If labtimemi.Caption <> "" Then
             If Val(labtimemi.Caption) > 0 Then
                data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labtimemi.Caption, "Standard")
                If labdeudaemi.Caption <> "" Then
                   If Val(labdeudaemi.Caption) > 0 Then
                      data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labdeudaemi.Caption, "Standard") - Format(labtimemi.Caption, "Standard")
                   End If
                End If
             Else
                If labdeudaemi.Caption <> "" Then
                   If Val(labdeudaemi.Caption) > 0 Then
                      data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labdeudaemi.Caption, "Standard")
                   Else
                      data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
                   End If
                Else
                   data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
                End If
             End If
          Else
             If labdeudaemi.Caption <> "" Then
                If Val(labdeudaemi.Caption) > 0 Then
                   data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard") - Format(labdeudaemi.Caption, "Standard")
                Else
                   data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
                End If
             Else
                data_cabeza2.Recordset("saldo_doc2") = Format(labtot.Caption, "Standard") - Format(Label8.Caption, "Standard")
             End If
          End If
       End If
       data_cabeza2.Recordset("cl_atrasoa") = 0
       data_cabeza2.Recordset("cl_atrasop") = Xlatasa
       data_cabeza2.Recordset("cl_decuota") = Xlatasa22
       data_cabeza2.Recordset("saldo_cc") = Format(Label8.Caption, "Standard")
       data_cabeza2.Recordset("saldo_cc2") = 0 'iva básico
       data_cabeza2.Recordset("saldo_doc") = Format(labtot.Caption, "Standard")
       data_cabeza2.Recordset("cl_grupo") = data_lineas.Recordset.RecordCount
       data_cabeza2.Recordset("saldo_chc") = Format(labtot.Caption, "Standard")
       If Trim(labtimemi.Caption) <> "" Then
          If Val(labtimemi.Caption) > 0 Then
             data_cabeza2.Recordset("cl_cedula") = Format(labtimemi.Caption, "0.00")
    ''         MsgBox ("Guarda:" & Format(labtimemi.Caption, "0.00"))
          Else
             If Trim(labtim.Caption) <> "" Then
                If Val(labtim.Caption) <= 0 Then
                   data_cabeza2.Recordset("cl_cedula") = Format(labtim.Caption, "0.00")
                End If
             Else
                data_cabeza2.Recordset("cl_cedula") = Format(labtim.Caption, "0.00")
             End If
          End If
          If Trim(labdeudaemi.Caption) <> "" Then
             If Val(labdeudaemi.Caption) > 0 Then
                If Val(labtimemi.Caption) > 0 Then
                   data_cabeza2.Recordset("cl_cedula") = data_cabeza2.Recordset("cl_cedula") + Format(labdeudaemi.Caption, "0.00")
                Else
                   data_cabeza2.Recordset("cl_cedula") = Format(labdeudaemi.Caption, "0.00")
                End If
             Else
                If Val(labtimemi.Caption) > 0 Then
                   data_cabeza2.Recordset("cl_cedula") = data_cabeza2.Recordset("cl_cedula") + Format(labdeudaemi.Caption, "0.00")
                Else
                   data_cabeza2.Recordset("cl_cedula") = Format(labdeudaemi.Caption, "0.00")
                End If
             End If
          End If
       Else
          If Trim(labdeudaemi.Caption) <> "" Then
             If Val(labdeudaemi.Caption) > 0 Then
                data_cabeza2.Recordset("cl_cedula") = Format(labdeudaemi.Caption, "0.00")
             Else
                data_cabeza2.Recordset("cl_cedula") = Format(labdeudaemi.Caption, "0.00")
             End If
          End If
       End If
       If data_lineas.Recordset.RecordCount = 1 Then
          If data_lineas.Recordset("cod_prod") = 995 Or data_lineas.Recordset("cod_prod") = 990 Then
             data_cabeza2.Recordset("cl_cedula") = Format(labtot.Caption, "Standard")
             data_cabeza2.Recordset("saldo_doc2") = 0
          End If
       End If
       If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "ND E-FACTURA" Or Label7.Caption = "ND E-TICKET" Or Label7.Caption = "DEV.RECIBO" Then
          If IsNull(data_lineas.Recordset("nrofactref")) = False Then
             If data_lineas.Recordset("nrofactref") > 0 Then
                If IsNull(data_lineas.Recordset("serieref")) = False Then
                   data_cabeza2.Recordset("obsp") = "CANCELA FACT: Serie:" & data_lineas.Recordset("serieref") & " Nro." & data_lineas.Recordset("nrofactref") & Chr(13) _
                   & "CAJA Usuario: " & data_lineas.Recordset("operador") & Chr(13) & " HORA:" & Format(Time, "HH:mm") & " BASE: " & frm_menu.data_parse.Recordset("base") & Chr(13) _
                   & "Aceptamos VISA, MASTERCARD, OCA, CABAL, C.Directos,Passcard. Débito por BROU." & Chr(13) _
                   & "Empresa afiliada al Clearing de Informes."
                Else
                   data_cabeza2.Recordset("obsp") = "CANCELA FACT: Serie:A Nro." & data_lineas.Recordset("nrofactref") & Chr(13) _
                   & "CAJA Usuario: " & data_lineas.Recordset("operador") & Chr(13) & " HORA:" & Format(Time, "HH:mm") & " BASE: " & frm_menu.data_parse.Recordset("base") & Chr(13) _
                   & "Aceptamos VISA, MASTERCARD, OCA, CABAL, C.Directos,Passcard. Débito por BROU." & Chr(13) _
                   & "Empresa afiliada al Clearing de Informes."
                End If
             Else
                data_cabeza2.Recordset("obsp") = "CAJA Usuario: " & data_lineas.Recordset("operador") & Chr(13) & " HORA:" & Format(Time, "HH:mm") & " BASE: " & frm_menu.data_parse.Recordset("base") & Chr(13) _
                & "Aceptamos VISA, MASTERCARD, OCA, CABAL, C.Directos,Passcard. Débito por BROU." & Chr(13) _
                & "Empresa afiliada al Clearing de Informes."
             End If
          Else
             data_cabeza2.Recordset("obsp") = "CAJA Usuario: " & data_lineas.Recordset("operador") & Chr(13) & " HORA:" & Format(Time, "HH:mm") & " BASE: " & frm_menu.data_parse.Recordset("base") & Chr(13) _
             & "Aceptamos VISA, MASTERCARD, OCA, CABAL, C.Directos,Passcard. Débito por BROU." & Chr(13) _
             & "Empresa afiliada al Clearing de Informes."
          End If
       Else
          If data_lineas.Recordset("cod_prod") = 999 Or data_lineas.Recordset("cod_prod") = 997 Or _
             data_lineas.Recordset("cod_prod") = 993 Or data_lineas.Recordset("cod_prod") = 994 Then
             data_cabeza2.Recordset("obsp") = "CAJA Usuario: " & data_lineas.Recordset("operador") & Chr(13) & " HORA:" & Format(Time, "HH:mm") & " BASE: " & frm_menu.data_parse.Recordset("base") & Chr(13) _
             & "Aceptamos VISA, MASTERCARD, OCA, CABAL, C.Directos,Passcard. Débito por BROU." & Chr(13) _
             & "El pago de este recibo no cancela deudas anteriores." & Chr(13) _
             & "Empresa afiliada al Clearing de Informes."
          Else
             data_cabeza2.Recordset("obsp") = "CAJA Usuario: " & data_lineas.Recordset("operador") & Chr(13) & " HORA:" & Format(Time, "HH:mm") & " BASE: " & frm_menu.data_parse.Recordset("base") & Chr(13) _
             & "Aceptamos VISA, MASTERCARD, OCA, CABAL, C.Directos,Passcard. Débito por BROU." & Chr(13) _
             & "Empresa afiliada al Clearing de Informes."
          End If
       End If
       If XMensaFertilab = 9 Then
          If XMensaFertilab2 <> 9 Then
            data_cabeza2.Recordset("obsp") = "Concurrir a Laboratorio FERTILAB. Sede Central: Canelones 2297 Telef: 24020041 -Lunes a sábados-." & vbCrLf _
            & "FERTILAB Platinum Carrasco: Costa Rica 1649 esq. Schroeder Telef: 26046776 " & vbCrLf _
            & "PARA EVITAR DEMORAS, recomendamos agendarse desde www.fertilab.com.uy" & vbCrLf _
            & "CAJA Usuario: " & data_lineas.Recordset("operador") & vbCrLf & " HORA:" & Format(Time, "HH:mm") & " BASE: " & frm_menu.data_parse.Recordset("base") & vbCrLf _
            & "Aceptamos VISA, MASTERCARD, OCA, CABAL, C.Directos,Passcard. Débito por BROU." & vbCrLf _
            & "Empresa afiliada al Clearing de Informes."
          End If
''''            & "SIGA LAS INDICACIONES DE LA ORDEN MÉDICA." & vbCrLf _

       End If
       If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "ND E-FACTURA" Or Label7.Caption = "ND E-TICKET" Then
          data_cabeza2.Recordset("cl_cuopaga") = data_lineas.Recordset("tipodocref")
          data_cabeza2.Recordset("codmotbaja") = data_lineas.Recordset("serieref")
          data_cabeza2.Recordset("ultanopmut") = data_lineas.Recordset("nrofactref")
          data_cabeza2.Recordset("cl_fultvta") = data_lineas.Recordset("fechafact")
          data_cabeza2.Recordset("cl_entre") = data_lineas.Recordset("motivoref")
       End If
       
       
       data_cabeza2.Recordset.Update
       data_cabeza2.Refresh
       data_cabeza2.Recordset.MoveFirst
       If Label7.Caption = "E-TICKET" Then
          b_etck_Click
'          labserie.Caption = "A"
'          labfac.Caption = 939
'          Command1_Click
       Else
          If Label7.Caption = "E-FACTURA" Then
             Command3_Click
          Else
             If Label7.Caption = "NC E-TICKET" Then
                If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
                   Xtotverif = data_cabeza2.Recordset("saldo_doc2") + data_cabeza2.Recordset("saldo_cc") + data_cabeza2.Recordset("cl_cedula")
                Else
                   Xtotverif = data_cabeza2.Recordset("saldo_doc2") + data_cabeza2.Recordset("saldo_cc")
                End If
                Ximpivaverif = data_cabeza2.Recordset("saldo_doc2") + data_cabeza2.Recordset("saldo_cc")
                Xivaverif = Ximpivaverif / 1.1 * 0.1
                If data_lineas.Recordset("nrofactref") = 0 Then
                   data_param.Recordset.Edit
                   data_param.Recordset("nro_rec") = data_param.Recordset("nro_rec") + 1
                   data_param.Recordset.Update
                   labfac.Caption = data_param.Recordset("nro_rec")
                   labserie.Caption = data_param.Recordset("serie_fact")
                   Command1_Click
                Else
                   If Format(Xtotverif, "Standard") = Format(data_cabeza2.Recordset("saldo_doc"), "Standard") Then
                      If Format(data_cabeza2.Recordset("saldo_cc"), "Standard") = Format(XivaVer, "Standard") Then
                         b_ncetck_Click
                      Else
                         MsgBox "Hay un error en los totales de la devolución, VERIFIQUE o pruebe realizarla manual " + Format(Xivaverif, "Standard") + " -" & Format(data_cabeza2.Recordset("saldo_cc"), "Standard"), vbExclamation
                         frmabm.btn_fact.Enabled = True
                         Unload Me
                         Exit Sub
                      End If
                   Else
                      MsgBox "Hay un error en los totales de la devolución, VERIFIQUE o pruebe realizarla manual " & Xtotverif & " - " & data_cabeza2.Recordset("saldo_doc"), vbExclamation
                      frmabm.btn_fact.Enabled = True
                      Unload Me
                      Exit Sub
                   End If
                End If
             Else
                If Label7.Caption = "NC E-FACTURA" Then
                   b_ncefct_Click
                Else
                   If Label7.Caption = "ND E-TICKET" Then
                      b_ncetck_Click
                   Else
                      If Label7.Caption = "ND E-FACTURA" Then
                         b_ncefct_Click
                      Else
                         If Label7.Caption = "RECIBO" Or Label7.Caption = "REG." Or Label7.Caption = "DEV.RECIBO" Then
                            data_param.Recordset.Edit
                            data_param.Recordset("nro_rec") = data_param.Recordset("nro_rec") + 1
                            data_param.Recordset.Update
                            labfac.Caption = data_param.Recordset("nro_rec")
                            labserie.Caption = data_param.Recordset("serie_fact")
                            Command1_Click
                         End If
                      End If
                   End If
                End If
             End If
          End If
       End If
    Else
       MsgBox "No hay lineas de factura para procesar", vbInformation
       Unload Me
    End If
End If

Xestaok = 0
XAlta = 0

Exit Sub
Elerrdefact:
            If Err.Number = 3155 Then
               MsgBox "ERROR al grabar :" & Err.Number & " " & Err.Description
               data_errfact.Recordset.AddNew
               data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
               data_errfact.Recordset("fecha") = Date
               data_errfact.Recordset("hora") = Format(Time, "HH:mm")
               data_errfact.Recordset("nroerr") = Err.Number
               data_errfact.Recordset("desc") = Mid(Err.Description, 1, 110)
               data_errfact.Recordset.Update
            Else
               MsgBox "ERROR al grabar :" & Err.Number & " " & Err.Description
               data_errfact.Recordset.AddNew
               data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
               data_errfact.Recordset("fecha") = Date
               data_errfact.Recordset("hora") = Format(Time, "HH:mm")
               data_errfact.Recordset("nroerr") = Err.Number
               data_errfact.Recordset("desc") = Mid(Err.Description, 1, 110)
               data_errfact.Recordset.Update
            
            End If
            
End Sub

Private Sub btn_graba_Click()
Dim Xrub As Long
Dim Xiva, Xvaltimme As Double
On Error GoTo Vererror
Dim XValtim As Long
Dim Xlafenaci As String
Dim Xnomedica, Xlin997, Xcuotabase, Xquerr As Integer
Dim Xnroform As String
Dim Xestco, Xestfl As Long
Dim Xmotivoref As String
Dim Xtelcmt As String
Dim MensajeCMT As String
MensajeCMT = vbNo
Dim Xcantveces As Integer
Xmotivoref = ""

Xtelcmt = ""

DBCombo1.Enabled = True

If Val(Label5.Caption) = 997 And data_lineas.Recordset.RecordCount > 0 Then
   MsgBox "Solo se puede facturar una línea para cobranza de deudas.", vbCritical
   b_cance_Click
Else
    If t_cant.Text <> "" Then
       If t_cant.Text > 0 Then
       Else
          t_cant.Text = 1
       End If
    Else
       t_cant.Text = 1
    End If
    
    If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "DEV.RECIBO" Then
       If Label5.Caption <> "" Then
          data_estudio.Recordset.FindFirst "codest =" & Label5.Caption
       End If
    End If
    Xvaltimme = 0
    Xlin997 = 0
    Xestco = data_estudio.Recordset("codest")
    Xestfl = data_estudio.Recordset("flia")
    Xcuotabase = 0
    Xnomedica = 0
    Xquerr = 0
    Xcantveces = 1
    
    If dbcboprom.Visible = True Then
       If dbcboprom.Text <> "" Then
          If labcodpro.Caption <> "" Then
             data_func.RecordSource = "select * from vende_func where idfunc =" & Val(labcodpro.Caption) & " and nombre ='" & dbcboprom.Text & "'"
             data_func.Refresh
             If data_func.Recordset.RecordCount > 0 Then
             Else
                Xnomedica = 33
             End If
          Else
             labcodpro.Caption = "0"
             data_func.RecordSource = "select * from vende_func where idfunc =" & Val(labcodpro.Caption) & " and nombre ='" & dbcboprom.Text & "'"
             data_func.Refresh
             If data_func.Recordset.RecordCount > 0 Then
             Else
                Xnomedica = 33
             End If
          End If
       Else
          Xnomedica = 33
       End If
    End If
                 
    If IsNull(frmabm.data_clientes.Recordset("saldo_chc2")) = False Then
       If frmabm.data_clientes.Recordset("saldo_chc2") = 1 Then
          If data_estudio.Recordset("codest") = 10003 Or _
             data_estudio.Recordset("codest") = 10004 Or _
             data_estudio.Recordset("codest") = 10005 Or _
             data_estudio.Recordset("codest") = 10006 Or _
             data_estudio.Recordset("flia") = 9 Or _
             data_estudio.Recordset("flia") = 19 Or _
             data_estudio.Recordset("codest") = 997 Or _
             data_estudio.Recordset("codest") = 999 Or _
             data_estudio.Recordset("codest") = 993 Or _
             data_estudio.Recordset("codest") = 994 Then
          Else
             Xnomedica = 31
          End If
       End If
    End If
    
    Xconvprom = ""
    If data_estudio.Recordset("codest") = 60103 Or _
       data_estudio.Recordset("codest") = 60105 Or _
       data_estudio.Recordset("codest") = 60106 Or _
       data_estudio.Recordset("codest") = 60108 Or _
       data_estudio.Recordset("codest") = 60107 Or _
       data_estudio.Recordset("codest") = 60109 Or _
       data_estudio.Recordset("codest") = 3 Or _
       data_estudio.Recordset("codest") = 10050 Or _
       data_estudio.Recordset("codest") = 14005 Then
       If data_estudio.Recordset("codest") = 60103 Or _
          data_estudio.Recordset("codest") = 60105 Or _
          data_estudio.Recordset("codest") = 60106 Or _
          data_estudio.Recordset("codest") = 60108 Or _
          data_estudio.Recordset("codest") = 60107 Or _
          data_estudio.Recordset("codest") = 60109 Then
          
          If Trim(labmedicacion.Caption) <> "" Then
             Xdescmedic = labmedicacion.Caption
          Else
'''             frm_buscamedica.Show vbModal
             Xdescmedic = InputBox("INGRESE NOMBRE DE MEDICAMENTO", "Medicación", Xconvprom)
             If Len(Trim(Xdescmedic)) > 3 Then
                Xnomedica = 0
             Else
                Xnomedica = 7
             End If
          End If
       Else
          Xtelcmt = InputBox("INGRESE CELULAR Y/O TELÉFONOS DE CONTACTO")
          If Len(Trim(Xtelcmt)) > 3 Then
             Xnomedica = 0
          Else
             Xnomedica = 7
          End If
          MensajeCMT = MsgBox("ES REPETICIÓN DE MEDICACIÓN?", vbInformation + vbYesNo, "Facturación")
          
       End If
    End If
    Xconvprom = ""
    If data_estudio.Recordset("flia") = 19 Then
       If frmabm.txt_nac.Text = "__/__/____" Then
          If frmabm.data_clientes.Recordset("cl_codigo") = labmatri.Caption Then
             If IsNull(frmabm.data_clientes.Recordset("cl_fultpag")) = True Then
                MsgBox "No figura fecha de nacimiento en el padrón, INGRESE DATOS DE NACIMIENTO", vbInformation, "Facturación"
                Xlafenaci = InputBox("Ingrese fecha de nacimiento en formato dd/mm/aaaa")
                If Xlafenaci <> "" Then
                   frmabm.data_clientes.Recordset.Edit
                   frmabm.data_clientes.Recordset("cl_fultpag") = Format(Xlafenaci, "dd/mm/yyyy")
                   frmabm.data_clientes.Recordset.Update
                Else
                   MsgBox "No ingresó fecha de nacimiento, se cancelará la acción", vbCritical
                   b_cance_Click
                End If
             End If
          End If
       End If
    End If
    If data_estudio.Recordset("codest") = 190011 Or _
       data_estudio.Recordset("codest") = 190013 Or _
       data_estudio.Recordset("codest") = 190018 Or _
       data_estudio.Recordset("codest") = 190020 Then
       Xnroform = InputBox("Ingrese número de FORMULARIO")
    End If
    
    If data_estudio.Recordset("flia") = 19 Or data_estudio.Recordset("codest") = 10021 Or _
       data_estudio.Recordset("codest") = 10022 Or data_estudio.Recordset("codest") = 10023 Or _
       data_estudio.Recordset("codest") = 10024 Or data_estudio.Recordset("codest") = 10025 Or _
       data_estudio.Recordset("codest") = 10026 Or data_estudio.Recordset("codest") = 10027 Or _
       data_estudio.Recordset("codest") = 10028 Or data_estudio.Recordset("codest") = 10029 Or _
       data_estudio.Recordset("codest") = 10030 Or data_estudio.Recordset("codest") = 10031 Or _
       data_estudio.Recordset("codest") = 10032 Or data_estudio.Recordset("codest") = 10033 Or _
       data_estudio.Recordset("codest") = 10034 Or data_estudio.Recordset("codest") = 10035 Or _
       data_estudio.Recordset("codest") = 10036 Or data_estudio.Recordset("codest") = 10037 Or _
       data_estudio.Recordset("codest") = 10038 Or data_estudio.Recordset("codest") = 10039 Or _
       data_estudio.Recordset("codest") = 10040 Or data_estudio.Recordset("codest") = 10041 Or _
       data_estudio.Recordset("codest") = 10042 Or data_estudio.Recordset("codest") = 10043 Or _
       data_estudio.Recordset("codest") = 10044 Or data_estudio.Recordset("codest") = 10045 Or _
       data_estudio.Recordset("codest") = 10046 Or data_estudio.Recordset("codest") = 10047 Or _
       data_estudio.Recordset("codest") = 10048 Or data_estudio.Recordset("codest") = 80004 Or _
       data_estudio.Recordset("flia") = 10 Or data_estudio.Recordset("codest") = 10001 Or _
       data_estudio.Recordset("codest") = 10002 Or data_estudio.Recordset("codest") = 10003 Or _
       data_estudio.Recordset("codest") = 10004 Or data_estudio.Recordset("codest") = 10005 Or _
       data_estudio.Recordset("flia") = 19 Or data_estudio.Recordset("flia") = 5 Or _
       data_estudio.Recordset("codest") = 13056 Then
       If dbcbomed.Text = "" Then
          Xnomedica = 8
       End If
    End If
    
    If labtimemi.Caption <> "" Then
       If Val(labtimemi.Caption) > 0 Then
          If txt_precio.Text <> "" Then
             If Val(txt_precio.Text) > 0 Then
                txt_precio.Text = Format(txt_precio.Text, "Standard") - Format(labtimemi.Caption, "Standard")
             End If
          End If
       End If
    End If
    If labdeudaemi.Caption <> "" Then
       If Val(labdeudaemi.Caption) > 0 Then
          If txt_precio.Text <> "" Then
             If Val(txt_precio.Text) > 0 Then
                txt_precio.Text = Format(txt_precio.Text, "Standard") - Format(labdeudaemi.Caption, "Standard")
             End If
          End If
       End If
    End If
    
    If Label5.Caption = "" Then
       MsgBox "No ingresó servicio a facturar!", vbInformation
    End If
    
    XValtim = 0
    If data_lineas.Recordset.RecordCount >= 1 And data_estudio.Recordset("codest") = 997 Then
       Xlin997 = 8
    Else
       Xlin997 = 0
    End If
    If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "ND E-FACTURA" Or Label7.Caption = "ND E-TICKET" Then
       Xnomedica = 46
    End If
    If data_estudio.Recordset("codest") = 992 Then
       If dbcboprom.Text = "" Then
          Xnomedica = 59
       End If
    End If
    
    If data_lineas.Recordset.RecordCount >= 150 Or Xnomedica = 7 Or Xnomedica = 8 Or Xcoddeu = 9 Or Xlin997 = 8 Or Xnomedica = 31 Or Xnomedica = 59 Or Xnomedica = 33 Then
       If Xnomedica = 7 Or Xnomedica = 33 Then
          If Xnomedica = 7 Then
             MsgBox "No ingresó dato solicitado, verifique!", vbCritical, "Facturación"
             frmabm.btn_fact.Enabled = True
             DBCombo1.Text = ""
             DBCombo1.SetFocus
          Else
             MsgBox "No ingresó Promotor de la Afiliación correctamente.", vbCritical, "Facturación"
             frmabm.btn_fact.Enabled = True
             DBCombo1.Text = ""
             DBCombo1.SetFocus
             Unload Me
          End If
       Else
          If Xnomedica = 8 Then
             MsgBox "No ingresó MEDICO que realiza", vbCritical, "Facturación"
             frmabm.btn_fact.Enabled = True
             DBCombo1.SetFocus
          Else
             If Xnomedica = 31 Then
                MsgBox "ATENCION!! Usuario con servicios restringidos. Verifique con administración al 097215419", vbInformation
                frmabm.btn_fact.Enabled = True
                Unload Me
             Else
                If Xnomedica = 59 Then
                   MsgBox "ATENCION!! No ingresó PROMOTOR de la afiliación!", vbInformation
                   frmabm.btn_fact.Enabled = True
                   Unload Me
                Else
                    MsgBox "ATENCION!! alcanzó el límite de líneas por factura", vbCritical
                    DBCombo1.Text = ""
                    txt_precio.Text = 0
                End If
             End If
          End If
       End If
    Else
       Do While Xcantveces <= t_cant.Text
           If data_estudio.Recordset("codest") = 60106 Or _
              data_estudio.Recordset("codest") = 60108 Or _
              data_estudio.Recordset("codest") = 993 Or _
              data_estudio.Recordset("codest") = 994 Or _
              data_estudio.Recordset("codest") = 997 Or _
              data_estudio.Recordset("codest") = 996 Or _
              data_estudio.Recordset("codest") = 60105 Or _
              data_estudio.Recordset("codest") = 999 Or _
              data_estudio.Recordset("codest") = 60109 Or _
              data_estudio.Recordset("codest") = 80011 Or _
              data_estudio.Recordset("codest") = 80012 Or _
              data_estudio.Recordset("codest") = 80013 Or _
              data_estudio.Recordset("codest") = 80014 Or _
              data_estudio.Recordset("codest") = 80016 Or _
              data_estudio.Recordset("codest") = 80015 Then
    
              If XQuefac = 4 Or XQuefac = 21 Then
                 Xquerr = 0
              Else
                 Xquerr = 1
              End If
           Else
              If XQuefac <> 4 Then
                 Xquerr = 0
              Else
                 Xquerr = 1
              End If
           End If
           If Xquerr = 0 Then
              If data_estudio.Recordset("flia") = 1 Or _
                 data_estudio.Recordset("flia") = 14 Or _
                 data_estudio.Recordset("flia") = 10 Or _
                 data_estudio.Recordset("flia") = 9 Then
                 If dbcbomed.Text = "" And Xnomedica <> 46 Then
                    MsgBox "Debe Ingresar Médico para éste servicio", vbInformation, "Facturación"
                    DBCombo1.SetFocus
                 Else
                    Xcandelin = Xcandelin + 1
                    data_lineas.Recordset.AddNew
                    If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "ND E-FACTURA" Or Label7.Caption = "ND E-TICKET" Then
                       ''Xmotivoref = InputBox("Ingrese motivo de la modificación")
                       Xmotivoref = "SIN DATOS"
                       If Xmotivoref <> "" Then
                          labmotivo.Caption = Xmotivoref
                          If labseriecance.Caption = "XX" Then
                             data_lincance.RecordSource = "Select * from linmmdd where factura =" & Val(labfaccance.Caption) & " and fecha =#" & Format(labfeccance.Caption, "yyyy-mm-dd") & "#"
                             data_lincance.Refresh
                          Else
                             If labidemi.Caption = 5 Then
                                data_lincance.RecordSource = "Select * from deudas where documento =" & Val(labfaccance.Caption)
                                data_lincance.Refresh
                             Else
                                data_lincance.RecordSource = "Select * from clirespl where cl_numero =" & Val(labfaccance.Caption) & " and cl_socmnro ='" & Trim(labseriecance.Caption) & "' and cl_fnac =#" & Format(labfeccance.Caption, "yyyy-mm-dd") & "#"
                                data_lincance.Refresh
                             End If
                          End If
                          If data_lincance.Recordset.RecordCount > 0 And Val(labidemi.Caption) <> 5 Then
                             If labseriecance.Caption <> "XX" Then
                                If Val(labtot.Caption) >= data_lincance.Recordset("saldo_doc") Then
                                   MsgBox "Ha excedido el importe de la factura", vbCritical
                                   b_cance_Click
                                End If
                             End If
                             If labseriecance.Caption = "XX" Then
                                data_lineas.Recordset("tipodocref") = 101
                             Else
                                data_lineas.Recordset("tipodocref") = data_lincance.Recordset("cl_tipocli")
                             End If
                             data_lineas.Recordset("serieref") = labseriecance.Caption
                             If Len(labfaccance.Caption) > 7 Then
                                data_lineas.Recordset("nrofactref") = Val(Mid(labfaccance.Caption, 1, 7))
                             Else
                                data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
                             End If
                             If labseriecance.Caption = "XX" Then
                                data_lineas.Recordset("fechafact") = CDate(labfeccance.Caption)
                             Else
                                data_lineas.Recordset("fechafact") = data_lincance.Recordset("cl_fnac")
                             End If
                             data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                             data_lineas.Recordset("linearef") = Val(lablinea.Caption)
                          Else
                             If labidemi.Caption = 5 Then
                                If Label7.Caption = "NC E-FACTURA" Then
                                   data_lineas.Recordset("tipodocref") = 111
                                Else
                                   data_lineas.Recordset("tipodocref") = 101
                                End If
                                data_lineas.Recordset("serieref") = labseriecance.Caption
                                If Len(labfaccance.Caption) > 7 Then
                                   data_lineas.Recordset("nrofactref") = Val(Mid(labfaccance.Caption, 1, 7))
                                Else
                                   data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
                                End If
                                data_lineas.Recordset("fechafact") = CDate(labfeccance.Caption)
                                data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                                data_lineas.Recordset("linearef") = 1
                             Else
                                If labseriecance.Caption = "XX" Then
                                   If Label7.Caption = "NC E-FACTURA" Then
                                      data_lineas.Recordset("tipodocref") = 111
                                   Else
                                      data_lineas.Recordset("tipodocref") = 101
                                   End If
                                   data_lineas.Recordset("serieref") = "A"
                                   If Len(labfaccance.Caption) > 7 Then
                                      data_lineas.Recordset("nrofactref") = Val(Mid(labfaccance.Caption, 1, 7))
                                   Else
                                      data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
                                   End If
                                   data_lineas.Recordset("fechafact") = CDate(labfeccance.Caption)
                                   data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                                   data_lineas.Recordset("linearef") = 1
                                Else
                                   MsgBox "No se encuentra número de factura a cancelar"
                                   b_cance_Click
                                End If
                             End If
                          End If
                       Else
                          MsgBox "No ingresó motivo de cancelación"
                          b_cance_Click
                       End If
                       btn_fin.SetFocus
                    End If
                        data_lineas.Recordset("univta") = 0
                        If txt_precio.Text <> "" Then
                           If txt_precio.Text > 0 Then
                              If Val(Label5.Caption) = 995 Or Val(Label5.Caption) = 990 Then
                                 data_lineas.Recordset("tipo_mov") = 1 'tipo de iva (indic de fact)
                              Else
                                 data_lineas.Recordset("tipo_mov") = 2 'tipo de iva (indic de fact)
                              End If
                           Else
                              data_lineas.Recordset("tipo_mov") = 5
                           End If
                        Else
                           data_lineas.Recordset("tipo_mov") = 5
                        End If
                        data_lineas.Recordset("factura") = 0
                        data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                        data_lineas.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
                        data_lineas.Recordset("cod_cli") = labmatri.Caption
                        data_lineas.Recordset("nom_cli") = Mid(labnomb.Caption, 1, 30)
                        data_lineas.Recordset("cod_prod") = data_estudio.Recordset("codest")
                        data_lineas.Recordset("nom_prod") = Mid(data_estudio.Recordset("descrip"), 1, 50)
                        data_lineas.Recordset("cantidad") = 1
                        data_lineas.Recordset("moneda") = "SR" 'Serie
                        If txt_rut.Visible = True Then
                           If Trim(txt_rut.Text) <> "" Then
                              Command2_Click
                              data_lineas.Recordset("ruc") = txt_rut.Text
                           End If
                        End If
                        data_lineas.Recordset("operador") = WElusuario
                        data_lineas.Recordset("hora") = Format(Time, "HH:mm")
                        data_lineas.Recordset("nro_flia") = data_estudio.Recordset("flia")
                        data_lineas.Recordset("nom_flia") = data_estudio.Recordset("nomflia")
                        If MensajeCMT = vbYes Then
                           data_lineas.Recordset("nom_superv") = "SI"
                        End If
                        data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
                        data_lineas.Recordset("ced_socio") = frmabm.data_clientes.Recordset("cl_cedula")
                        data_lineas.Recordset("fact") = frmabm.data_clientes.Recordset("cl_codced")
                        If Trim(Xtelcmt) <> "" Then
                           data_lineas.Recordset("contact_tel") = Mid(Trim(Xtelcmt), 1, 50)
                        End If
                        If data_estudio.Recordset("codest") = 997 Then
                           data_lineas.Recordset("rub_cont") = 113003
                           data_lineas.Recordset("tipo_mov") = 1
                           Xrub = 113003
                        Else
                           If data_estudio.Recordset("codest") = 993 Or _
                              data_estudio.Recordset("codest") = 994 Then
                              data_lineas.Recordset("rub_cont") = 112022
                              data_lineas.Recordset("tipo_mov") = 1
                              Xrub = 112022
                           Else
                              If data_estudio.Recordset("codest") = 999 Then
                                 data_lineas.Recordset("tipo_mov") = 1
                                 If IsNull(frmabm.data_clientes.Recordset("cl_nrocobr")) = False Then
                                    If frmabm.data_clientes.Recordset("cl_nrocobr") = 615 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 616 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 635 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 602 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 113 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 653 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 672 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 1 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 10 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 201 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 512 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 636 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 685 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 208 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 209 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 8 Or _
                                        frmabm.data_clientes.Recordset("cl_nrocobr") = 0 Then
                                        data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                                        Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                                        data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                                    Else
                                        data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                                        Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                                        data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                                        Xcuotabase = 9
                                    End If
                                 Else
                                    data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                                    Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                                    data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                                 End If
                              Else
                                 If data_estudio.Recordset("codest") = 992 Then
                                    data_lineas.Recordset("rub_cont") = 513007
                                    Xrub = 513007
                                 Else
                                    If Xfpago = 2 Then
                                       data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcrd")
                                       Xrub = frmabm.data_parsec.Recordset("srvcrd")
                                    Else
                                       data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcnt")
                                       Xrub = frmabm.data_parsec.Recordset("srvcnt")
                                    End If
                                 End If
                              End If
                           End If
                        End If
                        If data_estudio.Recordset("codest") = 996 Then
                           data_lineas.Recordset("rub_cont") = 211473
                           data_lineas.Recordset("tipo_mov") = 1
                           Xrub = 211473
                        End If
                        data_codcaja.Recordset.FindFirst "numero =" & Xrub
                        If Not data_codcaja.Recordset.NoMatch Then
                           data_lineas.Recordset("rub_nomb") = data_codcaja.Recordset("nombre")
                        Else
                           data_lineas.Recordset("rub_nomb") = "NO REG."
                        End If
                        If data_estudio.Recordset("codest") = 60106 Then
                           data_lineas.Recordset("rub_cont") = 211397
                           data_lineas.Recordset("rub_nomb") = "M. SMI"
                        End If
                        If data_estudio.Recordset("codest") = 60105 Then
                           data_lineas.Recordset("rub_cont") = 211397
                           data_lineas.Recordset("rub_nomb") = "M. SMI"
                        End If
                        If data_estudio.Recordset("codest") = 60108 Then
                           data_lineas.Recordset("rub_cont") = 211302
                           data_lineas.Recordset("rub_nomb") = "M.UNIVERSAL"
                        End If
                        data_lineas.Recordset("arancel") = Format(txt_precio.Text, "Standard")
                        data_lineas.Recordset("tot_lin") = Format(txt_precio.Text, "Standard")
                        Xiva = data_lineas.Recordset("tot_lin") / 1.1
                        Xiva = Xiva * 0.1
                        If Label5.Caption <> "" Then
                           If Val(Label5.Caption) = 995 Or Val(Label5.Caption) = 990 Then
                              Xiva = 0
                           End If
                        End If
                        data_lineas.Recordset("imp_iva") = Format(Xiva, "Standard")
                        If Label8.Caption = "" Then
                           Label8.Caption = Format(Xiva, "Standard")
                        Else
                           Label8.Caption = CDbl(Label8.Caption) + Xiva
                           Label8.Caption = Format(Label8.Caption, "Standard")
                        End If
                        If data_lineas.Recordset("cod_prod") = 60106 Or _
                           data_lineas.Recordset("cod_prod") = 60108 Or _
                           data_lineas.Recordset("cod_prod") = 994 Or _
                           data_lineas.Recordset("cod_prod") = 993 Or _
                           data_lineas.Recordset("cod_prod") = 996 Or _
                           data_lineas.Recordset("cod_prod") = 60105 Or _
                           data_lineas.Recordset("cod_prod") = 999 Or _
                           data_lineas.Recordset("cod_prod") = 60109 Then
                           data_lineas.Recordset("imp_iva") = 0
                           data_lineas.Recordset("porce_est") = Xelnrodeuda
                           Xiva = 0
                           Label8.Caption = 0
                           Label8.Caption = Format(Label8.Caption, "Standard")
                        Else
                           data_lineas.Recordset("porce_est") = 0
                        End If
                        If Xcuotabase = 9 Then
                           data_lineas.Recordset("imp_iva") = 0
                           Xiva = 0
                           Label8.Caption = 0
                           Label8.Caption = Format(Label8.Caption, "Standard")
                        End If
                        If dbcbomed.Text <> "" Then
                           data_lineas.Recordset("nro_med_a") = labmed.Caption
                           data_lineas.Recordset("nom_med_a") = dbcbomed.Text
                        End If
                        If dbcbomedo.Text <> "" Then
                           data_lineas.Recordset("nro_med_s") = labmedo.Caption
                           data_lineas.Recordset("nom_med_s") = dbcbomedo.Text
                        End If
                        data_lineas.Recordset("precio_est") = Format(txt_precio.Text, "Standard")
                        If data_estudio.Recordset("codest") = 993 Or data_estudio.Recordset("codest") = 999 Or data_estudio.Recordset("codest") = 994 Then
                           data_lineas.Recordset("mes_paga") = Val(txt_mes.Text)
                           data_lineas.Recordset("ano_paga") = Val(txt_ano.Text)
                        End If
                        data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
                        If Xfpago = 2 Then
                           If frmabm.data_parsec.Recordset("base") = 20 And Val(txt_precio.Text) <= 0 And Val(Label5.Caption) = 10050 Then
                              Xfpago = 1
                              data_lineas.Recordset("tipo") = "CONTADO"
                           Else
                              data_lineas.Recordset("tipo") = "CREDITO"
                           End If
                        Else
                           data_lineas.Recordset("tipo") = "CONTADO"
                        End If
                        If labtot.Caption = "" Then
                        Else
                           If Xiva = 0 Then
                              data_lineas.Recordset("costo") = txt_precio.Text
                           Else
                              data_lineas.Recordset("costo") = txt_precio.Text - Xiva
                           End If
                        End If
                        data_lineas.Recordset("linea") = Xcandelin
                        data_lineas.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
                        data_lineas.Recordset("in_unid") = "INT1"
                        If labtot.Caption <> "" Then
                           labtot.Caption = Val(labtot.Caption) + Format(txt_precio.Text, "Standard")
                           labtot.Caption = Format(labtot.Caption, "Standard")
                        Else
                           labtot.Caption = Format(txt_precio.Text, "Standard")
                           labtot.Caption = Format(labtot.Caption, "Standard")
                        End If
                        data_lineas.Recordset.Update
                        data_lineas.Refresh
                        If labtimemi.Caption <> "" Then
                           If Val(labtimemi.Caption) > 0 Then
                              Command4_Click
                           End If
                        End If
                        If labdeudaemi.Caption <> "" Then
                           If Val(labdeudaemi.Caption) > 0 Then
                              Command5_Click
                           End If
                        End If
                        If cbotim.Text = "SI" And Label7.Caption <> "NC E-TICKET" Then
                           data_estudiobus.RecordSource = "select * from estudios where codest =" & 995
                           data_estudiobus.Refresh
                           If data_estudiobus.Recordset.RecordCount > 0 Then
                              XValtim = data_estudiobus.Recordset("cons")
                           Else
                              XValtim = 85
                           End If
                           data_lineas.Recordset.FindFirst "cod_prod =" & 995
                           If Not data_lineas.Recordset.NoMatch Then
                              data_lineas.Recordset.Edit
                              data_lineas.Recordset("tot_lin") = data_lineas.Recordset("tot_lin") + XValtim
                              data_lineas.Recordset("cantidad") = data_lineas.Recordset("cantidad") + 1
                              data_lineas.Recordset.Update
                              labtot.Caption = Val(labtot.Caption) + XValtim
                              labtot.Caption = Format(labtot.Caption, "Standard")
                              labtim.Caption = Val(labtim.Caption) + XValtim
                           Else
                             Xcandelin = Xcandelin + 1
                             data_lineas.Recordset.AddNew
                             data_lineas.Recordset("reg_cab") = 0
                             data_lineas.Recordset("factura") = 0
                             data_lineas.Recordset("tipo_mov") = 1
                             data_lineas.Recordset("realizada") = Format(mf.Text, "dd/mm/yyyy")
                             data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                             data_lineas.Recordset("cod_cli") = labmatri.Caption
                             data_lineas.Recordset("nom_cli") = labnomb.Caption
                             data_lineas.Recordset("cod_prod") = data_estudiobus.Recordset("codest")
                             data_lineas.Recordset("nom_prod") = data_estudiobus.Recordset("descrip")
                             If txt_rut.Visible = True Then
                                If Trim(txt_rut.Text) <> "" Then
                                   data_lineas.Recordset("ruc") = txt_rut.Text
                                End If
                             End If
                             data_lineas.Recordset("cantidad") = 1
                             data_lineas.Recordset("moneda") = "SR"
                             data_lineas.Recordset("operador") = WElusuario
                             data_lineas.Recordset("hora") = Format(Time, "HH:mm")
                             data_lineas.Recordset("nro_flia") = data_estudiobus.Recordset("flia")
                             data_lineas.Recordset("nom_flia") = data_estudiobus.Recordset("nomflia")
                             data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
                             data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
                             data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
                             data_lineas.Recordset("usa_timbre") = "S"
                             If Xestco = 13010 Or Xestco = 13014 Or Xestco = 13017 Or _
                                Xestco = 13017 Or Xestco = 13022 Or Xestco = 13034 Or Xestco = 13026 Or Xestfl = 3 Then
                                data_lineas.Recordset("rub_cont") = 211332
                                data_lineas.Recordset("rub_nomb") = "FERTILAB"
                             Else
                                If Xestco = 13019 Or Xestco = 13021 Or Xestco = 13024 Or _
                                   Xestco = 13029 Or Xestco = 13037 Or Xestfl = 5 Then
                                   data_lineas.Recordset("rub_cont") = 211587
                                   data_lineas.Recordset("rub_nomb") = "ECOGRAFISTA"
                                Else
                                   If Xestco = 13009 Or Xestco = 13013 Or Xestco = 13023 Or _
                                      Xestco = 13027 Or Xestco = 13035 Or Xestfl = 7 Then
                                      data_lineas.Recordset("rub_cont") = 211313
                                      data_lineas.Recordset("rub_nomb") = "RADIOLOGOS"
                                   Else
                                      If Xestco = 13012 Or Xestco = 13016 Or Xestco = 13038 Or _
                                         Xestfl = 11 Then
                                         data_lineas.Recordset("rub_cont") = 211586
                                         data_lineas.Recordset("rub_nomb") = "SERV.CARDIOLOGICOS"
                                      Else
                                         data_lineas.Recordset("rub_cont") = 213076
                                         data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
                                      End If
                                   End If
                                End If
                             End If
                             data_lineas.Recordset("arancel") = XValtim
                             data_lineas.Recordset("tot_lin") = XValtim
                             If labtot.Caption <> "" Then
                                labtot.Caption = Val(labtot.Caption) + XValtim
                                labtot.Caption = Format(labtot.Caption, "Standard")
                             Else
                                labtot.Caption = XValtim
                                labtot.Caption = Format(labtot.Caption, "Standard")
                             End If
                             data_lineas.Recordset("precio_est") = XValtim
                             data_lineas.Recordset("porce_est") = 0
                             data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
                             If Xfpago = 2 Then
                                data_lineas.Recordset("tipo") = "CREDITO"
                             Else
                                data_lineas.Recordset("tipo") = "CONTADO"
                             End If
                             data_lineas.Recordset("linea") = Xcandelin
                             data_lineas.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
                             data_lineas.Recordset("in_unid") = "INT1"
                             data_lineas.Recordset.Update
                             labtim.Caption = XValtim
                             data_lineas.Refresh
                             If data_estudio.Recordset("codest") = 997 Then
                                Xcoddeu = 9
                             End If
                           End If
                        End If
                   End If
                Else ' aca es medicacion y más
                    Xcandelin = Xcandelin + 1
                    data_lineas.Recordset.AddNew
                    If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "ND E-FACTURA" Or Label7.Caption = "ND E-TICKET" Then
                       ''Xmotivoref = InputBox("Ingrese motivo de la modificación")
                       Xmotivoref = "SIN DATOS"
                       If Xmotivoref <> "" Then
                          labmotivo.Caption = Xmotivoref
                          If labseriecance.Caption = "XX" Then
                             data_lincance.RecordSource = "Select * from linmmdd where factura =" & Val(labfaccance.Caption) & " and fecha =#" & Format(labfeccance.Caption, "yyyy-mm-dd") & "#"
                             data_lincance.Refresh
                          Else
                             If labidemi.Caption = 5 Then
                                data_lincance.RecordSource = "Select * from deudas where documento =" & Val(labfaccance.Caption)
                                data_lincance.Refresh
                             Else
                                data_lincance.RecordSource = "Select * from clirespl where cl_numero =" & Val(labfaccance.Caption) & " and cl_socmnro ='" & Trim(labseriecance.Caption) & "' and cl_fnac =#" & Format(labfeccance.Caption, "yyyy-mm-dd") & "#"
                                data_lincance.Refresh
                             End If
                          End If
                          If data_lincance.Recordset.RecordCount > 0 And Val(labidemi.Caption) <> 5 Then
                             If labseriecance.Caption <> "XX" Then
                                If Val(labtot.Caption) >= data_lincance.Recordset("saldo_doc") Then
                                   MsgBox "Ha excedido el importe de la factura", vbCritical
                                   b_cance_Click
                                End If
                             End If
                             If labseriecance.Caption = "XX" Then
                                data_lineas.Recordset("tipodocref") = 101
                             Else
                                data_lineas.Recordset("tipodocref") = data_lincance.Recordset("cl_tipocli")
                             End If
                             data_lineas.Recordset("serieref") = labseriecance.Caption
                             If Len(labfaccance.Caption) > 7 Then
                                data_lineas.Recordset("nrofactref") = Val(Mid(labfaccance.Caption, 1, 7))
                             Else
                                data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
                             End If
                             If labseriecance.Caption = "XX" Then
                                data_lineas.Recordset("fechafact") = CDate(labfeccance.Caption)
                             Else
                                data_lineas.Recordset("fechafact") = data_lincance.Recordset("cl_fnac")
                             End If
                             data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                             data_lineas.Recordset("linearef") = Val(lablinea.Caption)
                          Else
                             If labidemi.Caption = 5 Then
                                If Label7.Caption = "NC E-FACTURA" Then
                                   data_lineas.Recordset("tipodocref") = 111
                                Else
                                   data_lineas.Recordset("tipodocref") = 101
                                End If
                                data_lineas.Recordset("serieref") = labseriecance.Caption
                                If Len(labfaccance.Caption) > 7 Then
                                   data_lineas.Recordset("nrofactref") = Val(Mid(labfaccance.Caption, 1, 7))
                                Else
                                   data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
                                End If
                                data_lineas.Recordset("fechafact") = CDate(labfeccance.Caption)
                                data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                                data_lineas.Recordset("linearef") = 1
                             Else
                                MsgBox "No se encuentra número de factura a cancelar"
                                b_cance_Click
                             End If
                          End If
                       Else
                          MsgBox "No ingresó motivo de cancelación"
                          b_cance_Click
                       End If
                       btn_fin.SetFocus
                    Else
                       If Label7.Caption = "DEV.RECIBO" Then
                          If labfaccance.Caption <> "" Then
                             data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
                          End If
                       End If
                    End If
                    If labcodpro.Caption <> "" Then
                       data_lineas.Recordset("numero") = Val(labcodpro.Caption)
                    End If
                    data_lineas.Recordset("reg_cab") = 99
                    data_lineas.Recordset("factura") = 0
                    If txt_precio.Text <> "" Then
                       If txt_precio.Text > 0 Then
                          If Val(Label5.Caption) = 995 Or Val(Label5.Caption) = 990 Then
                             data_lineas.Recordset("tipo_mov") = 1 'tipo de iva (indic de fact)
                          Else
                             data_lineas.Recordset("tipo_mov") = 2 'tipo de iva (indic de fact)
                          End If
                       Else
                          data_lineas.Recordset("tipo_mov") = 5
                       End If
                    Else
                       data_lineas.Recordset("tipo_mov") = 5
                    End If
                    data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                    data_lineas.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
                    data_lineas.Recordset("cod_cli") = labmatri.Caption
                    data_lineas.Recordset("nom_cli") = Mid(labnomb.Caption, 1, 30)
                    data_lineas.Recordset("cod_prod") = data_estudio.Recordset("codest")
                    data_lineas.Recordset("nom_prod") = Mid(data_estudio.Recordset("descrip"), 1, 50)
                    data_lineas.Recordset("cantidad") = 1
                    data_lineas.Recordset("moneda") = "SR" 'Serie
                    If txt_rut.Visible = True Then
                       If Trim(txt_rut.Text) <> "" Then
                          Command2_Click
                          data_lineas.Recordset("ruc") = txt_rut.Text
                       End If
                    End If
                    If data_estudio.Recordset("flia") = 3 Then
                       data_lineas.Recordset("tcambio") = 8
                    End If
                    If data_estudio.Recordset("codest") = 60103 Or _
                       data_estudio.Recordset("codest") = 60105 Or _
                       data_estudio.Recordset("codest") = 60106 Or _
                       data_estudio.Recordset("codest") = 60107 Or _
                       data_estudio.Recordset("codest") = 60108 Or _
                       data_estudio.Recordset("codest") = 60109 Then
                       If data_estudio.Recordset("codest") <> 60107 Then
                          If data_estudio.Recordset("codest") <> 60103 Then
                             data_lineas.Recordset("tipo_mov") = 1
                          End If
                       End If
                       data_lineas.Recordset("codelmedic") = XcodelMedicamento
                       data_lineas.Recordset("idtablapres") = IdTablaPres
                       If Xdescmedic <> "" Then
                          data_lineas.Recordset("nom_medic") = Mid(UCase(Xdescmedic), 1, 50)
                       End If
                       If txt_precio.Text <> "" Then
                          If txt_precio.Text < 0 Then
                             data_lineas.Recordset("dias") = 1
                          Else
                             data_lineas.Recordset("dias") = 0
                          End If
                       Else
                          data_lineas.Recordset("dias") = 0
                       End If
                    End If
                    data_lineas.Recordset("operador") = WElusuario
                    data_lineas.Recordset("hora") = Format(Time, "HH:mm")
                    data_lineas.Recordset("nro_flia") = data_estudio.Recordset("flia")
                    data_lineas.Recordset("nom_flia") = data_estudio.Recordset("nomflia")
                    data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
                    data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
                    data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
                    data_lineas.Recordset("ced_socio") = frmabm.data_clientes.Recordset("cl_cedula")
                    data_lineas.Recordset("fact") = frmabm.data_clientes.Recordset("cl_codced")
                    If Trim(Xtelcmt) <> "" Then
                       data_lineas.Recordset("contact_tel") = Mid(Trim(Xtelcmt), 1, 50)
                    End If
                    
                    If data_estudio.Recordset("codest") = 997 Then
                       data_lineas.Recordset("rub_cont") = 113003
                       data_lineas.Recordset("tipo_mov") = 1
                       Xrub = 113003
                    Else
                       If data_estudio.Recordset("codest") = 993 Or _
                          data_estudio.Recordset("codest") = 994 Then
                          data_lineas.Recordset("rub_cont") = 112022
                          data_lineas.Recordset("tipo_mov") = 1
                          Xrub = 112022
                       Else
                          If data_estudio.Recordset("codest") = 999 Then
                             data_lineas.Recordset("tipo_mov") = 1
                             If IsNull(frmabm.data_clientes.Recordset("cl_nrocobr")) = False Then
                                If frmabm.data_clientes.Recordset("cl_nrocobr") = 615 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 616 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 635 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 602 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 113 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 653 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 672 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 1 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 10 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 201 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 512 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 636 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 685 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 208 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 209 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 8 Or _
                                    frmabm.data_clientes.Recordset("cl_nrocobr") = 0 Then
                                    data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                                    Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                                    data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                                Else
                                    data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                                    Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                                    data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                                    Xcuotabase = 9
                                End If
                             Else
                                data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                                Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                                data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                             End If
                          Else
                             If data_estudio.Recordset("codest") = 992 Then
                                data_lineas.Recordset("rub_cont") = 513007
                                Xrub = 513007
                             Else
                                If Xfpago = 2 Then
                                   data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcrd")
                                   Xrub = frmabm.data_parsec.Recordset("srvcrd")
                                Else
                                   data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcnt")
                                   Xrub = frmabm.data_parsec.Recordset("srvcnt")
                                End If
                             End If
                          End If
                       End If
                    End If
                    If data_estudio.Recordset("codest") = 996 Then
                       data_lineas.Recordset("rub_cont") = 211473
                       data_lineas.Recordset("tipo_mov") = 1
                       Xrub = 211473
                    End If
                    data_codcaja.Recordset.FindFirst "numero =" & Xrub
                    If Not data_codcaja.Recordset.NoMatch Then
                       data_lineas.Recordset("rub_nomb") = data_codcaja.Recordset("nombre")
                    End If
                    If data_estudio.Recordset("codest") = 60106 Then
                       data_lineas.Recordset("rub_cont") = 211397
                       data_lineas.Recordset("rub_nomb") = "M. SMI"
                    End If
                    If data_estudio.Recordset("codest") = 60105 Then
                       data_lineas.Recordset("rub_cont") = 211397
                       data_lineas.Recordset("rub_nomb") = "M. SMI"
                    End If
                    If data_estudio.Recordset("codest") = 60108 Then
                       data_lineas.Recordset("rub_cont") = 211302
                       data_lineas.Recordset("rub_nomb") = "M.UNIVERSAL"
                    End If
                    If data_estudio.Recordset("codest") = 60109 Then
                       data_lineas.Recordset("rub_cont") = 211372
                       data_lineas.Recordset("rub_nomb") = "M.CGAL"
                    End If
                    
                    If (Xestco = 60107 Or Xestco = 60103) And Label7.Caption = "E-TICKET" Then
                       If txt_precio.Text > 0 Then
    '                      data_estudio.RecordSource = "estudios"
    '                      data_estudio.Refresh
                          data_estudiobus.RecordSource = "Select * from estudios where codest =" & 990
                          data_estudiobus.Refresh
                          If data_estudiobus.Recordset.RecordCount > 0 Then
                             Xvaltimme = data_estudiobus.Recordset("cons")
                             If labtimme.Caption = "" Then
                                labtimme.Caption = Xvaltimme
                             Else
                                labtimme.Caption = Val(labtimme.Caption) + Xvaltimme
                             End If
                          Else
                             Xvaltimme = 18
                             If labtimme.Caption = "" Then
                                labtimme.Caption = Xvaltimme
                             Else
                                labtimme.Caption = Val(labtimme.Caption) + Xvaltimme
                             End If
                          End If
                          If CDbl(txt_precio.Text) >= CDbl(Xvaltimme) Then
                             data_lineas.Recordset("arancel") = Format(txt_precio.Text - Xvaltimme, "Standard")
                             data_lineas.Recordset("tot_lin") = Format(txt_precio.Text - Xvaltimme, "Standard")
                          Else
                             MsgBox "Importe menor al timbre, VERIFIQUE!!!", vbCritical
                             frmabm.btn_fact.Enabled = True
                             Unload Me
                             Exit Sub
                             Xvaltimme = 0
                             labtimme.Caption = 0
                             data_lineas.Recordset("arancel") = Format(txt_precio.Text, "Standard")
                             data_lineas.Recordset("tot_lin") = Format(txt_precio.Text, "Standard")
                          End If
                       Else
                          data_lineas.Recordset("arancel") = Format(txt_precio.Text, "Standard")
                          data_lineas.Recordset("tot_lin") = Format(txt_precio.Text, "Standard")
                       End If
                    Else
                       data_lineas.Recordset("arancel") = Format(txt_precio.Text, "Standard")
                       data_lineas.Recordset("tot_lin") = Format(txt_precio.Text, "Standard")
                    End If
        'hasta aqui
                    Xiva = data_lineas.Recordset("tot_lin") / 1.1
                    Xiva = Xiva * 0.1
                    If Label5.Caption <> "" Then
                       If Val(Label5.Caption) = 995 Or Val(Label5.Caption) = 990 Then
                          Xiva = 0
                       End If
                    End If
                    data_lineas.Recordset("imp_iva") = Format(Xiva, "Standard")
                    If Label8.Caption = "" Then
                       Label8.Caption = Format(Xiva, "Standard")
                    Else
                       Label8.Caption = CDbl(Label8.Caption) + Xiva
                       Label8.Caption = Format(Label8.Caption, "Standard")
                    End If
                    If data_lineas.Recordset("cod_prod") = 60106 Or _
                       data_lineas.Recordset("cod_prod") = 60108 Or _
                       data_lineas.Recordset("cod_prod") = 994 Or _
                       data_lineas.Recordset("cod_prod") = 993 Or _
                       data_lineas.Recordset("cod_prod") = 996 Or _
                       data_lineas.Recordset("cod_prod") = 60105 Or _
                       data_lineas.Recordset("cod_prod") = 60109 Or _
                       data_lineas.Recordset("cod_prod") = 999 Or _
                       data_lineas.Recordset("cod_prod") = 80011 Or _
                       data_lineas.Recordset("cod_prod") = 80012 Or _
                       data_lineas.Recordset("cod_prod") = 80013 Or _
                       data_lineas.Recordset("cod_prod") = 80014 Or _
                       data_lineas.Recordset("cod_prod") = 80016 Or _
                       data_lineas.Recordset("cod_prod") = 80015 Then
    
                       data_lineas.Recordset("imp_iva") = 0
                       Xiva = 0
                       Label8.Caption = 0
                       Label8.Caption = Format(Label8.Caption, "Standard")
                    End If
                    If Xcuotabase = 9 Then
                       data_lineas.Recordset("imp_iva") = 0
                       Xiva = 0
                       Label8.Caption = 0
                       Label8.Caption = Format(Label8.Caption, "Standard")
                    End If
                    If dbcbomed.Text <> "" Then
                       data_lineas.Recordset("nro_med_a") = labmed.Caption
                       data_lineas.Recordset("nom_med_a") = dbcbomed.Text
                    End If
                    If dbcbomedo.Text <> "" Then
                       data_lineas.Recordset("nro_med_s") = labmedo.Caption
                       data_lineas.Recordset("nom_med_s") = dbcbomedo.Text
                    End If
                    If (Xestco = 60107 Or Xestco = 60103) And Label7.Caption = "E-TICKET" Then
                       If txt_precio.Text > 0 Then
                          If Val(txt_precio.Text) >= Xvaltimme Then
                             data_lineas.Recordset("precio_est") = Format(txt_precio.Text - Xvaltimme, "Standard")
                          Else
                             data_lineas.Recordset("precio_est") = Format(txt_precio.Text, "Standard")
                          End If
                       Else
                          data_lineas.Recordset("precio_est") = Format(txt_precio.Text, "Standard")
                       End If
                    Else
                       data_lineas.Recordset("precio_est") = Format(txt_precio.Text, "Standard")
                    End If
                    If Xestco = 993 Or Xestco = 999 Or Xestco = 994 Or Xestco = 997 Then
                       If Xestco = 997 Then
                          data_lineas.Recordset("porce_est") = Xelnrodeuda
                       Else
                          data_lineas.Recordset("mes_paga") = Val(txt_mes.Text)
                          data_lineas.Recordset("ano_paga") = Val(txt_ano.Text)
                          data_lineas.Recordset("porce_est") = Xelnrodeuda
                       End If
                    Else
                       data_lineas.Recordset("porce_est") = 0
                    End If
                    data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
                    If labtot.Caption = "" Then
                    Else
                       If (Xestco = 60107 Or Xestco = 60103) And Label7.Caption = "E-TICKET" Then
                          If txt_precio.Text > 0 Then
                             If Val(txt_precio.Text) >= Xvaltimme Then
                                If Xiva = 0 Then
                                   data_lineas.Recordset("costo") = txt_precio.Text - Xvaltimme
                                Else
                                   data_lineas.Recordset("costo") = txt_precio.Text - Xvaltimme - Xiva
                                End If
                             Else
                                If Xiva = 0 Then
                                   data_lineas.Recordset("costo") = txt_precio.Text
                                Else
                                   data_lineas.Recordset("costo") = txt_precio.Text - Xiva
                                End If
                             End If
                          Else
                             If Xiva = 0 Then
                                data_lineas.Recordset("costo") = txt_precio.Text
                             Else
                                data_lineas.Recordset("costo") = txt_precio.Text - Xiva
                             End If
                          End If
                       Else
                          If Xiva = 0 Then
                             data_lineas.Recordset("costo") = txt_precio.Text
                          Else
                             data_lineas.Recordset("costo") = txt_precio.Text - Xiva
                          End If
                       End If
                    End If
                    If Xfpago = 2 Then
                       data_lineas.Recordset("tipo") = "CREDITO"
                    Else
                       data_lineas.Recordset("tipo") = "CONTADO"
                    End If
                    If labtot.Caption <> "" Then
                       labtot.Caption = Val(labtot.Caption) + Val(txt_precio.Text)
                       labtot.Caption = Format(labtot.Caption, "Standard")
                    Else
                       labtot.Caption = Format(txt_precio.Text, "Standard")
                       labtot.Caption = Format(labtot.Caption, "Standard")
                    End If
                    data_lineas.Recordset("linea") = Xcandelin
                    data_lineas.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
                    data_lineas.Recordset("in_unid") = "INT1"
                    data_lineas.Recordset.Update
                    data_lineas.Refresh
                    If labtimemi.Caption <> "" Then
                       If Format(labtimemi.Caption, "Standard") > 0 Then
                          Command4_Click
                       End If
                    End If
                    If labdeudaemi.Caption <> "" Then
                       If Format(labdeudaemi.Caption, "Standard") > 0 Then
                          Command5_Click
                       End If
                    End If
                    If (Xestco = 60107 Or Xestco = 60103) And txt_precio.Text > 0 And Label7.Caption = "E-TICKET" Then
                       data_lineas.Recordset.FindFirst "cod_prod =" & 990
                       If Not data_lineas.Recordset.NoMatch Then
                          data_lineas.Recordset.Edit
                          data_lineas.Recordset("tot_lin") = data_lineas.Recordset("tot_lin") + Xvaltimme
                          data_lineas.Recordset("cantidad") = data_lineas.Recordset("cantidad") + 1
                          data_lineas.Recordset.Update
                          labtim.Caption = Val(labtim.Caption) + Xvaltimme
                       Else
                          Xcandelin = Xcandelin + 1
                          data_lineas.Recordset.AddNew
                          data_lineas.Recordset("reg_cab") = 0
                          data_lineas.Recordset("factura") = 0
                          data_lineas.Recordset("tipo_mov") = 1
                          data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                          data_lineas.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
                          data_lineas.Recordset("cod_cli") = labmatri.Caption
                          data_lineas.Recordset("nom_cli") = labnomb.Caption
                          data_lineas.Recordset("cod_prod") = 990
                          data_lineas.Recordset("nom_prod") = "TIMBRES PROFESIONAL M"
                          data_lineas.Recordset("usa_timbre") = "M"
                          data_lineas.Recordset("moneda") = "SR" 'Serie
                          If txt_rut.Visible = True Then
                             If Trim(txt_rut.Text) <> "" Then
                                Command2_Click
                                data_lineas.Recordset("ruc") = txt_rut.Text
                             End If
                          End If
                          data_lineas.Recordset("cantidad") = 1
                          data_lineas.Recordset("operador") = WElusuario
                          data_lineas.Recordset("hora") = Format(Time, "HH:mm")
                          data_lineas.Recordset("nro_flia") = 8
                          data_lineas.Recordset("nom_flia") = "OTROS SERVICIOS"
                          data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
                          data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
                          data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
                          data_lineas.Recordset("rub_cont") = 213076
                          data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
                          data_lineas.Recordset("arancel") = Xvaltimme
                          data_lineas.Recordset("tot_lin") = Xvaltimme
                          data_lineas.Recordset("precio_est") = Xvaltimme
                          data_lineas.Recordset("porce_est") = 0
                          data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
                          If Xfpago = 2 Then
                             data_lineas.Recordset("tipo") = "CREDITO"
                          Else
                             data_lineas.Recordset("tipo") = "CONTADO"
                          End If
                          data_lineas.Recordset("linea") = Xcandelin
                          data_lineas.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
                          data_lineas.Recordset("in_unid") = "INT1"
                          data_lineas.Recordset.Update
                          data_lineas.Refresh
                          labtim.Caption = Xvaltimme
                       End If
                    End If
                    If data_estudio.Recordset("codest") = 997 Then
                       Xcoddeu = 9
                    End If
                    If cbotim.Text = "SI" And Label7.Caption <> "NC E-TICKET" Then
                       data_estudiobus.RecordSource = "select * from estudios where codest =" & 995
                       data_estudiobus.Refresh
                       If data_estudiobus.Recordset.RecordCount > 0 Then
                          XValtim = data_estudiobus.Recordset("cons")
                       Else
                          XValtim = 59
                       End If
                       data_lineas.Recordset.FindFirst "cod_prod =" & 995
                       If Not data_lineas.Recordset.NoMatch Then
                          data_lineas.Recordset.Edit
                          data_lineas.Recordset("tot_lin") = data_lineas.Recordset("tot_lin") + XValtim
                          data_lineas.Recordset("cantidad") = data_lineas.Recordset("cantidad") + 1
                          data_lineas.Recordset.Update
                          If labtot.Caption <> "" Then
                             labtot.Caption = Val(labtot.Caption) + XValtim
                             labtot.Caption = Format(labtot.Caption, "Standard")
                             labtim.Caption = Val(labtim.Caption) + XValtim
                          Else
                             labtot.Caption = XValtim
                             labtot.Caption = Format(labtot.Caption, "Standard")
                             labtim.Caption = Val(labtim.Caption) + XValtim
                          End If
                       Else
                         Xcandelin = Xcandelin + 1
                         data_lineas.Recordset.AddNew
                         data_lineas.Recordset("reg_cab") = 0
                         data_lineas.Recordset("factura") = 0
                         data_lineas.Recordset("tipo_mov") = 1
                         data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                         data_lineas.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
                         data_lineas.Recordset("cod_cli") = labmatri.Caption
                         data_lineas.Recordset("nom_cli") = labnomb.Caption
                         data_lineas.Recordset("cod_prod") = data_estudiobus.Recordset("codest")
                         data_lineas.Recordset("nom_prod") = data_estudiobus.Recordset("descrip")
                         If txt_rut.Visible = True Then
                            If Trim(txt_rut.Text) <> "" Then
                               Command2_Click
                               data_lineas.Recordset("ruc") = txt_rut.Text
                            End If
                         End If
                         data_lineas.Recordset("cantidad") = 1
                         data_lineas.Recordset("moneda") = "SR"
                         data_lineas.Recordset("operador") = WElusuario
                         data_lineas.Recordset("hora") = Format(Time, "HH:mm")
                         data_lineas.Recordset("nro_flia") = data_estudiobus.Recordset("flia")
                         data_lineas.Recordset("nom_flia") = data_estudiobus.Recordset("nomflia")
                         data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
                         data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
                         data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
                         data_lineas.Recordset("usa_timbre") = "S"
                         If Xestco = 13010 Or Xestco = 13014 Or Xestco = 13017 Or _
                            Xestco = 13017 Or Xestco = 13022 Or Xestco = 13034 Or Xestco = 13026 Or Xestfl = 3 Then
                            data_lineas.Recordset("rub_cont") = 211332
                            data_lineas.Recordset("rub_nomb") = "FERTILAB"
                         Else
                            If Xestco = 13019 Or Xestco = 13021 Or Xestco = 13024 Or _
                               Xestco = 13029 Or Xestco = 13037 Or Xestfl = 5 Then
                               data_lineas.Recordset("rub_cont") = 211587
                               data_lineas.Recordset("rub_nomb") = "ECOGRAFISTA"
                            Else
                               If Xestco = 13009 Or Xestco = 13013 Or Xestco = 13023 Or _
                                  Xestco = 13027 Or Xestco = 13035 Or Xestfl = 7 Then
                                  data_lineas.Recordset("rub_cont") = 211313
                                  data_lineas.Recordset("rub_nomb") = "RADIOLOGOS"
                               Else
                                  If Xestco = 13012 Or Xestco = 13016 Or Xestco = 13038 Or _
                                     Xestfl = 11 Then
                                     data_lineas.Recordset("rub_cont") = 211586
                                     data_lineas.Recordset("rub_nomb") = "SERV.CARDIOLOGICOS"
                                  Else
                                     data_lineas.Recordset("rub_cont") = 213076
                                     data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
                                  End If
                               End If
                            End If
                         End If
                         data_lineas.Recordset("arancel") = XValtim
                         data_lineas.Recordset("tot_lin") = XValtim
                         If labtot.Caption <> "" Then
                            labtot.Caption = Val(labtot.Caption) + XValtim
                            labtot.Caption = Format(labtot.Caption, "Standard")
                         Else
                            labtot.Caption = XValtim
                            labtot.Caption = Format(labtot.Caption, "Standard")
                         End If
                         data_lineas.Recordset("precio_est") = XValtim
                         data_lineas.Recordset("porce_est") = 0
                         data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
                         If Xfpago = 2 Then
                            data_lineas.Recordset("tipo") = "CREDITO"
                         Else
                            data_lineas.Recordset("tipo") = "CONTADO"
                         End If
                         data_lineas.Recordset("linea") = Xcandelin
                         data_lineas.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
                         data_lineas.Recordset("in_unid") = "INT1"
                         data_lineas.Recordset.Update
           
    '                     Xcantveces = Xcantveces + 1
                         
                         data_lineas.Refresh
                         labtim.Caption = XValtim
                         If data_estudiobus.Recordset("codest") = 997 Then
                            Xcoddeu = 9
                         End If
                       End If
                    End If
                End If
           Else
              MsgBox "Debe facturar cómo RECIBO!!", vbCritical, "FACTURAR"
              Xquerr = 0
              b_cance_Click
           End If
           Xcantveces = Xcantveces + 1
        Loop
        labtimme.Caption = ""
        labmed.Caption = ""
        txt_precio.Text = 0
        cbotim.ListIndex = 0
        dbcbomed.Text = ""
        DBCombo1.Text = ""
        txt_mes.Text = ""
        txt_ano.Text = ""
        Label12.Visible = False
        If DBCombo1.Enabled = True Then
           DBCombo1.SetFocus
        End If
    
    End If
    t_cant.Text = 1
    Xcantveces = 1
    
    Xcuotabase = 0
    
    If data_lineas.Recordset.RecordCount > 0 Then
       data_lineas.Recordset.MoveFirst
       labtot.Caption = 0
       Label8.Caption = 0
       Do While Not data_lineas.Recordset.EOF
          labtot.Caption = Val(labtot.Caption) + data_lineas.Recordset("tot_lin")
          Label8.Caption = Val(Label8.Caption) + data_lineas.Recordset("imp_iva")
          data_lineas.Recordset.MoveNext
       Loop
       labtot.Caption = Format(labtot.Caption, "Standard")
       Label8.Caption = Format(Label8.Caption, "Standard")
    End If
    
    'If Xnomedica = 8 Or Xnomedica = 7 Then
    'Else
    Label5.Caption = ""
    'End If
    data_lindbgri.RecordSource = "select * from lineas"
    data_lindbgri.Refresh
    data_lindbgri.RecordSource = ""
End If

Exit Sub

Vererror:
         If Err.Number = 3421 Then
            MsgBox "Verifique datos ingresados", vbCritical, "Mensaje"
            If DBCombo1.Enabled = True Then
               DBCombo1.SetFocus
            End If
         Else
            MsgBox "Hay un error en los datos de la factura ANOTE ERROR Y ENVIE A COMPUTOS!" & Err.Description, vbCritical, "Mensaje"
            If DBCombo1.Enabled = True Then
               DBCombo1.SetFocus
            End If
         End If


End Sub

Private Sub cbotim_KeyPress(KeyAscii As Integer)
On Error GoTo Altimbre

If KeyAscii = 13 Then
   If txt_mes.Enabled = True Then
      txt_mes.SetFocus
   Else
      dbcbomed.SetFocus
   End If
End If

Exit Sub

Altimbre:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al timbre"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al timbre"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub cbotim_LostFocus()
On Error GoTo Alcbotimbre

   If data_estudio.Recordset("codest") = 997 Then
      cbotim.ListIndex = 0
   End If

Exit Sub

Alcbotimbre:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbo timbre"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbo timbre"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub



Private Sub Command1_Click()
Dim Xsaldocaj As Long
Dim Xquetot As Long
Dim Xquees As Integer
Dim Xsaldofac As Double
Dim Xlincan As Long
Dim Xnromedic As Long
Dim Xenquelugar As Integer
Dim Xesmedevang As Integer

Xenquelugar = 0
Xesmedevang = 0

On Error GoTo Xelerrfactura

Xcoddeu = 0
labtimme.Caption = ""

'data_lincab.RecordSource = "Select * from clirespl where cl_numero =" & labfac.Caption & " and cl_socmnro ='" & labserie.Caption & "'"
'data_lincab.Refresh

data_cabeza2.Recordset.Edit
data_cabeza2.Recordset("cl_socmnro") = labserie.Caption
data_cabeza2.Recordset("cl_numero") = Val(labfac.Caption)
data_cabeza2.Recordset.Update
data_cabeza2.Refresh

If data_cabeza2.Recordset.RecordCount > 0 Then
   data_cabeza2.Recordset.MoveFirst
End If

'fin de cabezal
If data_lineas.Recordset.RecordCount > 0 Then
   data_lineas.Recordset.MoveFirst
End If
If Xfpago = 2 And (Label7.Caption = "E-TICKET" Or Label7.Caption = "E-FACTURA") Then
   frm_credito.Show vbModal
   If Xestaok = 1 Then
   Else
      MsgBox "No ingresó todos los datos del crédito. Avise a Administración.", vbCritical
'      b_cance_Click
   End If
Else
   If data_lineas.Recordset("cod_prod") = 991 Then
      frm_solhc.Show vbModal
      If Xestaok = 1 Then
      Else
         MsgBox "No ingresó todos los datos de la solicitud. Avise a Registros médicos", vbCritical
'         b_cance_Click
      End If
   End If
End If
If data_caja.Recordset.RecordCount > 0 Then
   Xsaldocaj = data_caja.Recordset("saldo")
Else
   Xsaldocaj = 0
End If

Do While Not data_lineas.Recordset.EOF
   If data_lineas.Recordset("cod_prod") = 997 And data_lineas.Recordset("tot_lin") > 0 Then
      If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "DEV.RECIBO" Then
         If frmabm.labdeudap.Caption <> "" Then
            If Val(frmabm.labdeudap.Caption) > 0 Then
               frmabm.labdeudap.Caption = Val(frmabm.labdeudap.Caption) + Val(data_lineas.Recordset("tot_lin"))
            Else
               frmabm.labdeudap.Caption = data_lineas.Recordset("tot_lin")
            End If
         Else
            frmabm.labdeudap.Caption = data_lineas.Recordset("tot_lin")
         End If
                     
         data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " and documento =" & Xelnrodeuda & " and fecha_pago is not null"
         data_deudas.Refresh
         If data_deudas.Recordset.RecordCount > 0 Then
            If IsNull(data_deudas.Recordset("fecha_pago")) = False Then
               data_deudas.Recordset.Edit
               data_deudas.Recordset("fecha_pago") = Null
               data_deudas.Recordset.Update
            End If
         Else
         End If
      Else
         If frmabm.data_clientes.Recordset("saldo_cc") >= 1 Then
            Xsaldofac = frmabm.data_clientes.Recordset("saldo_cc") - data_lineas.Recordset("tot_lin")
         Else
            Xsaldofac = 0
         End If
         If frmabm.labdeudap.Caption <> "" Then
            If frmabm.labdeudap.Caption > 0 Then
               frmabm.labdeudap.Caption = Val(frmabm.labdeudap.Caption) - Val(data_lineas.Recordset("tot_lin"))
            End If
         Else
            frmabm.labdeudap.Caption = 0
         End If
         data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " and documento =" & Xelnrodeuda & " and fecha_pago is null"
         data_deudas.Refresh
         If data_deudas.Recordset.RecordCount > 0 Then
            If IsNull(data_deudas.Recordset("fecha_pago")) = True Then
               data_deudas.Recordset.Edit
               data_deudas.Recordset("fecha_pago") = Date
               data_deudas.Recordset.Update
            End If
         End If
      End If
   End If
   If data_lineas.Recordset("cod_prod") = 999 Then
      If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "DEV.RECIBO" Then
         If frmabm.labatra.Caption <> "" Then
            If frmabm.labatra.Caption > 0 Then
               frmabm.labatra.Caption = Val(frmabm.labatra.Caption) + 1
            End If
         Else
            frmabm.labatra.Caption = 1
         End If
         If frmabm.labdeudap.Caption <> "" Then
            If frmabm.labdeudap.Caption > 0 Then
               frmabm.labdeudap.Caption = Val(frmabm.labdeudap.Caption) + data_lineas.Recordset("tot_lin")
            Else
               frmabm.labdeudap.Caption = data_lineas.Recordset("tot_lin")
            End If
         Else
            frmabm.labdeudap.Caption = data_lineas.Recordset("tot_lin")
         End If
         If frmabm.labump.Caption <> "" Then
            If frmabm.labump.Caption = 1 Then
               frmabm.labump.Caption = 12
               frmabm.labuap.Caption = data_lineas.Recordset("ano_paga") - 1
            Else
               frmabm.labump.Caption = data_lineas.Recordset("mes_paga") - 1
               frmabm.labuap.Caption = data_lineas.Recordset("ano_paga")
            End If
         Else
            frmabm.labump.Caption = data_lineas.Recordset("mes_paga") - 1
            frmabm.labuap.Caption = data_lineas.Recordset("ano_paga")
         End If
         data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " and mes =" & data_lineas.Recordset("mes_paga") & " and ano =" & data_lineas.Recordset("ano_paga") & " and fecha_pago is not null"
         data_deudas.Refresh
         If data_deudas.Recordset.RecordCount > 0 Then
            If IsNull(data_deudas.Recordset("fecha_pago")) = False Then
               data_deudas.Recordset.Edit
               data_deudas.Recordset("fecha_pago") = Null
               data_deudas.Recordset.Update
            End If
         End If
         If IsNull(frmabm.data_clientes.Recordset("cl_nrocobr")) = False Then
            If frmabm.data_clientes.Recordset("cl_nrocobr") = 615 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 616 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 635 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 602 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 113 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 653 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 672 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 1 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 10 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 201 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 512 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 636 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 685 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 208 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 209 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 8 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 0 Then
            Else
               data_arq.RecordSource = "Select * from arqueo where matricula =" & labmatri.Caption & " and mes =" & data_lineas.Recordset("mes_paga") & " and ano =" & data_lineas.Recordset("ano_paga")
               data_arq.Refresh
               If data_arq.Recordset.RecordCount > 0 Then
                  If IsNull(data_arq.Recordset("arqueo")) = False Then
                     If data_arq.Recordset("arqueo") = "D" Then
                        data_arq.Recordset.Edit
                        data_arq.Recordset("arqueo") = "P"
                        data_arq.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                        data_arq.Recordset("usuar") = WElusuario
                        data_arq.Recordset.Update
                     End If
                  End If
               End If
            End If
         End If
      Else
         If frmabm.labatra.Caption <> "" Then
            If frmabm.labatra.Caption > 0 Then
               frmabm.labatra.Caption = Val(frmabm.labatra.Caption) - 1
            End If
         Else
            frmabm.labatra.Caption = 0
         End If
         If frmabm.labdeudap.Caption <> "" Then
            If frmabm.labdeudap.Caption > 0 Then
               frmabm.labdeudap.Caption = Val(frmabm.labdeudap.Caption) - data_lineas.Recordset("tot_lin")
            End If
         Else
            frmabm.labdeudap.Caption = 0
         End If
         frmabm.labump.Caption = data_lineas.Recordset("mes_paga")
         frmabm.labuap.Caption = data_lineas.Recordset("ano_paga")
'aquí
         data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " and mes =" & data_lineas.Recordset("mes_paga") & " and ano =" & data_lineas.Recordset("ano_paga") & " and fecha_pago is null"
         data_deudas.Refresh
         If data_deudas.Recordset.RecordCount > 0 Then
            data_deudas.Recordset.Edit
            data_deudas.Recordset("fecha_pago") = Date
            data_deudas.Recordset.Update
         End If
         If IsNull(frmabm.data_clientes.Recordset("cl_nrocobr")) = False Then
            If frmabm.data_clientes.Recordset("cl_nrocobr") = 615 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 616 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 635 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 602 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 113 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 653 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 672 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 1 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 10 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 201 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 512 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 636 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 685 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 208 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 209 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 8 Or _
               frmabm.data_clientes.Recordset("cl_nrocobr") = 0 Then
            Else
               data_arq.RecordSource = "Select * from arqueo where matricula =" & labmatri.Caption & " and mes =" & data_lineas.Recordset("mes_paga") & " and ano =" & data_lineas.Recordset("ano_paga")
               data_arq.Refresh
               If data_arq.Recordset.RecordCount > 0 Then
                  If IsNull(data_arq.Recordset("arqueo")) = False Then
                     If data_arq.Recordset("arqueo") = "D" Then
                     Else
                        data_arq.Recordset.Edit
                        data_arq.Recordset("arqueo") = "D"
                        data_arq.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                        data_arq.Recordset("usuar") = WElusuario
                        data_arq.Recordset.Update
                     End If
                  Else
                     data_arq.Recordset.Edit
                     data_arq.Recordset("arqueo") = "D"
                     data_arq.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                     data_arq.Recordset("usuar") = WElusuario
                     data_arq.Recordset.Update
                  End If
               End If
            End If
         End If
      End If
   End If
   
   data_linmmdd.RecordSource = "select * from linmmdd where factura =" & 223595
   data_linmmdd.Refresh
   
   data_linmmdd.Recordset.AddNew
   data_linmmdd.Recordset("fecha") = data_lineas.Recordset("fecha")
   data_linmmdd.Recordset("reg_cab") = data_lineas.Recordset("reg_cab")
   data_linmmdd.Recordset("factura") = labfac.Caption
   data_linmmdd.Recordset("moneda") = labserie.Caption
   If data_cabeza2.Recordset("cl_tipocli") = 111 Then
      data_linmmdd.Recordset("pendiente") = "F"
   Else
      If data_cabeza2.Recordset("cl_tipocli") = 101 Then
         data_linmmdd.Recordset("pendiente") = "T"
      Else
         If data_cabeza2.Recordset("cl_tipocli") = 112 Then
            data_linmmdd.Recordset("pendiente") = "N" 'NC de E-FACT
         Else
            If data_cabeza2.Recordset("cl_tipocli") = 102 Then
               data_linmmdd.Recordset("pendiente") = "C" 'NC de E-TCK
            Else
               If data_cabeza2.Recordset("cl_tipocli") = 113 Then
                  data_linmmdd.Recordset("pendiente") = "A" 'ND de E-FACT
               Else
                  If data_cabeza2.Recordset("cl_tipocli") = 103 Then
                     data_linmmdd.Recordset("pendiente") = "B" 'ND de E-TCK
                  Else
                     If Label7.Caption = "REG." Then
                        data_linmmdd.Recordset("pendiente") = "X"
                     Else
                        If Label7.Caption = "DEV.RECIBO" Then
                           data_linmmdd.Recordset("pendiente") = "R"
                        Else
                           data_linmmdd.Recordset("pendiente") = "Z"
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   data_linmmdd.Recordset("servicio") = data_lineas.Recordset("servicio")
   data_linmmdd.Recordset("tipo") = labfpago.Caption
   data_linmmdd.Recordset("realizada") = data_lineas.Recordset("realizada")
   data_linmmdd.Recordset("cod_cli") = data_lineas.Recordset("cod_cli")
   If frmabm.t_rs.Text <> "" Then
      If Check1.Value = 1 Then
         data_linmmdd.Recordset("nom_cli") = Mid(frmabm.t_rs.Text, 1, 30)
      Else
         data_linmmdd.Recordset("nom_cli") = Mid(data_lineas.Recordset("nom_cli"), 1, 30)
      End If
   Else
      data_linmmdd.Recordset("nom_cli") = Mid(data_lineas.Recordset("nom_cli"), 1, 30)
   End If
   data_linmmdd.Recordset("ced_socio") = data_lineas.Recordset("ced_socio")
   data_linmmdd.Recordset("tcambio") = data_lineas.Recordset("tcambio")
   data_linmmdd.Recordset("fact") = data_lineas.Recordset("fact")
   If data_lineas.Recordset("cod_prod") = 993 Or _
      data_lineas.Recordset("cod_prod") = 994 Then
      Xquees = 1
   Else
      Xquees = 0
   End If
   data_linmmdd.Recordset("cod_prod") = data_lineas.Recordset("cod_prod")
   data_linmmdd.Recordset("nom_prod") = Mid(data_lineas.Recordset("nom_prod"), 1, 45)
   data_linmmdd.Recordset("cantidad") = data_lineas.Recordset("cantidad")
   data_linmmdd.Recordset("operador") = data_lineas.Recordset("operador")
   data_linmmdd.Recordset("hora") = data_lineas.Recordset("hora")
   If txt_rut.Visible = True Then
      If Trim(txt_rut.Text) <> "" Then
         data_linmmdd.Recordset("ruc") = data_lineas.Recordset("ruc")
      End If
   End If
   data_linmmdd.Recordset("nro_flia") = data_lineas.Recordset("nro_flia")
   data_linmmdd.Recordset("nom_flia") = data_lineas.Recordset("nom_flia")
   data_linmmdd.Recordset("nro_superv") = data_lineas.Recordset("nro_superv")
   data_linmmdd.Recordset("nom_superv") = data_lineas.Recordset("nom_superv")
   If IsNull(data_lineas.Recordset("nom_superv")) = False Then
      If Trim(data_lineas.Recordset("nom_superv")) = "SI" Then
         data_linmmdd.Recordset("repetir") = "S"
      End If
   End If
   If frmabm.t_queconv.Text = "" Then
      data_linmmdd.Recordset("convenio") = data_lineas.Recordset("convenio")
   Else
      data_linmmdd.Recordset("convenio") = data_lineas.Recordset("convenio")
      data_linmmdd.Recordset("unidad") = frmabm.t_queconv.Text
   End If
   data_linmmdd.Recordset("grupo") = data_lineas.Recordset("grupo")
   data_linmmdd.Recordset("rub_cont") = data_lineas.Recordset("rub_cont")
   data_linmmdd.Recordset("arancel") = data_lineas.Recordset("arancel")
   data_linmmdd.Recordset("usa_timbre") = data_lineas.Recordset("usa_timbre")
   data_linmmdd.Recordset("imp_timbre") = data_lineas.Recordset("imp_timbre")
   data_linmmdd.Recordset("tot_lin") = data_lineas.Recordset("tot_lin")
   data_linmmdd.Recordset("rub_nomb") = data_lineas.Recordset("rub_nomb")
   data_linmmdd.Recordset("nro_med_a") = data_lineas.Recordset("nro_med_a")
   If IsNull(data_lineas.Recordset("nro_med_a")) = False Then
      Xnromedic = data_lineas.Recordset("nro_med_a")
   Else
      Xnromedic = 0
   End If
   data_linmmdd.Recordset("nom_med_a") = data_lineas.Recordset("nom_med_a")
   data_linmmdd.Recordset("nro_med_s") = data_lineas.Recordset("nro_med_s")
   data_linmmdd.Recordset("nom_med_s") = data_lineas.Recordset("nom_med_s")
   data_linmmdd.Recordset("precio_est") = data_lineas.Recordset("precio_est")
   data_linmmdd.Recordset("mes_paga") = data_lineas.Recordset("mes_paga")
   data_linmmdd.Recordset("ano_paga") = data_lineas.Recordset("ano_paga")
   data_linmmdd.Recordset("base") = data_lineas.Recordset("base")
   If IsNull(data_lineas.Recordset("imp_iva")) = False Then
      data_linmmdd.Recordset("imp_iva") = Format(data_lineas.Recordset("imp_iva"), "0.00")
   Else
      data_linmmdd.Recordset("imp_iva") = 0
   End If
   data_linmmdd.Recordset("linea") = data_lineas.Recordset("linea")
   data_linmmdd.Recordset("dias") = data_lineas.Recordset("dias")
   data_linmmdd.Recordset("cod_medic") = data_lineas.Recordset("cod_medic")
   data_linmmdd.Recordset("nom_medic") = data_lineas.Recordset("nom_medic")
   data_linmmdd.Recordset("pre_civa") = data_lineas.Recordset("pre_civa")
   If data_lineas.Recordset("cod_prod") = 997 Or data_lineas.Recordset("cod_prod") = 999 Then
''      If IsNull(data_lineas.Recordset("nrofactref")) = False Then
      If IsNull(data_lineas.Recordset("porce_est")) = False Then
         data_linmmdd.Recordset("porce_est") = data_lineas.Recordset("porce_est")
''         data_linmmdd.Recordset("porce_est") = data_lineas.Recordset("nrofactref")
      End If
   Else
      data_linmmdd.Recordset("porce_est") = 0
   End If
   If IsNull(data_lineas.Recordset("contact_tel")) = False Then
      data_linmmdd.Recordset("contact_tel") = data_lineas.Recordset("contact_tel")
   End If
   If t_codaut.Text = "" Then
      t_codaut.Text = 0
   End If
   data_linmmdd.Recordset("rub_nomb") = t_codaut.Text
'   data_linmmdd.Recordset("libro_rub") = data_lineas.Recordset("libro_rub") 'descrip tipo fact
   If IsNull(data_lineas.Recordset("tipo_mov")) = False Then
      data_linmmdd.Recordset("tipo_mov") = data_lineas.Recordset("tipo_mov")
   End If
   If XMensaFertilab = 9 Then
      data_linmmdd.Recordset("esfertilab") = 1
   End If
   If IsNull(data_lineas.Recordset("nro_pedido")) = False Then
      data_linmmdd.Recordset("nro_pedido") = data_lineas.Recordset("nro_pedido")
   End If
   data_linmmdd.Recordset.Update
   If data_lineas.Recordset("cod_prod") = 60103 Or _
      data_lineas.Recordset("cod_prod") = 60106 Or _
      data_lineas.Recordset("cod_prod") = 60107 Or _
      data_lineas.Recordset("cod_prod") = 60108 Then
      If IsNull(data_lineas.Recordset("nro_pedido")) = False Then
         Actualiza_Pedidos
      End If
   End If
   
   If Xfpago = 2 Then
   Else
      data_caja.Recordset.AddNew
      data_caja.Recordset("fecha") = data_lineas.Recordset("fecha")
      data_caja.Recordset("numero") = data_lineas.Recordset("rub_cont")
      data_caja.Recordset("nombre") = Mid(data_lineas.Recordset("rub_nomb"), 1, 35)
      data_caja.Recordset("moneda") = "$"
      data_caja.Recordset("movimiento") = "INGRESO"
      If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "DEV.RECIBO" Then
         data_caja.Recordset("imp_fact") = data_lineas.Recordset("tot_lin") * -1
      Else
         data_caja.Recordset("imp_fact") = data_lineas.Recordset("tot_lin")
      End If
      data_caja.Recordset("documento") = Val(labfac.Caption)
      data_caja.Recordset("observ") = data_lineas.Recordset("tipo") & " " & Trim(labserie.Caption) & " " & Trim(str(labfac.Caption))
      data_caja.Recordset("saldo") = Xsaldocaj + data_lineas.Recordset("tot_lin")
      data_caja.Recordset("usuario") = data_lineas.Recordset("operador")
      data_caja.Recordset("hora") = data_lineas.Recordset("hora")
      data_caja.Recordset("base") = data_lineas.Recordset("base")
      data_caja.Recordset("cod_serv") = data_lineas.Recordset("cod_prod")
      data_caja.Recordset("nom_serv") = Mid(data_lineas.Recordset("nom_prod"), 1, 50)
      data_caja.Recordset("cod_socio") = data_lineas.Recordset("cod_cli")
      If frmabm.t_rs.Text <> "" Then
         If Check1.Value = 1 Then
            data_caja.Recordset("nom_socio") = Mid(frmabm.t_rs.Text, 1, 30)
         Else
            data_caja.Recordset("nom_socio") = Mid(data_lineas.Recordset("nom_cli"), 1, 30)
         End If
      Else
         data_caja.Recordset("nom_socio") = Mid(data_lineas.Recordset("nom_cli"), 1, 30)
      End If
      data_caja.Recordset("caja_mesp") = data_lineas.Recordset("mes_paga")
      data_caja.Recordset("caja_anop") = data_lineas.Recordset("ano_paga")
      If IsNull(data_lineas.Recordset("imp_iva")) = False Then
         data_caja.Recordset("imp_iva") = Format(data_lineas.Recordset("imp_iva"), "0.00")
      Else
         data_caja.Recordset("imp_iva") = 0
      End If
      If IsNull(data_lineas.Recordset("imp_iva")) = True Then
         data_caja.Recordset("opiva") = 0
      Else
         If data_lineas.Recordset("imp_iva") = 0 Then
            data_caja.Recordset("opiva") = 0
         Else
            data_caja.Recordset("opiva") = 1
         End If
     End If
     data_caja.Recordset.Update
   End If
   
'promoción por afiliaciones
   Dim NroAfilia, XimpAfilia As Long
   NroAfilia = 0
   XimpAfilia = 0
   If data_lineas.Recordset("cod_prod") = 992 Or _
      data_lineas.Recordset("cod_prod") = 984 Or _
      data_lineas.Recordset("cod_prod") = 985 Or _
      data_lineas.Recordset("cod_prod") = 986 Or _
      data_lineas.Recordset("cod_prod") = 987 Or _
      data_lineas.Recordset("cod_prod") = 989 Or _
      data_lineas.Recordset("cod_prod") = 802 Or _
      data_lineas.Recordset("cod_prod") = 803 Or _
      data_lineas.Recordset("cod_prod") = 804 Or _
      data_lineas.Recordset("cod_prod") = 805 Or _
      data_lineas.Recordset("cod_prod") = 806 Then

      data_facafil.RecordSource = "select * from linmmdd_afil where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
      data_facafil.Refresh
      data_facafil.Recordset.AddNew
      data_facafil.Recordset("fecha") = data_lineas.Recordset("fecha")
      data_facafil.Recordset("factura") = Val(labfac.Caption)
      data_facafil.Recordset("codfunc") = data_lineas.Recordset("numero")
      data_facafil.Recordset("nombre") = labnomprom.Caption
      data_facafil.Recordset.Update
   End If
   
   
   data_ctr.Recordset.Edit
   data_ctr.Recordset("fecha") = Date
   data_ctr.Recordset.Update

   If Label7.Caption = "DEV.RECIBO" Then
      If data_lineas.Recordset("cod_prod") = 999 Then
         data_arq.RecordSource = "Select * from arqueo where matricula =" & frmabm.data_clientes.Recordset("cl_codigo") & " and nrorec =" & Xelnrodeuda
         data_arq.Refresh
         If data_arq.Recordset.RecordCount > 0 Then
            If IsNull(data_arq.Recordset("arqueo")) = False Then
               If data_arq.Recordset("arqueo") <> "P" Then
                  data_arq.Recordset.Edit
                  data_arq.Recordset("arqueo") = "P"
                  data_arq.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                  data_arq.Recordset("usuar") = WElusuario
                  data_arq.Recordset.Update
               End If
            Else
               data_arq.Recordset.Edit
               data_arq.Recordset("arqueo") = "P"
               data_arq.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
               data_arq.Recordset("usuar") = WElusuario
               data_arq.Recordset.Update
            End If
         End If
      End If
   End If
   If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Then
      If data_lineas.Recordset("cod_prod") = 60107 Or data_lineas.Recordset("cod_prod") = 995 Or _
         data_lineas.Recordset("cod_prod") = 990 Or data_lineas.Recordset("cod_prod") = 60103 Then
      Else
        data_arq.RecordSource = "Select * from arqueo where matricula =" & frmabm.data_clientes.Recordset("cl_codigo") & " and nrorec =" & data_lineas.Recordset("nrofactref")
        data_arq.Refresh
        If data_arq.Recordset.RecordCount > 0 Then
           If data_arq.Recordset("arqueo") <> "D" Then
              data_arq.Recordset.Edit
              data_arq.Recordset("arqueo") = "D"
              data_arq.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
              data_arq.Recordset("usuar") = WElusuario
              data_arq.Recordset.Update
           End If
        End If
        If data_lineas.Recordset("nrofactref") > 0 Then
           data_deudas.RecordSource = "Select * from deudas where documento =" & data_lineas.Recordset("nrofactref") & " and cliente =" & data_lineas.Recordset("cod_cli") & " and fecha_pago is null"
           data_deudas.Refresh
           If data_deudas.Recordset.RecordCount > 0 Then
              data_deudas.Recordset.Edit
              data_deudas.Recordset("fecha_pago") = Date
              data_deudas.Recordset.Update
              If frmabm.labdeudap.Caption = "" Then
                 frmabm.labdeudap.Caption = 0
              Else
                 frmabm.labdeudap.Caption = Val(frmabm.labdeudap.Caption) - data_lineas.Recordset("tot_lin")
              End If
           End If
        Else
           data_deudas.RecordSource = "Select * from deudas where fecha_pago is null and documento =" & data_lineas.Recordset("nrofactref") & " and cliente =" & data_lineas.Recordset("cod_cli") & " and fecha =#" & Format(data_lineas.Recordset("fechafact"), "yyyy/mm/dd") & "#"
           data_deudas.Refresh
           If data_deudas.Recordset.RecordCount > 0 Then
              data_deudas.Recordset.Edit
              data_deudas.Recordset("fecha_pago") = Date
              data_deudas.Recordset.Update
              If frmabm.labdeudap.Caption = "" Then
                 frmabm.labdeudap.Caption = 0
              Else
                 frmabm.labdeudap.Caption = Val(frmabm.labdeudap.Caption) - data_lineas.Recordset("tot_lin")
              End If
           End If
        End If
      
      End If
   End If
   If Xfpago = 2 Then
      If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Then
      Else
         If data_lineas.Recordset("tot_lin") > 0 Then
            data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption
            data_deudas.Refresh
            data_deudas.Recordset.AddNew
            data_deudas.Recordset("cod_cnv") = frmabm.txt_codcnv.Text
            data_deudas.Recordset("nom_cnv") = Mid(frmabm.txt_nomcnv.Text, 1, 20)
            data_deudas.Recordset("cliente") = data_lineas.Recordset("cod_cli")
            data_deudas.Recordset("nombre") = data_lineas.Recordset("nom_cli")
            data_deudas.Recordset("fecha") = Date
            data_deudas.Recordset("tipodoc") = "CRE"
            data_deudas.Recordset("nro_superv") = 30
            data_deudas.Recordset("documento") = Val(labfac.Caption)
            data_deudas.Recordset("tipocta") = labserie.Caption
            data_deudas.Recordset("importe") = data_lineas.Recordset("tot_lin")
            data_deudas.Recordset("moneda") = 1
            data_deudas.Recordset("origen") = data_lineas.Recordset("libro_rub") & " NRO." & data_lineas.Recordset("moneda") & Trim(str(data_lineas.Recordset("factura")))
'data_lineas.Recordset("libro_rub")
            data_deudas.Recordset("saldo_cc") = data_lineas.Recordset("tot_lin")
            data_deudas.Recordset("mes") = 0
            data_deudas.Recordset("ano") = 0
            data_deudas.Recordset("estado_cta") = 1
            data_deudas.Recordset("tiquet") = 0
            data_deudas.Recordset("deudas") = 0
            data_deudas.Recordset("total") = data_lineas.Recordset("tot_lin")
            data_deudas.Recordset("iva") = data_lineas.Recordset("imp_iva")
            data_deudas.Recordset("servi") = 0
            data_deudas.Recordset("nro_vende") = data_lineas.Recordset("linea")
            data_deudas.Recordset.Update
         End If
         If frmabm.labdeudap.Caption = "" Then
            frmabm.labdeudap.Caption = data_lineas.Recordset("tot_lin")
         Else
            frmabm.labdeudap.Caption = Val(frmabm.labdeudap.Caption) + data_lineas.Recordset("tot_lin")
         End If
      End If
   End If
   data_lineas.Recordset.Edit
   data_lineas.Recordset("unidad") = "S"
   data_lineas.Recordset.Update
   data_lineas.Recordset.MoveNext
Loop

If labtot.Caption <> "" Then
   If Val(labtot.Caption) > 0 Then
      Xquetot = Val(labtot.Caption)
   Else
      Xquetot = 0
   End If
Else
   Xquetot = 0
End If

frmabm.btn_fact.Enabled = True

data_lineas.RecordSource = "Select * from lineas"
data_lineas.Refresh
If data_lineas.Recordset.RecordCount > 0 Then
   data_lineas.Recordset.MoveFirst
   Do While Not data_lineas.Recordset.EOF
      data_lineas.Recordset.Edit
      data_lineas.Recordset("cl_nrotarj") = data_cabeza2.Recordset("cl_nrotarj")
      data_lineas.Recordset("cl_referen") = data_cabeza2.Recordset("cl_referen")
      data_lineas.Recordset("cl_tjemi_c") = data_cabeza2.Recordset("cl_tjemi_c")
      data_lineas.Recordset("cl_diacobr") = data_cabeza2.Recordset("cl_diacobr")
      data_lineas.Recordset("cl_telefon") = data_cabeza2.Recordset("cl_telefon")
      data_lineas.Recordset("obsp") = data_cabeza2.Recordset("obsp")
      data_lineas.Recordset("qr") = data_cabeza2.Recordset("qr")
      data_lineas.Recordset("cl_fax") = data_cabeza2.Recordset("cl_fax")
      data_lineas.Recordset("cl_socmnro") = data_cabeza2.Recordset("cl_socmnro")
      data_lineas.Recordset("cl_numero") = data_cabeza2.Recordset("cl_numero")
      data_lineas.Recordset("cl_celular") = data_cabeza2.Recordset("cl_celular")
      If IsNull(data_cabeza2.Recordset("cl_fnac")) = False Then
         data_lineas.Recordset("cl_fnac") = Format(data_cabeza2.Recordset("cl_fnac"), "dd/mm/yyyy")
      End If
      data_lineas.Recordset("usu_baja") = data_cabeza2.Recordset("usu_baja")
      data_lineas.Recordset("info_debit") = data_cabeza2.Recordset("info_debit")
      data_lineas.Recordset("cl_nrocobr") = data_cabeza2.Recordset("cl_nrocobr")
      data_lineas.Recordset("cl_medflia") = data_cabeza2.Recordset("cl_medflia")
      data_lineas.Recordset("hora_baja") = data_cabeza2.Recordset("hora_baja")
      data_lineas.Recordset("cl_nomcobr") = data_cabeza2.Recordset("cl_nomcobr")
      data_lineas.Recordset("cl_nom_sup") = data_cabeza2.Recordset("cl_nom_sup")
      data_lineas.Recordset("saldo_cc") = data_cabeza2.Recordset("saldo_cc")
      data_lineas.Recordset("saldo_doc") = data_cabeza2.Recordset("saldo_doc")
      If IsNull(data_cabeza2.Recordset("cl_fultpag")) = False Then
         data_lineas.Recordset("cl_fultpag") = Format(data_cabeza2.Recordset("cl_fultpag"), "dd/mm/yyyy")
      End If
      data_lineas.Recordset("cl_nombre") = data_cabeza2.Recordset("cl_nombre")
      If data_lineas.Recordset("cod_prod") = 60103 Or _
         data_lineas.Recordset("cod_prod") = 60105 Or _
         data_lineas.Recordset("cod_prod") = 60106 Or _
         data_lineas.Recordset("cod_prod") = 60108 Or _
         data_lineas.Recordset("cod_prod") = 60107 Or _
         data_lineas.Recordset("cod_prod") = 60109 Then
         data_lineas.Recordset("nom_prod") = data_lineas.Recordset("nom_medic")
      End If
      data_lineas.Recordset.Update
      data_lineas.Recordset.MoveNext
   Loop
End If

If data_cabeza2.Recordset.RecordCount > 0 Then
    data_lincab.Recordset.AddNew
    '           data_cabezal.Recordset("id") = 1
    data_lincab.Recordset("cl_tipcli") = "1.0"
    data_lincab.Recordset("cl_tipocli") = data_cabeza2.Recordset("cl_tipocli")
    data_lincab.Recordset("cl_socmnro") = labserie.Caption
    data_lincab.Recordset("cl_numero") = Val(labfac.Caption)
    data_lincab.Recordset("cl_fnac") = data_cabeza2.Recordset("cl_fnac")
    data_lincab.Recordset("fecha_reac") = Format(mf.Text, "dd/mm/yyyy")
    data_lincab.Recordset("cl_tj_venc") = Format(mf.Text, "dd/mm/yyyy")
    data_lincab.Recordset("cl_nrovend") = data_cabeza2.Recordset("cl_nrovend")
    data_lincab.Recordset("cl_forpago") = data_cabeza2.Recordset("cl_forpago")
    data_lincab.Recordset("cl_celular") = data_cabeza2.Recordset("cl_celular") 'descripcion f.pago
    data_lincab.Recordset("fecha_modi") = Format(mf.Text, "dd/mm/yyyy")
    data_lincab.Recordset("cl_diacobr") = Trim(str(data_param.Recordset("ruc")))
    data_lincab.Recordset("cl_nrotarj") = data_param.Recordset("nombre")
    data_lincab.Recordset("cl_tjemi_n") = data_param.Recordset("nombre")
    data_lincab.Recordset("cl_tjemi_c") = data_param.Recordset("codsuc")
    data_lincab.Recordset("cl_referen") = data_param.Recordset("domic")
    data_lincab.Recordset("tit_tarj") = data_param.Recordset("ciudad")
    data_lincab.Recordset("cl_nomconv") = data_param.Recordset("dpto")
    'receptor
    data_lincab.Recordset("cl_nro_sup") = data_cabeza2.Recordset("cl_nro_sup")
    data_lincab.Recordset("hora_baja") = "UY"
    data_lincab.Recordset("cl_nom_sup") = data_cabeza2.Recordset("cl_nom_sup")
    data_lincab.Recordset("info_debit") = data_cabeza2.Recordset("info_debit")
    data_lincab.Recordset("cl_direcci") = data_cabeza2.Recordset("cl_direcci")
    data_lincab.Recordset("cl_zona") = data_cabeza2.Recordset("cl_zona")
    data_lincab.Recordset("cl_localid") = "URUGUAY" 'opcional
    data_lincab.Recordset("cl_codigo") = data_cabeza2.Recordset("cl_codigo")
    data_lincab.Recordset("usu_baja") = data_cabeza2.Recordset("usu_baja") 'moneda
    data_lincab.Recordset("saldo_chc2") = data_cabeza2.Recordset("saldo_chc2") 'valor dolar
    data_lincab.Recordset("saldo_cc") = data_cabeza2.Recordset("saldo_cc")  'iva minimo
    data_lincab.Recordset("saldo_cc2") = data_cabeza2.Recordset("saldo_cc2") 'iva básico
    data_lincab.Recordset("cl_atrasoa") = data_cabeza2.Recordset("cl_atrasoa") 'subtot iva 22
    data_lincab.Recordset("cl_cedula") = data_cabeza2.Recordset("cl_cedula") 'subtot iva cero
    data_lincab.Recordset("saldo_doc2") = data_cabeza2.Recordset("saldo_doc2")
    data_lincab.Recordset("cl_atrasop") = data_cabeza2.Recordset("cl_atrasop")
    data_lincab.Recordset("cl_decuota") = data_cabeza2.Recordset("cl_decuota")
    data_lincab.Recordset("saldo_doc") = data_cabeza2.Recordset("saldo_doc")
    data_lincab.Recordset("cl_grupo") = data_cabeza2.Recordset("cl_grupo")
    data_lincab.Recordset("saldo_chc") = data_cabeza2.Recordset("saldo_chc")
    data_lincab.Recordset("cl_telefon") = data_cabeza2.Recordset("cl_telefon")
    data_lincab.Recordset("cl_nombre") = data_cabeza2.Recordset("cl_nombre")
    If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "ND E-FACTURA" Or Label7.Caption = "ND E-TICKET" Then
       data_lincab.Recordset("cl_cuopaga") = data_cabeza2.Recordset("cl_cuopaga")
       data_lincab.Recordset("codmotbaja") = data_cabeza2.Recordset("codmotbaja")
       data_lincab.Recordset("ultanopmut") = data_cabeza2.Recordset("ultanopmut")
       data_lincab.Recordset("cl_fultvta") = data_cabeza2.Recordset("cl_fultvta")
       data_lincab.Recordset("cl_entre") = data_cabeza2.Recordset("cl_entre")
    End If
    If Label7.Caption = "DEV.RECIBO" Then
       data_lincab.Recordset("codmotbaja") = labseriecance.Caption
       data_lincab.Recordset("ultanopmut") = Val(labfaccance.Caption)
       data_lincab.Recordset("cl_fultvta") = CDate(labfeccance.Caption)
       If labdevol.Caption <> "" Then
          data_lincab.Recordset("cl_entre") = labdevol.Caption
       End If
    Else
       data_lincab.Recordset("cl_fultpag") = data_cabeza2.Recordset("cl_fultpag")
       data_lincab.Recordset("cl_ultmesp") = data_cabeza2.Recordset("cl_ultmesp")
       data_lincab.Recordset("cl_nomvend") = data_cabeza2.Recordset("cl_nomvend")
       data_lincab.Recordset("cl_fax") = data_cabeza2.Recordset("cl_fax")
    End If
    data_lincab.Recordset.Update

End If

If data_lineas.Recordset.RecordCount > 0 Then
   data_lineas.Recordset.MoveFirst
   If data_lineas.Recordset("cod_prod") = 60103 Or _
      data_lineas.Recordset("cod_prod") = 60105 Or _
      data_lineas.Recordset("cod_prod") = 60106 Or _
      data_lineas.Recordset("cod_prod") = 60108 Or _
      data_lineas.Recordset("cod_prod") = 60107 Or _
      data_lineas.Recordset("cod_prod") = 60109 Then
      data_deudas.RecordSource = "select * from pedidos_medic where matricula =" & data_lineas.Recordset("cod_cli") & " and fec_ent =#" & Format(mf.Text, "yyyy/mm/dd") & "# and estado ='" & "Entregado" & "' and tot_pesos =" & Val(labtot.Caption)
      data_deudas.Refresh
      If data_deudas.Recordset.RecordCount > 0 Then
         data_deudas.Recordset.MoveFirst
         If IsNull(data_deudas.Recordset("nro_factura")) = True Then
            data_deudas.Recordset.Edit
            data_deudas.Recordset("nro_factura") = Val(labfac.Caption)
            data_deudas.Recordset("fecha_fac") = data_lineas.Recordset("fecha")
            data_deudas.Recordset.Update
         End If
      End If
   End If
End If
If data_lineas.Recordset.RecordCount > 0 Then
   If data_lineas.Recordset("cod_prod") = 30081 Or data_lineas.Recordset("cod_prod") = 30084 Or data_lineas.Recordset("cod_prod") = 30085 Then
      data_deudas.RecordSource = "select * from sol_hisopos where cedula=" & data_lineas.Recordset("ced_socio") & " and fecha_fact is null and fecha <=#" & Format(Date, "yyyy/mm/dd") & "#"
      data_deudas.Refresh
      If data_deudas.Recordset.RecordCount > 0 Then
         data_deudas.Recordset.Edit
         data_deudas.Recordset("fecha_fact") = data_lineas.Recordset("fecha")
         data_deudas.Recordset("base_fact") = data_lineas.Recordset("base")
         data_deudas.Recordset("nom_prod") = data_lineas.Recordset("nom_prod")
         data_deudas.Recordset.Update
      End If
   End If
End If

If data_lineas.Recordset.RecordCount > 0 Then
   If data_lineas.Recordset("cod_prod") = 60103 Or _
      data_lineas.Recordset("cod_prod") = 60105 Or _
      data_lineas.Recordset("cod_prod") = 60106 Or _
      data_lineas.Recordset("cod_prod") = 60108 Or _
      data_lineas.Recordset("cod_prod") = 60107 Or _
      data_lineas.Recordset("cod_prod") = 60109 Then
      data_eror.DatabaseName = App.path & "\selec.mdb"
      data_eror.RecordSource = "selec"
      data_eror.Refresh
      If data_eror.Recordset.RecordCount > 0 Then
         data_eror.Recordset.MoveFirst
         Do While Not data_eror.Recordset.EOF
            If IsNull(data_eror.Recordset("idsel")) = False Then
               If IsNull(data_eror.Recordset("mat")) = False Then
                  data_deudas.RecordSource = "select * from hc_prescrip where hc_mat=" & data_eror.Recordset("mat") & " and id =" & data_eror.Recordset("idsel") & " and hc_fecentrega is null"
                  data_deudas.Refresh
                  If data_deudas.Recordset.RecordCount > 0 Then
                     data_deudas.Recordset.Edit
                     data_deudas.Recordset("hc_fecentrega") = Date
                     data_deudas.Recordset("hc_baseent") = data_lineas.Recordset("base")
                     data_deudas.Recordset("hc_usuarioent") = WElusuario
                     data_deudas.Recordset.Update
                  End If
               End If
            End If
            data_eror.Recordset.MoveNext
         Loop
         data_eror.Recordset.MoveFirst
         Do While Not data_eror.Recordset.EOF
            data_eror.Recordset.Delete
            data_eror.Recordset.MoveNext
         Loop
      End If
      data_eror.DatabaseName = App.path & "\erores.mdb"
      data_eror.RecordSource = "erores"
      data_eror.Refresh
   End If
End If

If Label7.Caption = "NC E-TICKET" Or Label7.Caption = "NC E-FACTURA" Or _
   Label7.Caption = "E-TICKET" Or Label7.Caption = "E-FACTURA" Then
    data_imagen.Recordset.AddNew
    data_imagen.Recordset("fecha") = Date
    data_imagen.Recordset("nrofact") = labfac.Caption
    data_imagen.Recordset("serie") = labserie.Caption
    Picture1.Picture = LoadPicture(App.path & "\qr.bmp")
    data_imagen.Recordset.Update
    data_imagen.Refresh
End If

data_imagen.RecordSource = "Select * from qr where nrofact =" & labfac.Caption & " and serie ='" & labserie.Caption & "'"
data_imagen.Refresh
If data_imagen.Recordset.RecordCount > 0 Then
   data_lineas.RecordSource = "Select * from lineas"
   data_lineas.Refresh
   If data_lineas.Recordset.RecordCount > 0 Then
      data_lineas.Recordset.MoveFirst
      Do While Not data_lineas.Recordset.EOF
         data_lineas.Recordset.Edit
         data_lineas.Recordset("qr") = data_imagen.Recordset("qr")
         data_lineas.Recordset.Update
         If data_lineas.Recordset("cod_prod") = 60107 Or data_lineas.Recordset("cod_prod") = 60103 Then
            Xesmedevang = 9
         End If
         data_lineas.Recordset.MoveNext
      Loop
   End If
End If

data_lineas.RecordSource = "Select * from lineas"
data_lineas.Refresh
If data_lineas.Recordset.RecordCount > 0 Then
   data_lineas.Recordset.MoveFirst
   If XMensaFertilab = 9 Then
      labcorreo.Caption = "****PACIENTE CONCURRE A REALIZAR EXAMENES EN FERTILAB****" & vbCrLf
      labcorreo.Caption = vbCrLf & labcorreo.Caption & "Nombre del Paciente: " & labnomb.Caption & vbCrLf
      labcorreo.Caption = vbCrLf & labcorreo.Caption & "CEDULA: " & Trim(str(data_lineas.Recordset("ced_socio"))) & "-" & Trim(str(data_lineas.Recordset("fact"))) & " ---NACIMIENTO: " & Format(frmabm.txt_nac.Text, "dd/mm/yyyy") & vbCrLf
      labcorreo.Caption = vbCrLf & labcorreo.Caption & "CORREO ELECTRONICO: " & XcorreoFertilab & vbCrLf
      labcorreo.Caption = vbCrLf & labcorreo.Caption & "EXAMENES A REALIZAR:" & vbCrLf
      Do While Not data_lineas.Recordset.EOF
         If data_lineas.Recordset("cod_prod") = 30090 Or data_lineas.Recordset("cod_prod") = 80018 Or _
            data_lineas.Recordset("cod_prod") = 995 Or data_lineas.Recordset("nro_flia") = 13 Then
         Else
            labcorreo.Caption = labcorreo.Caption & vbCrLf & "- " & data_lineas.Recordset("nom_prod")
         End If
         data_lineas.Recordset.MoveNext
      Loop
      data_lineas.Recordset.MoveFirst
      EnviarCorreoFert
   End If
End If
If Xquetot > 0 And data_lineas.Recordset("base") <> 20 Then
   Dim Sionoimp As String
'   Sionoimp = MsgBox("Desea Imprimir el documento?", vbExclamation + vbYesNo)
'   If Sionoimp = vbYes Then
      If Label7.Caption = "RECIBO" Then
         cr1.ReportFileName = App.path & "\infticksapp4.rpt"
         If data_lineas.Recordset("cod_prod") = 999 Or data_lineas.Recordset("cod_prod") = 997 Or _
            data_lineas.Recordset("cod_prod") = 993 Or data_lineas.Recordset("cod_prod") = 994 Then
         Else
            cr1.CopiesToPrinter = 2
         End If
      Else
         cr1.ReportFileName = App.path & "\infticksapp3.rpt"
         If Xesmedevang = 9 Then
            cr1.CopiesToPrinter = 2
         End If
      End If
      cr1.Action = 1
   'End If
Else
   
End If

frm_factura.Enabled = True
btn_fin.Enabled = True
'b_borr.Enabled = True
b_cance.Enabled = True
btn_graba.Enabled = True
If labtim.Enabled = True Then
   labtim.Caption = ""
End If
If labmed.Enabled = True Then
   labmed.Caption = ""
End If
If txt_precio.Enabled = True Then
   txt_precio.Text = 0
End If
If cbotim.Enabled = True Then
   cbotim.ListIndex = 0
End If
If dbcbomed.Enabled = True Then
   dbcbomed.Text = ""
End If
If DBCombo1.Enabled = True Then
   DBCombo1.Text = ""
End If
If txt_mes.Enabled = True Then
   txt_mes.Text = ""
End If
If txt_ano.Enabled = True Then
   txt_ano.Text = ""
End If
XQuefac = 0

Unload Me

Exit Sub

Xelerrfactura:
              
              If Err.Number = 3155 Then
                 MsgBox "Error al terminar la factura: " & Err.Number & " " & Err.Description, vbInformation
                 MsgBox "Si vuelve a mostrar el mismo error, pruebe cancelar la factura y volver a facturar", vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = Err.Number
                 data_errfact.Recordset("desc") = Trim(str(Xenquelugar)) & Mid(Err.Description, 1, 125)
                 data_errfact.Recordset.Update
                 Unload Me
              Else
                 If Err.Number = 20545 Then
                    MsgBox "No se pudo imprimir la factura, pruebe reimprimir.", vbInformation
                    data_errfact.Recordset.AddNew
                    data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                    data_errfact.Recordset("fecha") = Date
                    data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                    data_errfact.Recordset("nroerr") = Err.Number
                    data_errfact.Recordset("desc") = Trim(str(Xenquelugar)) & Mid(Err.Description, 1, 125)
                    data_errfact.Recordset.Update
                    Unload Me
                 Else
                    MsgBox "Error al terminar la factura:" & Err.Number & " " & Err.Description, vbInformation
                    MsgBox "Si vuelve a mostrar el mismo error, pruebe cancelar la factura y volver a facturar", vbInformation
                    data_errfact.Recordset.AddNew
                    data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                    data_errfact.Recordset("fecha") = Date
                    data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                    data_errfact.Recordset("nroerr") = Err.Number
                    data_errfact.Recordset("desc") = Trim(str(Xenquelugar)) & Mid(Err.Description, 1, 125)
                    data_errfact.Recordset.Update
                    Unload Me
                 End If
              End If

End Sub

Private Sub Command2_Click()
Dim Xdig, Xrut, Xtot, Xtot2, Xfactor, i As Integer
Dim Xelrut As String
On Error GoTo Alrut

Xelrut = ""
Xelrut = txt_rut.Text
Xtot2 = 0
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
      Else
         MsgBox "El RUT ingresado NO ES CORRECTO! Verifique antes de GRABAR!", vbCritical
         If txt_rut.Visible = True Then
            txt_rut.SetFocus
         End If
         b_cance_Click
      End If
   Else
      MsgBox "El RUT ingresado debe contener solo números"
      b_cance_Click
   End If
Else
   MsgBox "La cantidad de dígitos del RUT no es correcta"
   b_cance_Click
End If

Exit Sub

Alrut:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al RUT"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al RUT"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub Command3_Click()

Dim strIdTransac As String

Set objPosCfe = New PosCfe
    
Dim objresultado As Resultado

If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
   Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
Else
    If frm_menu.data_parse.Recordset("base") = 1 Then
       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-301", vbNullString)
    Else
       If frm_menu.data_parse.Recordset("base") = 2 Then
          Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-002", vbNullString)
       Else
          If frm_menu.data_parse.Recordset("base") = 3 Then
             Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-001", vbNullString)
          Else
             If frm_menu.data_parse.Recordset("base") = 4 Then
                Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-004", vbNullString)
             Else
                If frm_menu.data_parse.Recordset("base") = 6 Then
                   Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-006", vbNullString)
                Else
                   If frm_menu.data_parse.Recordset("base") = 8 Then
                      Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-008", vbNullString)
                   Else
                      If frm_menu.data_parse.Recordset("base") = 10 Then
                         Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-010", vbNullString)
                      Else
                         If frm_menu.data_parse.Recordset("base") = 13 Then
                            Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-013", vbNullString)
                         Else
                            If frm_menu.data_parse.Recordset("base") = 16 Then
                               Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-116", vbNullString)
                            Else
                               If frm_menu.data_parse.Recordset("base") = 91 Then
                                  Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-216", vbNullString)
                               Else
                                  If frm_menu.data_parse.Recordset("base") = 17 Then
                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-017", vbNullString)
                                  Else
                                     If frm_menu.data_parse.Recordset("base") = 18 Then
                                        Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-018", vbNullString)
                                     Else
                                        If frm_menu.data_parse.Recordset("base") = 96 Then
                                           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-218", vbNullString)
                                        Else
                                           If frm_menu.data_parse.Recordset("base") = 38 Then
                                              Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
                                           Else
                                              If frm_menu.data_parse.Recordset("base") = 11 Then
                                                 Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-111", vbNullString)
                                              Else
                                                 If frm_menu.data_parse.Recordset("base") = 33 Or frm_menu.data_parse.Recordset("base") = 34 Then  ' B3 adm
                                                    If frm_menu.data_parse.Recordset("base") = 33 Then
                                                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-333", vbNullString)
                                                    Else
                                                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-334", vbNullString)
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
                   End If
                End If
             End If
          End If
       End If
    End If
End If
'Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
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
    
'MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'      "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
    'Enviando
If Not EstaInicializado() Then Exit Sub

Dim objCfe As CFE
Set objCfe = New CFE

Dim objCf As ClassFactory

Set objCf = New ClassFactory

Set objCfe.EFact = New EFact
    
With objCfe.EFact.Encabezado.IdDoc
     .TipoCFE = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
     .FchEmis.SetDate Year(data_cabeza2.Recordset("cl_fnac")), Month(data_cabeza2.Recordset("cl_fnac")), Day(data_cabeza2.Recordset("cl_fnac"))
     .IsValidMntBruto = True
     .MntBruto = IdDoc_Tck_MntBruto_1
     If data_cabeza2.Recordset("cl_forpago") = 1 Then
        .FmaPago = IdDoc_Fact_FmaPago_1
     Else
        .FmaPago = IdDoc_Fact_FmaPago_2
     End If
End With
With objCfe.EFact.Encabezado.Emisor
     .RUCEmisor = data_param.Recordset("ruc")
     .RznSoc = data_param.Recordset("nomc")
     .CdgDGISucur.FromString Trim(str(data_param.Recordset("codsuc")))
     .DomFiscal = data_param.Recordset("domic")
     .Ciudad = data_param.Recordset("ciudad")
     .Departamento = data_param.Recordset("dpto")
End With
With objCfe.EFact.Encabezado.Receptor
     If data_cabeza2.Recordset("cl_nro_sup") = 2 Then
        .TipoDocRecep = DocType_2
     Else
        If data_cabeza2.Recordset("cl_nro_sup") = 3 Then
           .TipoDocRecep = DocType_3
        Else
           If data_cabeza2.Recordset("cl_nro_sup") = 5 Then
              .TipoDocRecep = DocType_5
           Else
              If data_cabeza2.Recordset("cl_nro_sup") = 6 Then
                 .TipoDocRecep = DocType_6
              Else
                 .TipoDocRecep = DocType_4
              End If
           End If
        End If
     End If
     .CodPaisRecep = CodPaisType_UY
     .DocRecep = data_cabeza2.Recordset("cl_nom_sup")
     .RznSocRecep = data_cabeza2.Recordset("info_debit")
     .DirRecep = data_cabeza2.Recordset("cl_direcci")
     .CiudadRecep = data_cabeza2.Recordset("cl_zona")
End With
With objCfe.EFact.Encabezado.Totales
     .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
     .IsValidTpoCambio = True
'     If data_cabeza2.Recordset("usu_baja") = "USD" Then
'        .TpoCambio.FromString Format(data_cabeza2.Recordset("saldo_chc2"), "0.00")
'     Else
     .TpoCambio.FromString "1"
'     End If
     .IsValidMntNetoIvaTasaMin = True
     .IsValidMntNetoIVATasaBasica = True
     .IsValidMntIVATasaMin = True
     .IsValidMntIVATasaBasica = True
     .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
     .MntNetoIVATasaBasica.FromString Format(data_cabeza2.Recordset("cl_atrasoa"), "0.00")
     If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
        If data_cabeza2.Recordset("cl_cedula") > 0 Then
           .IsValidMntNoGrv = True
           .MntNoGrv.FromString Format(data_cabeza2.Recordset("cl_cedula"), "0.00")
        End If
     End If
     .IVATasaMin = TasaIVAType_10FullStop000
     .IVATasaBasica = TasaIVAType_22FullStop000
     .MntIVATasaMin.FromString Format(data_cabeza2.Recordset("saldo_cc"), "0.00")
     .MntIVATasaBasica.FromString Format(data_cabeza2.Recordset("saldo_cc2"), "0.00")
     .CantLinDet.FromString data_cabeza2.Recordset("cl_grupo")
     .MntTotal.FromString Format(data_cabeza2.Recordset("saldo_doc"), "0.00")
     .MntPagar.FromString Format(data_cabeza2.Recordset("saldo_chc"), "0.00")
End With

Do While Not data_lineas.Recordset.EOF
   With objCfe.EFact.Detalle.Item.AddNew
       .NroLinDet.FromString Trim(str(data_lineas.Recordset("linea")))
       .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_lineas.Recordset("tipo_mov"))))
       .NomItem = data_lineas.Recordset("nom_prod")
       .cantidad.FromString Trim(str(data_lineas.Recordset("cantidad")))
       .UniMed = "N/A"
       .PrecioUnitario.FromString Format(data_lineas.Recordset("arancel"), "0.00")
       .MontoItem.FromString Format(data_lineas.Recordset("tot_lin"), "0.00")
   End With
   data_lineas.Recordset.MoveNext
Loop
Dim s As String
s = objCfe.ToXml(True, XmlFormatting_Indented)
'        Open App.Path & "\sapp.xml" For Output As #1
'        Print #1, s

'        Text1.Text = s
Dim strGuid As String
strGuid = objPosCfe.CrearGuid()
Dim objResultadoCfe As ResultadoCfe
'Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))

Set objUltimaSerieNumero = Nothing
DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
If Not objUltimaSerieNumero Is Nothing Then _
   ' cmdFirmarNc.Enabled = True
'           MsgBox "firmar NC"
End If
If frm_menu.data_parse.Recordset("base") = 38 Then
   MsgBox "Terminado"
   Unload Me
Else
   Command1_Click
End If


End Sub

Private Sub Command4_Click()
On Error GoTo Alcomman4

data_lineas.Recordset.AddNew
If Label7.Caption = "NC E-FACTURA" Then
   data_lineas.Recordset("tipodocref") = 111
Else
   data_lineas.Recordset("tipodocref") = 101
End If
data_lineas.Recordset("serieref") = labseriecance.Caption
If Len(labfaccance.Caption) > 7 Then
   data_lineas.Recordset("nrofactref") = Val(Mid(labfaccance.Caption, 1, 7))
Else
   data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
End If
data_lineas.Recordset("fechafact") = CDate(labfeccance.Caption)
data_lineas.Recordset("motivoref") = Mid(labmotivo.Caption, 1, 90)
data_lineas.Recordset("linearef") = 2

data_lineas.Recordset("reg_cab") = 0
data_lineas.Recordset("factura") = 0
data_lineas.Recordset("tipo_mov") = 1
data_lineas.Recordset("realizada") = Format(mf.Text, "dd/mm/yyyy")
data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
data_lineas.Recordset("cod_cli") = labmatri.Caption
data_lineas.Recordset("nom_cli") = labnomb.Caption
data_lineas.Recordset("cod_prod") = 995
data_lineas.Recordset("nom_prod") = "TIMBRE PROFESIONAL"
If txt_rut.Visible = True Then
   If Trim(txt_rut.Text) <> "" Then
      data_lineas.Recordset("ruc") = txt_rut.Text
   End If
End If
data_lineas.Recordset("cantidad") = 1
data_lineas.Recordset("moneda") = "SR"
data_lineas.Recordset("operador") = WElusuario
data_lineas.Recordset("hora") = Format(Time, "HH:mm")
data_lineas.Recordset("nro_flia") = 8
data_lineas.Recordset("nom_flia") = "OTROS SERVICIOS"
data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
data_lineas.Recordset("rub_cont") = 213076
data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
data_lineas.Recordset("arancel") = Val(labtimemi.Caption)
data_lineas.Recordset("tot_lin") = Val(labtimemi.Caption)
data_lineas.Recordset("precio_est") = Val(labtimemi.Caption)
data_lineas.Recordset("porce_est") = 0
data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
data_lineas.Recordset("tipo") = labfpago.Caption
data_lineas.Recordset("linea") = 2
data_lineas.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
data_lineas.Recordset("in_unid") = "INT1"
data_lineas.Recordset.Update
data_lineas.Refresh
If labtot.Caption <> "" Then
   labtot.Caption = Val(labtot.Caption) + Val(labtimemi.Caption)
Else
   labtot.Caption = Val(labtimemi.Caption)
End If

Exit Sub

Alcomman4:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al Comman4"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al Comman4"
            data_errfact.Recordset.Update
            Unload Me
         End If


End Sub

Private Sub Command5_Click()
On Error GoTo Alcomman5

data_lineas.Recordset.AddNew
If Label7.Caption = "NC E-FACTURA" Then
   data_lineas.Recordset("tipodocref") = 111
Else
   data_lineas.Recordset("tipodocref") = 101
End If
data_lineas.Recordset("serieref") = labseriecance.Caption
If Len(labfaccance.Caption) > 7 Then
   data_lineas.Recordset("nrofactref") = Val(Mid(labfaccance.Caption, 1, 7))
Else
   data_lineas.Recordset("nrofactref") = Val(labfaccance.Caption)
End If
data_lineas.Recordset("fechafact") = CDate(labfeccance.Caption)
data_lineas.Recordset("motivoref") = Mid(labmotivo.Caption, 1, 90)
If labtimemi.Caption <> "" Then
   If Format(labtimemi.Caption, "Standard") > 0 Then
      data_lineas.Recordset("linearef") = 3
   Else
      data_lineas.Recordset("linearef") = 2
   End If
Else
   data_lineas.Recordset("linearef") = 2
End If
data_lineas.Recordset("reg_cab") = 0
data_lineas.Recordset("factura") = 0
data_lineas.Recordset("tipo_mov") = 1
data_lineas.Recordset("realizada") = Format(mf.Text, "dd/mm/yyyy")
data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
data_lineas.Recordset("cod_cli") = labmatri.Caption
data_lineas.Recordset("nom_cli") = labnomb.Caption
data_lineas.Recordset("cod_prod") = 882
data_lineas.Recordset("nom_prod") = "DEUDAS POR SERVICIOS"
If txt_rut.Visible = True Then
   If Trim(txt_rut.Text) <> "" Then
      data_lineas.Recordset("ruc") = txt_rut.Text
   End If
End If
data_lineas.Recordset("cantidad") = 1
data_lineas.Recordset("moneda") = "SR"
data_lineas.Recordset("operador") = WElusuario
data_lineas.Recordset("hora") = Format(Time, "HH:mm")
data_lineas.Recordset("nro_flia") = 8
data_lineas.Recordset("nom_flia") = "OTROS SERVICIOS"
data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
data_lineas.Recordset("rub_cont") = 213041
data_lineas.Recordset("rub_nomb") = "PROVISORIOS"
data_lineas.Recordset("arancel") = Format(labdeudaemi.Caption, "Standard")
data_lineas.Recordset("tot_lin") = Format(labdeudaemi.Caption, "Standard")
data_lineas.Recordset("precio_est") = Format(labdeudaemi.Caption, "Standard")
data_lineas.Recordset("porce_est") = 0
data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
data_lineas.Recordset("tipo") = labfpago.Caption
data_lineas.Recordset("linea") = 2
data_lineas.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
data_lineas.Recordset("in_unid") = "INT1"
data_lineas.Recordset.Update
data_lineas.Refresh
Dim Xelivaemi As Double
Xelivaemi = Format(labdeudaemi.Caption, "Standard") / 1.1 * 0.1

If labtot.Caption <> "" Then
   labtot.Caption = Format(labtot.Caption, "Standard") + Format(labdeudaemi.Caption, "Standard")
   Label8.Caption = Format(Label8.Caption, "Standard") + Xelivaemi
   Label8.Caption = Format(Label8.Caption, "Standard")
Else
   labtot.Caption = Format(labdeudaemi.Caption, "Standard")
   Label8.Caption = Format(Label8.Caption, "Standard") + Xelivaemi
   Label8.Caption = Format(Label8.Caption, "Standard")
End If

Exit Sub

Alcomman5:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al comman5"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al comman5"
            data_errfact.Recordset.Update
            Unload Me
         End If


End Sub

Private Sub dbcbomed_KeyPress(KeyAscii As Integer)
On Error GoTo Alcbomedpres

If KeyAscii = 13 Then
   If dbcbomed.Text <> "" Then
        dbcbomed.ListField = "med_nombre"
        dbcbomed.BoundColumn = "med_nombre"
        If IsNumeric(dbcbomed.Text) Then
           data_medicos.Recordset.FindFirst "med_cod =" & dbcbomed.Text
           If Not data_medicos.Recordset.NoMatch Then
              dbcbomed.ListField = ""
              dbcbomed.BoundColumn = ""
              dbcbomed.Text = data_medicos.Recordset("med_nombre")
              dbcbomed.Height = 400
              labmed.Caption = data_medicos.Recordset("med_cod")
'              btn_graba.SetFocus
              If dbcbomedo.Enabled = True Then
                 dbcbomedo.SetFocus
              End If
           Else
              dbcbomed.ListField = "med_nombre"
              dbcbomed.BoundColumn = "med_nombre"
              data_medicos.RecordSource = "select * from medicos where med_cod >=" & dbcbomed.Text
              data_medicos.Refresh
              If data_medicos.Recordset.RecordCount > 0 Then
                 dbcbomed.Height = 1250
              Else
                 MsgBox "No se encontraron registros", vbCritical, "Factura"
                 dbcbomed.Height = 400
'                 labmed.Caption = data_medicos.Recordset("med_cod")
                 If dbcbomed.Enabled = True Then
                    dbcbomed.SetFocus
                 End If
              End If
           End If
        Else
           data_medicos.Recordset.FindFirst "med_nombre ='" & dbcbomed.Text & "'"
           If Not data_medicos.Recordset.NoMatch Then
              dbcbomed.ListField = ""
              dbcbomed.BoundColumn = ""
              dbcbomed.Text = data_medicos.Recordset("med_nombre")
              labmed.Caption = data_medicos.Recordset("med_cod")
              dbcbomed.Height = 400
'              btn_graba.SetFocus
              If dbcbomedo.Enabled = True Then
                 dbcbomedo.SetFocus
              End If
           Else
              dbcbomed.ListField = "med_nombre"
              dbcbomed.BoundColumn = "med_nombre"
              data_medicos.RecordSource = "select * from medicos where med_nombre >='" & dbcbomed.Text & "' order by med_nombre"
              data_medicos.Refresh
              dbcbomed.Height = 1250
           End If
        End If
   Else
'       btn_graba.SetFocus
        If dbcbomedo.Enabled = True Then
           dbcbomedo.SetFocus
        End If
   End If
End If

Exit Sub

Alcbomedpres:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbomedpres"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbomedpres"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub



Private Sub dbcbomedo_KeyPress(KeyAscii As Integer)
On Error GoTo Alcbomeddos

If KeyAscii = 13 Then
   If dbcbomedo.Text <> "" Then
        dbcbomedo.ListField = "med_nombre"
        dbcbomedo.BoundColumn = "med_nombre"
        If IsNumeric(dbcbomedo.Text) Then
           data_medicos.Recordset.FindFirst "med_cod =" & dbcbomedo.Text
           If Not data_medicos.Recordset.NoMatch Then
              dbcbomedo.ListField = ""
              dbcbomedo.BoundColumn = ""
              dbcbomedo.Text = data_medicos.Recordset("med_nombre")
              dbcbomedo.Height = 400
              labmedo.Caption = data_medicos.Recordset("med_cod")
              If btn_graba.Enabled = True Then
                 btn_graba.SetFocus
              End If
           Else
              dbcbomedo.ListField = "med_nombre"
              dbcbomedo.BoundColumn = "med_nombre"
              data_medicos.RecordSource = "select * from medicos where med_cod >=" & dbcbomedo.Text
              data_medicos.Refresh
              If data_medicos.Recordset.RecordCount > 0 Then
                 dbcbomed.Height = 1250
              Else
                 MsgBox "No se encontraron registros", vbCritical, "Factura"
                 dbcbomedo.Height = 400
'                 labmed.Caption = data_medicos.Recordset("med_cod")
                 If dbcbomedo.Enabled = True Then
                    dbcbomedo.SetFocus
                 End If
              End If
           End If
        Else
           data_medicos.Recordset.FindFirst "med_nombre ='" & dbcbomedo.Text & "'"
           If Not data_medicos.Recordset.NoMatch Then
              dbcbomedo.ListField = ""
              dbcbomedo.BoundColumn = ""
              dbcbomedo.Text = data_medicos.Recordset("med_nombre")
              labmedo.Caption = data_medicos.Recordset("med_cod")
              dbcbomedo.Height = 400
              If btn_graba.Enabled = True Then
                 btn_graba.SetFocus
              End If
           Else
              dbcbomedo.ListField = "med_nombre"
              dbcbomedo.BoundColumn = "med_nombre"
              data_medicos.RecordSource = "select * from medicos where med_nombre >='" & dbcbomedo.Text & "' order by med_nombre"
              data_medicos.Refresh
              dbcbomedo.Height = 1250
           End If
        End If
   Else
       If btn_graba.Enabled = True Then
          btn_graba.SetFocus
       End If
   End If
End If

Exit Sub

Alcbomeddos:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbomeddos"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbomeddos"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub dbcboprom_KeyPress(KeyAscii As Integer)
On Error GoTo Alcbopromo

If KeyAscii = 13 Then
   If dbcboprom.Text <> "" Then
        dbcboprom.ListField = "nombre"
        dbcboprom.BoundColumn = "nombre"
        If IsNumeric(dbcboprom.Text) Then
           data_func.Recordset.FindFirst "idfunc =" & dbcboprom.Text
           If Not data_func.Recordset.NoMatch Then
              dbcboprom.ListField = ""
              dbcboprom.BoundColumn = ""
              dbcboprom.Text = data_func.Recordset("nombre")
              dbcboprom.Height = 400
              labcodpro.Caption = data_func.Recordset("idfunc")
              labnomprom.Caption = data_func.Recordset("nombre")
'              btn_graba.SetFocus
              If btn_graba.Enabled = True Then
                 btn_graba.SetFocus
              End If
           Else
              dbcboprom.ListField = "nombre"
              dbcboprom.BoundColumn = "nombre"
              data_func.RecordSource = "select * from vende_func where idfunc >=" & dbcboprom.Text
              data_func.Refresh
              If data_func.Recordset.RecordCount > 0 Then
                 dbcboprom.Height = 1250
              Else
                 MsgBox "No se encontraron registros de promotor", vbCritical, "Factura"
                 dbcboprom.Height = 400
'                 labmed.Caption = data_medicos.Recordset("med_cod")
                 If dbcboprom.Enabled = True Then
                    dbcboprom.SetFocus
                 End If
              End If
           End If
        Else
           data_func.Recordset.FindFirst "nombre ='" & dbcboprom.Text & "'"
           If Not data_func.Recordset.NoMatch Then
              dbcboprom.ListField = ""
              dbcboprom.BoundColumn = ""
              dbcboprom.Text = data_func.Recordset("nombre")
              labcodpro.Caption = data_func.Recordset("idfunc")
              labnomprom.Caption = data_func.Recordset("nombre")
              dbcboprom.Height = 400
'              btn_graba.SetFocus
              If btn_graba.Enabled = True Then
                 btn_graba.SetFocus
              End If
           
           Else
              dbcboprom.ListField = "nombre"
              dbcboprom.BoundColumn = "nombre"
              data_func.RecordSource = "select * from vende_func where nombre >='" & dbcboprom.Text & "' order by nombre"
              data_func.Refresh
              dbcboprom.Height = 1250
           End If
        End If
   Else
'       btn_graba.SetFocus
        If btn_graba.Enabled = True Then
           btn_graba.SetFocus
        End If
   End If
End If

Exit Sub

Alcbopromo:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbopromo"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al cbopromo"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub



Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
Dim Xcuota, Xdescuento As Double
Dim Xquerr As Integer
Xquerr = 0
'terminado
'211372 'cgalicia
On Error GoTo Aldbcbo1
'992 afiliación
If KeyAscii = 13 Then
   If DBCombo1.Text <> "" Then
        DBCombo1.ListField = "DESCRIP"
        DBCombo1.BoundColumn = "DESCRIP"
        If IsNumeric(DBCombo1.Text) Then
           If DBCombo1.Text = 60106 Or _
              DBCombo1.Text = 60108 Or _
              DBCombo1.Text = 993 Or _
              DBCombo1.Text = 994 Or _
              DBCombo1.Text = 997 Or _
              DBCombo1.Text = 996 Or _
              DBCombo1.Text = 60105 Or _
              DBCombo1.Text = 999 Or _
              DBCombo1.Text = 60109 Or _
              DBCombo1.Text = 80011 Or _
              DBCombo1.Text = 80012 Or _
              DBCombo1.Text = 80013 Or _
              DBCombo1.Text = 80014 Or _
              DBCombo1.Text = 80015 Or _
              DBCombo1.Text = 80016 Then
              If XQuefac = 4 Or XQuefac = 21 Then
                 Xquerr = 0
              Else
                 MsgBox "Debe de facturar cómo RECIBO", vbCritical, "Mensaje"
                 Xquerr = 1
                 b_cance_Click
                 Exit Sub
              End If
           Else
              If XQuefac = 4 Then
                 MsgBox "Debe de facturar cómo E-TICKET o E-FACTURA", vbCritical, "Mensaje"
                 Xquerr = 1
                 b_cance_Click
                 Exit Sub
              Else
                 Xquerr = 0
              End If
           End If
           
           If DBCombo1.Text <> "" Then
              data_estudio.Recordset.FindFirst "codest =" & DBCombo1.Text
           End If
           If Not data_estudio.Recordset.NoMatch Then
              Label5.Caption = data_estudio.Recordset("codest")
              DBCombo1.Text = data_estudio.Recordset("descrip")
              '''''xop1 es la var que tiene el valor del id de aran-gpos (lo toma en frmquefact)
''aca nuevo
              data_arancel.RecordSource = "Select * from Aran_servicios where id_gpo =" & Xop1 & " and id_serv =" & data_estudio.Recordset("codest")
              data_arancel.Refresh
              If data_arancel.Recordset.RecordCount > 0 Then
                 If data_arancel.Recordset("prec_serv") > 0 Then
                    txt_precio.Text = data_arancel.Recordset("prec_serv")
                    txt_precio.Text = Format(txt_precio.Text, "Standard")
                 Else
                    If data_arancel.Recordset("por_serv") = 100 Then
                       txt_precio.Text = 0
                       txt_precio.Text = 0
                    Else
                       If data_arancel.Recordset("por_serv") = 0 Then
                          txt_precio.Text = Format(data_estudio.Recordset("cons"), "Standard")
                          txt_precio.Text = Format(txt_precio.Text, "Standard")
                       Else
                          Xdescuento = data_arancel.Recordset("por_serv") * data_estudio.Recordset("cons") / 100
                          txt_precio.Text = data_estudio.Recordset("cons") - Xdescuento
                          txt_precio.Text = Format(txt_precio.Text, "Standard")
                       End If
                    End If
                 End If
              Else
                 txt_precio.Text = Format(data_estudio.Recordset("part"), "Standard")
                 txt_precio.Text = Format(txt_precio.Text, "Standard")
              End If
             If frmabm.data_clientes.Recordset("cl_codconv") = "PART" Then
                txt_precio.Text = Format(data_estudio.Recordset("part"), "Standard")
                txt_precio.Text = Format(txt_precio.Text, "Standard")
             End If
              
              If data_estudio.Recordset("codest") = 993 Or _
                 data_estudio.Recordset("codest") = 994 Then
                 cbotim.Enabled = False
                 txt_mes.Text = ""
                 txt_ano.Text = ""
                 txt_mes.Enabled = True
                 txt_ano.Enabled = True
                 dbcbomed.Enabled = False
              End If
              If data_estudio.Recordset("codest") = 999 Then
                 cbotim.Enabled = False
                 txt_mes.Enabled = True
                 txt_ano.Enabled = True
                 dbcbomed.Enabled = False
                 data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " And fecha_pago is null and mes >" & 0 & " order by ano,mes"
                 data_deudas.Refresh
                 If data_deudas.Recordset.RecordCount > 0 Then
                    frm_veodeuda.Show vbModal
                    txt_mes.Enabled = False
                    txt_ano.Enabled = False
                 Else
                    MsgBox "No figura Deuda, no se puede realizar cobro"
                    Unload Me
                    frmabm.btn_fact.Enabled = True
                    Exit Sub
                 End If
              End If
              If data_estudio.Recordset("codest") = 997 Then
                 cbotim.Enabled = False
                 txt_mes.Enabled = True
                 txt_ano.Enabled = True
                 txt_mes.Text = ""
                 txt_ano.Text = ""
                 dbcbomed.Enabled = False
                 data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " And fecha_pago is null and mes =" & 0
                 data_deudas.Refresh
                 If data_deudas.Recordset.RecordCount > 0 Then
                    frm_veodeuda.Show vbModal
                 Else
                    Xcuota = 0
                    MsgBox "Socio sin FACTURAS A CREDITO PENDIENTES...$: " & Trim(str(Xcuota)), vbInformation, "Mensaje"
                    txt_precio.Text = Xcuota
                    If txt_precio.Enabled = True Then
                       txt_precio.SetFocus
                    End If
                    frmabm.btn_fact.Enabled = True
                    Unload Me
                    Exit Sub
                 End If
                 txt_mes.Text = ""
                 txt_ano.Text = ""
                 txt_mes.Enabled = False
                 txt_ano.Enabled = False
              End If
                            
              If Xquerr <> 1 Then
                 If data_estudio.Recordset("codest") = 60103 Or data_estudio.Recordset("codest") = 60106 Or _
                    data_estudio.Recordset("codest") = 60107 Or data_estudio.Recordset("codest") = 60108 Then
                    btn_fin.SetFocus
                 Else
                    If txt_precio.Enabled = True Then
                       txt_precio.SetFocus
                    End If
                 End If
                DBCombo1.Height = 500
                DBCombo1.ListField = ""
                DBCombo1.BoundColumn = ""
              Else
                DBCombo1.Height = 500
                DBCombo1.ListField = ""
                DBCombo1.BoundColumn = ""
                DBCombo1.Text = ""
                txt_precio.Text = ""
              End If
           Else
              data_estudio.RecordSource = "select * from estudios where codest >=" & DBCombo1.Text
              data_estudio.Refresh
              DBCombo1.Height = 1350
           End If
        Else
           data_estudio.Recordset.FindFirst "descrip ='" & DBCombo1.Text & "'"
           If Not data_estudio.Recordset.NoMatch Then
              DBCombo1.Text = data_estudio.Recordset("descrip")
              Label5.Caption = data_estudio.Recordset("codest")
              
              ''''xop1 es la var que tiene el valor del id de aran-gpos que se toma el valor en frmquefact
              
              data_arancel.RecordSource = "Select * from Aran_servicios where id_gpo =" & Xop1 & " and id_serv =" & data_estudio.Recordset("codest")
              data_arancel.Refresh
              If data_arancel.Recordset.RecordCount > 0 Then
                 If data_arancel.Recordset("prec_serv") > 0 Then
                    txt_precio.Text = data_arancel.Recordset("prec_serv")
                    txt_precio.Text = Format(txt_precio.Text, "Standard")
                 Else
                    If data_arancel.Recordset("por_serv") = 100 Then
                       txt_precio.Text = 0
                       txt_precio.Text = 0
                    Else
                       If data_arancel.Recordset("por_serv") = 0 Then
                          txt_precio.Text = Format(data_estudio.Recordset("cons"), "Standard")
                          txt_precio.Text = Format(txt_precio.Text, "Standard")
                       Else
                          Xdescuento = data_arancel.Recordset("por_serv") * data_estudio.Recordset("cons") / 100
                          txt_precio.Text = data_estudio.Recordset("cons") - Xdescuento
                          txt_precio.Text = Format(txt_precio.Text, "Standard")
                       End If
                    End If
                 End If
              Else
                 txt_precio.Text = Format(data_estudio.Recordset("part"), "Standard")
                 txt_precio.Text = Format(txt_precio.Text, "Standard")
              End If

              If frmabm.data_clientes.Recordset("cl_codconv") = "PART" Then
                 txt_precio.Text = Format(data_estudio.Recordset("part"), "Standard")
                 txt_precio.Text = Format(txt_precio.Text, "Standard")
              End If
          
                If DBCombo1.Text = "60106" Or _
                   DBCombo1.Text = "60108" Or _
                   DBCombo1.Text = "993" Or _
                   DBCombo1.Text = "994" Or _
                   DBCombo1.Text = "997" Or _
                   DBCombo1.Text = "996" Or _
                   DBCombo1.Text = "60105" Or DBCombo1.Text = "80016" Or _
                   DBCombo1.Text = "IMPASA M." Or DBCombo1.Text = "RESERVA PSIQUIATRA DR.FIELIZ PABLO" Or _
                   DBCombo1.Text = "SMI M." Or _
                   DBCombo1.Text = "80011" Or DBCombo1.Text = "RESERVA SICOLOGÍA" Or _
                   DBCombo1.Text = "80012" Or DBCombo1.Text = "RESERVA NUTRICIONISTA" Or _
                   DBCombo1.Text = "80013" Or DBCombo1.Text = "RESERVA ODONTOLOGÍA" Or _
                   DBCombo1.Text = "80014" Or DBCombo1.Text = "RESERVA CARNET DE SALUD" Or _
                   DBCombo1.Text = "80015" Or DBCombo1.Text = "RESERVA FISIOTERAPIA" Or _
                   DBCombo1.Text = "UNIVERSAL M." Or DBCombo1.Text = "COBRANZA VIGILIA" Or _
                   DBCombo1.Text = "999" Or DBCombo1.Text = "COBRANZA MUTUAL" Or _
                   DBCombo1.Text = "PAGO DE CUOTA EN BASE" Or _
                   DBCombo1.Text = "60109" Or _
                   DBCombo1.Text = "CASA DE GALICIA M." Then
                   If XQuefac = 4 Or XQuefac = 21 Then
                      Xquerr = 0
                   Else
                      MsgBox "Debe de facturar cómo RECIBO", vbCritical, "Mensaje"
                      Xquerr = 1
                      b_cance_Click
                      Exit Sub
                   End If
                Else
                   If XQuefac = 4 Then
                      MsgBox "Debe de facturar cómo E-TICKET o E-FACTURA", vbCritical, "Mensaje"
                      Xquerr = 1
                      b_cance_Click
                      Exit Sub
                   Else
                      Xquerr = 0
                   End If
                End If
              If data_estudio.Recordset("codest") = 993 Or _
                 data_estudio.Recordset("codest") = 994 Then
                 cbotim.Enabled = False
                 txt_mes.Text = ""
                 txt_ano.Text = ""
                 txt_mes.Enabled = True
                 txt_ano.Enabled = True
                 dbcbomed.Enabled = False
              End If
              If data_estudio.Recordset("codest") = 999 Then
                 cbotim.Enabled = False
                 txt_mes.Enabled = True
                 txt_ano.Enabled = True
                 dbcbomed.Enabled = False
                 data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " And fecha_pago is null and mes >" & 0 & " order by ano,mes"
                 data_deudas.Refresh
                 If data_deudas.Recordset.RecordCount > 0 Then
                    frm_veodeuda.Show vbModal
                    txt_mes.Enabled = False
                    txt_ano.Enabled = False
                 Else
                    MsgBox "No figura Deuda, no se puede realizar cobro"
                    Unload Me
                    frmabm.btn_fact.Enabled = True
                    Exit Sub
                 End If
              End If
              If data_estudio.Recordset("codest") = 997 Then
                 cbotim.Enabled = False
                 txt_mes.Enabled = True
                 txt_ano.Enabled = True
                 txt_mes.Text = ""
                 txt_ano.Text = ""
                 dbcbomed.Enabled = False
                 data_deudas.RecordSource = "Select * from deudas where cliente =" & labmatri.Caption & " And fecha_pago is null and mes =" & 0 & " order by ano,mes"
                 data_deudas.Refresh
                 If data_deudas.Recordset.RecordCount > 0 Then
                    frm_veodeuda.Show vbModal
                 Else
                    Xcuota = 0
                    MsgBox "Socio sin FACTURAS A CREDITO PENDIENTES...$: " & Trim(str(Xcuota)), vbInformation, "Mensaje"
                    txt_precio.Text = Xcuota
                    If txt_precio.Enabled = True Then
                       txt_precio.SetFocus
                    End If
                 End If
                 txt_mes.Text = ""
                 txt_ano.Text = ""
                 txt_mes.Enabled = False
                 txt_ano.Enabled = False
              End If
              
              If Xquerr <> 1 Then
                 If data_estudio.Recordset("codest") = 60103 Or data_estudio.Recordset("codest") = 60106 Or _
                    data_estudio.Recordset("codest") = 60107 Or data_estudio.Recordset("codest") = 60108 Then
                    btn_fin.SetFocus
                 Else
                    If txt_precio.Enabled = True Then
                       txt_precio.SetFocus
                    End If
                 End If
                DBCombo1.Height = 500
                DBCombo1.ListField = ""
                DBCombo1.BoundColumn = ""
              Else
                DBCombo1.Height = 500
                DBCombo1.ListField = ""
                DBCombo1.BoundColumn = ""
                DBCombo1.Text = ""
                txt_precio.Text = ""
              End If
           Else
              data_estudio.RecordSource = "select * from estudios where descrip >='" & DBCombo1.Text & "' order by descrip"
              data_estudio.Refresh
              DBCombo1.Height = 1350
           End If
        End If
   Else
       If btn_fin.Enabled = True Then
          btn_fin.SetFocus
       End If
   End If
End If

Exit Sub

Aldbcbo1:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al dbcbo1"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al dbcbo1"
            data_errfact.Recordset.Update
            Unload Me
         End If


End Sub



Private Sub DBCombo1_LostFocus()
Dim Xc As Long
Dim Xlafv As Date
Dim Xelcodigoaut, Xlapersona As String
Dim XCGalicia, Xtienecmt As Integer
XCGalicia = 0
Xtienecmt = 0

'terminado
On Error GoTo Aldbcbo1los

If data_estudio.Recordset("codest") = 999 Or _
   data_estudio.Recordset("codest") = 993 Or _
   data_estudio.Recordset("codest") = 994 Then
   cbotim.Enabled = False
   txt_mes.Enabled = True
   txt_ano.Enabled = True
   dbcbomed.Enabled = False
Else
   If data_estudio.Recordset("codest") = 995 And Trim(DBCombo1.Text) <> "" Then
      If Label7.Caption = "E-TICKET" Then
         MsgBox "No se puede facturar timbres desde ésta opción, seleccione la opción TIMBRE SI", vbInformation
         frmabm.btn_fact.Enabled = True
         End
      Else
         cbotim.Enabled = True
         txt_mes.Text = ""
         txt_ano.Text = ""
         txt_mes.Enabled = False
         txt_ano.Enabled = False
         dbcbomed.Enabled = True
      End If
   Else
      cbotim.Enabled = True
      txt_mes.Text = ""
      txt_ano.Text = ""
      txt_mes.Enabled = False
      txt_ano.Enabled = False
      dbcbomed.Enabled = True
   End If
End If
'If frm_menu.data_parse.Recordset("base") = 12 And data_estudio.Recordset("codest") <> 30081 Then
'   If data_estudio.Recordset("codest") = 30084 Or data_estudio.Recordset("codest") = 30085 Then
'   Else
'      MsgBox "Servicio no autorizado a facturar", vbCritical
'      End
'   End If
'End If
   
If DBCombo1.Text <> "" Then
    If Xquehag <> 9 Then
       If labfpago.Caption = "CREDITO" Or labfpago.Caption = "CONTADO" Then
          If data_estudio.Recordset("sin_deuda") = 1 Or data_estudio.Recordset("flia") = 8 Then
             Xquehag = 0
          Else
             Xdeb = 2
             Wopszond = ""
             data_consdeu.RecordSource = "Select * from deudas where cliente =" & Val(labmatri.Caption) & " and tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanciacion" & "' and mes =" & 0
             data_consdeu.Refresh
             If data_consdeu.Recordset.RecordCount > 0 Then
                data_consdeu.Recordset.MoveFirst
                Do While Not data_consdeu.Recordset.EOF
                   If IsNull(data_consdeu.Recordset("nro_superv")) = False Then
                      Xlafv = data_consdeu.Recordset("fecha") + data_consdeu.Recordset("nro_superv")
                   Else
                      Xlafv = data_consdeu.Recordset("fecha") + 30
                   End If
                   If Format(Xlafv, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                      If XQuefac = 102 Or XQuefac = 112 Then
                         Xquehag = 0
                         Wxquepreg = 0
                      Else
                         Xquehag = 9
                         Wxquepreg = 1 'es deuda por servicio
                      End If
                   End If
                   data_consdeu.Recordset.MoveNext
                Loop
             Else
                Xquehag = 0
             End If
                     
             data_consdeu.RecordSource = "Select * from deudas where cliente =" & Val(labmatri.Caption) & " and mes >" & 0 & " and fecha_pago is null and origen <='" & "Refinan" & "' order by ano,mes"
             data_consdeu.Refresh
             If data_consdeu.Recordset.RecordCount > 0 Then
                data_consdeu.Recordset.MoveLast
                If data_consdeu.Recordset.RecordCount > 2 Then
                   Xop4 = data_consdeu.Recordset("mes")
                   Xop5 = data_consdeu.Recordset("ano")
                   Xquehag = 9
                   If Wxquepreg = 0 Then
                      Wxquepreg = 2 'es por cuota
                   End If
                End If
             End If
                     
             data_consdeu.RecordSource = "Select * from deudas where cliente =" & Val(labmatri.Caption) & " and fecha_pago is null and origen >='" & "Refinan" & "'"
             data_consdeu.Refresh
             If data_consdeu.Recordset.RecordCount > 0 Then
                data_consdeu.Recordset.MoveFirst
                Do While Not data_consdeu.Recordset.EOF
                   If IsNull(data_consdeu.Recordset("nro_superv")) = False Then
                      Xlafv = data_consdeu.Recordset("fecha") + data_consdeu.Recordset("nro_superv")
                   Else
                      Xlafv = data_consdeu.Recordset("fecha") + 30
                   End If
                   If Format(Xlafv, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                      Xquehag = 9
                      Wxquepreg = 3 'es por refinanc
                   End If
                   data_consdeu.Recordset.MoveNext
                Loop
             End If
          End If
          If Xquehag = 9 Then
             If data_estudio.Recordset("codest") = 30084 Or data_estudio.Recordset("codest") = 30085 Or data_estudio.Recordset("codest") = 30081 Then
                Xquehag = 0
             Else
                MsgBox "Socio moroso, se pasará al sistema de autorización automática", vbCritical
                Xhab = Val(labmatri.Caption)
                frm_autoriza.Show vbModal
                Xelcodigoaut = InputBox("INGRESE CÓDIGO DE AUTORIZACIÓN:", "AUTORIZACIÓN", Wopszond)
                If Trim(Xelcodigoaut) <> "" Then
                   data_aut.RecordSource = "select * from Codigos_aut where codaut ='" & Trim(Xelcodigoaut) & "' and socio =" & Val(labmatri.Caption)
                   data_aut.Refresh
                   If data_aut.Recordset.RecordCount > 0 Then
                      Xquehag = 0
                   Else
                      MsgBox "ATENCION! No se encuentra código de autorización, realice nuevamente la autorización o comunique a Administración", vbCritical
                      Xquehag = 9
                   End If
                Else
                   MsgBox "Socio moroso, debe ingresar autorización.", vbCritical
                   Xquehag = 9
                End If
                If Xquehag = 9 Then
                   frmabm.btn_fact.Enabled = True
                   Unload Me
                End If
             End If
          End If
       End If
    End If
''30084/84/85
    
    If Xestaok = 22 Or Xestaok = 19 Then
       If Xestaok = 19 And data_estudio.Recordset("codest") = 30081 Then
'       If Xestaok = 19 Then
       Else
          If data_estudio.Recordset("codest") = 30084 Or data_estudio.Recordset("codest") = 30085 Then
          Else
            If data_estudio.Recordset("sin_deuda") = 1 Or data_estudio.Recordset("flia") = 8 Then
            Else
               MsgBox "Realice carta mutual para poder facturar", vbCritical
               frmabm.btn_fact.Enabled = True
               Unload Me
            End If
          End If
       End If
    End If
End If
If data_estudio.Recordset("codest") = 13010 Or _
   data_estudio.Recordset("codest") = 13014 Or _
   data_estudio.Recordset("codest") = 13017 Or _
   data_estudio.Recordset("codest") = 13022 Or _
   data_estudio.Recordset("codest") = 13034 Then
   cbotim.ListIndex = 1
Else
   cbotim.ListIndex = 0
End If
If data_estudio.Recordset("codest") = 999 Or data_estudio.Recordset("codest") = 997 Then
   txt_mes.Enabled = False
   txt_ano.Enabled = False
   txt_precio.Enabled = False
Else
   If data_estudio.Recordset("codest") = 994 Or data_estudio.Recordset("codest") = 993 Then
      txt_mes.Enabled = True
      txt_ano.Enabled = True
      txt_precio.Enabled = True
   Else
      txt_mes.Enabled = False
      txt_ano.Enabled = False
      txt_precio.Enabled = True
   End If
End If
If data_estudio.Recordset("codest") = 992 Or data_estudio.Recordset("codest") = 984 Or _
   data_estudio.Recordset("codest") = 985 Or data_estudio.Recordset("codest") = 986 Or _
   data_estudio.Recordset("codest") = 987 Or data_estudio.Recordset("codest") = 989 Or _
   data_estudio.Recordset("codest") = 802 Or data_estudio.Recordset("codest") = 803 Or _
   data_estudio.Recordset("codest") = 804 Or data_estudio.Recordset("codest") = 805 Or _
   data_estudio.Recordset("codest") = 806 Then
   Label14.Visible = True
   labcodpro.Visible = True
   dbcboprom.Visible = True
Else
   Label14.Visible = False
   labcodpro.Visible = False
   dbcboprom.Visible = False
End If
If frmabm.txt_codcnv.Text = "APS" Then
   If data_estudio.Recordset("codest") = 10005 Then
   Else
      MsgBox "ATENCION! Paciente no habilitado para servicios. POR AGRESIONES! Consulte con Administración!", vbCritical
      Unload Me
   End If
End If

If data_estudio.Recordset("codest") = 10018 Or data_estudio.Recordset("codest") = 10050 Then
   If Label7.Caption = "NC E-TICKET" Then
   Else
      Xtienecmt = YatieneCMT()
      If Xtienecmt = 1 Then
         MsgBox "ATENCION!! Paciente ya figura con un CMT registrado en policlínica o Despacho. VERIFIQUE!", vbCritical
         Unload Me
      End If
   End If
End If

If data_estudio.Recordset("codest") = 30090 Or data_estudio.Recordset("codest") = 80018 Then
   XMensaFertilab = 9
   If frmabm.txt_nac.Text = "__/__/____" Then
      MsgBox "No figura ingresada la fecha de nacimiento. Modifique la ficha para poder facturar!", vbCritical
      frmabm.btn_fact.Enabled = True
      Unload Me
   Else
'   frm_mensaauto.Show vbModal
      If data_estudio.Recordset("codest") = 80018 Then
         XMensaFertilab2 = 9
         MsgBox "ATENCION!! DEBERA INGRESAR CORREO ELECTRONICO DEL PACIENTE.", vbCritical
      Else
         MsgBox "ATENCION!! el paciente debe concurrir al laboratorio Fertilab a realizarse el Examen.", vbCritical
      End If
      XcorreoFertilab = InputBox("INGRESE CORREO ELECTRONICO DEL PACIENTE:", "FACTURACION")
      If ValidarCorreoFerti() = 1 Then
      Else
         MsgBox "ATENCION! Debe ingresar una dirección de correo válida!!!", vbCritical
         frmabm.btn_fact.Enabled = True
         Unload Me
      End If
   End If
End If

If data_estudio.Recordset("codest") = 60103 Or _
   data_estudio.Recordset("codest") = 60105 Or _
   data_estudio.Recordset("codest") = 60106 Or _
   data_estudio.Recordset("codest") = 60108 Or _
   data_estudio.Recordset("codest") = 60107 Or _
   data_estudio.Recordset("codest") = 60109 Then
   If Label7.Caption = "E-TICKET" Or Label7.Caption = "RECIBO" Then
      frm_factselectm.Show vbModal
   End If
End If


Exit Sub

Aldbcbo1los:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al dbcbo1 los"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al dbcbo1 los"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'17005009

On Error GoTo ParaSalir

If KeyAscii = vbKeyEscape Then 'Se ha pulsado ESC
   b_cance.SetFocus
   b_cance_Click
End If
'211372 'cgalicia
Exit Sub

ParaSalir:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al salir"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al salir"
            data_errfact.Recordset.Update
            Unload Me
         End If
            
End Sub

Private Sub Form_Load()
Dim Xlafechacaja As Date
Dim Xnosepuede As Integer
Xnosepuede = 0
On Error GoTo Xelerralcomienzo

Xlafechacaja = Date - 10

'data_conv.DatabaseName = App.Path & "\sapp.mdb"
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_aut.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_facafil.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_estudiobus.Connect = "odbc;dsn=" & Xconexrmt & ";"

'data_conv.RecordSource = "convenio"
'data_conv.Refresh
Xcandelin = 0

Xelnrodeuda = 0
XMensaFertilab = 0
XMensaFertilab2 = 0
XcorreoFertilab = ""

data_u.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ctr.DatabaseName = App.path & "\ctrf.mdb"
data_ctr.RecordSource = "ctrf"
data_ctr.Refresh

data_eror.DatabaseName = App.path & "\selec.mdb"
data_eror.RecordSource = "selec"
data_eror.Refresh
If data_eror.Recordset.RecordCount > 0 Then
   data_eror.Recordset.MoveFirst
   Do While Not data_eror.Recordset.EOF
      data_eror.Recordset.Delete
      data_eror.Recordset.MoveNext
   Loop
End If

'data_cablocal.DatabaseName = App.Path & "\cablocal.mdb"

'data_cablocal.RecordSource = "cabezados"
'data_cablocal.Refresh

If frm_menu.data_parse.Recordset("base") = 38 Then
   data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

data_imagen.DatabaseName = App.path & "\imagen.mdb"
data_imagen.RecordSource = "qr"
data_imagen.Refresh
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_ui.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_ui.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
data_ui.RecordSource = "hc_frecresp"
data_ui.Refresh

data_eror.DatabaseName = App.path & "\erores.mdb"
data_eror.RecordSource = "erores"
data_eror.Refresh

data_errfact.DatabaseName = App.path & "\errores.mdb"
data_errfact.RecordSource = "errores"
data_errfact.Refresh

data_arancel.Connect = "odbc;dsn=" & Xconexrmt & ";"

mf.Text = Format(Date, "dd/mm/yyyy")

data_lindbgri.DatabaseName = App.path & "\factura.mdb"

data_cabeza2.DatabaseName = App.path & "\factura.mdb"
data_cabeza2.RecordSource = "cabezados"
data_cabeza2.Refresh

'data_caja.DatabaseName = App.Path & "\sapp.mdb"
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_caja.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_caja.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_lincance.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_lincance.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

data_codcaja.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_codcaja.RecordSource = "cod_caja"
data_codcaja.Refresh

'data_deudas.DatabaseName = App.Path & "\sapp.mdb"
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_deudas.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_deudas.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

If frm_menu.data_parse.Recordset("base") = 38 Then
   data_lincab.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_lincab.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
data_lincab.RecordSource = "Select * from clirespl where cl_codigo =" & 25048
data_lincab.Refresh

data_estudio.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_estudio.RecordSource = "estudios"
data_estudio.Refresh

data_lineas.DatabaseName = App.path & "\factura.mdb"
data_lineas.RecordSource = "lineas"
data_lineas.Refresh

'data_linmmdd.DatabaseName = App.Path & "\sapp.mdb"
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_linmmdd.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_linmmdd.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

data_medicos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_medicos.RecordSource = "select * from medicos order by med_nombre"
data_medicos.Refresh

data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
cbotim.ListIndex = 0
labmatri.Caption = frmabm.txt_mat.Caption
labnomb.Caption = frmabm.txt_apellid.Text

labfac.Caption = ""
If Xfpago = 1 Then
   labfpago.Caption = "CONTADO"
Else
   If Xfpago = 2 Then
      labfpago.Caption = "CREDITO"
   Else
      labfpago.Caption = "CONTADO"
   End If
End If

If XQuefac = 101 Then
   Label7.Caption = "E-TICKET"
   Xfaccancerecep = 0
Else
   If XQuefac = 111 Then
      Label7.Caption = "E-FACTURA"
      Check1.Visible = True
      Check1.Value = 1
      txt_rut.Visible = True
      Xfaccancerecep = 0
   Else
      If XQuefac = 102 Then
         Label7.Caption = "NC E-TICKET"
         b_verfaccance.Visible = True
         Xfaccancerecep = 1
         DBCombo1.Enabled = False
         txt_precio.Enabled = False
      Else
         If XQuefac = 103 Then
            Label7.Caption = "ND E-TICKET"
            b_verfaccance.Visible = True
            Xfaccancerecep = 1
'            DBCombo1.Enabled = False
'            txt_precio.Enabled = False
         Else
            If XQuefac = 112 Then
               Label7.Caption = "NC E-FACTURA"
               Check1.Visible = True
               Check1.Value = 1
               txt_rut.Visible = True
               b_verfaccance.Visible = True
               Xfaccancerecep = 1
               DBCombo1.Enabled = False
               txt_precio.Enabled = False
            Else
               If XQuefac = 113 Then
                  Label7.Caption = "ND E-FACTURA"
                  Check1.Visible = True
                  Check1.Value = 1
                  txt_rut.Visible = True
                  b_verfaccance.Visible = True
                  Xfaccancerecep = 1
'                  DBCombo1.Enabled = False
'                  txt_precio.Enabled = False
               Else
                  If XQuefac = 4 Then
                     Label7.Caption = "RECIBO"
                     Xfaccancerecep = 0
                  Else
                     If XQuefac = 21 Then
                        Label7.Caption = "DEV.RECIBO"
                        b_verfaccance.Visible = True
                        Xfaccancerecep = 1
                        DBCombo1.Enabled = False
                        txt_precio.Enabled = False
                     Else
                        MsgBox "Hay un error en tipo de factura, cierre el programa", vbCritical, "SAPP"
                        Xfaccancerecep = 0
                        End
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If


If data_cabeza2.Recordset.RecordCount > 0 Then
   data_cabeza2.Recordset.MoveFirst
   Do While Not data_cabeza2.Recordset.EOF
      data_cabeza2.Recordset.Delete
      data_cabeza2.Recordset.MoveNext
   Loop
   data_cabeza2.Refresh
End If

If data_lineas.Recordset.RecordCount > 0 Then
   data_lineas.Recordset.MoveFirst
   Do While Not data_lineas.Recordset.EOF
'      If IsNull(data_lineas.Recordset("unidad")) = False Then
'         If data_lineas.Recordset("unidad") = "N" Then
'            Xnosepuede = 9
'         Else
'            data_lineas.Recordset.Delete
'         End If
'      Else
      data_lineas.Recordset.Delete
'      End If
      data_lineas.Recordset.MoveNext
   Loop
   data_lineas.Refresh
End If

'If Xnosepuede = 9 Then
'   MsgBox "Hay un registro que se envió a DGI y no quedó registrado en SAPP, Avise a informática para poder continuar.", vbCritical
'   Xnosepuede = 0
'   End
'End If

data_func.Connect = "odbc;dsn=sappnew;"
data_func.RecordSource = "select * from vende_func order by nombre"
data_func.Refresh

'SelectLimit 10
data_caja.RecordSource = "select * from caja where id=" & 2328952
data_caja.Refresh
'SelectLimit 10

'SelectLimit 0
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_consdeu.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_consdeu.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

data_param.Connect = "odbc;dsn=sappfact;"
data_param.RecordSource = "paramsapp"
data_param.Refresh

If XQuefac = 111 Or XQuefac = 112 Or XQuefac = 113 Then
   txt_rut.Text = Xelrutfact
End If

labtot.Caption = 0

If Xfaccancerecep = 1 Then
   b_verfaccance_Click
End If


Exit Sub

Xelerralcomienzo:
                 If Err.Number = 3155 Then
                    MsgBox "No se pudo iniciar el cabezal de la factura. Avise a Informática", vbInformation
                    If Err.Number = 5 Then
                       data_errfact.Recordset.AddNew
                       data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                       data_errfact.Recordset("fecha") = Date
                       data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                       data_errfact.Recordset("nroerr") = Err.Number
                       data_errfact.Recordset("desc") = "cabezal fact"
                       data_errfact.Recordset.Update
                       Unload Me
                    Else
                       data_errfact.Recordset.AddNew
                       data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                       data_errfact.Recordset("fecha") = Date
                       data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                       data_errfact.Recordset("nroerr") = Err.Number
                       data_errfact.Recordset("desc") = "cabezal fact"
                       data_errfact.Recordset.Update
                       Unload Me
                    End If
                 
                 Else
                    MsgBox "No se pudo iniciar el cabezal de la factura. Avise a Informática", vbInformation
                    If Err.Number = 5 Then
                       data_errfact.Recordset.AddNew
                       data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                       data_errfact.Recordset("fecha") = Date
                       data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                       data_errfact.Recordset("nroerr") = Err.Number
                       data_errfact.Recordset("desc") = "cabezal fact"
                       data_errfact.Recordset.Update
                       Unload Me
                    Else
                       data_errfact.Recordset.AddNew
                       data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                       data_errfact.Recordset("fecha") = Date
                       data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                       data_errfact.Recordset("nroerr") = Err.Number
                       data_errfact.Recordset("desc") = "cabezal fact"
                       data_errfact.Recordset.Update
                       Unload Me
                    End If
                 
                 End If

                 
End Sub

Private Sub Form_Resize()
On Error GoTo Alimagen

With Image1
   .Left = 0
   .Top = 0
   .Width = Me.Width
   .Height = Me.Height
End With

Exit Sub

Alimagen:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al imagen"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al imagen"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub


Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btn_graba.SetFocus
End If

End Sub

Private Sub t_cant_LostFocus()
If t_cant.Text <> "" Then
   If t_cant.Text > 99 Then
      MsgBox "Verifique si está bien la cantidad", vbCritical
      t_cant.SetFocus
   Else
   
   End If
End If
   
End Sub

Private Sub txt_ano_KeyPress(KeyAscii As Integer)
On Error GoTo Altexano

If KeyAscii = 13 Then
   If dbcbomed.Enabled = True Then
      dbcbomed.SetFocus
   Else
      btn_graba.SetFocus
   End If
End If

Exit Sub

Altexano:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al TextAno"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al TextAno"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub txt_ano_LostFocus()
On Error GoTo Altxtano

If txt_precio.Text <> "" Then
   If txt_mes.Text <> "" Then
      If txt_mes.Text >= 1 And txt_mes.Text <= 12 Then
            If txt_ano.Text <> "" Then
               If txt_ano.Text <> 0 Then
                  If txt_ano.Text > 2000 Then
                     btn_graba.SetFocus
                  Else
                     MsgBox "Ingrese año válido", vbCritical, "Mensaje"
                     txt_ano.SetFocus
                  End If
               Else
                  MsgBox "Ingrese año válido", vbCritical, "Mensaje"
                  txt_ano.SetFocus
               End If
            Else
               MsgBox "Ingrese año válido", vbCritical, "Mensaje"
               txt_ano.SetFocus
            End If
       Else
          txt_mes.SetFocus
       End If
   Else
       txt_mes.SetFocus
   End If
Else
    txt_precio.SetFocus
End If

Exit Sub

Altxtano:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al TXTAno"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al TXTAno"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub txt_mes_KeyPress(KeyAscii As Integer)
On Error GoTo Altxtmes

If KeyAscii = 13 Then
   txt_mes_LostFocus
End If

Exit Sub

Altxtmes:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al txtMES"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al txtMES"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub txt_mes_LostFocus()
On Error GoTo AltxtmesLos

If txt_precio.Text <> "" Then
    If txt_mes.Text <> "" Then
       If txt_mes.Text <> 0 Then
          If txt_mes.Text >= 1 And txt_mes.Text <= 12 Then
             txt_ano.SetFocus
          Else
             MsgBox "Ingrese mes válido", vbCritical, "Mensaje"
             txt_mes.SetFocus
          End If
       Else
          MsgBox "Ingrese mes válido", vbCritical, "Mensaje"
          txt_mes.SetFocus
       End If
    Else
       MsgBox "Ingrese mes válido", vbCritical, "Mensaje"
       txt_mes.SetFocus
    End If
Else
    txt_precio.SetFocus
End If

Exit Sub

AltxtmesLos:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al TXTMES Los"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al TXTMES Los"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
On Error GoTo Altprecio

If KeyAscii = 13 Then
'   txt_precio.Text = Format(txt_precio.Text, "Standard")
   If Val(Label5.Caption) = 60103 Or Val(Label5.Caption) = 60106 Or _
      Val(Label5.Caption) = 60107 Or Val(Label5.Caption) = 60108 Then
      btn_graba.SetFocus
   Else
      If dbcboprom.Visible = True Then
         dbcboprom.SetFocus
      Else
         If cbotim.Enabled = True Then
            If cbotim.Enabled = True Then
               cbotim.SetFocus
            End If
         Else
            If txt_mes.Enabled = True Then
               If txt_mes.Enabled = True Then
                  txt_mes.SetFocus
               End If
            Else
               If dbcbomed.Enabled = True Then
                  If dbcbomed.Enabled = True Then
                     dbcbomed.SetFocus
                  End If
               Else
                  If btn_graba.Enabled = True Then
                     If Xdeb <> 2 Then
                        btn_graba.SetFocus
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If
Exit Sub

Altprecio:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al T Precio"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al T Precio"
            data_errfact.Recordset.Update
            Unload Me
         End If

End Sub

Private Sub txt_precio_LostFocus()
On Error GoTo AlprecioLos
If data_estudio.Recordset("codest") = 993 Or _
   data_estudio.Recordset("codest") = 994 Or _
   data_estudio.Recordset("codest") = 999 Then
   If txt_precio.Text <> "" Then
      txt_precio.Text = Format(txt_precio.Text, "Standard")
      If dbcboprom.Visible = True Then
         dbcboprom.SetFocus
      Else
        If cbotim.Enabled = True Then
           cbotim.SetFocus
        Else
           If txt_mes.Enabled = True Then
              txt_mes.SetFocus
           Else
              If btn_graba.Enabled = True Then
                 If Xdeb <> 2 Then
                    btn_graba.SetFocus
                 End If
              End If
           End If
        End If
     End If
   Else
      MsgBox "Debe ingresar importe", vbCritical, "Mensaje"
      txt_precio.SetFocus
   End If
Else
   txt_precio.Text = Format(txt_precio.Text, "Standard")
   If dbcboprom.Visible = True Then
      dbcboprom.SetFocus
   Else
        If cbotim.Enabled = True Then
           cbotim.SetFocus
        Else
           If txt_mes.Enabled = True Then
              txt_mes.SetFocus
           Else
              If Xdeb <> 2 Then
                 btn_graba.SetFocus
              End If
           End If
        End If
   End If
End If
If txt_precio.Text <> "" Then
   If txt_precio.Text < 0 Then
      MsgBox "No puede ingresar monto menor a cero", vbCritical
      b_cance_Click
   Else
      If txt_precio.Text >= 40000 Then
         MsgBox "Confirme si es correcto el importe!!", vbCritical
      End If
   End If
End If

Exit Sub

AlprecioLos:
         If Err.Number = 5 Then
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al precio Los"
            data_errfact.Recordset.Update
            Unload Me
         Else
            data_errfact.Recordset.AddNew
            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
            data_errfact.Recordset("fecha") = Date
            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
            data_errfact.Recordset("nroerr") = Err.Number
            data_errfact.Recordset("desc") = "Al precio Los"
            data_errfact.Recordset.Update
            Unload Me
         End If


End Sub


Private Function EstaInicializado() As Boolean

    EstaInicializado = False

    If objPosCfe Is Nothing Or Not objPosCfe.Inicializado Then
        MsgBox "Debe inicializar el POS"
        Set objPosCfe = Nothing
        
        Exit Function
    End If

    EstaInicializado = True
End Function

Private Sub DesplegarInfoEstadoCfe(Mensaje As String, ResultadoCfe As ResultadoCfe)

On Error GoTo XQuepasaalenviar


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
        MsgBox "El CFE no trae número de folio"
        
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.DatosCae Is Nothing Then
        MsgBox "El CFE no trae datos del CAE"
                
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
    
    If Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 0 Or _
       Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 11 Then
       labserie.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.serie
       labfac.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
       data_cabeza2.Recordset.Edit
       labvence.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Vencimiento)
       labautoriza.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Autorizacion)
       labdesde.Caption = labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroDesde) & " - " & labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroHasta)
       labhasta.Caption = CStr(ResultadoCfe.EstadoCfe.CodigoSeguridad)
       If Len(labvence.Caption) = 8 Then
          labvenceok.Caption = Mid(labvence.Caption, 7, 2) & "/" & Mid(labvence.Caption, 5, 2) & "/" & Mid(labvence.Caption, 1, 4)
       Else
          labvenceok.Caption = "31/12/2016"
       End If
       data_cabeza2.Recordset("cl_fultpag") = CDate(labvenceok.Caption)
       data_cabeza2.Recordset("cl_nrocobr") = Val(labautoriza.Caption)
       data_cabeza2.Recordset("cl_medflia") = Trim(labdesde.Caption)
       data_cabeza2.Recordset("cl_fax") = Trim(labhasta.Caption)
       data_cabeza2.Recordset.Update
    Else
       data_eror.Recordset.AddNew
       data_eror.Recordset("nro") = 11
       data_eror.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
       data_eror.Recordset("hora") = Format(Time, "HH:mm")
       data_eror.Recordset("obs") = "FACT CANCE"
       data_eror.Recordset.Update
       MsgBox "Comprobante RECHAZADO, NO FUE ACEPTADO, debe realizarlo nuevamente, verifique datos!", vbInformation
       
       End
    End If

    strUltimoGuid = ResultadoCfe.EstadoCfe.Guid
    Set objUltimaSerieNumero = ResultadoCfe.EstadoCfe.SerieNumeroCfe

Exit Sub

XQuepasaalenviar:
              If Err.Number = 3155 Then
                 MsgBox "Error al terminar la factura: " & Err.Number & " " & Err.Description, vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = 11
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = 11
                 data_errfact.Recordset("desc") = "FACT.REC"
                 data_errfact.Recordset.Update
              Else
                 MsgBox "Error al terminar la factura:" & Err.Number & " " & Err.Description, vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = 11
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = 12
                 data_errfact.Recordset("desc") = "FACT.CANCE"
                 data_errfact.Recordset.Update
              End If
                 
                 

End Sub

Public Function Cgalicia() As Integer

Dim XsqlpromoF As String
Dim XreccliiAvisoF As New ADODB.Recordset

ConectarAvisoF
ConbdSappAvisoF.Open

If frmabm.txt_codcnv.Text <> "" Then
   XsqlpromoF = "Select * from convenio where cnv_codigo ='" & frmabm.txt_codcnv.Text & "' and cnv_grupo in ('CASA DE GALICIA')"
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
   Cgalicia = 0
End If
XreccliiAvisoF.Close
ConbdSappAvisoF.Close


End Function
Public Function YatieneCMT() As Integer

Dim XsqlpromoF As String
Dim XreccliiAvisoF As New ADODB.Recordset
Dim Xlacedulahc As Long


ConectarAvisoF
ConbdSappAvisoF.Open

XsqlpromoF = "Select * from linmmdd where cod_cli =" & Val(frmabm.txt_mat.Caption) & " and fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cod_prod in (10050,10018)"
With XreccliiAvisoF
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
End With
If XreccliiAvisoF.RecordCount > 0 Then
   If IsNull(XreccliiAvisoF("ced_socio")) = False Then
      Xlacedulahc = XreccliiAvisoF("ced_socio")
   End If
   YatieneCMT = 1
Else
   YatieneCMT = 0
End If
XreccliiAvisoF.Close

XsqlpromoF = "Select * from llamado where matric =" & Val(frmabm.txt_mat.Caption) & " and fecha ='" & Format(Date, "yyyy-mm-dd") & "' and movilpas in (0,2015)"
With XreccliiAvisoF
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
End With
If XreccliiAvisoF.RecordCount > 0 Then
   If IsNull(XreccliiAvisoF("ci")) = False Then
      Xlacedulahc = XreccliiAvisoF("ci")
   End If
   If YatieneCMT = 0 Then
      YatieneCMT = 1
   End If
Else
   If YatieneCMT = 1 Then
   Else
      YatieneCMT = 0
   End If
End If
XreccliiAvisoF.Close

If YatieneCMT = 1 Then
   XsqlpromoF = "select * from cabezal_hcdig where fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cednum =" & Xlacedulahc & " and tipo_consd in ('Orientación Telefónica')"
   With XreccliiAvisoF
      .CursorLocation = adUseClient
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
   End With
   If XreccliiAvisoF.RecordCount > 0 Then
      YatieneCMT = 0
   Else
      YatieneCMT = 1
   End If
   XreccliiAvisoF.Close
End If

ConbdSappAvisoF.Close


End Function


Public Function ValidarCorreoFerti() As Integer
Dim XX As Integer
Dim textocorreo As String
textocorreo = ""
If UCase(Trim(XcorreoFertilab)) <> "NO APLICA" Then
   For XX = 1 To Len(Trim(XcorreoFertilab))
       If Mid(Trim(XcorreoFertilab), XX, 1) = "@" Then
          textocorreo = "@"
       Else
          If Mid(Trim(XcorreoFertilab), XX, 1) = "." Then
             If textocorreo = "@" Then
                textocorreo = textocorreo + "."
             End If
          End If
       End If
   Next
   If textocorreo = "@." Then
      ValidarCorreoFerti = 1
   Else
      ValidarCorreoFerti = 0
   End If
Else
   ValidarCorreoFerti = 0
End If

End Function

Public Sub EnviarCorreoFert()
Dim EnviarCorreo, CorreoAP, textocorreo As String
Dim Noenvia As Integer

EnviarCorreo = ""
textocorreo = ""
CorreoAP = ""

Dim MenCorreo2 As String
Dim oMail2 As Class1
Set oMail2 = New Class1
With oMail2
     .servidor = "smtp.office365.com"
     .puerto = 25
     .UseAuntentificacion = True
     .ssl = True
     .Usuario = "jefedepartamentoti@sapp.com.uy"
     .PassWord = "DptotiJunio2021"
     If Label7.Caption = "NC E-FACTURA" Or Label7.Caption = "NC E-TICKET" Then
        If data_estudio.Recordset("codest") = 80018 Then
           .Asunto = "ANULACION ORDEN DE LABORATORIO DESDE FERTILAB -REGISTRADO EN SAPP NRO.:" & labserie.Caption & "-" & labfac.Caption & " BASE:" & data_lineas.Recordset("base")
        Else
           .Asunto = "ANULACION ORDEN DE LABORATORIO DESDE SAPP Nro.:" & labserie.Caption & "-" & labfac.Caption & " BASE:" & data_lineas.Recordset("base")
        End If
     Else
        If data_estudio.Recordset("codest") = 80018 Then
           .Asunto = "ORDEN DE LABORATORIO DESDE FERTILAB -REGISTRADO EN SAPP NRO.:" & labserie.Caption & "-" & labfac.Caption & " BASE:" & data_lineas.Recordset("base")
        Else
           .Asunto = "ORDEN DE LABORATORIO DESDE SAPP Nro.:" & labserie.Caption & "-" & labfac.Caption & " BASE:" & data_lineas.Recordset("base")
        End If
     End If
     .de = "jefedepartamentoti@sapp.com.uy"
     .para = "fertilab@fertilab.com.uy; jefedepartamentoti@sapp.com.uy; jefeadministracion@sapp.com.uy"
     .Mensaje = labcorreo.Caption
     .Enviar_Backup ' manda el mail
End With
Set oMail2 = Nothing

End Sub

Public Sub Actualiza_Pedidos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from pedidos_facturar where id =" & data_lineas.Recordset("nro_pedido") & " and fecha_fact is null"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii("fecha_fact") = Date
   Xrecclii("nro_factura") = Val(labfac.Caption)
   Xrecclii("usuario") = WElusuario
   Xrecclii("base") = data_lineas.Recordset("base")
   Xrecclii.Update
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
