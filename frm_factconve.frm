VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_factconve22 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Facturacion a convenios"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Data data_timbre 
      Caption         =   "data_timbre"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_errfact 
      Caption         =   "data_errfact"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   3840
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowControlBox=   0   'False
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      DiscardSavedData=   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowZoomCtl=   0   'False
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.PictureBox Picture1 
      DataField       =   "qr"
      DataSource      =   "data_imagen"
      Height          =   1695
      Left            =   6840
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   71
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data data_imagen 
      Caption         =   "data_imagen"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_lincance 
      Caption         =   "data_lincance"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton b_vernc 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8280
      Picture         =   "frm_factconve.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Ver facturas del cliente"
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cbomon 
      Height          =   360
      ItemData        =   "frm_factconve.frx":058A
      Left            =   3840
      List            =   "frm_factconve.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "NC etck"
      Height          =   375
      Left            =   7560
      TabIndex        =   58
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Fin"
      Height          =   375
      Left            =   6360
      TabIndex        =   57
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Data data_cabeza2 
      Caption         =   "data_cabeza2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "etck"
      Height          =   375
      Left            =   4800
      TabIndex        =   56
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "efact"
      Height          =   375
      Left            =   3120
      TabIndex        =   55
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "nc efact"
      Height          =   375
      Left            =   5400
      TabIndex        =   54
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data data_ui 
      Caption         =   "data_ui"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_buscalafac 
      Caption         =   "data_buscalafac"
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
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSMask.MaskEdBox labvence 
      Height          =   375
      Left            =   4080
      TabIndex        =   53
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16711680
      ForeColor       =   16777215
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Data data_ctrolfact 
      Caption         =   "data_ctrolfact"
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7560
      TabIndex        =   52
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   360
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Data data_lin3 
      Caption         =   "data_lin3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_lin2 
      Caption         =   "data_lin2"
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox t_cant 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   49
      Top             =   3360
      Width           =   855
   End
   Begin VB.Data data_cabezal 
      Caption         =   "data_cabezal"
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
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox mfh 
      Height          =   375
      Left            =   7680
      TabIndex        =   46
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16711680
      ForeColor       =   16777215
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mfd 
      Height          =   375
      Left            =   6000
      TabIndex        =   45
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16711680
      ForeColor       =   16777215
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
   Begin VB.TextBox t_pie 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   960
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   7380
      Width           =   7575
   End
   Begin VB.Data data_verfac 
      Caption         =   "data_verfac"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FF80&
      Caption         =   "Consumidor Final"
      Height          =   255
      Left            =   6000
      TabIndex        =   37
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox t_desc3 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      MaxLength       =   62
      TabIndex        =   36
      Top             =   2880
      Width           =   6975
   End
   Begin VB.TextBox t_desc2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      MaxLength       =   70
      TabIndex        =   35
      Top             =   2520
      Width           =   6975
   End
   Begin VB.CommandButton b_calc 
      BackColor       =   &H0080FF80&
      Height          =   375
      Left            =   6600
      MaskColor       =   &H0000FFFF&
      Picture         =   "frm_factconve.frx":05A0
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Calcular importe"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton b_elilin 
      Caption         =   "Eliminar línea"
      Height          =   375
      Left            =   6960
      TabIndex        =   31
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox t_desc 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   30
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
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
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_temp 
      Caption         =   "data_temp"
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
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_factconve.frx":09E2
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frm_factconve.frx":09FA
      TabIndex        =   21
      ToolTipText     =   "SELECCIONE EL REGISTRO HACIENDO DOBLE CLICK"
      Top             =   5160
      Width           =   8775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar Factura"
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Terminar"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton b_graba 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox cboiva 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "frm_factconve.frx":1C0D
      Left            =   6120
      List            =   "frm_factconve.frx":1C1A
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox t_imp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1920
      TabIndex        =   14
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox t_ano 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox t_mes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "frm_factconve.frx":1C2C
      Left            =   1920
      List            =   "frm_factconve.frx":1C36
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   3855
   End
   Begin MSMask.MaskEdBox mf 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16711680
      ForeColor       =   16777215
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label labcanttiquet 
      Height          =   375
      Left            =   8280
      TabIndex        =   79
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labimptiquet 
      Height          =   375
      Left            =   6960
      TabIndex        =   78
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labfeccance 
      Height          =   375
      Left            =   5040
      TabIndex        =   77
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label labvenceok 
      Height          =   255
      Left            =   3360
      TabIndex        =   76
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labcodseg 
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   4440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label labrango 
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label labautoriza 
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label labvencecnv 
      Height          =   255
      Left            =   120
      TabIndex        =   72
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label labstot0 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   7440
      TabIndex        =   70
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF80&
      Caption         =   "S-Tot(0)"
      Height          =   255
      Left            =   6360
      TabIndex        =   69
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label labstot22 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   4560
      TabIndex        =   68
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF80&
      Caption         =   "S-Tot(22%):"
      Height          =   255
      Left            =   3240
      TabIndex        =   67
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label labnrofactnc 
      Height          =   375
      Left            =   5520
      TabIndex        =   65
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labseriefactnc 
      Height          =   375
      Left            =   4680
      TabIndex        =   64
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label labnrolin 
      Height          =   375
      Left            =   3840
      TabIndex        =   63
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label labiva22 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   4320
      TabIndex        =   62
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label labiv22 
      BackColor       =   &H00FFFF80&
      Caption         =   "IVA 22%:"
      Height          =   255
      Left            =   3240
      TabIndex        =   61
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Moneda"
      Height          =   255
      Left            =   3840
      TabIndex        =   60
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label labmontoit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   3000
      TabIndex        =   51
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF80&
      Caption         =   "Monto x ITEM:"
      Height          =   255
      Left            =   1560
      TabIndex        =   50
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cantidad:"
      Height          =   375
      Left            =   5160
      TabIndex        =   48
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label labserie 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   6600
      TabIndex        =   47
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080FF80&
      Caption         =   "Fechas de los servicios"
      Height          =   255
      Left            =   6000
      TabIndex        =   44
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF80&
      Caption         =   "ADENDA:"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lablocal 
      Height          =   255
      Left            =   2880
      TabIndex        =   41
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label labnrocli 
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label labdom 
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vencimiento:"
      Height          =   255
      Left            =   4080
      TabIndex        =   38
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label labcalc 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FF80&
      Caption         =   "Cálculo:"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Descripción:"
      Height          =   495
      Left            =   240
      TabIndex        =   29
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label labtot 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   7320
      TabIndex        =   28
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFF80&
      Caption         =   "Total:"
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label labiva 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "IVA 10%:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label labstot 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "S-Tot(10%)"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label labusuario 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   7440
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tasa IVA:"
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "IMPORTE:"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "MES/AÑO:"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Servicio:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FECHA:"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   -720
      X2              =   8400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label labnrofac 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SERIE/NUMERO:"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label labtipof 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FACTURA:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label labcli 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CLIENTE:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   240
      Picture         =   "frm_factconve.frx":1C4F
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   3135
   End
End
Attribute VB_Name = "frm_factconve22"
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

Public Xbb, Xconlin As Integer
Public Xivvva, Xtot, Xsubt As Double


Private Sub b_calc_Click()
'frm_calcula.Show vbModal

End Sub

Private Sub b_elilin_Click()
Dim Xloelimina As String
On Error GoTo Alelilin

'MsgBox "SELECCIONE EL REGISTRO A ELIMINAR HACIENDO DOBLE CLICK"
Xloelimina = MsgBox("Desea eliminar el registro seleccionado? " & data_temp.Recordset("cod_prod") & " " & data_temp.Recordset("solicitant") & " ??", vbCritical + vbYesNo, "BORRAR")
If Xloelimina = vbYes Then
   labstot.Caption = CDbl(labstot.Caption) - CDbl(data_temp.Recordset("tot_lin"))
   Xsubt = Xsubt - CDbl(data_temp.Recordset("tot_lin"))
   labiva.Caption = CDbl(labiva.Caption) - CDbl(data_temp.Recordset("pre_civa"))
   labtot.Caption = CDbl(labtot.Caption) - CDbl(data_temp.Recordset("tot_lin"))
   labtot.Caption = CDbl(labtot.Caption) - CDbl(data_temp.Recordset("pre_civa"))
   Xivvva = Xivvva - CDbl(data_temp.Recordset("pre_civa"))
   Xtot = Xtot - CDbl(data_temp.Recordset("tot_lin"))
   Xtot = Xtot - CDbl(data_temp.Recordset("pre_civa"))
   data_temp.Recordset.Delete
   data_temp.Refresh
   Xcandelin = Xcandelin - 1
   Xconlin = Xconlin - 1
End If
Exit Sub

Alelilin:
              If Err.Number = 3155 Then
                 MsgBox "Error al terminar la factura: " & Err.Number & " " & Err.Description, vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = Err.Number
                 data_errfact.Recordset("desc") = "Al eliminar"
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
                    data_errfact.Recordset("desc") = "Al eliminar"
                    data_errfact.Recordset.Update
                    Unload Me
                 Else
                    MsgBox "Error al terminar la factura:" & Err.Number & " " & Err.Description, vbInformation
                    data_errfact.Recordset.AddNew
                    data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                    data_errfact.Recordset("fecha") = Date
                    data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                    data_errfact.Recordset("nroerr") = Err.Number
                    data_errfact.Recordset("desc") = "Al eliminar"
                    data_errfact.Recordset.Update
                    Unload Me
                 End If
              End If

End Sub

Private Sub b_graba_Click()
Dim Xivauno As Double
Dim Xlafdesde As Date
Dim Xfeccontrolh As Date
Dim Xlf As Date
Dim Xelano As Integer
Dim Xmotivoref As String
Dim XivaInc, XtotInc As Double

On Error GoTo Algrabar

Xelano = Year(Date) + 1
Xlafdesde = Date - 400
Xfeccontrolh = CDate(mf.Text) - 60

If t_mes.Text <> "" And t_ano.Text <> "" And Combo1.ListIndex >= 0 And t_cant.Text <> "" Then
   If t_mes.Text >= 1 And t_mes.Text <= 12 Then
        data_verfac.RecordSource = "Select * from linmmdd where base =" & Xbb & " and tipo ='" & labtipof.Caption & "' and fecha >=#" & Format(Xlafdesde, "yyyy/mm/dd") & "# order by fecha"
        data_verfac.Refresh
        If data_verfac.Recordset.RecordCount > 0 Then
           data_verfac.Recordset.MoveLast
           Xlf = data_verfac.Recordset("fecha")
        Else
           Xlf = Date - 1
        End If
        Xlf = Xlf - 400
        If mf.Text < Xlf Then
           MsgBox "No puede facturar con una fecha anterior a la última realizada", vbInformation, "Facturación"
           mf.SetFocus
        Else
           If Format(Xfeccontrolh, "yyyy/mm/dd") >= Format(Date, "yyyy/mm/dd") Then
              MsgBox "La fecha no puede exceder los 60 días", vbInformation
              mf.SetFocus
           Else
                Xivauno = 0
                If t_imp.Text = "" Then
                   t_imp.Text = 0
                End If
                If t_mes.Text = "" Then
                   t_mes.Text = 0
                End If
                If t_ano.Text = "" Then
                   t_ano.Text = 0
                End If
                Xivvva = 0
                If mf.Text <> "__/__/____" Then
                   Xconlin = Xconlin + 1
                   If Xconlin > 5 Then
                      MsgBox "Cantidad de líneas por factura son 5, presione TERMINAR"
                   Else
                      If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Or Label3.Caption = "ND de E-TICKET" Then
                         Xmotivoref = InputBox("Ingrese motivo de la modificación")
                         If Xmotivoref <> "" Then
                            If labseriefactnc.Caption = "XX" Then
                               data_lin.RecordSource = "Select * from linmmdd where factura =" & Val(labnrofactnc.Caption)
                               data_lin.Refresh
                               Xcandelin = Xcandelin + 1
                               data_temp.Recordset.AddNew
                               If Label3.Caption = "NC de E-FACTURA" Then
                                  data_temp.Recordset("tipodocref") = 111
                               Else
                                  If Label3.Caption = "NC de E-TICKET" Then
                                     data_temp.Recordset("tipodocref") = 101
                                  Else
                                     data_temp.Recordset("tipodocref") = 111
                                  End If
                               End If
                               data_temp.Recordset("serieref") = labseriefactnc.Caption
                               If Len(labnrofactnc.Caption) > 7 Then
                                  data_temp.Recordset("nrofactref") = Val(Mid(labnrofactnc.Caption, 1, 7))
                               Else
                                  data_temp.Recordset("nrofactref") = Val(labnrofactnc.Caption)
                               End If
                               data_temp.Recordset("fechafact") = CDate(labfeccance.Caption)
                               data_temp.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                               data_temp.Recordset("linearef") = Val(labnrolin.Caption)
                               data_temp.Recordset("linea") = Xcandelin
    '                              data_temp.Recordset("unidad") = labserie.Caption
                               data_temp.Recordset("libro_rub") = Label3.Caption
                               data_temp.Recordset("in_unid") = "INT1"
                               data_temp.Recordset("in_mat") = 2
                               If cboiva.ListIndex = 0 Then
                                  data_temp.Recordset("indfact") = 2
                               Else
                                  If cboiva.ListIndex = 1 Then
                                     data_temp.Recordset("indfact") = 3
                                  Else
                                     If cboiva.ListIndex = 2 Then
                                        data_temp.Recordset("indfact") = 1
                                     Else
                                        data_temp.Recordset("indfact") = 2
                                     End If
                                  End If
                               End If
                               If t_cant.Text = "" Then
                                  t_cant.Text = 1
                                  data_temp.Recordset("cantidad") = 1
                               Else
                                  data_temp.Recordset("cantidad") = t_cant.Text
                               End If
                               If t_pie.Text <> "" Then
                                  data_temp.Recordset("in_obs") = t_pie.Text
                               End If
                               data_temp.Recordset("factura") = 0
                               data_temp.Recordset("tipo") = labtipof.Caption
                               data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                               data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
                               data_temp.Recordset("cod_cli") = frm_convenios.txt_cuenta.Text
                               data_temp.Recordset("nom_cli") = labcli.Caption
                               data_temp.Recordset("convenio") = frm_convenios.txt_cod.Text
                               If labdom.Caption <> "" Then
                                  data_temp.Recordset("nom_flia") = Mid(labdom.Caption, 1, 40)
                               End If
                               If labnrocli.Caption <> "" Then
                                  data_temp.Recordset("nro_superv") = labnrocli.Caption
                               End If
                               If lablocal.Caption <> "" Then
                                  data_temp.Recordset("nom_superv") = Mid(lablocal.Caption, 1, 25)
                               End If
                               If Combo1.ListIndex = 0 Then
                                  data_temp.Recordset("cod_prod") = 30
                                  data_temp.Recordset("in_usuario") = Trim(str(30))
                               Else
                                  If Combo1.ListIndex = 1 Then
                                     data_temp.Recordset("cod_prod") = 31
                                     data_temp.Recordset("in_usuario") = Trim(str(31))
                                  Else
                                     If Combo1.ListIndex = 2 Then
                                        data_temp.Recordset("cod_prod") = 32
                                        data_temp.Recordset("in_usuario") = Trim(str(32))
                                     Else
                                        data_temp.Recordset("cod_prod") = 33
                                        data_temp.Recordset("in_usuario") = Trim(str(30))
                                     End If
                                  End If
                               End If
                               data_temp.Recordset("nom_prod") = t_desc.Text
                               data_temp.Recordset("solicitant") = Combo1.Text
                               data_temp.Recordset("costo_prod") = cboiva.ListIndex
                               data_temp.Recordset("operador") = WElusuario
                               data_temp.Recordset("hora") = Format(Time, "HH:mm")
                               data_temp.Recordset("imp_timbre") = t_imp.Text
                               data_temp.Recordset("tot_lin") = t_imp.Text * t_cant.Text
                               labmontoit.Caption = t_imp.Text * t_cant.Text
                               labmontoit.Caption = Format(labmontoit.Caption, "Standard")
                               data_temp.Recordset("mes_paga") = t_mes.Text
                               data_temp.Recordset("ano_paga") = t_ano.Text
                               data_temp.Recordset("base") = Xbb
                               data_temp.Recordset("ruc") = frm_convenios.txt_ruc.Text
                               If frm_convenios.cbomut.Text = "UNIVERSAL" Or _
                                  frm_convenios.cbomut.Text = "CCOU" Or _
                                  frm_convenios.cbomut.Text = "SMI" Or _
                                  frm_convenios.cbomut.Text = "IMPASA" Or _
                                  frm_convenios.cbomut.Text = "CASA DE GALICIA" Or _
                                  frm_convenios.cbomut.Text = "H.EVANGELICO" Then
                                  data_temp.Recordset("rub_cont") = 511030
                               Else
                                  data_temp.Recordset("rub_cont") = 513008
                               End If
                               If Check1.Value = 1 Then
                                  data_temp.Recordset("tipo_mov") = "X"
                                  data_temp.Recordset("ruc") = ""
                               Else
                                  data_temp.Recordset("tipo_mov") = ""
                               End If
                               If cboiva.ListIndex = 0 Then
                                  Xsubt = Xsubt + CDbl(labmontoit.Caption)
                                  Xtot = Xsubt
                               Else
                                  If cboiva.ListIndex = 1 Then
                                     Xsubt22 = Xsubt22 + CDbl(labmontoit.Caption)
                                     Xtot = Xsubt22
                                  Else
                                     If cboiva.ListIndex = 2 Then
                                        Xsubt0 = Xsubt0 + CDbl(labmontoit.Caption)
                                        Xtot = Xsubt0
                                     Else
                                        Xsubt = Xsubt + CDbl(labmontoit.Caption)
                                        Xtot = Xsubt
                                     End If
                                  End If
                               End If
                               If t_imp.Text > 0 Then
                                  If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                     Xivauno = CDbl(labmontoit.Caption) * 0.1
                                     Xivvva = Xsubt * 0.1
                                     Xtot = Xtot + Xivvva
                                  Else
                                     If cboiva.ListIndex = 1 Then
                                        Xivauno = CDbl(labmontoit.Caption) * 0.22
                                        Xivvva = Xsubt22 * 0.22
                                        Xtot = Xtot + Xivvva
                                     Else
                                        If cboiva.ListIndex = 2 Then
                                           Xivauno = 0
                                           Xivvva = 0
                                           Xtot = Xtot + Xivvva
                                        Else
                                           Xivauno = CDbl(labmontoit.Caption) * 0.1
                                           Xivvva = Xsubt * 0.1
                                           Xtot = Xtot + Xivvva
                                        End If
                                     End If
                                  End If
                               Else
                                  Xivvva = 0
                               End If
                               If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                  data_temp.Recordset("tipo_mov") = Trim(str("2"))
                                  data_temp.Recordset("pre_civa") = Xivauno
                                  data_temp.Recordset("iva22") = 0
                               Else
                                  If cboiva.ListIndex = 1 Then
                                     data_temp.Recordset("tipo_mov") = Trim(str("3"))
                                     data_temp.Recordset("pre_civa") = 0
                                     data_temp.Recordset("iva22") = Xivauno
                                  Else
                                     data_temp.Recordset("tipo_mov") = Trim(str("1"))
                                     data_temp.Recordset("pre_civa") = 0
                                     data_temp.Recordset("iva22") = 0
                                  End If
                               End If
    '                           data_temp.Recordset("pre_prod") = Val(Xlafac) 'que era ?
                               data_temp.Recordset("grupo") = data_cli.Recordset("cl_nrocobr")
                               data_temp.Recordset("nom_med_s") = Mid(labcalc.Caption, 1, 40)
                               'If t_desc.Text <> "" Then
                                  If t_desc2.Text <> "" Then
                                     If t_desc3.Text <> "" Then
                                        data_temp.Recordset("obsp") = t_desc2.Text & " " & t_desc3.Text
                                     Else
                                        data_temp.Recordset("obsp") = t_desc2.Text
                                     End If
                                  Else
                                     If t_mes.Text <> "" Then
                                        data_temp.Recordset("obsp") = Trim(t_mes.Text) & "/" & t_ano.Text
                                     Else
                                        data_temp.Recordset("obsp") = Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
                                     End If
                                  End If
                               'Else
                               '   data_temp.Recordset("obsp") = Combo1.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text))
                               'End If
                               If cbomon.ListIndex = 1 Then
                                  data_temp.Recordset("libro") = "D"
                               Else
                                  data_temp.Recordset("libro") = "U"
                               End If
                               data_temp.Recordset.Update
                               labtot.Caption = Xtot
                               labtot.Caption = Format(labtot.Caption, "Standard")
                               If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                  labstot.Caption = Xsubt
                                  labstot.Caption = Format(Xsubt, "Standard")
                               Else
                                  If cboiva.ListIndex = 1 Then
                                     labstot22.Caption = Xsubt22
                                     labstot22.Caption = Format(Xsubt22, "Standard")
                                  Else
                                     If cboiva.ListIndex = 2 Then
                                        labstot0.Caption = Xsubt0
                                        labstot0.Caption = Format(Xsubt0, "Standard")
                                     Else
                                        labstot.Caption = Xsubt
                                        labstot.Caption = Format(Xsubt, "Standard")
                                     End If
                                  End If
                               End If
                               If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                  labiva.Caption = Format(Xivvva, "Standard")
                               Else
                                  If cboiva.ListIndex = 1 Then
                                     labiva22.Caption = Format(Xivvva, "Standard")
                                  End If
                               End If
                            Else
                               If Label3.Caption = "NC de E-FACTURA" Then
                                  data_lin.RecordSource = "Select * from clirespl where cl_numero =" & Val(labnrofactnc.Caption) & " and cl_socmnro ='" & Trim(labseriefactnc.Caption) & "' and cl_tipocli =" & 111
                               Else
                                  data_lin.RecordSource = "Select * from clirespl where cl_numero =" & Val(labnrofactnc.Caption) & " and cl_socmnro ='" & Trim(labseriefactnc.Caption) & "' and cl_tipocli =" & 101
                               End If
                               data_lin.Refresh
                               If data_lin.Recordset.RecordCount > 0 Then
                                  If Val(labtot.Caption) >= data_lin.Recordset("saldo_doc") Then
                                     MsgBox "El importe no puede superar el total de la factura", vbCritical
                                     Command4_Click
                                  Else
                                      Xcandelin = Xcandelin + 1
                                      data_temp.Recordset.AddNew
                                      data_temp.Recordset("tipodocref") = data_lin.Recordset("cl_tipocli")
                                      data_temp.Recordset("serieref") = labseriefactnc.Caption
                                      data_temp.Recordset("nrofactref") = Val(labnrofactnc.Caption)
                                      data_temp.Recordset("fechafact") = data_lin.Recordset("cl_fnac")
                                      data_temp.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                                      data_temp.Recordset("linearef") = Val(labnrolin.Caption)
                                      data_temp.Recordset("linea") = Xcandelin
        '                              data_temp.Recordset("unidad") = labserie.Caption
                                       data_temp.Recordset("libro_rub") = Label3.Caption
                                       data_temp.Recordset("in_unid") = "INT1"
                                       data_temp.Recordset("in_mat") = 2
                                       If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                          data_temp.Recordset("indfact") = 2
                                       Else
                                          If cboiva.ListIndex = 1 Then
                                             data_temp.Recordset("indfact") = 3
                                          Else
                                             If cboiva.ListIndex = 2 Then
                                                data_temp.Recordset("indfact") = 1
                                             Else
                                                data_temp.Recordset("indfact") = 2
                                             End If
                                          End If
                                       End If
                                       If t_cant.Text = "" Then
                                          t_cant.Text = 1
                                          data_temp.Recordset("cantidad") = 1
                                       Else
                                          data_temp.Recordset("cantidad") = t_cant.Text
                                       End If
                                       If t_pie.Text <> "" Then
                                          data_temp.Recordset("in_obs") = t_pie.Text
                                       End If
                                       data_temp.Recordset("factura") = 0
                                       data_temp.Recordset("tipo") = labtipof.Caption
                                       data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                                       data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
                                       data_temp.Recordset("cod_cli") = frm_convenios.txt_cuenta.Text
                                       data_temp.Recordset("nom_cli") = labcli.Caption
                                       data_temp.Recordset("convenio") = frm_convenios.txt_cod.Text
                                       If labdom.Caption <> "" Then
                                          data_temp.Recordset("nom_flia") = Mid(labdom.Caption, 1, 40)
                                       End If
                                       If labnrocli.Caption <> "" Then
                                          data_temp.Recordset("nro_superv") = labnrocli.Caption
                                       End If
                                       If lablocal.Caption <> "" Then
                                          data_temp.Recordset("nom_superv") = Mid(lablocal.Caption, 1, 25)
                                       End If
             '                          data_temp.Recordset("vence") = Format(labvence.Text, "dd/mm/yyyy")
                                       If t_desc2.Text <> "" Then
                                          data_temp.Recordset("nom_medic") = Mid(t_desc2.Text, 1, 50)
                                       End If
                                       If t_desc3.Text <> "" Then
                                          data_temp.Recordset("nom_med_a") = Mid(t_desc3.Text, 1, 40)
                                       End If
                                       If Combo1.ListIndex = 0 Then
                                          data_temp.Recordset("cod_prod") = 30
                                          data_temp.Recordset("in_usuario") = Trim(str(30))
                                       Else
                                          If Combo1.ListIndex = 1 Then
                                             data_temp.Recordset("cod_prod") = 31
                                             data_temp.Recordset("in_usuario") = Trim(str(31))
                                          Else
                                             If Combo1.ListIndex = 2 Then
                                                data_temp.Recordset("cod_prod") = 32
                                                data_temp.Recordset("in_usuario") = Trim(str(32))
                                             Else
                                                data_temp.Recordset("cod_prod") = 33
                                                data_temp.Recordset("in_usuario") = Trim(str(30))
                                             End If
                                          End If
                                       End If
                                       data_temp.Recordset("nom_prod") = t_desc.Text
                                       data_temp.Recordset("solicitant") = Combo1.Text
                                       data_temp.Recordset("costo_prod") = cboiva.ListIndex
                                       data_temp.Recordset("operador") = WElusuario
                                       data_temp.Recordset("hora") = Format(Time, "HH:mm")
                                       data_temp.Recordset("imp_timbre") = t_imp.Text
                                       data_temp.Recordset("tot_lin") = t_imp.Text * t_cant.Text
                                       labmontoit.Caption = t_imp.Text * t_cant.Text
                                       labmontoit.Caption = Format(labmontoit.Caption, "Standard")
                                       data_temp.Recordset("mes_paga") = t_mes.Text
                                       data_temp.Recordset("ano_paga") = t_ano.Text
                                       data_temp.Recordset("base") = Xbb
                                       data_temp.Recordset("ruc") = frm_convenios.txt_ruc.Text
                                       If frm_convenios.cbomut.Text = "UNIVERSAL" Or _
                                          frm_convenios.cbomut.Text = "CCOU" Or _
                                          frm_convenios.cbomut.Text = "SMI" Or _
                                          frm_convenios.cbomut.Text = "IMPASA" Or _
                                          frm_convenios.cbomut.Text = "CASA DE GALICIA" Or _
                                          frm_convenios.cbomut.Text = "H.EVANGELICO" Then
                                          data_temp.Recordset("rub_cont") = 511030
                                       Else
                                          data_temp.Recordset("rub_cont") = 513008
                                       End If
                                       If Check1.Value = 1 Then
                                          data_temp.Recordset("tipo_mov") = "X"
                                          data_temp.Recordset("ruc") = ""
                                       Else
                                          data_temp.Recordset("tipo_mov") = ""
                                       End If
                                       If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                          Xsubt = Xsubt + CDbl(labmontoit.Caption)
                                          Xtot = Xsubt
                                       Else
                                          If cboiva.ListIndex = 1 Then
                                             Xsubt22 = Xsubt22 + CDbl(labmontoit.Caption)
                                             Xtot = Xsubt22
                                          Else
                                             If cboiva.ListIndex = 2 Then
                                                Xsubt0 = Xsubt0 + CDbl(labmontoit.Caption)
                                                Xtot = Xsubt0
                                             Else
                                                Xsubt = Xsubt + CDbl(labmontoit.Caption)
                                                Xtot = Xsubt
                                             End If
                                          End If
                                       End If
                                       If t_imp.Text > 0 Then
                                          If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                             Xivauno = CDbl(labmontoit.Caption) * 0.1
                                             Xivvva = Xsubt * 0.1
                                             Xtot = Xtot + Xivvva
                                          Else
                                             If cboiva.ListIndex = 1 Then
                                                Xivauno = CDbl(labmontoit.Caption) * 0.22
                                                Xivvva = Xsubt22 * 0.22
                                                Xtot = Xtot + Xivvva
                                             Else
                                                If cboiva.ListIndex = 2 Then
                                                   Xivauno = 0
                                                   Xivvva = 0
                                                   Xtot = Xtot + Xivvva
                                                Else
                                                   Xivauno = CDbl(labmontoit.Caption) * 0.1
                                                   Xivvva = Xsubt * 0.1
                                                   Xtot = Xtot + Xivvva
                                                End If
                                             End If
                                          End If
                                       Else
                                          Xivvva = 0
                                       End If
                                       If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                          data_temp.Recordset("tipo_mov") = Trim(str("2"))
                                          data_temp.Recordset("pre_civa") = Xivauno
                                          data_temp.Recordset("iva22") = 0
                                       Else
                                          If cboiva.ListIndex = 1 Then
                                             data_temp.Recordset("tipo_mov") = Trim(str("3"))
                                             data_temp.Recordset("pre_civa") = 0
                                             data_temp.Recordset("iva22") = Xivauno
                                          Else
                                             data_temp.Recordset("tipo_mov") = Trim(str("5"))
                                             data_temp.Recordset("pre_civa") = 0
                                             data_temp.Recordset("iva22") = 0
                                          End If
                                       End If
            '                           data_temp.Recordset("pre_prod") = Val(Xlafac) 'que era ?
                                       data_temp.Recordset("grupo") = data_cli.Recordset("cl_nrocobr")
                                       data_temp.Recordset("nom_med_s") = Mid(labcalc.Caption, 1, 40)
                                       If t_desc.Text <> "" Then
                                          If t_desc2.Text <> "" Then
                                             If t_desc3.Text <> "" Then
                                                data_temp.Recordset("obsp") = t_desc.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text)) & " " & t_desc2.Text & " " & t_desc3.Text
                                             Else
                                                data_temp.Recordset("obsp") = t_desc.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text)) & " " & t_desc2.Text
                                             End If
                                          Else
                                             data_temp.Recordset("obsp") = t_desc.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text))
                                          End If
                                       Else
                                          data_temp.Recordset("obsp") = Combo1.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text))
                                       End If
                                       If cbomon.ListIndex = 1 Then
                                          data_temp.Recordset("libro") = "D"
                                       Else
                                          data_temp.Recordset("libro") = "U"
                                       End If
                                       data_temp.Recordset.Update
                                       labtot.Caption = Xtot
                                       labtot.Caption = Format(labtot.Caption, "Standard")
                                       If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                          labstot.Caption = Xsubt
                                          labstot.Caption = Format(Xsubt, "Standard")
                                       Else
                                          If cboiva.ListIndex = 1 Then
                                             labstot22.Caption = Xsubt22
                                             labstot22.Caption = Format(Xsubt22, "Standard")
                                          Else
                                             If cboiva.ListIndex = 2 Then
                                                labstot0.Caption = Xsubt0
                                                labstot0.Caption = Format(Xsubt0, "Standard")
                                             Else
                                                labstot.Caption = Xsubt
                                                labstot.Caption = Format(Xsubt, "Standard")
                                             End If
                                          End If
                                       End If
                                       If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                          labiva.Caption = Format(Xivvva, "Standard")
                                       Else
                                          If cboiva.ListIndex = 1 Then
                                             labiva22.Caption = Format(Xivvva, "Standard")
                                          End If
                                       End If
                                  End If
                               Else
                                  MsgBox "No se cuentra cabezal de factura"
                               End If
                            End If
                         Else
                            MsgBox "Falta ingresar motivo", vbInformation
                         End If
                      Else
                         Xcandelin = Xcandelin + 1
                         data_temp.Recordset.AddNew
                         data_temp.Recordset("linea") = Xcandelin
                         data_temp.Recordset("libro_rub") = Label3.Caption
                         data_temp.Recordset("in_unid") = "INT1"
                         data_temp.Recordset("in_mat") = 2
                         If cboiva.ListIndex = 0 Then
                            data_temp.Recordset("indfact") = 2
                         Else
                            If cboiva.ListIndex = 1 Then
                               data_temp.Recordset("indfact") = 3
                            Else
                               If cboiva.ListIndex = 2 Then
                                  data_temp.Recordset("indfact") = 1
                               Else
                                  data_temp.Recordset("indfact") = 2
                               End If
                            End If
                         End If
                         If t_imp.Text = 0 Then
                            data_temp.Recordset("indfact") = 5
                         End If
                         If t_cant.Text = "" Then
                            t_cant.Text = 1
                            data_temp.Recordset("cantidad") = 1
                         Else
                            data_temp.Recordset("cantidad") = t_cant.Text
                         End If
                         If t_pie.Text <> "" Then
                            data_temp.Recordset("in_obs") = t_pie.Text
                         End If
                         data_temp.Recordset("tipo") = labtipof.Caption
                         data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                         data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
                         data_temp.Recordset("cod_cli") = frm_convenios.txt_cuenta.Text
                         data_temp.Recordset("nom_cli") = labcli.Caption
                         data_temp.Recordset("convenio") = frm_convenios.txt_cod.Text
                         If labdom.Caption <> "" Then
                            data_temp.Recordset("nom_flia") = Mid(labdom.Caption, 1, 40)
                         End If
                         If labnrocli.Caption <> "" Then
                            data_temp.Recordset("nro_superv") = labnrocli.Caption
                         End If
                         If lablocal.Caption <> "" Then
                            data_temp.Recordset("nom_superv") = Mid(lablocal.Caption, 1, 25)
                         End If
                         If Combo1.ListIndex = 0 Then
                            data_temp.Recordset("cod_prod") = 30
                            data_temp.Recordset("in_usuario") = Trim(str(30))
                         Else
                            If Combo1.ListIndex = 1 Then
                               data_temp.Recordset("cod_prod") = 31
                               data_temp.Recordset("in_usuario") = Trim(str(31))
                            Else
                               If Combo1.ListIndex = 2 Then
                                  data_temp.Recordset("cod_prod") = 32
                                  data_temp.Recordset("in_usuario") = Trim(str(32))
                               Else
                                  data_temp.Recordset("cod_prod") = 33
                                  data_temp.Recordset("in_usuario") = Trim(str(30))
                               End If
                            End If
                         End If
                         data_temp.Recordset("nom_prod") = t_desc.Text
                         data_temp.Recordset("solicitant") = Combo1.Text
                         data_temp.Recordset("costo_prod") = cboiva.ListIndex
                         data_temp.Recordset("operador") = WElusuario
                         data_temp.Recordset("hora") = Format(Time, "HH:mm")
                         data_temp.Recordset("imp_timbre") = t_imp.Text
                         data_temp.Recordset("tot_lin") = t_imp.Text * t_cant.Text
                         labmontoit.Caption = t_imp.Text * t_cant.Text
                         labmontoit.Caption = Format(labmontoit.Caption, "Standard")
                         data_temp.Recordset("mes_paga") = t_mes.Text
                         data_temp.Recordset("ano_paga") = t_ano.Text
                         data_temp.Recordset("base") = Xbb
                         data_temp.Recordset("ruc") = frm_convenios.txt_ruc.Text
                         If frm_convenios.cbomut.Text = "UNIVERSAL" Or _
                            frm_convenios.cbomut.Text = "CCOU" Or _
                            frm_convenios.cbomut.Text = "SMI" Or _
                            frm_convenios.cbomut.Text = "IMPASA" Or _
                            frm_convenios.cbomut.Text = "CASA DE GALICIA" Or _
                            frm_convenios.cbomut.Text = "H.EVANGELICO" Then
                            data_temp.Recordset("rub_cont") = 511030
                         Else
                            data_temp.Recordset("rub_cont") = 513008
                         End If
                         If Check1.Value = 1 Then
                            data_temp.Recordset("tipo_mov") = "X"
                            data_temp.Recordset("ruc") = ""
                         Else
                            data_temp.Recordset("tipo_mov") = ""
                         End If
                         data_temp.Recordset("grupo") = data_cli.Recordset("cl_nrocobr")
                         data_temp.Recordset("nom_med_s") = Mid(labcalc.Caption, 1, 40)
                         If t_imp.Text = 0 Then
                         Else
                            If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                               Xsubt = Xsubt + CDbl(labmontoit.Caption)
                               Xtot = Xsubt
                            Else
                               If cboiva.ListIndex = 1 Then
                                  Xsubt22 = Xsubt22 + CDbl(labmontoit.Caption)
                                  Xtot = Xsubt22
                               Else
                                  If cboiva.ListIndex = 2 Then
                                     Xsubt0 = Xsubt0 + CDbl(labmontoit.Caption)
                                     Xtot = Xsubt0
                                  Else
                                     Xsubt = Xsubt + CDbl(labmontoit.Caption)
                                     Xtot = Xsubt
                                  End If
                               End If
                            End If
                            If t_imp.Text > 0 Then
                               If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                                  Xivauno = CDbl(labmontoit.Caption) * 0.1
                                  Xivvva = Xsubt * 0.1
                                  Xtot = Xtot + Xivvva
                               Else
                                  If cboiva.ListIndex = 1 Then
                                     Xivauno = CDbl(labmontoit.Caption) * 0.22
                                     Xivvva = Xsubt22 * 0.22
                                     Xtot = Xtot + Xivvva
                                  Else
                                     If cboiva.ListIndex = 2 Then
                                        Xivauno = 0
                                        Xivvva = 0
                                        Xtot = Xtot + Xivvva
                                     Else
                                        Xivauno = CDbl(labmontoit.Caption) * 0.1
                                        Xivvva = Xsubt * 0.1
                                        Xtot = Xtot + Xivvva
                                     End If
                                  End If
                               End If
                            Else
                               Xivvva = 0
                            End If
                         End If
                         If t_imp.Text = 0 Then
                         Else
                            If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                               data_temp.Recordset("tipo_mov") = Trim(str("2"))
                               data_temp.Recordset("pre_civa") = Xivauno
                               data_temp.Recordset("iva22") = 0
                            Else
                               If cboiva.ListIndex = 1 Then
                                  data_temp.Recordset("tipo_mov") = Trim(str("3"))
                                  data_temp.Recordset("pre_civa") = 0
                                  data_temp.Recordset("iva22") = Xivauno
                               Else
                                  data_temp.Recordset("tipo_mov") = Trim(str("5"))
                                  data_temp.Recordset("pre_civa") = 0
                                  data_temp.Recordset("iva22") = 0
                               End If
                            End If
                         End If
                         
                         If t_desc.Text <> "" Then
                            If t_desc2.Text <> "" Then
                               If t_desc3.Text <> "" Then
                                  data_temp.Recordset("obsp") = t_desc.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text)) & " " & t_desc2.Text & " " & t_desc3.Text
                               Else
                                  data_temp.Recordset("obsp") = t_desc.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text)) & " " & t_desc2.Text
                               End If
                            Else
                               data_temp.Recordset("obsp") = t_desc.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text))
                            End If
                         Else
                            data_temp.Recordset("obsp") = Combo1.Text & " " & Trim(str(t_mes.Text)) & "/" & Trim(str(t_ano.Text))
                         End If
                         If cbomon.ListIndex = 1 Then
                            data_temp.Recordset("libro") = "D"
                         Else
                            data_temp.Recordset("libro") = "U"
                         End If
                         data_temp.Recordset.Update
                         mf.Enabled = False
                         If t_imp.Text = 0 Then
                         Else
                            labtot.Caption = Xtot
                            labtot.Caption = Format(labtot.Caption, "Standard")
                            If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                               labstot.Caption = Xsubt
                               labstot.Caption = Format(Xsubt, "Standard")
                            Else
                               If cboiva.ListIndex = 1 Then
                                  labstot22.Caption = Xsubt22
                                  labstot22.Caption = Format(Xsubt22, "Standard")
                               Else
                                  If cboiva.ListIndex = 2 Then
                                     labstot0.Caption = Xsubt0
                                     labstot0.Caption = Format(Xsubt0, "Standard")
                                  Else
                                     labstot.Caption = Xsubt
                                     labstot.Caption = Format(Xsubt, "Standard")
                                  End If
                               End If
                            End If
                            If cboiva.ListIndex = 0 Or cboiva.ListIndex = 3 Then
                               labiva.Caption = Format(Xivvva, "Standard")
                            Else
                               If cboiva.ListIndex = 1 Then
                                  labiva22.Caption = Format(Xivvva, "Standard")
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
            End If
        
        End If
'      Else
'        MsgBox "Año inválido, verifique!"
'      End If
    Else
        MsgBox "Mes inválido, verifique!!"
    End If
Else
    MsgBox "Mes/Año/Cantidad no puede estar vacío, o verifique Servicio a facturar"
End If

Exit Sub

Algrabar:
              If Err.Number = 3155 Then
                 MsgBox "Error al terminar la factura. AVISE A INFORMATICA : " & Err.Number & " " & Err.Description, vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = Err.Number
                 data_errfact.Recordset("desc") = "Al grabar"
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
                    data_errfact.Recordset("desc") = "Al grabar"
                    data_errfact.Recordset.Update
                    Unload Me
                 Else
                    MsgBox "Error al terminar la factura. AVISE A INFORMATICA! " & Err.Number & " " & Err.Description, vbInformation
                    data_errfact.Recordset.AddNew
                    data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                    data_errfact.Recordset("fecha") = Date
                    data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                    data_errfact.Recordset("nroerr") = Err.Number
                    data_errfact.Recordset("desc") = "Al grabar"
                    data_errfact.Recordset.Update
                    Unload Me
                 End If
              End If

   
End Sub

Private Sub b_vernc_Click()
On Error GoTo ErrorCance

frm_factcancela.Show vbModal
Exit Sub

ErrorCance:
            MsgBox ("Error al cancelar")
            
End Sub

Private Sub cboiva_KeyPress(KeyAscii As Integer)
On Error GoTo Errorcboiva

If KeyAscii = 13 Then
   b_graba.SetFocus
End If

Exit Sub

Errorcboiva:
            MsgBox ("Error al cboIVA")
            
End Sub

Private Sub Combo1_Click()
On Error GoTo Alcombo1clic

If Combo1.ListIndex >= 0 And Combo1.ListIndex <= 1 Then
   data_lin.RecordSource = "Select * from linmmdd where solicitant ='" & Combo1.Text & "' and cod_cli =" & data_cli.Recordset("cl_codigo") & " and fecha <#" & Format(mf.Text, "yyyy/mm/dd") & "# order by fecha"
   data_lin.Refresh
   If data_lin.Recordset.RecordCount > 0 Then
      data_lin.Recordset.MoveLast
      If IsNull(data_lin.Recordset("nom_prod")) = False Then
         t_desc.Text = Combo1.Text & " " & data_lin.Recordset("nom_prod") & Month(Date) & "/" & Year(Date)
      Else
         t_desc.Text = "Servicios " & Month(Date) & "/" & Year(Date)
      End If
      If IsNull(data_lin.Recordset("nom_medic")) = False Then
         t_desc2.Text = data_lin.Recordset("nom_medic")
      Else
         t_desc2.Text = ""
      End If
      If IsNull(data_lin.Recordset("nom_med_a")) = False Then
         t_desc3.Text = data_lin.Recordset("nom_med_a")
      Else
         t_desc3.Text = ""
      End If
      t_imp.Text = Format(data_lin.Recordset("imp_timbre"), "Standard")
      If IsNull(data_lin.Recordset("nom_med_s")) = False Then
         labcalc.Caption = data_lin.Recordset("nom_med_s")
      Else
         labcalc.Caption = ""
      End If
      If IsNull(data_lin.Recordset("costo_prod")) = False Then
         If cboiva.Enabled = True Then
            cboiva.ListIndex = data_lin.Recordset("costo_prod")
         End If
      End If
      If IsNull(data_lin.Recordset("mes_paga")) = False Then
         If data_lin.Recordset("mes_paga") > 0 Then
            If data_lin.Recordset("mes_paga") = 12 Then
               t_mes.Text = 1
               t_ano.Text = data_lin.Recordset("ano_paga") + 1
            Else
               t_mes.Text = data_lin.Recordset("mes_paga") + 1
               t_ano.Text = data_lin.Recordset("ano_paga")
            End If
         End If
      End If
   Else
      labcalc.Caption = ""
      t_desc.Text = Combo1.Text & " " & Month(Date) & "/" & Year(Date)
   End If
Else
   If Combo1.Text = "COSTO POR LLAMADOS" Then
      If Trim(labimptiquet.Caption) <> "" Then
         t_desc.Text = "Llamados fuera de tope contrato"
         t_mes.Text = Month(Date)
         t_ano.Text = Year(Date)
         t_imp.Text = Val(labimptiquet.Caption)
         t_cant.Text = Val(labcanttiquet.Caption)
      End If
   End If
End If

Exit Sub

Alcombo1clic:
             MsgBox "Error al terminar la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al combo1 click"
             data_errfact.Recordset.Update
             Unload Me


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo combo1Err
If KeyAscii = 13 Then
   t_mes.SetFocus
End If

Exit Sub

combo1Err:
            MsgBox ("Error al combo1Err")
            
End Sub

Private Sub Command1_Click()
On Error GoTo Comman1document

Dim oDom As New MSXML2.DOMDocument30
Dim Xelrut As Double
'Dim xmlsc As FileSystemObject
 
Exit Sub

Comman1document:
                MsgBox ("Error al comman1Document")
                
  
End Sub




Private Sub Command2_Click()
On Error GoTo Alcoman2cli

If mf.Enabled = True Then
   mf.SetFocus
End If

Combo1.ListIndex = -1
t_imp.Text = 0
t_mes.Text = 0
t_ano.Text = 0
cboiva.ListIndex = 0
labcalc.Caption = ""
t_desc.Text = ""
t_desc2.Text = ""
t_desc3.Text = ""
Check1.Value = 0

Exit Sub

Alcoman2cli:
             MsgBox "Error al terminar la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al coman2click click"
             data_errfact.Recordset.Update
             Unload Me

End Sub

Private Sub Command3_Click()
Dim Xlatasa, Xlatasa22, Xmasdiezui, Xlaui As Double
Dim Xveo As Integer

On Error GoTo Verquepa

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")
Xveo = 0
If data_temp.Recordset.RecordCount > 0 Then
   data_lin.RecordSource = "Select * from linmmdd where factura =" & 223595
   data_lin.Refresh
   data_temp.Recordset.MoveFirst
      data_cabeza2.Recordset.AddNew
 '           data_cabezal.Recordset("id") = 1
      data_cabeza2.Recordset("cl_tipcli") = "1.0"
      If Label3.Caption = "E-FACTURA" Then
         data_cabeza2.Recordset("cl_tipocli") = 111
         data_cabeza2.Recordset("cl_telefon") = "e-Factura"
      Else
         If Label3.Caption = "E-TICKET" Then
            data_cabeza2.Recordset("cl_tipocli") = 101
            data_cabeza2.Recordset("cl_telefon") = "e-Ticket"
         Else
            If Label3.Caption = "NC de E-FACTURA" Then
               data_cabeza2.Recordset("cl_tipocli") = 112
               data_cabeza2.Recordset("cl_telefon") = "NC e-Factura"
            Else
               If Label3.Caption = "NC de E-TICKET" Then
                  data_cabeza2.Recordset("cl_tipocli") = 102
                  data_cabeza2.Recordset("cl_telefon") = "NC e-Ticket"
               Else
                  If Label3.Caption = "ND de E-FACTURA" Then
                     data_cabeza2.Recordset("cl_tipocli") = 113
                     data_cabeza2.Recordset("cl_telefon") = "ND e-Factura"
                  Else
                     If Label3.Caption = "ND de E-TICKET" Then
                        data_cabeza2.Recordset("cl_tipocli") = 103
                        data_cabeza2.Recordset("cl_telefon") = "ND e-Ticket"
                     Else
                        data_cabeza2.Recordset("cl_tipocli") = 111
                     End If
                  End If
               End If
            End If
         End If
      End If
'      data_cabezal.Recordset("cl_socmnro") = labserie.Caption
'      data_cabezal.Recordset("cl_numero") = Val(labnrofac.Caption)
      data_cabeza2.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
      Xveo = 1
      If mfd.Text <> "__/__/____" Then
         data_cabeza2.Recordset("fecha_reac") = Format(mfd.Text, "dd/mm/yyyy")
      End If
      If mfh.Text <> "__/__/____" Then
         data_cabeza2.Recordset("cl_tj_venc") = Format(mfh.Text, "dd/mm/yyyy")
      End If
      data_cabeza2.Recordset("cl_nrovend") = 0
      If labtipof.Caption = "CONTADO" Then
         data_cabeza2.Recordset("cl_forpago") = 1
      Else
         If labtipof.Caption = "CREDITO" Then
            data_cabeza2.Recordset("cl_forpago") = 2
         Else
            data_cabeza2.Recordset("cl_forpago") = 2
        End If
      End If
      data_cabeza2.Recordset("cl_celular") = labtipof.Caption
      Xveo = 2
      If labvence.Text <> "__/__/____" Then
         data_cabeza2.Recordset("fecha_modi") = Format(labvence.Text, "dd/mm/yyyy")
      End If
      data_cabeza2.Recordset("cl_diacobr") = Trim(str(data_par.Recordset("ruc")))
      data_cabeza2.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
      data_cabeza2.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
      data_cabeza2.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
      data_cabeza2.Recordset("cl_referen") = data_par.Recordset("domic")
      data_cabeza2.Recordset("tit_tarj") = data_par.Recordset("ciudad")
      data_cabeza2.Recordset("cl_nomconv") = data_par.Recordset("dpto")
      Xveo = 3
    'receptor
'      If Label3.Caption = "E-TICKET" Then
      data_cabeza2.Recordset("cl_nro_sup") = Xtipodedocumento
'      Else
'         data_cabeza2.Recordset("cl_nro_sup") = 2
'      End If
      data_cabeza2.Recordset("hora_baja") = "UY"
      If Xtipodedocumento = 2 Then
         data_cabeza2.Recordset("cl_nom_sup") = frm_convenios.txt_ruc.Text
      Else
         If Xtipodedocumento = 3 Then
            data_cabeza2.Recordset("cl_nom_sup") = Trim(str(data_cli.Recordset("cl_cedula"))) & Trim(str(data_cli.Recordset("cl_codced")))
         Else
            data_cabeza2.Recordset("cl_nom_sup") = Trim(str(data_cli.Recordset("cl_codigo")))
         End If
      End If
      'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
      Xveo = 4
      data_cabeza2.Recordset("info_debit") = frm_convenios.t_razon.Text
      data_cabeza2.Recordset("cl_direcci") = frm_convenios.txt_direc.Text
      Xveo = 5
      If frm_convenios.txt_localid.Text <> "" Then
         data_cabeza2.Recordset("cl_zona") = frm_convenios.txt_localid.Text
      End If
      If frm_convenios.t_dpto.Text <> "" Then
         data_cabeza2.Recordset("cl_email") = frm_convenios.t_dpto.Text 'opcional
      End If
      data_cabeza2.Recordset("cl_localid") = "URUGUAY" 'opcional
      data_cabeza2.Recordset("cl_codigo") = Val(frm_convenios.txt_cuenta.Text)
      Xveo = 6
      data_temp.Refresh
      data_temp.Recordset.MoveFirst
      
      If IsNull(data_temp.Recordset("libro")) = False Then
         If data_temp.Recordset("libro") = "D" Then
            data_cabeza2.Recordset("usu_baja") = "USD"
         Else
            data_cabeza2.Recordset("usu_baja") = "UYU"
         End If
      Else
         data_cabeza2.Recordset("usu_baja") = "UYU"
      End If
      data_cabeza2.Recordset("saldo_chc2") = Xeldolar
      data_cabeza2.Recordset("saldo_doc2") = Format(labstot.Caption, "Standard") 'subtot iva 10
      data_cabeza2.Recordset("cl_atrasoa") = Format(labstot22.Caption, "Standard") 'subtot iva 22
      data_cabeza2.Recordset("cl_cedula") = Format(labstot0.Caption, "Standard") 'subtot iva cero
      Xveo = 7
      data_cabeza2.Recordset("cl_atrasop") = Xlatasa
      data_cabeza2.Recordset("cl_decuota") = Xlatasa22
      data_cabeza2.Recordset("saldo_cc") = Format(labiva.Caption, "Standard") 'iva10
      data_cabeza2.Recordset("saldo_cc2") = Format(labiva22.Caption, "Standard") 'iva22
      Xveo = 8
      data_cabeza2.Recordset("saldo_doc") = Format(labtot.Caption, "Standard")
      data_cabeza2.Recordset("cl_grupo") = data_temp.Recordset.RecordCount
      data_cabeza2.Recordset("saldo_chc") = Format(labtot.Caption, "Standard")
      If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Or Label3.Caption = "ND de E-TICKET" Then
         data_cabeza2.Recordset("cl_cuopaga") = data_temp.Recordset("tipodocref")
         data_cabeza2.Recordset("codmotbaja") = data_temp.Recordset("serieref")
         data_cabeza2.Recordset("ultanopmut") = data_temp.Recordset("nrofactref")
         data_cabeza2.Recordset("cl_fultvta") = data_temp.Recordset("fechafact")
         data_cabeza2.Recordset("cl_entre") = data_temp.Recordset("motivoref")
      End If
      Xveo = 9
      If t_pie.Text <> "" Then
         If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Or Label3.Caption = "ND de E-TICKET" Then
            If Xlafac <> "" Then
               data_buscalafac.RecordSource = "Select * from linmmdd where factura =" & Val(Xlafac)
               data_buscalafac.Refresh
               If data_buscalafac.Recordset.RecordCount > 0 Then
                  data_cabeza2.Recordset("obsp") = t_pie.Text & chr(10) & chr(13) & "REFERENCIA A FACTURA NRO:" & Xlafac & " DE FECHA:" & data_buscalafac.Recordset("fecha")
               Else
                  data_cabeza2.Recordset("obsp") = t_pie.Text
               End If
            Else
               data_cabeza2.Recordset("obsp") = t_pie.Text
            End If
         Else
            data_cabeza2.Recordset("obsp") = t_pie.Text
         End If
      Else
         If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Or Label3.Caption = "ND de E-TICKET" Then
            If labnrofactnc.Caption <> "" Then
               data_buscalafac.RecordSource = "Select * from linmmdd where factura =" & labnrofactnc.Caption
               data_buscalafac.Refresh
               If data_buscalafac.Recordset.RecordCount > 0 Then
                  data_cabeza2.Recordset("obsp") = "REFERENCIA A FACTURA NRO:" & Xlafac & " DE FECHA:" & data_buscalafac.Recordset("fecha")
               End If
            End If
         Else
            data_cabeza2.Recordset("obsp") = "Cuentas SAPP: BROU 073-2052"
         End If
      End If
      If frm_convenios.t_nrocompra.Text <> "" Then
         data_cabeza2.Recordset("cl_nomcobr") = Trim(Mid(frm_convenios.t_nrocompra.Text, 1, 25))
      End If
      Xveo = 10
      data_cabeza2.Recordset.Update
      data_cabeza2.Refresh
      'fin de cabezal
      Command6_Click
'      Command1_Click
'      cr1.ReportFileName = App.Path & "\faccnvpru.rpt"
'      cr1.Action = 1
'      MsgBox "Proceso terminado"
'hasta aqui
   
Else
   MsgBox "No hay lineas en la factura"
   
End If
Unload Me

Exit Sub

Verquepa:
          If Err.Number = 3155 Then
             MsgBox "Error en la factura. AVISE A INFORMATICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al terminar la factura " & Trim(str(Xveo))
             data_errfact.Recordset.Update
          Else
             MsgBox "Error en la factura. AVISE A INFORMATICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al terminar la factura " & Trim(str(Xveo))
             data_errfact.Recordset.Update
          End If
          Unload Me

End Sub

Private Sub Command4_Click()
On Error GoTo Alcoman4clic

Combo1.ListIndex = -1
t_imp.Text = 0
t_mes.Text = 0
t_ano.Text = 0
cboiva.ListIndex = 0
labcalc.Caption = ""
t_desc.Text = ""
t_desc2.Text = ""
t_desc3.Text = ""
Check1.Value = 0
labmontoit.Caption = 0
labiva.Caption = 0
labstot.Caption = 0
labtot.Caption = 0

Unload Me

Exit Sub

Alcoman4clic:
             MsgBox "Error en la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al coman4 click"
             data_errfact.Recordset.Update
             Unload Me


End Sub

Private Sub Command5_Click()
Dim strIdTransac As String
On Error GoTo Alcoman5clic

If Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Then

    Set objPosCfe = New PosCfe
    
    Dim objresultado As Resultado
'    Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
    If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
       If frm_menu.data_parse.Recordset("base") = 78 Then
          Set objresultado = objPosCfe.Inicializar("SAPP-105", frm_menu.data_parse.Recordset("contacto"), vbNullString)
       Else
          Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
       End If
    Else
        If frm_menu.data_parse.Recordset("base") = 1 Then
           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-011", vbNullString)
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
                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)
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
                                                  If frm_menu.data_parse.Recordset("base") = 44 Then
                                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-266", vbNullString)
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
    
    data_temp.Recordset.MoveFirst
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
    
    '    EnableButtons True
    '    lblLibUcfeServiceHostPort.Caption = objPosCfe.ServicioLibUcfeHostPort
    
    strIdTransac = objPosCfe.CrearGuid
    '    lblIdTransaccionPOS2000.Caption = strIdTransaccionPos2000
    
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
    
'        MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'            "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
    
    'Enviando
    If Not EstaInicializado() Then Exit Sub
    
    Dim objCfe As CFE
    Set objCfe = New CFE

    Dim objCf As ClassFactory

    Set objCf = New ClassFactory
       
'    data_cabezal.RecordSource = "Select * from clirespl where cl_numero =" & labnrofac.Caption
'    data_cabezal.Refresh

    Set objCfe.EFact = New EFact
    With objCfe.EFact.Encabezado.IdDoc
         .TipoCFE = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
         .FchEmis.SetDate Year(data_temp.Recordset("fecha")), Month(data_temp.Recordset("fecha")), Day(data_temp.Recordset("fecha"))
         If data_cabeza2.Recordset("cl_forpago") = 1 Then
          .FmaPago = IdDoc_Fact_FmaPago_1
         Else
          .FmaPago = IdDoc_Fact_FmaPago_2
         End If
    End With
    With objCfe.EFact.Encabezado.Emisor
        .RUCEmisor = data_par.Recordset("ruc")
        .RznSoc = data_par.Recordset("nomc")
        .CdgDGISucur.FromString Trim(str(data_par.Recordset("codsuc")))
        .DomFiscal = data_par.Recordset("domic")
        .Ciudad = data_par.Recordset("ciudad")
        .Departamento = data_par.Recordset("dpto")
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
         If IsNull(data_cabeza2.Recordset("cl_nomcobr")) = False Then
            .IsValidCompraID = True
            .CompraID = data_cabeza2.Recordset("cl_nomcobr")
         End If
    End With
    With objCfe.EFact.Encabezado.Totales
         .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
         .IsValidTpoCambio = True
         If data_cabeza2.Recordset("usu_baja") = "USD" Then
            .TpoCambio.FromString Format(data_cabeza2.Recordset("saldo_chc2"), "0.00")
         Else
            .TpoCambio.FromString "1"
         End If
         .IsValidMntNetoIvaTasaMin = True
         .IsValidMntNetoIVATasaBasica = True
         .IsValidMntIVATasaMin = True
         .IsValidMntIVATasaBasica = True
         .IsValidMntNoGrv = True
         .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
         .MntNetoIVATasaBasica.FromString Format(data_cabeza2.Recordset("cl_atrasoa"), "0.00")
         If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
            If data_cabeza2.Recordset("cl_cedula") > 0 Then
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
    
    Do While Not data_temp.Recordset.EOF
       With objCfe.EFact.Detalle.Item.AddNew
          .NroLinDet.FromString Trim(str(data_temp.Recordset("linea")))
          .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_temp.Recordset("indfact"))))
          .NomItem = data_temp.Recordset("solicitant")
          .IsValidDscItem = True
          .DscItem = data_temp.Recordset("obsp")
          .cantidad.FromString Trim(str(data_temp.Recordset("cantidad")))
          .UniMed = "N/A"
          .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00")
          .MontoItem.FromString Format(data_temp.Recordset("tot_lin"), "0.00")
       End With
       data_temp.Recordset.MoveNext
    Loop
    data_temp.Recordset.MoveFirst
    Set objCfe.EFact.Referencia = New Referencia
    Do While Not data_temp.Recordset.EOF
       With objCfe.EFact.Referencia.ReferenciaA.AddNew
           .NroLinRef.FromString Trim(str(data_temp.Recordset("linearef")))
           .IsValidIndGlobal = False
           .IsValidTpoDocRef = True
           .TpoDocRef = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(data_temp.Recordset("tipodocref"))))
           .IsValidSerie = True
           .serie = Trim(data_temp.Recordset("serieref"))
           .IsValidNroCFERef = True
           .NroCFERef.FromLong data_temp.Recordset("nrofactref")
           .IsValidFechaCFEref = True
           .FechaCFEref.SetDate Year(data_temp.Recordset("fechafact")), Month(data_temp.Recordset("fechafact")), Day(data_temp.Recordset("fechafact"))
       End With
       If Label3.Caption = "NC de E-FACTURA" Then
          data_lincance.RecordSource = "Select * from linmmdd where factura =" & data_temp.Recordset("nrofactref") & " and moneda ='" & data_temp.Recordset("serieref") & "' and linea =" & data_temp.Recordset("linearef")
          data_lincance.Refresh
          If data_lincance.Recordset.RecordCount > 0 Then
             If IsNull(data_lincance.Recordset("descuento")) = True Then
                data_lincance.Recordset.Edit
                data_lincance.Recordset("descuento") = 1
                data_lincance.Recordset.Update
             End If
          End If
       End If
       
       data_temp.Recordset.MoveNext
    Loop
    
    Dim s As String
    s = objCfe.ToXml(True, XmlFormatting_Indented)

    Dim strGuid As String
    strGuid = objPosCfe.CrearGuid()
    Dim objResultadoCfe As ResultadoCfe
'    Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    If IsNull(data_cabeza2.Recordset("obsp")) = False Then
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))
    Else
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    End If
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
       Command8_Click
    End If
Else

End If

Exit Sub

Alcoman5clic:
             MsgBox "Error en la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al coman5 click"
             data_errfact.Recordset.Update
             Unload Me


End Sub

Private Sub Command6_Click()


Dim strIdTransac As String
On Error GoTo Alcoman6clic

If Label3.Caption = "E-FACTURA" Then

    Set objPosCfe = New PosCfe
    
    Dim objresultado As Resultado
'    Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
    
    If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
       If frm_menu.data_parse.Recordset("base") = 78 Then
          Set objresultado = objPosCfe.Inicializar("SAPP-105", frm_menu.data_parse.Recordset("contacto"), vbNullString)
       Else
          Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
       End If
       'Set objresultado = objPosCfe.Inicializar("SAPP-105", frm_menu.data_parse.Recordset("contacto"), vbNullString)
    Else
        If frm_menu.data_parse.Recordset("base") = 1 Then
           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-011", vbNullString)
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
                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)
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
                                   If frm_menu.data_parse.Recordset("base") = 91 Then 'Farmacia B16
                                      Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-216", vbNullString)
                                   Else
                                      If frm_menu.data_parse.Recordset("base") = 17 Then
                                         Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-017", vbNullString)
                                      Else
                                         If frm_menu.data_parse.Recordset("base") = 18 Then
                                            Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-018", vbNullString)
                                         Else
                                            If frm_menu.data_parse.Recordset("base") = 96 Then 'Despacho
                                               Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-218", vbNullString)
                                            Else
                                               If frm_menu.data_parse.Recordset("base") = 38 Then 'Cómputos
                                                  Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
                                               Else
                                                  If frm_menu.data_parse.Recordset("base") = 44 Then ' Gustavo
                                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-266", vbNullString)
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
    data_temp.Recordset.MoveFirst
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
    
    '    EnableButtons True
    '    lblLibUcfeServiceHostPort.Caption = objPosCfe.ServicioLibUcfeHostPort
    
    strIdTransac = objPosCfe.CrearGuid
    '    lblIdTransaccionPOS2000.Caption = strIdTransaccionPos2000
    
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
    
    '    MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
    '        "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
    
    'Enviando
    
        If Not EstaInicializado() Then Exit Sub
        
        Dim objCfe As CFE
        Set objCfe = New CFE
    
        Dim objCf As ClassFactory
    
        Set objCf = New ClassFactory
        
        Set objCfe.EFact = New EFact
    
    '    objCfe.FromXmlFile App.Path & "\sapp.xml"
        
'       data_cabeza2.RecordSource = "Select * from cabezados where cl_numero =" & labnrofac.Caption
'       data_cabeza2.Refresh
       With objCfe.EFact.Encabezado.IdDoc
            .TipoCFE = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
            .FchEmis.SetDate Year(data_temp.Recordset("fecha")), Month(data_temp.Recordset("fecha")), Day(data_temp.Recordset("fecha"))
            If IsNull(data_cabeza2.Recordset("fecha_modi")) = False Then
               .IsValidFchVenc = True
               .FchVenc.SetDate Year(data_cabeza2.Recordset("fecha_modi")), Month(data_cabeza2.Recordset("fecha_modi")), Day(data_cabeza2.Recordset("fecha_modi"))
            End If
            If data_cabeza2.Recordset("cl_forpago") = 1 Then
             .FmaPago = IdDoc_Fact_FmaPago_1
            Else
             .FmaPago = IdDoc_Fact_FmaPago_2
            End If
        End With
        With objCfe.EFact.Encabezado.Emisor
            .RUCEmisor = data_par.Recordset("ruc")
            .RznSoc = data_par.Recordset("nomc")
            .CdgDGISucur.FromString Trim(str(data_par.Recordset("codsuc")))
            .DomFiscal = data_par.Recordset("domic")
            .Ciudad = data_par.Recordset("ciudad")
            .Departamento = data_par.Recordset("dpto")
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
            If IsNull(data_cabeza2.Recordset("cl_nomcobr")) = False Then
               .IsValidCompraID = True
               .CompraID = data_cabeza2.Recordset("cl_nomcobr")
            End If
        End With
        With objCfe.EFact.Encabezado.Totales
            .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
            .IsValidTpoCambio = True
            If data_cabeza2.Recordset("usu_baja") = "USD" Then
               .TpoCambio.FromString Format(data_cabeza2.Recordset("saldo_chc2"), "0.00")
            Else
               .TpoCambio.FromString "1"
            End If
'      data_cabeza2.Recordset("saldo_doc2") = Format(labstot.Caption, "Standard") 'subtot iva 10
'      data_cabeza2.Recordset("cl_atrasoa") = Format(labstot22.Caption, "Standard") 'subtot iva 22
'      data_cabeza2.Recordset("cl_cedula") = Format(labstot0.Caption, "Standard") 'subtot iva cero
'      data_cabeza2.Recordset("saldo_cc")
'      data_cabeza2.Recordset("saldo_cc2") = Format(labiva22.
            .IsValidMntNetoIvaTasaMin = True
            .IsValidMntNetoIVATasaBasica = True
            .IsValidMntIVATasaMin = True
            .IsValidMntIVATasaBasica = True
            .IsValidMntNoGrv = True
            .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
            .MntNetoIVATasaBasica.FromString Format(data_cabeza2.Recordset("cl_atrasoa"), "0.00")
            If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
               If data_cabeza2.Recordset("cl_cedula") > 0 Then
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
        Do While Not data_temp.Recordset.EOF
           With objCfe.EFact.Detalle.Item.AddNew
              .NroLinDet.FromString Trim(str(data_temp.Recordset("linea")))
              .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_temp.Recordset("indfact"))))
              .NomItem = data_temp.Recordset("solicitant")
              .IsValidDscItem = True
'              .DscItem = data_temp.Recordset("nom_prod")
              .DscItem = data_temp.Recordset("obsp")
              .cantidad.FromString Trim(str(data_temp.Recordset("cantidad")))
              .UniMed = "N/A"
              .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00")
              .MontoItem.FromString Format(data_temp.Recordset("tot_lin"), "0.00")
           End With
           data_temp.Recordset.MoveNext
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
        If IsNull(data_cabeza2.Recordset("obsp")) = False Then
           Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))
        Else
           Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
        End If
        Set objUltimaSerieNumero = Nothing
        DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
        If Not objUltimaSerieNumero Is Nothing Then _
           ' cmdFirmarNc.Enabled = True
'           MsgBox "firmar NC"
        End If

        If frm_menu.data_parse.Recordset("base") = 38 Then
           MsgBox "Terminado"
'           Unload Me
        Else
           Command8_Click
        End If
Else
   If Label3.Caption = "E-TICKET" Then
      Command7_Click
   Else
      If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "ND de E-TICKET" Then
         Command9_Click
      Else
         If Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Then
            Command5_Click
         End If
      End If
   End If
End If

Exit Sub

Alcoman6clic:
             MsgBox "Error en la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al coman6 click"
             data_errfact.Recordset.Update
             Unload Me

End Sub

Private Sub Command7_Click()
Dim strIdTransac As String
On Error GoTo Alcoman7clic

If Label3.Caption = "E-TICKET" Then

    Set objPosCfe = New PosCfe
    
    Dim objresultado As Resultado
'    Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
    
    If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
       If frm_menu.data_parse.Recordset("base") = 78 Then 'Notebook JF
          Set objresultado = objPosCfe.Inicializar("SAPP-105", frm_menu.data_parse.Recordset("contacto"), vbNullString)
       Else
          Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
       End If
    Else
        If frm_menu.data_parse.Recordset("base") = 1 Then
           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-011", vbNullString)
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
                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)
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
                                                  If frm_menu.data_parse.Recordset("base") = 44 Then
                                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-266", vbNullString)
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
    data_temp.Recordset.MoveFirst
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
    
    '    EnableButtons True
    '    lblLibUcfeServiceHostPort.Caption = objPosCfe.ServicioLibUcfeHostPort
    
    strIdTransac = objPosCfe.CrearGuid
    '    lblIdTransaccionPOS2000.Caption = strIdTransaccionPos2000
    
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
    
'        MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'            "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
    
    'Enviando
    If Not EstaInicializado() Then Exit Sub
    
    Dim objCfe As CFE
    Set objCfe = New CFE

    Dim objCf As ClassFactory

    Set objCf = New ClassFactory
       
'    data_cabezal.RecordSource = "Select * from clirespl where cl_numero =" & labnrofac.Caption
'    data_cabezal.Refresh

    Set objCfe.ETck = New ETck
    With objCfe.ETck.Encabezado.IdDoc
        .TipoCFE = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
        .FchEmis.SetDate Year(data_temp.Recordset("fecha")), Month(data_temp.Recordset("fecha")), Day(data_temp.Recordset("fecha"))
        If IsNull(data_cabeza2.Recordset("fecha_modi")) = False Then
           .IsValidFchVenc = True
           .FchVenc.SetDate Year(data_cabeza2.Recordset("fecha_modi")), Month(data_cabeza2.Recordset("fecha_modi")), Day(data_cabeza2.Recordset("fecha_modi"))
        End If
        If data_cabeza2.Recordset("cl_forpago") = 1 Then
           .FmaPago = IdDoc_Tck_FmaPago_1
        Else
           .FmaPago = IdDoc_Tck_FmaPago_2
        End If
    End With
    With objCfe.ETck.Encabezado.Emisor
        .RUCEmisor = data_par.Recordset("ruc")
        .RznSoc = data_par.Recordset("nomc")
        .CdgDGISucur.FromString Trim(str(data_par.Recordset("codsuc")))
        .DomFiscal = data_par.Recordset("domic")
        .Ciudad = data_par.Recordset("ciudad")
        .Departamento = data_par.Recordset("dpto")
    End With
    Set objCfe.ETck.Encabezado.Receptor = New Receptor_Tck
    Set objCfe.ETck.Encabezado.Receptor.Receptor_Tck_Choice = New Receptor_Tck_Choice
    With objCfe.ETck.Encabezado.Receptor
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
        .Receptor_Tck_Choice.DocRecepExt = data_cabeza2.Recordset("cl_nom_sup")
'        If IsNull(data_cabezal.Recordset("cl_nom_sup")) = False Then
'           .Receptor_Tck_Choice.DocRecep = data_cabezal.Recordset("cl_nom_sup")
'        Else
'           .Receptor_Tck_Choice.DocRecep = "0"
'        End If
'        .Receptor_Tck_Choice.DocRecepExt = data_cabezal.Recordset("cl_nom_sup")
        .RznSocRecep = data_cabeza2.Recordset("info_debit")
        .DirRecep = data_cabeza2.Recordset("cl_direcci")
        .CiudadRecep = data_cabeza2.Recordset("cl_zona")
        If IsNull(data_cabeza2.Recordset("cl_nomcobr")) = False Then
           .IsValidCompraID = True
           .CompraID = data_cabeza2.Recordset("cl_nomcobr")
        End If
    End With
    With objCfe.ETck.Encabezado.Totales
         .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
         .IsValidTpoCambio = True
         If data_cabeza2.Recordset("usu_baja") = "USD" Then
            .TpoCambio.FromString Format(data_cabeza2.Recordset("saldo_chc2"), "0.00")
         Else
            .TpoCambio.FromString "1"
         End If
         .IsValidMntNetoIvaTasaMin = True
         .IsValidMntNetoIVATasaBasica = True
         .IsValidMntIVATasaMin = True
         .IsValidMntIVATasaBasica = True
         .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
         .MntNetoIVATasaBasica.FromString Format(data_cabeza2.Recordset("cl_atrasoa"), "0.00")
         If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
            If data_cabeza2.Recordset("cl_cedula") > 0 Then
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
    Do While Not data_temp.Recordset.EOF
       With objCfe.ETck.Detalle.Item.AddNew
          .NroLinDet.FromString Trim(str(data_temp.Recordset("linea")))
          .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_temp.Recordset("indfact"))))
          .NomItem = data_temp.Recordset("solicitant")
          .IsValidDscItem = True
          .DscItem = data_temp.Recordset("obsp")
          .cantidad.FromString Trim(str(data_temp.Recordset("cantidad")))
          .UniMed = "N/A"
          .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00")
          .MontoItem.FromString Format(data_temp.Recordset("tot_lin"), "0.00")
       End With
       data_temp.Recordset.MoveNext
    Loop
    
    Dim s As String
    s = objCfe.ToXml(True, XmlFormatting_Indented)

'    Open App.Path & "\sapp.xml" For Output As #1
'    Print #1, s

    Dim strGuid As String
    strGuid = objPosCfe.CrearGuid()
    Dim objResultadoCfe As ResultadoCfe
'    Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    If IsNull(data_cabeza2.Recordset("obsp")) = False Then
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))
    Else
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    End If
    Set objUltimaSerieNumero = Nothing
    DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
    If Not objUltimaSerieNumero Is Nothing Then _
        ' cmdFirmarNc.Enabled = True
'       MsgBox "firmar NC"
    End If
    If frm_menu.data_parse.Recordset("base") = 38 Then
       MsgBox "Terminado"
       Unload Me
    Else
       Command8_Click
    End If
End If

Exit Sub

Alcoman7clic:
             MsgBox "Error en la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al coman7 click"
             data_errfact.Recordset.Update
             Unload Me


End Sub

Private Sub Command8_Click()
Dim Xlatasa, Xlatasa22, Xmasdiezui As Double
Dim Xquehace As Integer
Dim Xenquelugar, XcantTimbre As Integer
Dim Imprime As String
Xenquelugar = 0
Xquehace = 0
XcantTimbre = 1
On Error GoTo Xquepasoalgrabar

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")

If data_temp.Recordset.RecordCount > 0 Then
   data_lin.RecordSource = "Select * from linmmdd where factura =" & 223595
   data_lin.Refresh
   data_temp.Recordset.MoveFirst
   data_cabezal.RecordSource = "Select * from clirespl where cl_codigo =" & labnrofac.Caption
''''data_cabezal.Recordset("cl_socmnro")   ''serie
   data_cabezal.Refresh
   Xenquelugar = 1
   data_lin2.RecordSource = "Select * from hc_torax where id =" & 5
   data_lin2.Refresh
   data_lin3.RecordSource = "Select * from indica_enfc where id =" & 50
   data_lin3.Refresh
   data_cabeza2.Refresh
   Xenquelugar = 2
   data_cabeza2.Recordset.MoveFirst
   data_cabezal.Recordset.AddNew
'           data_cabezal.Recordset("id") = 1
   data_cabezal.Recordset("cl_tipcli") = "1.0"
   data_cabezal.Recordset("cl_tipocli") = data_cabeza2.Recordset("cl_tipocli")
   data_cabezal.Recordset("cl_telefon") = data_cabeza2.Recordset("cl_telefon")
   data_cabezal.Recordset("cl_socmnro") = labserie.Caption
   data_cabezal.Recordset("cl_numero") = Val(labnrofac.Caption)
   data_cabezal.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
   If mfd.Text <> "__/__/____" Then
      data_cabezal.Recordset("fecha_reac") = Format(mfd.Text, "dd/mm/yyyy")
   End If
   If mfh.Text <> "__/__/____" Then
      data_cabezal.Recordset("cl_tj_venc") = Format(mfh.Text, "dd/mm/yyyy")
   End If
   data_cabezal.Recordset("cl_nrovend") = 0
   If labtipof.Caption = "CONTADO" Then
      data_cabezal.Recordset("cl_forpago") = 1
   Else
      If labtipof.Caption = "CREDITO" Then
         data_cabezal.Recordset("cl_forpago") = 2
      Else
         data_cabezal.Recordset("cl_forpago") = 2
      End If
   End If
   data_cabezal.Recordset("cl_celular") = data_cabeza2.Recordset("cl_celular") 'descripcion f.pago
   If labvence.Text <> "__/__/____" Then
      data_cabezal.Recordset("fecha_modi") = Format(labvence.Text, "dd/mm/yyyy")
   End If
   data_cabezal.Recordset("cl_diacobr") = Trim(str(data_par.Recordset("ruc")))
   data_cabezal.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
   data_cabezal.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
   data_cabezal.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
   data_cabezal.Recordset("cl_referen") = data_par.Recordset("domic")
   data_cabezal.Recordset("tit_tarj") = data_par.Recordset("ciudad")
   data_cabezal.Recordset("cl_nomconv") = data_par.Recordset("dpto")
'receptor
   If Label3.Caption = "E-TICKET" Then
      data_cabezal.Recordset("cl_nro_sup") = 0
   Else
      data_cabezal.Recordset("cl_nro_sup") = 2
   End If
   data_cabezal.Recordset("hora_baja") = "UY"
   data_cabezal.Recordset("cl_nom_sup") = frm_convenios.txt_ruc.Text
    'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
   data_cabezal.Recordset("info_debit") = frm_convenios.t_razon.Text
   data_cabezal.Recordset("cl_direcci") = frm_convenios.txt_direc.Text
   If frm_convenios.txt_localid.Text <> "" Then
      data_cabezal.Recordset("cl_zona") = frm_convenios.txt_localid.Text
   End If
   If frm_convenios.t_dpto.Text <> "" Then
      data_cabezal.Recordset("cl_email") = frm_convenios.t_dpto.Text 'opcional
   End If
   data_cabezal.Recordset("cl_localid") = "URUGUAY" 'opcional
   data_cabezal.Recordset("cl_codigo") = Val(frm_convenios.txt_cuenta.Text)
   data_cabezal.Recordset("usu_baja") = data_cabeza2.Recordset("usu_baja") 'moneda
   data_cabezal.Recordset("saldo_chc2") = data_cabeza2.Recordset("saldo_chc2") 'valor dolar
   data_cabezal.Recordset("saldo_cc") = data_cabeza2.Recordset("saldo_cc")  'iva minimo
   data_cabezal.Recordset("saldo_cc2") = data_cabeza2.Recordset("saldo_cc2") 'iva básico
   data_cabezal.Recordset("cl_atrasoa") = data_cabeza2.Recordset("cl_atrasoa") 'subtot iva 22
   data_cabezal.Recordset("cl_cedula") = data_cabeza2.Recordset("cl_cedula") 'subtot iva cero
'   data_cabezal.Recordset("saldo_doc2") = Format(labstot.Caption, "Standard")
   data_cabezal.Recordset("saldo_doc2") = data_cabeza2.Recordset("saldo_doc2") 'subtot grav iva 10
   data_cabezal.Recordset("cl_atrasop") = Xlatasa
   data_cabezal.Recordset("cl_decuota") = Xlatasa22
   data_cabezal.Recordset("saldo_doc") = Format(labtot.Caption, "Standard")
   data_cabezal.Recordset("cl_grupo") = data_temp.Recordset.RecordCount
   data_cabezal.Recordset("saldo_chc") = Format(labtot.Caption, "Standard")
   If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Or Label3.Caption = "ND de E-TICKET" Then
      data_cabezal.Recordset("cl_cuopaga") = data_temp.Recordset("tipodocref")
      data_cabezal.Recordset("codmotbaja") = data_temp.Recordset("serieref")
      data_cabezal.Recordset("ultanopmut") = data_temp.Recordset("nrofactref")
      data_cabezal.Recordset("cl_fultvta") = data_temp.Recordset("fechafact")
      data_cabezal.Recordset("cl_entre") = data_temp.Recordset("motivoref")
   End If
   If IsNull(data_cabeza2.Recordset("cl_fultpag")) = False Then
      data_cabezal.Recordset("cl_fultpag") = data_cabeza2.Recordset("cl_fultpag")
   End If
   data_cabezal.Recordset("cl_nomcobr") = data_cabeza2.Recordset("cl_nomcobr")
   data_cabezal.Recordset.Update
   Xenquelugar = 3
'fin de cabezal
   Do While Not data_temp.Recordset.EOF
      Xcandelin = Xcandelin + 1
      data_lin.Recordset.AddNew
      data_lin.Recordset("linea") = data_temp.Recordset("linea")
      data_lin.Recordset("factura") = Val(labnrofac.Caption)
      data_lin.Recordset("tipo") = data_temp.Recordset("tipo")
      data_lin.Recordset("realizada") = Format(data_temp.Recordset("realizada"), "dd/mm/yyyy")
      data_lin.Recordset("fecha") = Format(data_temp.Recordset("fecha"), "dd/mm/yyyy")
      data_lin.Recordset("cod_cli") = data_temp.Recordset("cod_cli")
      data_lin.Recordset("nom_cli") = Mid(data_temp.Recordset("nom_cli"), 1, 30)
      data_lin.Recordset("convenio") = data_temp.Recordset("convenio")
      data_lin.Recordset("cod_prod") = data_temp.Recordset("cod_prod")
      data_lin.Recordset("nom_prod") = data_temp.Recordset("nom_prod")
      data_lin.Recordset("costo_prod") = data_temp.Recordset("costo_prod")
      data_lin.Recordset("operador") = data_temp.Recordset("operador")
      data_lin.Recordset("hora") = data_temp.Recordset("hora")
      data_lin.Recordset("imp_timbre") = data_temp.Recordset("imp_timbre") ' sub total de la línea
      data_lin.Recordset("tot_lin") = data_temp.Recordset("tot_lin") ' total de la linea de la factura
   '      data_lin.Recordset("costo") = CDbl(labstot.Caption) ' sub total de la factura
      If IsNull(data_temp.Recordset("iva22")) = False Then
         If data_temp.Recordset("iva22") > 0 Then
            data_lin.Recordset("valor_iva") = data_temp.Recordset("iva22")
         Else
            data_lin.Recordset("valor_iva") = data_temp.Recordset("pre_civa")
         End If
      Else
         data_lin.Recordset("valor_iva") = data_temp.Recordset("pre_civa")
      End If
      data_lin.Recordset("mes_paga") = data_temp.Recordset("mes_paga")
      data_lin.Recordset("ano_paga") = data_temp.Recordset("ano_paga")
      data_lin.Recordset("base") = data_temp.Recordset("base")
      data_lin.Recordset("ruc") = data_temp.Recordset("ruc")
      data_lin.Recordset("grupo") = data_temp.Recordset("grupo") 'cobrador
      data_lin.Recordset("solicitant") = data_temp.Recordset("solicitant")
      data_lin.Recordset("nom_med_s") = data_temp.Recordset("nom_med_s")
      data_lin.Recordset("nom_medic") = data_temp.Recordset("nom_medic")
      data_lin.Recordset("nom_med_a") = data_temp.Recordset("nom_med_a")
'               data_lin.Recordset("vto") = data_temp.Recordset("vto")
      data_lin.Recordset("rub_cont") = data_temp.Recordset("rub_cont")
      data_lin.Recordset("nom_flia") = data_temp.Recordset("nom_flia")
      data_lin.Recordset("nro_superv") = data_temp.Recordset("nro_superv")
      data_lin.Recordset("nom_superv") = data_temp.Recordset("nom_superv")
      data_lin.Recordset("pre_prod") = data_temp.Recordset("pre_prod")
      data_lin.Recordset("recargo") = data_temp.Recordset("linearef")
      data_lin.Recordset("moneda") = labserie.Caption
      data_lin.Recordset("tipo_mov") = data_temp.Recordset("tipo_mov")
      If data_cabeza2.Recordset("cl_tipocli") = 111 Then
         data_lin.Recordset("pendiente") = "F"
      Else
         If data_cabeza2.Recordset("cl_tipocli") = 101 Then
            data_lin.Recordset("pendiente") = "T"
         Else
            If data_cabeza2.Recordset("cl_tipocli") = 112 Then
               data_lin.Recordset("pendiente") = "N" 'NC de E-FACT
            Else
               If data_cabeza2.Recordset("cl_tipocli") = 102 Then
                  data_lin.Recordset("pendiente") = "C" 'NC de E-TCK
               Else
                  If data_cabeza2.Recordset("cl_tipocli") = 113 Then
                     data_lin.Recordset("pendiente") = "A" 'ND de E-FACT
                  Else
                     If data_cabeza2.Recordset("cl_tipocli") = 103 Then
                        data_lin.Recordset("pendiente") = "B" 'ND de E-TCK
                     Else
                        If Label7.Caption = "REG." Then
                           data_lin.Recordset("pendiente") = "X"
                        Else
                           data_lin.Recordset("pendiente") = "Z"
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      data_lin.Recordset.Update
      If data_temp.Recordset("nom_prod") = "Costo por llamados" Then
         data_timbre.RecordSource = "select * from convenio_tiquets where nom_grupo ='" & frm_convenios.txt_cod.Text & "' and fecha_pago is null and importe =" & data_temp.Recordset("imp_timbre")
         data_timbre.Refresh
         If data_timbre.Recordset.RecordCount > 0 Then
            data_timbre.Recordset.MoveFirst
            Do While XcantTimbre <= data_temp.Recordset("cantidad")
               data_timbre.Recordset.Edit
               data_timbre.Recordset("fecha_pago") = Date
               data_timbre.Recordset("nro_doc") = Val(labnrofac.Caption)
               data_timbre.Recordset.Update
               XcantTimbre = XcantTimbre + 1
            Loop
         End If
      End If
      Xenquelugar = 4
      data_lin2.Recordset.AddNew
      data_lin2.Recordset("hora") = "INT1"
      data_lin2.Recordset("descrip") = labserie.Caption  'serie del comprobante
      data_lin2.Recordset("hc_nro") = 2 'tasa minima
      data_lin2.Recordset("hc_cod") = Val(labnrofac.Caption)
      data_lin2.Recordset("hc_mat") = data_temp.Recordset("linea")
      data_lin2.Recordset.Update
      Xenquelugar = 5
'      data_lin2.Refresh

'es acaaaaaa
      If IsNull(data_temp.Recordset("obsp")) = False Then
    '   If t_pie.Text <> "" Then
         data_lin3.Recordset.AddNew
         data_lin3.Recordset("id") = data_par.Recordset("varios") + 1
         data_lin3.Recordset("idhc") = Val(labnrofac.Caption)
         data_lin3.Recordset("in_dosis") = 3
         data_lin3.Recordset("in_obs") = data_temp.Recordset("obsp")
         data_lin3.Recordset("in_hora") = labserie.Caption
         data_lin3.Recordset("in_uni") = data_temp.Recordset("linea")
         data_lin3.Recordset.Update
         data_lin3.Refresh
         data_par.Recordset.Edit
         data_par.Recordset("varios") = data_par.Recordset("varios") + 1
         data_par.Recordset.Update
         data_par.Refresh
      End If
      Xenquelugar = 6
      data_temp.Recordset.MoveNext
   Loop
   Xcandelin = 0
   Xconlin = 0
   If IsNull(data_cabeza2.Recordset("obsp")) = False Then
'   If t_pie.Text <> "" Then
      data_lin3.Recordset.AddNew
      data_lin3.Recordset("id") = data_par.Recordset("varios") + 1
      data_lin3.Recordset("idhc") = Val(labnrofac.Caption)
      data_lin3.Recordset("in_dosis") = 1
      data_lin3.Recordset("in_obs") = t_pie.Text
      data_lin3.Recordset("in_hora") = labserie.Caption
      data_lin3.Recordset.Update
      data_lin3.Refresh
      data_par.Recordset.Edit
      data_par.Recordset("varios") = data_par.Recordset("varios") + 1
      data_par.Recordset.Update
      data_par.Refresh
   End If
   
   Xenquelugar = 7
   Imprime = MsgBox("Desea imprimir la factura?", vbInformation + vbYesNo)
   If Imprime = vbYes Then
        data_temp.Recordset.MoveFirst
        cr2.ReportFileName = App.path & "\faccnvnew.rpt"
        cr2.CopiesToPrinter = 2
        cr2.Action = 1
   Else
        cr1.ReportFileName = App.path & "\faccnvnew.rpt"
        cr1.Action = 1
   End If
'   cr1.ReportFileName = App.path & "\faccnvnew.rpt"
'   cr1.Action = 1
   MsgBox "Terminado"


'            frm_impfaccnv2.Show vbModal
'   MsgBox "Terminado"
'   frm_impfaccnv.Show vbModal
   
Else
   MsgBox "No hay lineas en la factura"
   
End If

Exit Sub

Xquepasoalgrabar:
              If Err.Number = 3155 Then
                 MsgBox "Error al terminar la factura. AVISE A INFORMATICA! " & Err.Number & " " & Err.Description, vbInformation
                 data_errfact.Recordset.AddNew
                 data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                 data_errfact.Recordset("fecha") = Date
                 data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                 data_errfact.Recordset("nroerr") = Err.Number
                 data_errfact.Recordset("desc") = "Al coman8 grabar " & Trim(str(Xenquelugar))
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
                    MsgBox "Error al terminar la factura. AVISE A INFORMATICA! " & Err.Number & " " & Err.Description, vbInformation
                    data_errfact.Recordset.AddNew
                    data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
                    data_errfact.Recordset("fecha") = Date
                    data_errfact.Recordset("hora") = Format(Time, "HH:mm")
                    data_errfact.Recordset("nroerr") = Err.Number
                    data_errfact.Recordset("desc") = "Al coman8 grabar " & Trim(str(Xenquelugar))
                    data_errfact.Recordset.Update
                    Unload Me
                 End If
              End If


'Unload Me

End Sub

Private Sub Command9_Click()
Dim strIdTransac As String
Dim tipo As Integer
tipo = 0
On Error GoTo Alcoman9clic

If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "ND de E-TICKET" Then

    Set objPosCfe = New PosCfe
    
    Dim objresultado As Resultado
'    Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
    If IsNull(frm_menu.data_parse.Recordset("contacto")) = False Then
       Set objresultado = objPosCfe.Inicializar("SAPP0001", frm_menu.data_parse.Recordset("contacto"), vbNullString)
    Else
        If frm_menu.data_parse.Recordset("base") = 1 Then
           Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-011", vbNullString)
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
                       Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)
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
                                                  If frm_menu.data_parse.Recordset("base") = 44 Then
                                                     Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-266", vbNullString)
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
    data_temp.Recordset.MoveFirst
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
    
    '    EnableButtons True
    '    lblLibUcfeServiceHostPort.Caption = objPosCfe.ServicioLibUcfeHostPort
    
    strIdTransac = objPosCfe.CrearGuid
    '    lblIdTransaccionPOS2000.Caption = strIdTransaccionPos2000
    
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
    
'        MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'            "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
    
    'Enviando
    If Not EstaInicializado() Then Exit Sub
    
    Dim objCfe As CFE
    Set objCfe = New CFE

    Dim objCf As ClassFactory

    Set objCf = New ClassFactory
       
'    data_cabezal.RecordSource = "Select * from clirespl where cl_numero =" & labnrofac.Caption
'    data_cabezal.Refresh

    Set objCfe.ETck = New ETck
    With objCfe.ETck.Encabezado.IdDoc
         .TipoCFE = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
         .FchEmis.SetDate Year(data_temp.Recordset("fecha")), Month(data_temp.Recordset("fecha")), Day(data_temp.Recordset("fecha"))
         If data_cabeza2.Recordset("cl_forpago") = 1 Then
          .FmaPago = IdDoc_Tck_FmaPago_1
         Else
          .FmaPago = IdDoc_Tck_FmaPago_2
         End If
    End With

    With objCfe.ETck.Encabezado.Emisor
        .RUCEmisor = data_par.Recordset("ruc")
        .RznSoc = data_par.Recordset("nomc")
        .CdgDGISucur.FromString Trim(str(data_par.Recordset("codsuc")))
        .DomFiscal = data_par.Recordset("domic")
        .Ciudad = data_par.Recordset("ciudad")
        .Departamento = data_par.Recordset("dpto")
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
         If IsNull(data_cabeza2.Recordset("cl_nomcobr")) = False Then
            .IsValidCompraID = True
            .CompraID = data_cabeza2.Recordset("cl_nomcobr")
         End If
    End With
    With objCfe.ETck.Encabezado.Totales
         .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
         .IsValidTpoCambio = True
         If data_cabeza2.Recordset("usu_baja") = "USD" Then
            .TpoCambio.FromString Format(data_cabeza2.Recordset("saldo_chc2"), "0.00")
         Else
            .TpoCambio.FromString "1"
         End If
         .IsValidMntNetoIvaTasaMin = True
         .IsValidMntNetoIVATasaBasica = True
         .IsValidMntIVATasaMin = True
         .IsValidMntIVATasaBasica = True
         .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
         .MntNetoIVATasaBasica.FromString Format(data_cabeza2.Recordset("cl_atrasoa"), "0.00")
         If IsNull(data_cabeza2.Recordset("cl_cedula")) = False Then
            If data_cabeza2.Recordset("cl_cedula") > 0 Then
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
    
    Do While Not data_temp.Recordset.EOF
       With objCfe.ETck.Detalle.Item.AddNew
          .NroLinDet.FromString Trim(str(data_temp.Recordset("linea")))
          .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_temp.Recordset("indfact"))))
          .NomItem = data_temp.Recordset("solicitant")
          .IsValidDscItem = True
          .DscItem = data_temp.Recordset("obsp")
          .cantidad.FromString Trim(str(data_temp.Recordset("cantidad")))
          .UniMed = "N/A"
          .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00")
          .MontoItem.FromString Format(data_temp.Recordset("tot_lin"), "0.00")
       End With
       data_temp.Recordset.MoveNext
    Loop
    data_temp.Recordset.MoveFirst
    Set objCfe.ETck.Referencia = New Referencia
    Do While Not data_temp.Recordset.EOF
       With objCfe.ETck.Referencia.ReferenciaA.AddNew
           .NroLinRef.FromString Trim(str(data_temp.Recordset("linearef")))
           .IsValidIndGlobal = False
           .IsValidTpoDocRef = True
           .TpoDocRef = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_temp.Recordset("tipodocref"))))
           .IsValidSerie = True
           .serie = Trim(data_temp.Recordset("serieref"))
           .IsValidNroCFERef = True
           .NroCFERef.FromLong data_temp.Recordset("nrofactref")
           .IsValidFechaCFEref = True
           .FechaCFEref.SetDate Year(data_temp.Recordset("fechafact")), Month(data_temp.Recordset("fechafact")), Day(data_temp.Recordset("fechafact"))
       End With
       If Label3.Caption = "NC de E-TICKET" Then
          data_lincance.RecordSource = "Select * from linmmdd where factura =" & data_temp.Recordset("nrofactref") & " and moneda ='" & data_temp.Recordset("serieref") & "' and linea =" & data_temp.Recordset("linearef")
          data_lincance.Refresh
          If data_lincance.Recordset.RecordCount > 0 Then
             If IsNull(data_lincance.Recordset("descuento")) = True Then
                data_lincance.Recordset.Edit
                data_lincance.Recordset("descuento") = 1
                data_lincance.Recordset.Update
             End If
          End If
       End If
       data_temp.Recordset.MoveNext
    Loop
    
    Dim s As String
    s = objCfe.ToXml(True, XmlFormatting_Indented)

    Dim strGuid As String
    strGuid = objPosCfe.CrearGuid()
    Dim objResultadoCfe As ResultadoCfe
'    Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    If IsNull(data_cabeza2.Recordset("obsp")) = False Then
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfeConAdenda(strGuid, objCfe, Trim(data_cabeza2.Recordset("obsp")))
    Else
       Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    End If
    Set objUltimaSerieNumero = Nothing
    DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
    If Not objUltimaSerieNumero Is Nothing Then _
        ' cmdFirmarNc.Enabled = True
    '   MsgBox "firmar NC"
    End If
    If frm_menu.data_parse.Recordset("base") = 38 Then
       MsgBox "TErminado"
       Unload Me
    Else
       Command8_Click
    End If
End If

Exit Sub

Alcoman9clic:
             MsgBox "Error en la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Al coman9 click"
             data_errfact.Recordset.Update
             Unload Me

End Sub

Private Sub Form_Load()
Dim Banderacuenta As Integer

'''On Error GoTo Alloafrm


Banderacuenta = 0
Xcandelin = 0

Xtot = 0
Xsubt = 0
Xivvva = 0
Xivauno = 0
cbomon.ListIndex = 0
Xsubt22 = 0
Xsubt0 = 0

If Month(Date) > 9 Then
   mfd.Text = "01/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
Else
   mfd.Text = "01/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
End If

mfh.Text = DateSerial(Year(Now), Month(Now) + 1, 0)
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_verfac.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_verfac.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

If frm_menu.data_parse.Recordset("base") = 38 Then
   data_lincance.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_lincance.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

data_errfact.DatabaseName = App.path & "\errores.mdb"
data_errfact.RecordSource = "errores"
data_errfact.Refresh

data_imagen.DatabaseName = App.path & "\imagen.mdb"
data_imagen.RecordSource = "qr"
data_imagen.Refresh

data_timbre.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_ui.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ui.RecordSource = "hc_frecresp"
data_ui.Refresh
labseriefactnc.Caption = ""
labnrofactnc.Caption = ""
labnrolin.Caption = ""

If data_ui.Recordset.RecordCount > 0 Then
   Xeldolar = CDbl(data_ui.Recordset("hora"))
Else
   Xeldolar = 0
End If

data_cabeza2.DatabaseName = App.path & "\factura.mdb"
data_cabeza2.RecordSource = "cabezados"
data_cabeza2.Refresh

If data_cabeza2.Recordset.RecordCount > 0 Then
   data_cabeza2.Recordset.MoveFirst
   Do While Not data_cabeza2.Recordset.EOF
      data_cabeza2.Recordset.Delete
      data_cabeza2.Recordset.MoveNext
   Loop
End If
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_buscalafac.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_buscalafac.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_verfac.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_verfac.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

If frm_menu.data_parse.Recordset("base") = 38 Then
   data_cabezal.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_cabezal.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_lin2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_lin2.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

data_ctrolfact.DatabaseName = App.path & "\ctrf.mdb"
data_ctrolfact.RecordSource = "ctrf"
data_ctrolfact.Refresh

If frm_menu.data_parse.Recordset("base") = 38 Then
   data_lin3.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_lin3.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
Else
   data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
data_par.Connect = "ODBC;DSN=sappfact;"
data_par.RecordSource = "paramsapp"
data_par.Refresh

If IsNull(data_par.Recordset("adenda")) = False Then
   t_pie.Text = data_par.Recordset("adenda")
Else
   t_pie.Text = ""
End If
t_pie.Text = "La presente factura podrá ser abonada de la siguiente manera:" & chr(13) & chr(10)
t_pie.Text = t_pie.Text & "Depósito en pesos CTA.CTE. 001552096-00005 NRo.CTA.anterior:073-0002052. A Nombre de SAPP SA" & chr(13) & chr(10)
t_pie.Text = t_pie.Text & "Depósito en pesos CTA.CTE. Santander Sucursal 15 Nro.5908 a nombre de SAPP SA" & chr(13) & chr(10)
t_pie.Text = t_pie.Text & "Enviar comprobante de depósito a tesoreria@sapp.com.uy" & chr(13) & chr(10)
t_pie.Text = t_pie.Text & "Facturación Teléfono 097-318598  Cobranza Teléfono 097-215423" & chr(13) & chr(10)
t_pie.Text = t_pie.Text & "Empresa adherida al CLEARING DE INFORMES."
t_pie.Text = t_pie.Text & "* 1- No gravado, 2- IVA tasa min, 3- IVA tasa básica"

If frm_convenios.txt_cuenta.Text = "" Then
   Banderacuenta = 1
Else
   If frm_convenios.txt_cuenta.Text > 0 Then
      Banderacuenta = 0
   Else
      Banderacuenta = 1
   End If
End If

labnrofac.Caption = ""
labserie.Caption = ""

labusuario.Caption = WElusuario

data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
If Banderacuenta = 0 Then
    data_cli.RecordSource = "Select * from clientes where cl_codigo =" & frm_convenios.txt_cuenta.Text
    data_cli.Refresh
    If data_cli.Recordset.RecordCount > 0 Then
       If XcomoFactura = 2 Then
          Xtipodedocumento = 4
       Else
          Xtipodedocumento = 2
       End If
          
       labcli.Caption = frm_convenios.t_razon.Text
       labnrocli.Caption = data_cli.Recordset("cl_codigo")
           labdom.Caption = frm_convenios.txt_direc.Text
           lablocal.Caption = frm_convenios.txt_localid.Text
        If frm_convenios.cbovenc.ListIndex > 0 Then
           labvence.Text = Date + Val(frm_convenios.cbovenc.Text)
        Else
           labvence.Text = Date
        End If
        If Xformapcnv = 2 Then
           labtipof.Caption = "CONTADO"
        Else
           If Xformapcnv = 1 Then
              labtipof.Caption = "CREDITO"
           Else
              MsgBox "Error en la factura"
              End
           End If
        End If
        If XAlta = 10 Then
           Xbb = 101
           Label3.Caption = "E-FACTURA"
           Check1.Value = 0
           Check1.Enabled = False
        Else
           If XAlta = 11 Then
              Label3.Caption = "E-TICKET"
              Xbb = 101
              Check1.Value = 1
           Else
              If XAlta = 12 Then
                 Label3.Caption = "NC de E-TICKET"
                 Xbb = 101
                 b_vernc.Visible = True
              Else
                 If XAlta = 14 Then
                    Label3.Caption = "NC de E-FACTURA"
                    Xbb = 101
                    Check1.Value = 0
                    Check1.Enabled = False
                    b_vernc.Visible = True
                 Else
                    If XAlta = 15 Then
                       Label3.Caption = "ND de E-FACTURA"
                       Xbb = 101
                       Check1.Enabled = False
                       b_vernc.Visible = True
                    Else
                       If XAlta = 16 Then
                          Label3.Caption = "ND de E-TICKET"
                          Xbb = 101
                          b_vernc.Visible = True
                       Else
                          MsgBox "Error en número de factura, VERIFIQUE!!"
                          End
                       End If
                    End If
                 End If
              End If
           End If
        End If
        
        data_temp.DatabaseName = App.path & "\factura.mdb"
        data_temp.RecordSource = "lineas2"
        data_temp.Refresh
        If data_temp.Recordset.RecordCount > 0 Then
           data_temp.Recordset.MoveFirst
           Do While Not data_temp.Recordset.EOF
              data_temp.Recordset.Delete
              data_temp.Recordset.MoveNext
           Loop
        End If
        data_temp.Refresh
'        labguid.Caption = labnrofac.Caption
        cboiva.ListIndex = 0
        Xconlin = 0
        labstot.Caption = 0
        labstot22.Caption = 0
        labiva22.Caption = 0
        labstot0.Caption = 0
        labiva.Caption = 0
        labtot.Caption = 0
        mf.Text = Date
        If Label3.Caption = "NC de E-TICKET" Or Label3.Caption = "NC de E-FACTURA" Or Label3.Caption = "ND de E-FACTURA" Or Label3.Caption = "ND de E-TICKET" Then
           frm_factcancela.Show vbModal
        End If
    Else
        MsgBox "No se encuentra cliente, VERIFIQUE CUENTA", vbCritical
        Unload Me
    End If
Else
    MsgBox "No figura número de cuenta, verifique!!"
End If
labiva.Caption = 0
labiva22.Caption = 0
Consulta_tiquet
'Exit Sub

'Alloafrm:
'             MsgBox "Error en la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
'             data_errfact.Recordset.AddNew
''             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
'             data_errfact.Recordset("fecha") = Date
'             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
'             data_errfact.Recordset("nroerr") = Err.Number
'             data_errfact.Recordset("desc") = "Al LOAD del frm"
'             data_errfact.Recordset.Update
'             Unload Me

End Sub

Private Sub Form_Resize()
On Error GoTo ErrRessi

With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

Exit Sub

ErrRessi:
        MsgBox ("error al RESIRZE")
        
End Sub

Private Sub mf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub t_ano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_imp.SetFocus
End If

End Sub

Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboiva.SetFocus
End If

End Sub

Private Sub t_desc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_desc2.SetFocus
End If

End Sub

Private Sub t_desc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_desc3.SetFocus
End If

End Sub

Private Sub t_desc3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mes.SetFocus
End If

End Sub

Private Sub t_imp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cant.SetFocus
End If

End Sub

Private Sub t_imp_LostFocus()
If t_imp.Text = "" Then
   t_imp.Text = 0
End If
t_imp.Text = Format(t_imp.Text, "Standard")

End Sub

Private Sub t_mes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_ano.SetFocus
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
On Error GoTo Aldespinfcfe

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
        labnrofac.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
        data_imagen.Recordset.AddNew
        data_imagen.Recordset("fecha") = Date
        data_imagen.Recordset("nrofact") = labnrofac.Caption
        data_imagen.Recordset("serie") = labserie.Caption
        Picture1.Picture = LoadPicture(App.path & "\qr.bmp")
        data_imagen.Recordset.Update
        data_imagen.Refresh
        data_cabeza2.Recordset.Edit
        labvencecnv.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Vencimiento)
        labautoriza.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Autorizacion)
        labrango.Caption = labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroDesde) & " - " & labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroHasta)
        labcodseg.Caption = CStr(ResultadoCfe.EstadoCfe.CodigoSeguridad)
        If Len(labvencecnv.Caption) = 8 Then
           labvenceok.Caption = Mid(labvencecnv.Caption, 7, 2) & "/" & Mid(labvencecnv.Caption, 5, 2) & "/" & Mid(labvencecnv.Caption, 1, 4)
        Else
           labvenceok.Caption = "31/12/2016"
        End If
        data_cabeza2.Recordset("cl_fultpag") = CDate(labvenceok.Caption)
        data_cabeza2.Recordset("cl_nrocobr") = Val(labautoriza.Caption)
        data_cabeza2.Recordset("cl_medflia") = Trim(labrango.Caption)
        data_cabeza2.Recordset("cl_fax") = Trim(labcodseg.Caption)
        data_cabeza2.Recordset("cl_socmnro") = labserie.Caption
        data_cabeza2.Recordset("cl_numero") = Val(labnrofac.Caption)
        data_imagen.RecordSource = "Select * from qr where nrofact =" & labnrofac.Caption & " and serie ='" & labserie.Caption & "'"
        data_imagen.Refresh
        If data_imagen.Recordset.RecordCount > 0 Then
           data_cabeza2.Recordset("qr") = data_imagen.Recordset("qr")
        End If
        data_cabeza2.Recordset.Update
        strUltimoGuid = ResultadoCfe.EstadoCfe.Guid
        Set objUltimaSerieNumero = ResultadoCfe.EstadoCfe.SerieNumeroCfe
    Else
        MsgBox "El comprobante no pasa los controles de facturación, no se puede continuar.", vbInformation
        MsgBox "Pruebe realizar la factura nuevamente", vbExclamation
        End
    End If
'    MsgBox "SON:" & labserie.Caption & " " & labnrofac.Caption
'    MsgBox "Serie: " & ResultadoCfe.EstadoCfe.SerieNumeroCfe.serie & vbNewLine & _
'        "Numero: " & CStr(ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero) & vbNewLine & _
'        "CAE autorización: " & ResultadoCfe.EstadoCfe.DatosCae.Autorizacion & vbNewLine & _
'        "CAE vencimiento: " & ResultadoCfe.EstadoCfe.DatosCae.Vencimiento & vbNewLine & _
'        "CAE numero desde: " & ResultadoCfe.EstadoCfe.DatosCae.NumeroDesde & vbNewLine & _
'        "CAE numero desde: " & ResultadoCfe.EstadoCfe.DatosCae.NumeroHasta & vbNewLine & _
'        "Contenido QR: " & ResultadoCfe.EstadoCfe.DatosQr & vbNewLine & _
'        "Código de seguridad: " & ResultadoCfe.EstadoCfe.CodigoSeguridad & vbNewLine & _
'        "Código de respuesta: " & ResultadoCfe.EstadoCfe.CodigoRespuesta & vbNewLine & _
'        "Fecha de firma: " & ResultadoCfe.EstadoCfe.FechaFirma & vbNewLine & _
'        "GUID: " & ResultadoCfe.EstadoCfe.Guid & vbNewLine & _
'        "Mensaje: " & ResultadoCfe.EstadoCfe.Mensaje & vbNewLine & _
'        "Pendiente de envío: " & CStr(ResultadoCfe.EstadoCfe.PendienteDeEnvio) & vbNewLine


    'cmdConsultaXguid.Enabled = True
    'cmdConsultaXnumero.Enabled = True
Exit Sub

Aldespinfcfe:
             MsgBox "Error en la factura. AVISE A INFORMÀTICA! " & Err.Number & " " & Err.Description, vbInformation
             data_errfact.Recordset.AddNew
             data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
             data_errfact.Recordset("fecha") = Date
             data_errfact.Recordset("hora") = Format(Time, "HH:mm")
             data_errfact.Recordset("nroerr") = Err.Number
             data_errfact.Recordset("desc") = "Desplegar info CFE"
             data_errfact.Recordset.Update
             Unload Me

End Sub



Public Sub Consulta_tiquet()
Dim Xsqlpromo, XCofirmar As String
Dim Xrecclii As New ADODB.Recordset
Dim XnoHayDatos As Integer
Dim Xtotal, Xiva, Xtotconiva As Double
Xtotconiva = 0
Xtotal = 0
XnoHayDatos = 0
ConectarBD
ConbdSapp.Open
If Trim(frm_convenios.cbogrupoap.Text) <> "" Then
   Xsqlpromo = "Select * from convenio_tiquets where nom_grupo ='" & frm_convenios.cbogrupoap.Text & "' and fecha_pago is null"
Else
   If Trim(frm_convenios.t_cantlla.Text) <> "" Then
      Xsqlpromo = "Select * from convenio_tiquets where nom_grupo ='" & frm_convenios.txt_cod.Text & "' and fecha_pago is null"
   Else
      XnoHayDatos = 9
   End If
End If
If XnoHayDatos <> 9 Then
   With Xrecclii
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   If Xrecclii.RecordCount > 0 Then
      Xrecclii.MoveFirst
      Do While Not Xrecclii.EOF()
         Xtotal = Xtotal + Xrecclii("importe")
         Xrecclii.MoveNext
      Loop
      Xtotconiva = Xtotal
      Xiva = Xtotal * 0.1 / 1.1
      Xiva = Format(Xiva, "Standard")
      Xtotal = Format(Xtotal, "Standard") - Format(Xiva, "Standard")
      
      labcanttiquet.Caption = Xrecclii.RecordCount
      MsgBox "El convenio tiene " & Xrecclii.RecordCount & " asistencias pendiente de pago por un total de: " & Val(Xtotconiva) & " Pesos.", vbExclamation, "Facturación"
      labimptiquet.Caption = Format(Xtotal, "Standard")
      Combo1.AddItem "COSTO POR LLAMADOS"
      labimptiquet.Caption = Format(Xtotal, "Standard") / Val(labcanttiquet.Caption)
   Else
      labimptiquet.Caption = ""
      labcanttiquet.Caption = ""
   End If
   Xrecclii.Close
Else
   labimptiquet.Caption = ""
   labcanttiquet.Caption = ""
End If

ConbdSapp.Close

End Sub
