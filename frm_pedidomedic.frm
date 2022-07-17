VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_pedidomedic 
   BackColor       =   &H00404040&
   Caption         =   "Sistema de pedidos MEDICACION a DOMICILIO"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13245
   Icon            =   "frm_pedidomedic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13245
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_arancel 
      Caption         =   "data_arancel"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_pedlin 
      Caption         =   "data_pedlin"
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
      Top             =   8040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_pedido 
      Caption         =   "data_pedido"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8115
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_param 
      Caption         =   "data_param"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton b_inf 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      Picture         =   "frm_pedidomedic.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Informes"
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_pedidomedic.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Consultar pedidos ingresados"
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_pedidomedic.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Cancelar datos y no grabar."
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton b_grabar 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_pedidomedic.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Grabar datos"
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_pedidomedic.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Editar registro seleccionado"
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_pedidomedic.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Nuevo registro"
      Top             =   7800
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos del pedido"
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
      ForeColor       =   &H0000FF00&
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12735
      Begin VB.TextBox t_det 
         Height          =   1095
         Left            =   10320
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Data data_pedexcel 
         Caption         =   "data_pedexcel"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4800
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton b_conscli 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6000
         Picture         =   "frm_pedidomedic.frx":26C6
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Consultar socio por cédula o matrícula"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton b_elimmed 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         Picture         =   "frm_pedidomedic.frx":2C50
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Borrar el medicamento seleccionado en la lista"
         Top             =   6000
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   240
         TabIndex        =   71
         Top             =   4560
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   12632064
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "Medicamento"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "c"
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox t_recib2 
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
         MaxLength       =   60
         TabIndex        =   69
         Top             =   6960
         Width           =   3975
      End
      Begin VB.TextBox t_cadete 
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
         Left            =   2640
         MaxLength       =   60
         TabIndex        =   61
         Top             =   6960
         Width           =   3615
      End
      Begin MSMask.MaskEdBox mhestado 
         Height          =   375
         Left            =   7560
         TabIndex        =   57
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox mfestado 
         Height          =   375
         Left            =   6240
         TabIndex        =   56
         Top             =   6480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.ComboBox cboestado 
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
         ItemData        =   "frm_pedidomedic.frx":31DA
         Left            =   2640
         List            =   "frm_pedidomedic.frx":31EA
         TabIndex        =   54
         Text            =   "cboestado"
         Top             =   6480
         Width           =   2175
      End
      Begin VB.TextBox t_imp 
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
         Left            =   8880
         TabIndex        =   49
         Top             =   4200
         Width           =   1215
      End
      Begin VB.ComboBox cbofpago 
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
         ItemData        =   "frm_pedidomedic.frx":321C
         Left            =   10320
         List            =   "frm_pedidomedic.frx":3229
         TabIndex        =   47
         Text            =   "cbofpago"
         Top             =   5640
         Width           =   2055
      End
      Begin VB.CommandButton b_newmedic 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         Picture         =   "frm_pedidomedic.frx":3250
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Agregar medicamento a la lista"
         Top             =   5640
         Width           =   375
      End
      Begin VB.TextBox t_cant 
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
         Left            =   6960
         TabIndex        =   44
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox t_medic 
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
         MaxLength       =   80
         TabIndex        =   42
         Top             =   4200
         Width           =   3975
      End
      Begin VB.ComboBox cborecctrol 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         ItemData        =   "frm_pedidomedic.frx":37DA
         Left            =   10440
         List            =   "frm_pedidomedic.frx":37E7
         TabIndex        =   40
         Text            =   "cborecctrol"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.ComboBox cborece 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         ItemData        =   "frm_pedidomedic.frx":3809
         Left            =   6360
         List            =   "frm_pedidomedic.frx":3813
         TabIndex        =   38
         Text            =   "cborece"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox cboh2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4080
         TabIndex        =   36
         Text            =   "cboh2"
         Top             =   3600
         Width           =   855
      End
      Begin VB.ComboBox cboh1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3120
         TabIndex        =   35
         Text            =   "cboh1"
         Top             =   3600
         Width           =   855
      End
      Begin MSMask.MaskEdBox mfent 
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.TextBox t_recib1 
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
         MaxLength       =   60
         TabIndex        =   32
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox t_correo 
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
         MaxLength       =   150
         TabIndex        =   30
         Top             =   3000
         Width           =   4935
      End
      Begin VB.TextBox t_telfs 
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
         MaxLength       =   80
         TabIndex        =   28
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox cbozona 
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
         Left            =   8400
         TabIndex        =   26
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox t_direc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox t_nombre 
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
         Left            =   8160
         MaxLength       =   100
         TabIndex        =   17
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox t_mat 
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
         Left            =   4320
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox t_codced 
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
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox t_ced 
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
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalle recetas"
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
         Left            =   10320
         TabIndex        =   77
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label laberror 
         Height          =   375
         Left            =   2760
         TabIndex        =   76
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label labalta 
         Height          =   375
         Left            =   360
         TabIndex        =   75
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label labcodzon 
         Height          =   375
         Left            =   7680
         TabIndex        =   74
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label labcance 
         Height          =   735
         Left            =   840
         TabIndex        =   70
         Top             =   5880
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre de quién recibió"
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
         Left            =   6600
         TabIndex        =   68
         Top             =   6960
         Width           =   1815
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre de cadete:"
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
         Left            =   240
         TabIndex        =   60
         Top             =   6960
         Width           =   2415
      End
      Begin VB.Label labusuario 
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
         Height          =   375
         Left            =   9960
         TabIndex        =   59
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario:"
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
         Left            =   8760
         TabIndex        =   58
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha/hora:"
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
         Left            =   5040
         TabIndex        =   55
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Estado actual del pedido:"
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
         Left            =   240
         TabIndex        =   53
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Label labtotp 
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
         Left            =   5880
         TabIndex        =   52
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label labtotcant 
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
         Left            =   5040
         TabIndex        =   51
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Totales:"
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
         Left            =   4080
         TabIndex        =   50
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Importe unitario:"
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
         Left            =   7680
         TabIndex        =   48
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forma de pago:"
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
         Left            =   7680
         TabIndex        =   46
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cantidad:"
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
         Left            =   5880
         TabIndex        =   43
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Medicación:"
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
         Left            =   240
         TabIndex        =   41
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tiene recetas de medicación controlada?"
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
         Left            =   7920
         TabIndex        =   39
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recetas en:"
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
         Left            =   5160
         TabIndex        =   37
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha y hora de entrega:"
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
         Left            =   240
         TabIndex        =   33
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recibe:"
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
         Left            =   6960
         TabIndex        =   31
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Correo:"
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
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Teléfonos:"
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
         Left            =   6960
         TabIndex        =   27
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6960
         TabIndex        =   25
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label10 
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
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label labmutual 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9720
         TabIndex        =   22
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8160
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label labnomconv 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label labcodconv 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Convenio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   6600
         TabIndex        =   16
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label labpedido 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   11040
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedido Nro."
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
         Left            =   9480
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label labusua 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   7320
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario:"
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
         Left            =   6240
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label labbase 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   5520
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Base:"
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
         Left            =   4680
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Label labhor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hora:"
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
         Left            =   2760
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label labfec 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   10200
      Picture         =   "frm_pedidomedic.frx":3824
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   2775
   End
End
Attribute VB_Name = "frm_pedidomedic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
Des_boton
Frame1.Enabled = True
Borrar_campos
labfec.Caption = Format(Date, "dd/mm/yyyy")
labhor.Caption = Format(Time, "HH:mm")
labbase.Caption = frm_menu.data_parse.Recordset("base")
labusua.Caption = WElusuario
labpedido.Caption = data_param.Recordset("p_pedidom") + 1
data_param.Recordset.Edit
data_param.Recordset("p_pedidom") = data_param.Recordset("p_pedidom") + 1
data_param.Recordset.Update
labalta.Caption = 1

t_ced.SetFocus


End Sub

Private Sub b_busca_Click()

frm_pedidomedicb.Show

End Sub

Private Sub b_cancela_Click()

Borrar_campos
Hab_boton
labalta.Caption = ""

Frame1.Enabled = False

End Sub

Private Sub b_conscli_Click()
If t_ced.Text <> "" Then
   
   t_nombre.SetFocus
   Consulta_cliCed
   
'          If Xquehag = 9 Then
'             MsgBox "Socio moroso, se pasará al sistema de autorización automática", vbCritical
'             Xhab = Val(labmatri.Caption)
'             frm_autoriza.Show vbModal
'             Xelcodigoaut = InputBox("INGRESE CÓDIGO DE AUTORIZACIÓN:", "AUTORIZACIÓN", Wopszond)
'             If Trim(Xelcodigoaut) <> "" Then
'                data_aut.RecordSource = "select * from Codigos_aut where codaut ='" & Trim(Xelcodigoaut) & "' and socio =" & Val(labmatri.Caption)
'                data_aut.Refresh
'                If data_aut.Recordset.RecordCount > 0 Then
'                   Xquehag = 0
'                Else
'                   MsgBox "ATENCION! No se encuentra código de autorización, realice nuevamente la autorización o comunique a Administración", vbCritical
'                   Xquehag = 9
'                End If
'             Else
'                MsgBox "Socio moroso, debe ingresar autorización.", vbCritical
'                Xquehag = 9
 '            End If
   
   Consultar_pedidos
   Consulta_clideudas
   If Trim(labcodconv.Caption) <> "" Then
      Consulta_precio
   End If
   Consultar_pedidoApr
Else
   If t_mat.Text <> "" Then
      t_nombre.SetFocus
      Consulta_cliMat
      Consultar_pedidos
      Consulta_clideudas
      If Trim(labcodconv.Caption) <> "" Then
         Consulta_precio
      End If
      Consultar_pedidoApr
   End If
End If


End Sub

Private Sub b_edita_Click()

If Trim(labpedido.Caption) <> "" And Trim(labfec.Caption) <> "" And Trim(labhor.Caption) <> "" Then
   If cboestado.Text = "Entregado" Or cboestado.Text = "Cancelado" Then
      MsgBox "El pedido está cerrado, no se puede modificar.", vbInformation
   Else
      Frame1.Enabled = True
      labalta.Caption = 2
      Des_boton
   End If
Else
   MsgBox "No hay registro seleccionado.", vbCritical
End If


End Sub

Private Sub b_elimmed_Click()

Dim Xind, Xcant, Xnro As Long
Dim Xborralaconsulta As String
Dim Xmedicam As String
Dim Xcount As Integer
Dim Xcantmenos As Integer
Dim Xmontomenos As Integer

Xcount = 1
Xmedicam = ""

Xborralaconsulta = MsgBox("Desea borrar el dato seleccionado?", vbInformation + vbYesNo)
If Xborralaconsulta = vbYes Then
   Xind = 0
   Xnro = 0
   Xcant = 0
   Dim Xcountt As Long
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
             Xmedicam = ListView1.ListItems(Xind).Text
             Data_pedlin.RecordSource = "select * from pedidos_mediclin where cod_pedido =" & Val(labpedido.Caption) & " and nom_medic ='" & Xmedicam & "'"
             Data_pedlin.Refresh
             If Data_pedlin.Recordset.RecordCount > 0 Then
                Data_pedlin.Recordset.Delete
                Data_pedlin.Refresh
             End If
          End If
      Next
      Xind = 1
      labtotcant.Caption = "0"
      labtotp.Caption = "0"
      Data_pedlin.RecordSource = "select * from pedidos_mediclin where cod_pedido =" & Val(labpedido.Caption)
      Data_pedlin.Refresh
      ListView1.ListItems.Clear
      If Data_pedlin.Recordset.RecordCount > 0 Then
         Data_pedlin.Recordset.MoveFirst
         Do While Not Data_pedlin.Recordset.EOF
            ListView1.ListItems.Add Xcountt, , Data_pedlin.Recordset("nom_medic")
            ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Data_pedlin.Recordset("cant")
            ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Format(Data_pedlin.Recordset("tot_imp"), "Standard")
            labtotcant.Caption = Val(labtotcant.Caption) + Data_pedlin.Recordset("cant")
            labtotp.Caption = Val(labtotp.Caption) + Data_pedlin.Recordset("tot_imp")
            Data_pedlin.Recordset.MoveNext
            Xcountt = Xcountt + 1
         Loop
      End If
   Else
      MsgBox "Debe seleccionar un solo registro."
   End If
End If




End Sub

Private Sub b_grabar_Click()
laberror.Caption = 0

If Trim(cbozona.Text) <> "" Then
   Consulta_zona
   If Trim(labcodzon.Caption) = "" Then
      laberror.Caption = 3
   Else
   End If
End If

If cboh1.Text = "10:00" Or cboh1.Text = "10:30" Or cboh1.Text = "11:00" Or cboh1.Text = "11:30" Or cboh1.Text = "12:00" Or _
   cboh1.Text = "12:30" Or cboh1.Text = "13:00" Or cboh1.Text = "13:30" Or cboh1.Text = "14:00" Or cboh1.Text = "14:30" Or _
   cboh1.Text = "15:00" Or cboh1.Text = "15:30" Or cboh1.Text = "16:00" Or cboh1.Text = "16:30" Or cboh1.Text = "17:00" Or _
   cboh1.Text = "17:30" Or cboh1.Text = "18:00" Then
Else
   laberror.Caption = 3
End If
If cboh2.Text = "10:00" Or cboh2.Text = "10:30" Or cboh2.Text = "11:00" Or cboh2.Text = "11:30" Or cboh2.Text = "12:00" Or _
   cboh2.Text = "12:30" Or cboh2.Text = "13:00" Or cboh2.Text = "13:30" Or cboh2.Text = "14:00" Or cboh2.Text = "14:30" Or _
   cboh2.Text = "15:00" Or cboh2.Text = "15:30" Or cboh2.Text = "16:00" Or cboh2.Text = "16:30" Or cboh2.Text = "17:00" Or _
   cboh2.Text = "17:30" Or cboh2.Text = "18:00" Then
Else
   laberror.Caption = 3
End If
If cborece.Text = "Socio" Or cborece.Text = "SAPP" Then
Else
   laberror.Caption = 3
End If

If Trim(t_ced.Text) = "" Then
   laberror.Caption = 3
End If
If Trim(t_codced.Text) = "" Then
   laberror.Caption = 3
End If
If Trim(t_telfs.Text) = "" Then
   laberror.Caption = 3
End If
If Trim(t_nombre.Text) = "" Then
   laberror.Caption = 3
End If
If Trim(t_direc.Text) = "" Then
   laberror.Caption = 3
End If
If Trim(cbozona.Text) = "" Then
   laberror.Caption = 3
End If
If Trim(t_recib1.Text) = "" Then
   laberror.Caption = 3
End If
If mfent.Text = "__/__/____" Then
   laberror.Caption = 3
End If
If Trim(cborecctrol.Text) = "" Or cborecctrol.Text = "Verde" Or cborecctrol.Text = "Naranja" Or cborecctrol.Text = "Doble blanca" Then
Else
   laberror.Caption = 3
End If
If ListView1.ListItems.count <= 0 Then
   laberror.Caption = 3
End If
If t_medic.Text <> "" Then
   laberror.Caption = 3
End If
If cbofpago.Text = "Efectivo" Or cbofpago.Text = "Efectivo+Vales" Or cbofpago.Text = "Tarjeta" Then
Else
   laberror.Caption = 3
End If
If cboestado.Text = "Pendiente" Or cboestado.Text = "En tránsito" Or cboestado.Text = "Entregado" Or cboestado.Text = "Cancelado" Then
Else
   laberror.Caption = 3
End If
If cboestado.Text = "En tránsito" Or cboestado.Text = "Entregado" Then
   If cboestado.Text = "En tránsito" Then
      If mfestado.Text = "__/__/____" Or mfestado.Text = "__:__" Or Trim(t_cadete.Text) = "" Then
         laberror.Caption = 3
      End If
   Else
      If mfestado.Text = "__/__/____" Or mfestado.Text = "__:__" Or Trim(t_cadete.Text) = "" Or Trim(t_recib2.Text) = "" Then
         laberror.Caption = 3
      End If
   End If
End If
      
If Val(laberror.Caption) = 3 Then
   MsgBox "Faltan datos, verifique!", vbCritical
Else
   If Val(labalta.Caption) = 1 Then
      data_pedido.Recordset.AddNew
      data_pedido.Recordset("fecha") = CDate(labfec.Caption)
      data_pedido.Recordset("hora") = Format(labhor.Caption, "HH:mm")
      data_pedido.Recordset("base") = Val(labbase.Caption)
      data_pedido.Recordset("usuario") = WElusuario
      data_pedido.Recordset("pedido_nro") = Val(labpedido.Caption)
      data_pedido.Recordset("cedula") = t_ced.Text
      data_pedido.Recordset("codced") = t_codced.Text
      data_pedido.Recordset("matricula") = t_mat.Text
      data_pedido.Recordset("nombre") = t_nombre.Text
      data_pedido.Recordset("codconv") = labcodconv.Caption
      data_pedido.Recordset("mutual") = labmutual.Caption
      data_pedido.Recordset("direcc") = t_direc.Text
      If Trim(cbozona.Text) <> "" Then
         data_pedido.Recordset("zona") = Mid(cbozona.Text, 1, 60)
      End If
      data_pedido.Recordset("telefs") = t_telfs.Text
      If Trim(t_correo.Text) <> "" Then
         data_pedido.Recordset("correo") = t_correo.Text
      End If
      data_pedido.Recordset("recibe1") = t_recib1.Text
      data_pedido.Recordset("fec_ent") = mfent.Text
      data_pedido.Recordset("hor_ent1") = cboh1.Text
      data_pedido.Recordset("hor_ent2") = cboh2.Text
      data_pedido.Recordset("recetasen") = cborece.Text
      If cborecctrol.ListIndex >= 0 Then
         data_pedido.Recordset("recetascont") = cborecctrol.Text
      End If
      data_pedido.Recordset("tot_cant") = Val(labtotcant.Caption)
      data_pedido.Recordset("tot_pesos") = Val(labtotp.Caption)
      data_pedido.Recordset("forma_pago") = cbofpago.Text
      data_pedido.Recordset("estado") = cboestado.Text
      If mfestado.Text <> "__/__/____" Then
         data_pedido.Recordset("fec_estado") = mfestado.Text
      End If
      If mhestado.Text <> "__:__" Then
         data_pedido.Recordset("hor_estado") = mhestado.Text
      End If
      If labusuario.Caption <> "" Then
         data_pedido.Recordset("usua_estado") = labusuario.Caption
      End If
      If t_cadete.Text <> "" Then
         data_pedido.Recordset("nom_cadete") = t_cadete.Text
      End If
      If t_recib2.Text <> "" Then
         data_pedido.Recordset("nom_recibe") = t_recib2.Text
      End If
      If t_det.Text <> "" Then
         data_pedido.Recordset("recetas_obs") = t_det.Text
      End If
      data_pedido.Recordset.Update
      MsgBox "Registro grabado correctamente.", vbInformation, "Pedidos de medicación"
      labalta.Caption = ""
      Borrar_campos
      Hab_boton
      Frame1.Enabled = False
      b_busca_Click
   Else
      If Trim(labpedido.Caption) <> "" Then
         data_pedido.RecordSource = "select * from pedidos_medic where pedido_nro =" & Val(labpedido.Caption)
         data_pedido.Refresh
         If data_pedido.Recordset.RecordCount > 0 Then
            data_pedido.Recordset.MoveFirst
            data_pedido.Recordset.Edit
            data_pedido.Recordset("direcc") = t_direc.Text
            If Trim(cbozona.Text) <> "" Then
               data_pedido.Recordset("zona") = Mid(cbozona.Text, 1, 60)
            End If
            data_pedido.Recordset("telefs") = t_telfs.Text
            If Trim(t_correo.Text) <> "" Then
               data_pedido.Recordset("correo") = t_correo.Text
            End If
            data_pedido.Recordset("recibe1") = t_recib1.Text
            data_pedido.Recordset("fec_ent") = mfent.Text
            data_pedido.Recordset("hor_ent1") = cboh1.Text
            data_pedido.Recordset("hor_ent2") = cboh2.Text
            data_pedido.Recordset("recetasen") = cborece.Text
            If cborecctrol.ListIndex >= 0 Then
               data_pedido.Recordset("recetascont") = cborecctrol.Text
            End If
            data_pedido.Recordset("tot_cant") = Val(labtotcant.Caption)
            data_pedido.Recordset("tot_pesos") = Val(labtotp.Caption)
            data_pedido.Recordset("forma_pago") = cbofpago.Text
            data_pedido.Recordset("estado") = cboestado.Text
            If mfestado.Text <> "__/__/____" Then
               data_pedido.Recordset("fec_estado") = mfestado.Text
            End If
            If mhestado.Text <> "__:__" Then
               data_pedido.Recordset("hor_estado") = mhestado.Text
            End If
            If labusuario.Caption <> "" Then
               data_pedido.Recordset("usua_estado") = labusuario.Caption
            End If
            If t_cadete.Text <> "" Then
               data_pedido.Recordset("nom_cadete") = t_cadete.Text
            End If
            If t_recib2.Text <> "" Then
               data_pedido.Recordset("nom_recibe") = t_recib2.Text
            End If
            If t_det.Text <> "" Then
               data_pedido.Recordset("recetas_obs") = t_det.Text
            End If
            data_pedido.Recordset.Update
            MsgBox "Registro grabado correctamente.", vbInformation, "Pedidos de medicación"
            labalta.Caption = ""
            Borrar_campos
            Hab_boton
            Frame1.Enabled = False
            b_busca_Click
         End If
      End If
   End If
End If


End Sub

Private Sub b_inf_Click()
Dim desde, hasta, Promo As String
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook
Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir3 As New Excel.Application
Dim Fecha1, Fecha2 As String
If Month(Date) < 10 Then
   Fecha1 = "01/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   If Day(Date) < 10 Then
      Fecha2 = "0" & Trim(str(Day(Date))) & "/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   Else
      Fecha2 = Trim(str(Day(Date))) & "/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   End If
Else
   Fecha1 = "01/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   If Day(Date) < 10 Then
      Fecha2 = "0" & Trim(str(Day(Date))) & "/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   Else
      Fecha2 = Trim(str(Day(Date))) & "/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   End If

End If

desde = InputBox("Ingrese fecha de inicio (formato: DD/MM/AAAA):", "FECHA INICIAL", Fecha1)
hasta = InputBox("Ingrese fecha final (formato: DD/MM/AAAA):", "FECHA FINAL", Fecha2)
Promo = InputBox("INGRESE CERO PARA LISTAR TODO O UNO (1) PARA LISTADO PARA CADETE.", "PEDIDOS", 0)
If Trim(Promo) = "" Then
   Promo = "0"
End If
frm_pedidomedic.MousePointer = 11
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0

If desde <> "" And hasta <> "" Then
   If Val(Promo) = 0 Then
      data_pedexcel.RecordSource = "select * from pedidos_medic where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# order by fecha"
      data_pedexcel.Refresh
   Else
      data_pedexcel.RecordSource = "select * from pedidos_medic where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and estado in ('En tránsito') order by fecha"
      data_pedexcel.Refresh
   End If
   If data_pedexcel.Recordset.RecordCount > 0 Then
      data_pedexcel.Recordset.MoveFirst
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      Xarchexel22.Name = Trim("Pedidos")
      Xlibexel22.SaveAs ("C:\planillas\Pedidos.xls")
      Xarchtex = "C:\planillas\Pedidos.xls"
      Xarchexel22.Cells(Xlin, XCol) = "ECONOMATO SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "M" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      If Val(Promo) = 0 Then
         Xarchexel22.Cells(Xlin, XCol) = "INFORME DE PEDIDOS TOTALES DESDE: " & desde & " HASTA: " & hasta
      Else
         Xarchexel22.Cells(Xlin, XCol) = "INFORME DE PEDIDOS EN TRÁNSITO DESDE: " & desde & " HASTA: " & hasta
      End If
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
      Xarchexel22.Range("A" & Trim(str(Xlin)), "M" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "FECHA"
      XCol = XCol + 1
      Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
      XCol = XCol + 1
      Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
      XCol = XCol + 1
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "MUTUALISTA"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "CANTIDAD"
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 13
      Xarchexel22.Cells(Xlin, XCol) = "Sub-TOT. $."
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 13
      Xarchexel22.Cells(Xlin, XCol) = "COSTO ENVÍO"
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 16
      Xarchexel22.Cells(Xlin, XCol) = "TOTAL $."
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 16
      Xarchexel22.Cells(Xlin, XCol) = "F.PAGO"
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "BASE"
      XCol = XCol + 1
      Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 26
      Xarchexel22.Cells(Xlin, XCol) = "TELEFONOS"
      XCol = XCol + 1
      Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "ESTADO"
      XCol = XCol + 1
      Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "NRO.FACTURA"
      
      Xlin = Xlin + 1
      XCol = 1
        
      Do While Not data_pedexcel.Recordset.EOF
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_pedexcel.Recordset("fecha"), "dd/mm/yyyy")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("cedula")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("nombre")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("mutual")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("tot_cant")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("tot_pesos")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Devuelve_costoPed()
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("tot_pesos") + Devuelve_costoPed()
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("forma_pago")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("base")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("telefs")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("estado")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_pedexcel.Recordset("nro_factura")
         
         Xlin = Xlin + 1
         XCol = 1
         Xtotreg = Xtotreg + 1
         data_pedexcel.Recordset.MoveNext
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
      frm_pedidomedic.MousePointer = 0
      MsgBox "Terminado"
   Else
      frm_pedidomedic.MousePointer = 0
      MsgBox "No hay registros"
   End If
Else
   frm_pedidomedic.MousePointer = 0
   MsgBox "Faltan fechas"
End If

End Sub

Private Sub b_newmedic_Click()
Dim Xcount As Integer
Dim Xcant, Xyafigura As Integer
Dim Xtot, Xsubtot As Double
Dim Siyafigura As String
Siyafigura = ""
Xcant = 0
Xtot = 0
Xsubtot = 0
Xyafigura = 0

Xcount = ListView1.ListItems.count + 1
If labtotcant.Caption = "" Then
   labtotcant.Caption = "0"
End If
If labtotp.Caption = "" Then
   labtotp.Caption = "0"
End If
Data_pedlin.RecordSource = "select * from pedidos_mediclin where cod_pedido =" & Val(labpedido.Caption)
Data_pedlin.Refresh

If t_medic.Text <> "" Then
   Dim Xind22 As Long
   For Xind22 = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind22).Selected = True
       If Trim(ListView1.ListItems(Xind22).Text) = Trim(t_medic.Text) Then
          Siyafigura = MsgBox("El medicamento ya existe en la lista, desea agregar igual?", vbInformation + vbYesNo, "Pedidos")
          If Siyafigura = vbYes Then
             Xyafigura = 0
          Else
             Xyafigura = 1
          End If
       End If
   Next Xind22
End If

If Xyafigura = 0 Then
    If t_medic.Text <> "" Then
       If t_cant.Text <> "" Then
          If t_imp.Text <> "" Then
             ListView1.ListItems.Add Xcount, , t_medic.Text
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , t_cant.Text
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(t_imp.Text, "Standard")
             Xsubtot = Val(t_imp.Text) * Val(t_cant.Text)
             Xcant = Val(labtotcant.Caption)
             Xtot = Val(labtotp.Caption)
             Xcant = Xcant + Val(t_cant.Text)
             Xtot = Xtot + Xsubtot
             labtotcant.Caption = Xcant
             labtotp.Caption = Val(Xtot)
             Data_pedlin.Recordset.AddNew
             Data_pedlin.Recordset("cod_pedido") = Val(labpedido.Caption)
             Data_pedlin.Recordset("nom_medic") = Trim(t_medic.Text)
             Data_pedlin.Recordset("cant") = Val(t_cant.Text)
             Data_pedlin.Recordset("imp_unit") = Val(t_imp.Text)
             Data_pedlin.Recordset("tot_imp") = Val(Xsubtot)
             Data_pedlin.Recordset.Update
             
             t_medic.Text = ""
             t_cant.Text = 1
             t_medic.SetFocus
             
          End If
       End If
    End If
End If

         
End Sub

Private Sub cboestado_Click()
labcance.Caption = ""
If cboestado.Text = "Cancelado" Then
   labcance.Caption = InputBox("Ingrese motivo de cancelación:", "PEDIDOS")
   labusuario.Caption = WElusuario
Else
   If cboestado.Text = "Entregado" Then
      labusuario.Caption = WElusuario
   Else
      If cboestado.Text = "En tránsito" Then
         labusuario.Caption = WElusuario
      Else
         labusuario.Caption = WElusuario
      End If
   End If
End If
   
End Sub

Private Sub cboestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If mfestado.Text = "__/__/____" Then
      mfestado.Text = Format(Date, "dd/mm/yyyy")
   End If
   mfestado.SetFocus
End If

End Sub

Private Sub cbofpago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboestado.SetFocus
End If

End Sub

Private Sub cboh1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboh2.SetFocus
End If

End Sub

Private Sub cboh2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cborece.SetFocus
End If

End Sub

Private Sub cborecctrol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_medic.SetFocus
End If

End Sub

Private Sub cborece_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cborecctrol.SetFocus
End If

End Sub

Private Sub cbozona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_telfs.SetFocus
End If

End Sub

Private Sub Form_Load()

data_pedido.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_pedido.RecordSource = "pedidos_medic"
data_pedido.Refresh

Data_pedlin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_pedexcel.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_param.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_param.RecordSource = "param_gral"
data_param.Refresh


cboh1.AddItem "10:00"
cboh1.AddItem "10:30"
cboh1.AddItem "11:00"
cboh1.AddItem "11:30"
cboh1.AddItem "12:00"
cboh1.AddItem "12:30"
cboh1.AddItem "13:00"
cboh1.AddItem "13:30"
cboh1.AddItem "14:00"
cboh1.AddItem "14:30"
cboh1.AddItem "15:00"
cboh1.AddItem "15:30"
cboh1.AddItem "16:00"
cboh1.AddItem "16:30"
cboh1.AddItem "17:00"
cboh1.AddItem "17:30"
cboh1.AddItem "18:00"
cboh2.AddItem "10:00"
cboh2.AddItem "10:30"
cboh2.AddItem "11:00"
cboh2.AddItem "11:30"
cboh2.AddItem "12:00"
cboh2.AddItem "12:30"
cboh2.AddItem "13:00"
cboh2.AddItem "13:30"
cboh2.AddItem "14:00"
cboh2.AddItem "14:30"
cboh2.AddItem "15:00"
cboh2.AddItem "15:30"
cboh2.AddItem "16:00"
cboh2.AddItem "16:30"
cboh2.AddItem "17:00"
cboh2.AddItem "17:30"
cboh2.AddItem "18:00"

Carga_zonas

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub

Private Sub mfent_GotFocus()
If mfent.Text = "__/__/____" Then
   If Format(Time, "HH:mm") >= "14:00" Then
      mfent.Text = Format(Date + 1, "dd/mm/yyyy")
   Else
      mfent.Text = Format(Date, "dd/mm/yyyy")
   End If
End If

End Sub

Private Sub mfent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboh1.SetFocus
End If

End Sub

Private Sub mfestado_GotFocus()
mfestado.Text = Format(Date, "dd/mm/yyyy")

End Sub

Private Sub mfestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If mhestado.Text = "__:__" Then
      mhestado.Text = Format(Time, "HH:mm")
   End If
   mhestado.SetFocus
End If

End Sub

Private Sub mhestado_GotFocus()
mhestado.Text = Format(Time, "HH:mm")

End Sub

Private Sub mhestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cadete.SetFocus
End If

End Sub

Private Sub t_cadete_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_recib2.SetFocus
End If

End Sub

Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_imp.SetFocus
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_conscli.SetFocus
End If

End Sub


Private Sub t_codced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mat.SetFocus
End If

End Sub

Private Sub t_correo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_recib1.SetFocus
End If

End Sub

Private Sub t_direc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   cbozona.SetFocus
End If

End Sub

Private Sub t_imp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_newmedic.SetFocus
End If

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nombre.SetFocus
End If

End Sub

Private Sub t_medic_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_cant.SetFocus
End If

End Sub

Private Sub t_medic_LostFocus()
   If t_cant = "" Then
      t_cant.Text = 1
   End If

End Sub

Private Sub t_nombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_direc.SetFocus
End If

End Sub

Public Sub Des_boton()
b_alta.Enabled = False
b_edita.Enabled = False
b_grabar.Enabled = True
b_cancela.Enabled = True
b_inf.Enabled = False
b_busca.Enabled = False
Frame1.Enabled = True

End Sub

Public Sub Hab_boton()
b_alta.Enabled = True
b_edita.Enabled = True
b_grabar.Enabled = False
b_cancela.Enabled = False
b_inf.Enabled = True
b_busca.Enabled = True
Frame1.Enabled = False

End Sub

Public Sub Borrar_campos()
labfec.Caption = ""
labhor.Caption = ""
labbase.Caption = ""
labusua.Caption = ""
labpedido.Caption = ""
t_ced.Text = ""
t_codced.Text = ""
t_mat.Text = ""
t_nombre.Text = ""
labcodconv.Caption = ""
labnomconv.Caption = ""
labmutual.Caption = ""
t_direc.Text = ""
cbozona.Text = ""
t_telfs.Text = ""
t_correo.Text = ""
t_recib1.Text = ""
mfent.Text = "__/__/____"
cboh1.Text = ""
cboh2.Text = ""
cborece.Text = ""
cborecctrol.Text = ""
t_medic.Text = ""
t_cant.Text = ""
t_imp.Text = ""
labtotcant.Caption = ""
labtotp.Caption = ""
cbofpago.Text = ""
cboestado.Text = ""
mfestado.Text = "__/__/____"
mhestado.Text = "__:__"
labusuario.Caption = ""
t_cadete.Text = ""
t_recib2.Text = ""
ListView1.ListItems.Clear
labcodzon.Caption = ""
t_det.Text = ""

End Sub

Private Sub t_recib1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   mfent.SetFocus
End If

End Sub

Private Sub t_recib2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   b_grabar.SetFocus
End If

End Sub

Private Sub t_telfs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_correo.SetFocus
End If

End Sub

Public Sub Consulta_cliCed()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pregunta As String

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select clientes.cl_codigo,clientes.cl_codconv,clientes.cl_nomconv,clientes.cl_cedula,clientes.cl_codced,clientes.estado,clientes.saldo_chc2," & _
"clientes.cl_apellid,clientes.cl_direcci,clientes.cl_dpto,clientes.cl_telefon,clientes.cl_zona,clientes.cl_referen," & _
"convenio.cnv_codigo,convenio.cnv_grupo from clientes inner join convenio on clientes.cl_codconv=convenio.cnv_codigo where clientes.cl_cedula =" & t_ced.Text
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   If IsNull(Xrecclii("saldo_chc2")) = False Then
      Sinservicio = Xrecclii("saldo_chc2")
   End If
   If Xrecclii("cl_codconv") = "APS" Then
      Sinservicio = 1
   End If
   If Xrecclii("estado") = 2 Or Xrecclii("estado") = 3 Or Sinservicio = 1 Then
      If Sinservicio = 1 Then
         MsgBox "ATENCION! Socio figura con servicios RESTRINGIDOS! Consulte en administración.", vbCritical
      Else
         MsgBox "ATENCION! Socio figura de BAJA, no se puede ingresar pedido!", vbCritical
      End If
      b_cancela_Click
   Else
        t_codced.Text = Xrecclii("cl_codced")
        t_mat.Text = Xrecclii("cl_codigo")
        t_nombre.Text = Xrecclii("cl_apellid")
        labcodconv.Caption = Xrecclii("cl_codconv")
        labnomconv.Caption = Xrecclii("cl_nomconv")
        If IsNull(Xrecclii("cnv_grupo")) = False Then
           labmutual.Caption = Xrecclii("cnv_grupo")
        Else
           labmutual.Caption = ""
        End If
        t_direc.Text = Xrecclii("cl_direcci")
        If IsNull(Xrecclii("cl_referen")) = False Then
           t_correo.Text = Xrecclii("cl_referen")
        Else
           t_correo.Text = ""
        End If
        If IsNull(Xrecclii("cl_zona")) = False Then
           cbozona.Text = Xrecclii("cl_zona")
        Else
           cbozona.Text = ""
        End If
        If IsNull(Xrecclii("cl_dpto")) = False Then
           If Trim(Xrecclii("cl_dpto")) = "NO APLICA" Then
              If IsNull(Xrecclii("cl_telefon")) = False Then
                 If Trim(Xrecclii("cl_telefon")) = "NO APLICA" Then
                    t_telfs.Text = ""
                 Else
                    t_telfs.Text = Xrecclii("cl_telefon")
                 End If
              Else
                 t_telfs.Text = ""
              End If
           Else
              If IsNull(Xrecclii("cl_telefon")) = False Then
                 If Trim(Xrecclii("cl_telefon")) = "NO APLICA" Then
                    t_telfs.Text = Xrecclii("cl_dpto")
                 Else
                    t_telfs.Text = Xrecclii("cl_dpto") & " Tel:" & Xrecclii("cl_telefon")
                 End If
              Else
                 t_telfs.Text = Xrecclii("cl_dpto")
              End If
           End If
        Else
           If IsNull(Xrecclii("cl_telefon")) = False Then
              If Trim(Xrecclii("cl_telefon")) = "NO APLICA" Then
                 t_telfs.Text = ""
              Else
                 t_telfs.Text = Xrecclii("cl_telefon")
              End If
           Else
              t_telfs.Text = ""
           End If
        End If
   End If
Else
   MsgBox "Socio no encontrado.", vbCritical
   t_ced.Text = ""
   t_codced.Text = ""
   t_mat.Text = ""
   t_nombre.Text = ""
   labcodconv.Caption = ""
   labnomconv.Caption = ""
   labmutual.Caption = ""
   t_direc.Text = ""
   t_correo.Text = ""
   cbozona.Text = ""
   t_telfs.Text = ""
   b_cancela_Click
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_cliMat()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pregunta As String
Dim Sinservicio As Integer

ConectarBD
ConbdSapp.Open
Sinservicio = 0

Xsqlpromo = "Select clientes.cl_codigo,clientes.cl_codconv,clientes.cl_nomconv,clientes.cl_cedula,clientes.cl_codced,clientes.estado,clientes.saldo_chc2," & _
"clientes.cl_grupo,clientes.cl_apellid,clientes.cl_direcci,clientes.cl_dpto,clientes.cl_telefon,clientes.cl_zona,clientes.cl_referen," & _
"convenio.cnv_codigo,convenio.cnv_grupo from clientes inner join convenio on clientes.cl_codconv=convenio.cnv_codigo where clientes.cl_codigo =" & t_mat.Text
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   If IsNull(Xrecclii("saldo_chc2")) = False Then
      Sinservicio = Xrecclii("saldo_chc2")
   End If
   If Xrecclii("estado") = 2 Or Xrecclii("estado") = 3 Or Sinservicio = 1 Then
      If Sinservicio = 1 Then
         MsgBox "ATENCION! Socio figura con servicios RESTRINGIDOS!", vbCritical
      Else
         MsgBox "ATENCION! Socio figura de BAJA, no se puede ingresar pedido!", vbCritical
      End If
      b_cancela_Click
   Else
        t_ced.Text = Xrecclii("cl_cedula")
        t_codced.Text = Xrecclii("cl_codced")
        t_nombre.Text = Xrecclii("cl_apellid")
        labcodconv.Caption = Xrecclii("cl_codconv")
        labnomconv.Caption = Xrecclii("cl_nomconv")
        If IsNull(Xrecclii("cl_grupo")) = False Then
           labmutual.Caption = Xrecclii("cl_grupo")
        Else
           labmutual.Caption = ""
        End If
        t_direc.Text = Xrecclii("cl_direcci")
        If IsNull(Xrecclii("cl_referen")) = False Then
           t_correo.Text = Xrecclii("cl_referen")
        Else
           t_correo.Text = ""
        End If
        If IsNull(Xrecclii("cl_zona")) = False Then
           cbozona.Text = Xrecclii("cl_zona")
        Else
           cbozona.Text = ""
        End If
        If IsNull(Xrecclii("cl_dpto")) = False Then
           If Trim(Xrecclii("cl_dpto")) = "NO APLICA" Then
              If IsNull(Xrecclii("cl_telefon")) = False Then
                 If Trim(Xrecclii("cl_telefon")) = "NO APLICA" Then
                    t_telfs.Text = ""
                 Else
                    t_telfs.Text = Xrecclii("cl_telefon")
                 End If
              Else
                 t_telfs.Text = ""
              End If
           Else
              If IsNull(Xrecclii("cl_telefon")) = False Then
                 If Trim(Xrecclii("cl_telefon")) = "NO APLICA" Then
                    t_telfs.Text = Xrecclii("cl_dpto")
                 Else
                    t_telfs.Text = Xrecclii("cl_dpto") & " Tel:" & Xrecclii("cl_telefon")
                 End If
              Else
                 t_telfs.Text = Xrecclii("cl_dpto")
              End If
           End If
        Else
           If IsNull(Xrecclii("cl_telefon")) = False Then
              If Trim(Xrecclii("cl_telefon")) = "NO APLICA" Then
                 t_telfs.Text = ""
              Else
                 t_telfs.Text = Xrecclii("cl_telefon")
              End If
           Else
              t_telfs.Text = ""
           End If
        End If
   End If
Else
   MsgBox "Socio no encontrado.", vbCritical
   t_ced.Text = ""
   t_codced.Text = ""
   t_mat.Text = ""
   t_nombre.Text = ""
   labcodconv.Caption = ""
   labnomconv.Caption = ""
   labmutual.Caption = ""
   t_direc.Text = ""
   t_correo.Text = ""
   cbozona.Text = ""
   t_telfs.Text = ""
   b_cancela_Click
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_zona()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pregunta As String

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from zonas where zo_nombre ='" & Trim(cbozona.Text) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcodzon.Caption = Xrecclii("zo_grupo")
   cbozona.Text = Xrecclii("zo_nombre")
Else
   labcodzon.Caption = ""
   cbozona.Text = ""
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


Public Sub Consulta_precio()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xopcion As Integer

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from convenio where cnv_codigo ='" & Trim(labcodconv.Caption) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   If IsNull(Xrecclii("cnv_aran")) = False Then
      Xopcion = Xrecclii("cnv_aran")
   Else
      Xopcion = 0
   End If
Else
   Xopcion = 0
End If
Xrecclii.Close
If Xopcion <> 0 Then
   If labmutual.Caption = "CCOU" Then
      Xsqlpromo = "Select * from Aran_servicios where id_gpo =" & Xopcion & " and id_serv =" & 60103
   Else
      If labmutual.Caption = "SMI" Then
         Xsqlpromo = "Select * from Aran_servicios where id_gpo =" & Xopcion & " and id_serv =" & 60106
      Else
         If labmutual.Caption = "UNIVERSAL" Then
            Xsqlpromo = "Select * from Aran_servicios where id_gpo =" & Xopcion & " and id_serv =" & 60108
         Else
            If labmutual.Caption = "H.EVANGELICO" Then
               Xsqlpromo = "Select * from Aran_servicios where id_gpo =" & Xopcion & " and id_serv =" & 60107
            Else
               If labmutual.Caption = "CASA DE GALICIA" Then
                  Xsqlpromo = "Select * from Aran_servicios where id_gpo =" & Xopcion & " and id_serv =" & 60109
               Else
                  Xsqlpromo = "Select * from Aran_servicios where id_gpo =" & Xopcion & " and id_serv =" & 60103
               End If
            End If
         End If
      End If
   End If
   With Xrecclii
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   If Xrecclii.RecordCount > 0 Then
      t_imp.Text = Val(Xrecclii("prec_serv"))
   Else
      t_imp.Text = 0
   End If
Else
   t_imp.Text = 0
End If

Xrecclii.Close
ConbdSapp.Close


End Sub

Public Sub Consultar_pedidos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
If Trim(t_ced.Text) <> "" Then
   Xsqlpromo = "Select * from pedidos_medic where cedula =" & t_ced.Text & " and fecha ='" & Format(Date, "yyyy/mm/dd") & "'"
Else
   Xsqlpromo = "Select * from pedidos_medic where matricula =" & 0 & " and fecha ='" & Format(Date, "yyyy/mm/dd") & "'"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   MsgBox "ATENCION! ESTA CEDULA, Ya tiene un pedido ingresado para hoy, Verifique!!", vbCritical
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_clideudas()

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xfecvence As Date
Dim Vencido As Integer
Vencido = 0
Xfecvence = Date
If t_mat.Text <> "" Then

    ConectarBD
    ConbdSapp.Open
                 
    Xsqlpromo = "Select * from deudas where cliente =" & t_mat.Text & " and fecha_pago is null and mes =" & 0
    
    With Xrecclii
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open Xsqlpromo, ConbdSapp, , , adCmdText
    End With
    If Xrecclii.RecordCount > 0 Then
       Xrecclii.MoveFirst
       Do While Not Xrecclii.EOF
          Xfecvence = Xrecclii("fecha") + Xrecclii("nro_superv")
          If Format(Xfecvence, "yyyy/mm/dd") <= Format(Date, "yyyy/mm/dd") Then
             Vencido = 1
          End If
          Xrecclii.MoveNext
       Loop
       If Vencido = 1 Then
          MsgBox "ATENCION! Socio con deuda pendiente de pago. Verifique!", vbCritical
          b_cancela_Click
       Else
          Xrecclii.Close
          Xsqlpromo = "Select * from deudas where cliente =" & t_mat.Text & " and fecha_pago is null and mes not in (0)"
          With Xrecclii
              .CursorLocation = adUseClient
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open Xsqlpromo, ConbdSapp, , , adCmdText
          End With
          If Xrecclii.RecordCount > 0 Then
             Xrecclii.MoveLast
             If Xrecclii.RecordCount > 2 Then
                MsgBox "ATENCION! Socio con deuda de cuota pendiente de pago. Verifique!", vbCritical
                b_cancela_Click
             End If
          End If
       End If
    Else
       Xrecclii.Close
       Xsqlpromo = "Select * from deudas where cliente =" & t_mat.Text & " and fecha_pago is null and mes not in (0)"
       With Xrecclii
           .CursorLocation = adUseClient
           .CursorType = adOpenKeyset
           .LockType = adLockOptimistic
           .Open Xsqlpromo, ConbdSapp, , , adCmdText
       End With
       If Xrecclii.RecordCount > 0 Then
          Xrecclii.MoveLast
          If Xrecclii.RecordCount > 2 Then
             MsgBox "ATENCION! Socio con deuda de cuota pendiente de pago. Verifique!", vbCritical
             b_cancela_Click
          End If
       End If
    End If
    Xrecclii.Close
    ConbdSapp.Close
End If

End Sub
Public Function Devuelve_costoPed() As Integer

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from estudios where codest =" & 60110
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_costoPed = Xrecclii("cons")
Else
   Devuelve_costoPed = 0
End If

Xrecclii.Close
ConbdSapp.Close

End Function


Public Sub Consultar_pedidoApr()
  frm_factselectmd.Show vbModal
  
End Sub
