VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_despachofact2 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos para la factura"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   Icon            =   "frm_despachofact2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.TextBox t_imptimbre 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6720
      TabIndex        =   46
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Data data_cablocal 
      Caption         =   "data_cablocal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_erro 
      Caption         =   "data_erro"
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
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_eror 
      Caption         =   "data_eror"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_hist 
      Caption         =   "data_hist"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_parse 
      Caption         =   "data_parse"
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_conslla 
      Caption         =   "data_conslla"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_deudas 
      Caption         =   "data_deudas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data data_conve 
      Caption         =   "data_conve"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_llamado 
      Caption         =   "data_llamado"
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton b_findos 
      Caption         =   "FIN2"
      Height          =   375
      Left            =   6960
      TabIndex        =   42
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton b_efct 
      Caption         =   "efct"
      Height          =   495
      Left            =   5640
      TabIndex        =   39
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton b_etck 
      Caption         =   "etck"
      Height          =   375
      Left            =   4560
      TabIndex        =   37
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FIN costo"
      Height          =   495
      Left            =   3600
      TabIndex        =   36
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Data data_cabeza2 
      Caption         =   "data_cabeza2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1800
      TabIndex        =   35
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Data data_caja 
      Caption         =   "data_caja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox t_nomcnv 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   34
      Top             =   2040
      Width           =   5055
   End
   Begin VB.TextBox t_codcnv 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      TabIndex        =   33
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Data data_ctrlf 
      Caption         =   "data_ctrlf"
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
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data data_ui 
      Caption         =   "data_ui"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox cbofpago 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      ItemData        =   "frm_despachofact2.frx":058A
      Left            =   1800
      List            =   "frm_despachofact2.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   600
      Width           =   1935
   End
   Begin VB.Data data_arancel 
      Caption         =   "data_arancel"
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_estudios 
      Caption         =   "data_estudios"
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox t_otrodoc 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox t_pie 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4680
      Width           =   6255
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_lin3 
      Caption         =   "data_lin3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_lin2 
      Caption         =   "data_lin2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   -240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_cabezal 
      Caption         =   "data_cabezal"
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_verfac 
      Caption         =   "data_verfac"
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
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_temp 
      Caption         =   "data_temp"
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
      Top             =   5400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   6360
      Picture         =   "frm_despachofact2.frx":05AA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_despachofact2.frx":0B34
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox t_total 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1800
      TabIndex        =   18
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox t_iva 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox t_imp 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   3480
      Width           =   1935
   End
   Begin MSMask.MaskEdBox mf 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   16711680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
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
   Begin VB.TextBox t_nombre 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   6255
   End
   Begin VB.TextBox t_codced 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox t_ced 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox t_mat 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label labtimbre 
      Height          =   255
      Left            =   7800
      TabIndex        =   45
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label labcodtras 
      Height          =   255
      Left            =   4080
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labnrollamado 
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labdescest 
      Height          =   255
      Left            =   2880
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label labcodest 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labvence 
      Height          =   375
      Left            =   6480
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONVENIO:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label labhora 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HORA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forma Pago:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label labcodsrv 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ADENDA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   240
      TabIndex        =   21
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IVA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPORTE:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label labnrofact 
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label labserie 
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E-TICKET"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NOMBRE:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DOCUMENTO:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MATRICULA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "frm_despachofact2.frx":10BE
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "frm_despachofact2"
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

Public Xbb, Xconlin, Xtipodedocumento, Xcandelin As Integer
Public Xivvva, Xtot, Xsubt, Xivauno As Double

Private Sub b_efct_Click()
Dim strIdTransac As String

If Label7.Caption = "E-FACTURA" Then

    Set objPosCfe = New PosCfe
    
    Dim objresultado As Resultado
'    Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
    Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-218", vbNullString)
    
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
    
    strIdTransac = objPosCfe.CrearGuid
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
        
        Set objCfe.EFact = New EFact
    
       With objCfe.EFact.Encabezado.IdDoc
            .TipoCFE = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(Str(data_cabeza2.Recordset("cl_tipocli"))))
            .FchEmis.SetDate Year(data_temp.Recordset("fecha")), Month(data_temp.Recordset("fecha")), Day(data_temp.Recordset("fecha"))
            .IsValidMntBruto = True
            .MntBruto = IdDoc_Tck_MntBruto_1
            If data_cabeza2.Recordset("cl_forpago") = 1 Then
             .FmaPago = IdDoc_Fact_FmaPago_1
            Else
             .FmaPago = IdDoc_Fact_FmaPago_2
            End If
        End With
        With objCfe.EFact.Encabezado.Emisor
            .RUCEmisor = data_par.Recordset("ruc")
            .RznSoc = data_par.Recordset("nomc")
            .CdgDGISucur.FromString Trim(Str(data_par.Recordset("codsuc")))
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
            .IsValidMntIVATasaMin = True
            .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
            .IVATasaMin = TasaIVAType_10FullStop000
            .MntIVATasaMin.FromString Format(data_cabeza2.Recordset("saldo_cc"), "0.00")
            .CantLinDet.FromString data_cabeza2.Recordset("cl_grupo")
            .MntTotal.FromString Format(data_cabeza2.Recordset("saldo_doc"), "0.00")
            .MntPagar.FromString Format(data_cabeza2.Recordset("saldo_chc"), "0.00")
        End With
        Do While Not data_temp.Recordset.EOF
           With objCfe.EFact.Detalle.Item.AddNew
              .NroLinDet.FromString Trim(Str(data_temp.Recordset("linea")))
              .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(Str(data_temp.Recordset("in_mat"))))
              .NomItem = data_temp.Recordset("nom_flia")
              .Cantidad.FromString Trim(Str(data_temp.Recordset("cantidad")))
              .UniMed = "N/A"
              .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00")
              .MontoItem.FromString Format(data_temp.Recordset("tot_lin"), "0.00")
           End With
           data_temp.Recordset.MoveNext
        Loop
        Dim s As String
        s = objCfe.ToXml(True, XmlFormatting_Indented)

'        Text1.Text = s
        Dim strGuid As String
        strGuid = objPosCfe.CrearGuid()
        Dim objResultadoCfe As ResultadoCfe
        Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    
        Set objUltimaSerieNumero = Nothing
        DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
        If Not objUltimaSerieNumero Is Nothing Then _
           ' cmdFirmarNc.Enabled = True
'           MsgBox "firmar NC"
        End If
        data_temp.Recordset.MoveFirst
        Command4_Click
End If


End Sub

Private Sub b_etck_Click()
Dim strIdTransac As String

'''On Error GoTo Cierrosieser3

If Label7.Caption = "E-TICKET" Then
    
    Set objPosCfe = New PosCfe
        
    Dim objresultado As Resultado
    
    Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-218", vbNullString)
'Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
'    Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
    
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
'        MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'            "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
    'Enviando
    If Not EstaInicializado() Then Exit Sub
    
    Dim objCfe As CFE
    Set objCfe = New CFE

    Dim objCf As ClassFactory

    Set objCf = New ClassFactory
       
    Set objCfe.ETck = New ETck
    With objCfe.ETck.Encabezado.IdDoc
        .TipoCFE = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(Str(data_cabeza2.Recordset("cl_tipocli"))))
        .FchEmis.SetDate Year(data_temp.Recordset("fecha")), Month(data_temp.Recordset("fecha")), Day(data_temp.Recordset("fecha"))
        .IsValidMntBruto = True
        .MntBruto = IdDoc_Tck_MntBruto_1
        If data_cabeza2.Recordset("cl_forpago") = 1 Then
           .FmaPago = IdDoc_Tck_FmaPago_1
        Else
           .FmaPago = IdDoc_Tck_FmaPago_2
        End If
    End With
    With objCfe.ETck.Encabezado.Emisor
        .RUCEmisor = data_par.Recordset("ruc")
        .RznSoc = data_par.Recordset("nomc")
        .CdgDGISucur.FromString Trim(Str(data_par.Recordset("codsuc")))
        .DomFiscal = data_par.Recordset("domic")
        .Ciudad = data_par.Recordset("ciudad")
        .Departamento = data_par.Recordset("dpto")
    End With
    Set objCfe.ETck.Encabezado.Receptor = New Receptor_Tck
    Set objCfe.ETck.Encabezado.Receptor.Receptor_Tck_Choice = New Receptor_Tck_Choice
    With objCfe.ETck.Encabezado.Receptor
        .TipoDocRecep = DocType_4
        .CodPaisRecep = CodPaisType_UY
        .Receptor_Tck_Choice.DocRecepExt = data_cabeza2.Recordset("cl_nom_sup")
'        .CodPaisRecep = CodPaisType_UY
'        If IsNull(data_cabezal.Recordset("cl_nom_sup")) = False Then
'           .Receptor_Tck_Choice.DocRecep = data_cabezal.Recordset("cl_nom_sup")
'        Else
'           .Receptor_Tck_Choice.DocRecep = "0"
'        End If
'        .Receptor_Tck_Choice.DocRecepExt = data_cabezal.Recordset("cl_nom_sup")
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
        .IVATasaMin = TasaIVAType_10FullStop000
        .MntIVATasaMin.FromString Format(data_cabeza2.Recordset("saldo_cc"), "0.00")
        .CantLinDet.FromString Trim(Str(data_cabeza2.Recordset("cl_grupo")))
        .MntTotal.FromString Format(data_cabeza2.Recordset("saldo_doc"), "0.00")
        .MntPagar.FromString Format(data_cabeza2.Recordset("saldo_chc"), "0.00")
    End With
    Do While Not data_temp.Recordset.EOF
       With objCfe.ETck.Detalle.Item.AddNew
          .NroLinDet.FromString Trim(Str(data_temp.Recordset("linea")))
          .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(Str(data_temp.Recordset("in_mat"))))
          .NomItem = data_temp.Recordset("nom_flia")
          .Cantidad.FromString Trim(Str(data_temp.Recordset("cantidad")))
          .UniMed = "N/A"
          .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00")
          .MontoItem.FromString Format(data_temp.Recordset("tot_lin"), "0.00")
       End With
       data_temp.Recordset.MoveNext
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
'       MsgBox "firmar NC"
    End If
    data_temp.Recordset.MoveFirst
    Command4_Click
Else
    If Label7.Caption = "E-FACTURA" Then
       b_efct_Click
    End If
End If

'''Exit Sub

'Cierrosieser3:
'              If Err.Number = 3155 Then
'                 data_erro.Recordset.AddNew
'                 data_erro.Recordset("id") = 2
'                 data_erro.Recordset("fecha") = Date
'                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
'                 data_erro.Recordset("nroerr") = Err.Number
'                 data_erro.Recordset("desc") = "btn e-tck " & data_cabeza2.Recordset("cl_nom_sup")
'                 data_erro.Recordset.Update
'              Else
'                 data_erro.Recordset.AddNew
'                 data_erro.Recordset("id") = 2
'                 data_erro.Recordset("fecha") = Date
'                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
''                 data_erro.Recordset("nroerr") = Err.Number
'                 data_erro.Recordset("desc") = "btn e-tck " & data_cabeza2.Recordset("cl_nom_sup")
'                 data_erro.Recordset.Update
'              End If
'              End
              
              
End Sub

Private Sub b_findos_Click()
Dim Xivauno As Double

Dim Xlf As Date
Dim Xelano, Xdiasfact, Xcandelin, XX2 As Integer
Dim Xfecvence As Date
Dim Xlatasa, Xlatasa22 As Double
Dim Xelerrdonde As Integer

On Error GoTo Cierrosieser5

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")

Xcandelin = 0
Xelano = Year(Date) + 1
Xfecvence = Format(mf.Text, "dd/mm/yyyy")
Xfecvence = Xfecvence + 30
Xelerrdonde = 1
XX2 = 0
Dim XnroMed As Integer
Dim Xnommed As String
If frm_factedesp.labcodmed.Caption <> "" Then
   XnroMed = frm_factedesp.labcodmed.Caption
Else
   XnroMed = 440
End If
If frm_factedesp.labnommed.Caption <> "" Then
   Xnommed = frm_factedesp.labnommed.Caption
Else
   Xnommed = "S/D"
End If
If t_imp.Text = "" Then
   t_imp.Text = 0
End If
If t_total.Text = "" Then
   t_total.Text = 0
End If
If t_iva.Text = "" Then
   t_iva.Text = 0
End If

If mf.Text <> "__/__/____" And List1.ListCount >= 1 Then
   labvence.Caption = CDate(mf.Text) + 30
   List1.ListIndex = 0
   Xcandelin = Xcandelin + 1
   data_temp.Recordset.AddNew
   data_temp.Recordset("linea") = Xcandelin
   data_temp.Recordset("libro_rub") = "REG." ' tipo de documento (Ej.e-ticket)
'       data_temp.Recordset("unidad") = labserie.Caption
   data_temp.Recordset("in_unid") = "INT1"
   data_temp.Recordset("in_mat") = 5 'gravado a tasa mínima
   data_temp.Recordset("cantidad") = 1
'       data_temp.Recordset("factura") = labnrofact.Caption
   data_temp.Recordset("tipo") = "REG." 'contado/crédito
   data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
   data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
   data_temp.Recordset("cod_cli") = t_mat.Text
   data_temp.Recordset("nom_cli") = t_nombre.Text
   data_temp.Recordset("convenio") = t_codcnv.Text
   If frm_factedesp.labtrasla.Caption = 1 Or frm_factedesp.labtrasla.Caption = 2 Or _
      frm_factedesp.labtrasla.Caption = 14 Or frm_factedesp.labtrasla.Caption = 10 Or _
      frm_factedesp.labtrasla.Caption = 11 Or frm_factedesp.labtrasla.Caption = 9 Or _
      frm_factedesp.labtrasla.Caption = 15 Then
      If frm_factedesp.labcateg.Caption = "MSP" Then
         data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
         data_estudios.Refresh
      Else
         data_estudios.RecordSource = "Select * from estudios where codest =" & labcodtras.Caption
         data_estudios.Refresh
      End If
   Else
      data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
      data_estudios.Refresh
   End If
   If data_estudios.Recordset.RecordCount > 0 Then
      data_temp.Recordset("cod_prod") = data_estudios.Recordset("codest")
      data_temp.Recordset("in_usuario") = Trim(Str(labcodsrv.Caption))
      data_temp.Recordset("nom_prod") = data_estudios.Recordset("descrip")
      data_temp.Recordset("nom_flia") = List1.List(List1.ListIndex)
   End If
   data_temp.Recordset("operador") = frm_factedesp.labusuario.Caption
   data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
   data_temp.Recordset("imp_timbre") = t_imp.Text
   data_temp.Recordset("tot_lin") = t_total.Text
   data_temp.Recordset("base") = data_par.Recordset("codsuc")
   data_temp.Recordset("pre_civa") = CDbl(t_iva.Text)
   data_temp.Recordset("reg_cab") = 99
   data_temp.Recordset("servicio") = 0
   If t_ced.Text <> "" Then
      data_temp.Recordset("ced_socio") = t_ced.Text
   End If
   If t_codced.Text <> "" Then
      data_temp.Recordset("fact") = t_codced.Text  'codcedula
   End If
   data_temp.Recordset("moneda") = "UYU"
   data_temp.Recordset("nro_flia") = 1
'       data_temp.Recordset("grupo") = data_lineas.Recordset("grupo") 'nro de cobrador
   If cbofpago.ListIndex = 0 Then
      data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cred") 'rubro
   Else
      data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cdo") 'rubro
   End If
   data_temp.Recordset("arancel") = Format(t_total.Text, "Standard")
   data_temp.Recordset("nro_med_a") = XnroMed
   data_temp.Recordset("nom_med_a") = Xnommed
   data_temp.Recordset("precio_est") = t_total.Text
   data_temp.Recordset("imp_iva") = CDbl(t_iva.Text)
   data_temp.Recordset.Update
   Xelerrdonde = 2
   
   If frm_factedesp.labtrasla.Caption = 1 Or frm_factedesp.labtrasla.Caption = 2 Or _
      frm_factedesp.labtrasla.Caption = 14 Or frm_factedesp.labtrasla.Caption = 10 Or _
      frm_factedesp.labtrasla.Caption = 11 Or frm_factedesp.labtrasla.Caption = 9 Or _
      frm_factedesp.labtrasla.Caption = 15 Then
      If frm_factedesp.labcateg.Caption = "MSP" Then
      Else
        Xcandelin = Xcandelin + 1
        data_temp.Recordset.AddNew
        data_temp.Recordset("linea") = Xcandelin
        data_temp.Recordset("libro_rub") = "REG." ' tipo de documento (Ej.e-ticket)
        data_temp.Recordset("in_unid") = "INT1"
        data_temp.Recordset("in_mat") = 5 'gravado a tasa mínima
        data_temp.Recordset("cantidad") = 1
        data_temp.Recordset("tipo") = "REG." 'contado/crédito
        data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
        data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
        data_temp.Recordset("cod_cli") = t_mat.Text
        data_temp.Recordset("nom_cli") = t_nombre.Text
        data_temp.Recordset("convenio") = t_codcnv.Text
        If labcodest.Caption <> "" Then
            data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
            data_estudios.Refresh
            If data_estudios.Recordset.RecordCount > 0 Then
               data_temp.Recordset("cod_prod") = data_estudios.Recordset("codest")
               data_temp.Recordset("in_usuario") = "10017"
               data_temp.Recordset("nom_prod") = data_estudios.Recordset("descrip")
               data_temp.Recordset("nom_flia") = "SERVICIOS DOMICILIARIOS"
            End If
        Else
            labcodest.Caption = 10001
            data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
            data_estudios.Refresh
            If data_estudios.Recordset.RecordCount > 0 Then
               data_temp.Recordset("cod_prod") = data_estudios.Recordset("codest")
               data_temp.Recordset("in_usuario") = "10017"
               data_temp.Recordset("nom_prod") = data_estudios.Recordset("descrip")
               data_temp.Recordset("nom_flia") = "SERVICIOS DOMICILIARIOS"
            End If
        End If
        data_temp.Recordset("operador") = frm_factedesp.labusuario.Caption
        data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
        data_temp.Recordset("imp_timbre") = 0
        data_temp.Recordset("tot_lin") = 0
        data_temp.Recordset("base") = data_parse.Recordset("base")
        data_temp.Recordset("pre_civa") = 0
        data_temp.Recordset("reg_cab") = 99
        data_temp.Recordset("servicio") = 0
        If t_ced.Text <> "" Then
           data_temp.Recordset("ced_socio") = t_ced.Text
        End If
        If t_codced.Text <> "" Then
           data_temp.Recordset("fact") = t_codced.Text  'codcedula
        End If
        data_temp.Recordset("moneda") = "UYU"
        data_temp.Recordset("nro_flia") = 1
        If cbofpago.ListIndex = 0 Then
           data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cred") 'rubro
        Else
           data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cdo") 'rubro
        End If
        data_temp.Recordset("arancel") = 0
        data_temp.Recordset("nro_med_a") = XnroMed
        data_temp.Recordset("nom_med_a") = Xnommed
        data_temp.Recordset("precio_est") = 0
        data_temp.Recordset("imp_iva") = 0
        data_temp.Recordset.Update
      End If
   End If
   data_temp.Refresh
   Xelerrdonde = 3
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveLast
      data_temp.Recordset.MoveFirst
   End If
   data_cabeza2.Recordset.AddNew
   data_cabeza2.Recordset("cl_tipcli") = "1.0"
   data_cabeza2.Recordset("cl_telefon") = "REG."
   data_cabeza2.Recordset("cl_tipocli") = 9
   data_cabeza2.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
   data_cabeza2.Recordset("cl_nrovend") = 1
   data_cabeza2.Recordset("cl_forpago") = 1
   data_cabeza2.Recordset("fecha_modi") = Format(labvence.Caption, "dd/mm/yyyy")
   data_cabeza2.Recordset("cl_diacobr") = Trim(Str(data_par.Recordset("ruc")))
   data_cabeza2.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
   data_cabeza2.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
   data_cabeza2.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
   data_cabeza2.Recordset("cl_referen") = data_par.Recordset("domic")
   data_cabeza2.Recordset("tit_tarj") = data_par.Recordset("ciudad")
   data_cabeza2.Recordset("cl_nomconv") = data_par.Recordset("dpto")
    'receptor
   data_cabeza2.Recordset("cl_nro_sup") = Xtipodedocumento
   data_cabeza2.Recordset("hora_baja") = "UY"
   If Xtipodedocumento = 3 Then
      data_cabeza2.Recordset("cl_nom_sup") = t_ced.Text & t_codced.Text
   Else
       data_cabeza2.Recordset("cl_nom_sup") = t_mat.Text
   End If
      'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
   data_cabeza2.Recordset("info_debit") = t_nombre.Text
   data_cabeza2.Recordset("cl_direcci") = "S/D"
   data_cabeza2.Recordset("cl_zona") = "S/D"
   data_cabeza2.Recordset("cl_localid") = "URUGUAY" 'opcional
   data_cabeza2.Recordset("cl_codigo") = t_mat.Text
   data_cabeza2.Recordset("usu_baja") = "UYU"
   data_cabeza2.Recordset("saldo_doc2") = Format(t_total.Text, "Standard") - Format(t_iva.Text, "Standard")
   data_cabeza2.Recordset("cl_atrasop") = Xlatasa
   data_cabeza2.Recordset("cl_decuota") = Xlatasa22
   data_cabeza2.Recordset("saldo_cc") = Format(t_iva.Text, "Standard")
   data_cabeza2.Recordset("saldo_cc2") = 0 'iva básico
   data_cabeza2.Recordset("saldo_doc") = Format(t_total.Text, "Standard")
   data_cabeza2.Recordset("cl_grupo") = data_temp.Recordset.RecordCount
   data_cabeza2.Recordset("saldo_chc") = Format(t_total.Text, "Standard")
   data_cabeza2.Recordset.Update
   data_cabeza2.Refresh
   
   'fin de cabezal
   Xelerrdonde = 4

   Xlatasa = CDbl("10.000")
   Xlatasa22 = CDbl("22.000")
   labserie.Caption = data_par.Recordset("serie_desp")
   labnrofact.Caption = data_par.Recordset("nro_desp") + 1
   data_par.Recordset.Edit
   data_par.Recordset("nro_desp") = data_par.Recordset("nro_desp") + 1
   data_par.Recordset.Update
'labserie.Caption = "A"
'labnrofact.Caption = 100011
   If data_temp.Recordset.RecordCount > 0 Then
      data_lin.RecordSource = "select * from linmmdd where cod_cli =" & data_temp.Recordset("cod_cli")
      data_lin.Refresh
      data_lin2.RecordSource = "Select * from hc_torax"
      data_lin2.Refresh
      data_lin3.RecordSource = "Select * from indica_enfc"
      data_lin3.Refresh
      data_cablocal.Recordset.AddNew
     '           data_cabezal.Recordset("id") = 1
      data_cablocal.Recordset("cl_tipcli") = "1.0"
      data_cablocal.Recordset("cl_tipocli") = 9
      data_cablocal.Recordset("cl_socmnro") = labserie.Caption
      data_cablocal.Recordset("cl_numero") = Val(labnrofact.Caption)
      data_cablocal.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
      data_cablocal.Recordset("cl_nrovend") = 1 'linea de detalle iva incluido
      data_cablocal.Recordset("cl_forpago") = 1
      data_cablocal.Recordset("cl_diacobr") = Trim(Str(data_par.Recordset("ruc")))
      data_cablocal.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
      data_cablocal.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
      data_cablocal.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
      data_cablocal.Recordset("cl_referen") = data_par.Recordset("domic")
      data_cablocal.Recordset("tit_tarj") = data_par.Recordset("ciudad")
      data_cablocal.Recordset("cl_nomconv") = data_par.Recordset("dpto")
             'receptor
      data_cablocal.Recordset("cl_nro_sup") = Xtipodedocumento 'tipo de documento
      data_cablocal.Recordset("hora_baja") = "UY" 'codigo del pais del documento
      data_cablocal.Recordset("cl_codigo") = t_mat.Text
      data_cablocal.Recordset("usu_baja") = "UYU" 'tipo de moneda
      data_cablocal.Recordset("saldo_doc2") = Format(t_total.Text, "Standard") 'total monto neto iva mínimo
      data_cablocal.Recordset("cl_atrasop") = Xlatasa
      data_cablocal.Recordset("cl_decuota") = Xlatasa22
      data_cablocal.Recordset("saldo_cc") = Format(t_iva.Text, "Standard")
      data_cablocal.Recordset("saldo_cc2") = 0 'iva básico
      data_cablocal.Recordset("saldo_doc") = Format(t_total.Text, "Standard")
      data_cablocal.Recordset("cl_grupo") = data_temp.Recordset.RecordCount 'nro de líneas
      data_cablocal.Recordset("saldo_chc") = Format(t_total.Text, "Standard")
      data_cablocal.Recordset.Update
      data_cablocal.RecordSource = "Select * from cabezados where cl_codigo =" & t_mat.Text & " and cl_numero =" & Val(labnrofact.Caption)
      data_cablocal.Refresh
      If data_cablocal.Recordset.RecordCount > 0 Then
         data_cablocal.Recordset.MoveFirst
      End If
        'fin de cabezal
      Xcandelin = 0
      Do While Not data_temp.Recordset.EOF
         Xcandelin = Xcandelin + 1
         data_lin.Recordset.AddNew
         data_lin.Recordset("linea") = data_temp.Recordset("linea")
         data_lin.Recordset("factura") = labnrofact.Caption
         data_lin.Recordset("tipo") = data_temp.Recordset("tipo")
         data_lin.Recordset("realizada") = Format(data_temp.Recordset("realizada"), "dd/mm/yyyy")
         data_lin.Recordset("fecha") = Format(data_temp.Recordset("fecha"), "dd/mm/yyyy")
         data_lin.Recordset("cod_cli") = data_temp.Recordset("cod_cli")
         data_lin.Recordset("nom_cli") = Mid(data_temp.Recordset("nom_cli"), 1, 30)
         data_lin.Recordset("convenio") = data_temp.Recordset("convenio")
         data_lin.Recordset("cod_prod") = data_temp.Recordset("cod_prod")
         data_lin.Recordset("nom_prod") = data_temp.Recordset("nom_prod")
         data_lin.Recordset("operador") = data_temp.Recordset("operador")
         data_lin.Recordset("hora") = data_temp.Recordset("hora")
         data_lin.Recordset("imp_timbre") = data_temp.Recordset("imp_timbre") ' sub total de la línea
         data_lin.Recordset("tot_lin") = data_temp.Recordset("tot_lin") ' total de la linea de la factura
         data_lin.Recordset("valor_iva") = data_temp.Recordset("pre_civa")
         data_lin.Recordset("base") = data_temp.Recordset("base")
         data_lin.Recordset("nom_med_a") = data_temp.Recordset("nom_med_a")
         data_lin.Recordset("rub_cont") = data_temp.Recordset("rub_cont")
         data_lin.Recordset("nom_flia") = data_temp.Recordset("nom_flia")
              
         data_lin.Recordset("pre_civa") = data_temp.Recordset("pre_civa")
         data_lin.Recordset("reg_cab") = data_temp.Recordset("reg_cab") '=99
         data_lin.Recordset("servicio") = data_temp.Recordset("servicio")
         data_lin.Recordset("ced_socio") = data_temp.Recordset("ced_socio")
         data_lin.Recordset("fact") = data_temp.Recordset("fact") 'codced
         data_lin.Recordset("moneda") = data_temp.Recordset("moneda")
         data_lin.Recordset("nro_flia") = data_temp.Recordset("nro_flia")
         data_lin.Recordset("rub_cont") = data_temp.Recordset("rub_cont")
         data_lin.Recordset("arancel") = data_temp.Recordset("arancel")
         data_lin.Recordset("nro_med_a") = data_temp.Recordset("nro_med_a")
         data_lin.Recordset("precio_est") = data_temp.Recordset("precio_est")
         data_lin.Recordset("imp_iva") = data_temp.Recordset("imp_iva")
         data_lin.Recordset("moneda") = labserie.Caption
         data_lin.Recordset("tipo_mov") = Trim(Str("2"))
         data_lin.Recordset("pendiente") = "X"
         data_lin.Recordset.Update
                             
         data_temp.Recordset.MoveNext
      Loop
      
      data_lin.Recordset.Close
      
      Xcandelin = 0
      Xconlin = 0
      Xelerrdonde = 5
      
      data_ctrlf.Recordset.Edit
      data_ctrlf.Recordset("fecha") = Date
      data_ctrlf.Recordset("hora") = Format(Time, "HH:mm")
      data_ctrlf.Recordset.Update
      List1.Clear
      
      data_llamado.RecordSource = "Select * from llamado where nrolla =" & labnrollamado.Caption
      data_llamado.Refresh
      If data_llamado.Recordset.RecordCount > 0 Then
         If IsNull(data_llamado.Recordset("totend")) = False Then
            If data_llamado.Recordset("totend") <> "" Then
               If data_llamado.Recordset("totend") = "FACT" Then
               Else
                  data_llamado.Recordset.Edit
                  data_llamado.Recordset("totend") = "FACT"
                  data_llamado.Recordset.Update
               End If
            Else
               data_llamado.Recordset.Edit
               data_llamado.Recordset("totend") = "FACT"
               data_llamado.Recordset.Update
            End If
         Else
            data_llamado.Recordset.Edit
            data_llamado.Recordset("totend") = "FACT"
            data_llamado.Recordset.Update
         End If
         data_llamado.Refresh
      End If
       
    data_cabezal.Recordset.AddNew
    '           data_cabezal.Recordset("id") = 1
    data_cabezal.Recordset("cl_tipcli") = "1.0"
    data_cabezal.Recordset("cl_tipocli") = data_cablocal.Recordset("cl_tipocli")
    data_cabezal.Recordset("cl_socmnro") = data_cablocal.Recordset("cl_socmnro")
    data_cabezal.Recordset("cl_numero") = data_cablocal.Recordset("cl_numero")
    data_cabezal.Recordset("cl_fnac") = data_cablocal.Recordset("cl_fnac")
    data_cabezal.Recordset("fecha_reac") = data_cablocal.Recordset("fecha_reac")
    data_cabezal.Recordset("cl_tj_venc") = data_cablocal.Recordset("cl_tj_venc")
    data_cabezal.Recordset("cl_nrovend") = data_cablocal.Recordset("cl_nrovend")
    data_cabezal.Recordset("cl_forpago") = data_cablocal.Recordset("cl_forpago")
    data_cabezal.Recordset("cl_celular") = data_cablocal.Recordset("cl_celular") 'descripcion f.pago
    data_cabezal.Recordset("fecha_modi") = data_cablocal.Recordset("fecha_modi")
    data_cabezal.Recordset("cl_diacobr") = data_cablocal.Recordset("cl_diacobr")
    data_cabezal.Recordset("cl_nrotarj") = data_cablocal.Recordset("cl_nrotarj")
    data_cabezal.Recordset("cl_tjemi_n") = data_cablocal.Recordset("cl_tjemi_n")
    data_cabezal.Recordset("cl_tjemi_c") = data_cablocal.Recordset("cl_tjemi_c")
    data_cabezal.Recordset("cl_referen") = data_cablocal.Recordset("cl_referen")
    data_cabezal.Recordset("tit_tarj") = data_cablocal.Recordset("tit_tarj")
    data_cabezal.Recordset("cl_nomconv") = data_cablocal.Recordset("cl_nomconv")
    'receptor
    data_cabezal.Recordset("cl_nro_sup") = data_cablocal.Recordset("cl_nro_sup")
    data_cabezal.Recordset("hora_baja") = data_cablocal.Recordset("hora_baja")
    data_cabezal.Recordset("cl_nom_sup") = data_cablocal.Recordset("cl_nom_sup")
        'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
    data_cabezal.Recordset("info_debit") = data_cablocal.Recordset("info_debit")
    data_cabezal.Recordset("cl_direcci") = data_cablocal.Recordset("cl_direcci")
    data_cabezal.Recordset("cl_zona") = data_cablocal.Recordset("cl_zona")
    'data_cabezal.Recordset("cl_email") = frm_convenios.t_dpto.Text 'opcional
    data_cabezal.Recordset("cl_localid") = data_cablocal.Recordset("cl_localid") 'opcional
    data_cabezal.Recordset("cl_codigo") = data_cablocal.Recordset("cl_codigo")
    data_cabezal.Recordset("usu_baja") = data_cablocal.Recordset("usu_baja") 'moneda
    data_cabezal.Recordset("saldo_chc2") = data_cablocal.Recordset("saldo_chc2") 'valor dolar
    data_cabezal.Recordset("saldo_cc") = data_cablocal.Recordset("saldo_cc")  'iva minimo
    data_cabezal.Recordset("saldo_cc2") = data_cablocal.Recordset("saldo_cc2") 'iva básico
    data_cabezal.Recordset("cl_atrasoa") = data_cablocal.Recordset("cl_atrasoa") 'subtot iva 22
    data_cabezal.Recordset("cl_cedula") = data_cablocal.Recordset("cl_cedula") 'subtot iva cero
    data_cabezal.Recordset("saldo_doc2") = data_cablocal.Recordset("saldo_doc2")
    data_cabezal.Recordset("cl_atrasop") = data_cablocal.Recordset("cl_atrasop")
    data_cabezal.Recordset("cl_decuota") = data_cablocal.Recordset("cl_decuota")
    data_cabezal.Recordset("saldo_doc") = data_cablocal.Recordset("saldo_doc")
    data_cabezal.Recordset("cl_grupo") = data_cablocal.Recordset("cl_grupo")
    data_cabezal.Recordset("saldo_chc") = data_cablocal.Recordset("saldo_chc")
    data_cabezal.Recordset("cl_telefon") = data_cablocal.Recordset("cl_telefon")
    data_cabezal.Recordset("cl_nombre") = data_cablocal.Recordset("cl_nombre")
    data_cabezal.Recordset("cl_cuopaga") = data_cablocal.Recordset("cl_cuopaga")
    data_cabezal.Recordset("codmotbaja") = data_cablocal.Recordset("codmotbaja")
    data_cabezal.Recordset("ultanopmut") = data_cablocal.Recordset("ultanopmut")
    data_cabezal.Recordset("cl_fultvta") = data_cablocal.Recordset("cl_fultvta")
    data_cabezal.Recordset("cl_entre") = data_cablocal.Recordset("cl_entre")
    data_cabezal.Recordset("codmotbaja") = data_cablocal.Recordset("codmotbaja")
    data_cabezal.Recordset("ultanopmut") = data_cablocal.Recordset("ultanopmut")
    data_cabezal.Recordset("cl_fultvta") = data_cablocal.Recordset("cl_fultvta")
    data_cabezal.Recordset("cl_entre") = data_cablocal.Recordset("cl_entre")
    data_cabezal.Recordset("cl_fultpag") = data_cablocal.Recordset("cl_fultpag")
    data_cabezal.Recordset("cl_ultmesp") = data_cablocal.Recordset("cl_ultmesp")
    data_cabezal.Recordset("cl_nomvend") = data_cablocal.Recordset("cl_nomvend")
    data_cabezal.Recordset("cl_fax") = data_cablocal.Recordset("cl_fax")
    data_cabezal.Recordset.Update
    'fin de cabezal
    data_cablocal.Recordset.Edit
    data_cablocal.Recordset("cl_codced") = 1
    data_cablocal.Recordset.Update
    Xelerrdonde = 6
       
    data_estudios.Recordset.Close
    data_cabezal.Recordset.Close
    data_llamado.Recordset.Close
    data_lin2.Recordset.Close
    data_lin3.Recordset.Close
    
'      MsgBox "Proceso de facturación terminado", vbInformation
   
   Else
      MsgBox "No hay líneas de facturación"
   End If
Else
   MsgBox "Hay un error en fecha o datos de la factura, NO SE PUEDE FACTURAR, REINTENTE!", vbCritical
End If
'terminado



Exit Sub

Cierrosieser5:
              If Err.Number = 3155 Then
                 data_erro.Recordset.AddNew
                 data_erro.Recordset("id") = 3
                 data_erro.Recordset("fecha") = Date
                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
                 data_erro.Recordset("nroerr") = Err.Number
                 data_erro.Recordset("desc") = Trim(Str(Xelerrdonde)) & Trim(Str(data_temp.Recordset("cod_cli"))) & Mid(Err.Description, 1, 110)
                 data_erro.Recordset.Update
              Else
                 data_erro.Recordset.AddNew
                 data_erro.Recordset("id") = 3
                 data_erro.Recordset("fecha") = Date
                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
                 data_erro.Recordset("nroerr") = Err.Number
                 data_erro.Recordset("desc") = Trim(Str(Xelerrdonde)) & Trim(Str(data_temp.Recordset("cod_cli"))) & Mid(Err.Description, 1, 110)
                 data_erro.Recordset.Update
              End If
              End


End Sub

Private Sub Command1_Click()
Dim Xivauno As Double

Dim Xlf As Date
Dim Xelano, Xcandelin, XX2 As Integer
Dim Xfecvence As Date
Dim Xlatasa, Xlatasa22 As Double

'''On Error GoTo Cierrosieser2

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")

Xcandelin = 0
Xelano = Year(Date) + 1
Xfecvence = Format(mf.Text, "dd/mm/yyyy")
Xfecvence = Xfecvence + 30

XX2 = 0
Dim XnroMed As Integer
Dim Xnommed As String

If frm_factedesp.labcodmed.Caption <> "" Then
   XnroMed = frm_factedesp.labcodmed.Caption
Else
   XnroMed = 440
End If
If frm_factedesp.labnommed.Caption <> "" Then
   Xnommed = frm_factedesp.labnommed.Caption
Else
   Xnommed = "S/D"
End If

''''XnroMed = frm_factedesp.labcodced.Caption
'''''Xnommed = frm_factedesp.labnommed.Caption
'Seguir acá

If t_imp.Text = "" Then
   t_imp.Text = 0
End If
If t_total.Text = "" Then
   t_total.Text = 0
End If
If t_iva.Text = "" Then
   t_iva.Text = 0
End If
If mf.Text <> "__/__/____" And List1.ListCount >= 1 Then
   labvence.Caption = CDate(mf.Text) + 30
   List1.ListIndex = 0
   Xcandelin = Xcandelin + 1
   data_temp.Recordset.AddNew
   data_temp.Recordset("linea") = Xcandelin
   data_temp.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
'       data_temp.Recordset("unidad") = labserie.Caption
   data_temp.Recordset("in_unid") = "INT1"
   data_temp.Recordset("in_mat") = 2 'gravado a tasa mínima
   data_temp.Recordset("cantidad") = 1
'       data_temp.Recordset("factura") = labnrofact.Caption
   data_temp.Recordset("tipo") = cbofpago.Text 'contado/crédito
   data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
   data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
   data_temp.Recordset("cod_cli") = t_mat.Text
   data_temp.Recordset("nom_cli") = Mid(t_nombre.Text, 1, 100)
   data_temp.Recordset("convenio") = t_codcnv.Text
   If frm_factedesp.labtrasla.Caption = 1 Or frm_factedesp.labtrasla.Caption = 2 Or _
      frm_factedesp.labtrasla.Caption = 14 Or frm_factedesp.labtrasla.Caption = 10 Or _
      frm_factedesp.labtrasla.Caption = 11 Or frm_factedesp.labtrasla.Caption = 9 Or _
      frm_factedesp.labtrasla.Caption = 15 Then
      If frm_factedesp.labcateg.Caption = "MSP" Then
         data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
         data_estudios.Refresh
      Else
         data_estudios.RecordSource = "Select * from estudios where codest =" & labcodtras.Caption
         data_estudios.Refresh
      End If
   Else
      data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
      data_estudios.Refresh
   End If
   If data_estudios.Recordset.RecordCount > 0 Then
      data_temp.Recordset("cod_prod") = data_estudios.Recordset("codest")
      data_temp.Recordset("in_usuario") = Trim(Str(labcodsrv.Caption))
      data_temp.Recordset("nom_prod") = data_estudios.Recordset("descrip")
      data_temp.Recordset("nom_flia") = List1.List(List1.ListIndex)
   End If
   data_temp.Recordset("operador") = frm_factedesp.labusuario.Caption
   data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
   data_temp.Recordset("imp_timbre") = t_imp.Text
   data_temp.Recordset("tot_lin") = t_imp.Text 't_totl
   data_temp.Recordset("base") = data_par.Recordset("codsuc")
   data_temp.Recordset("pre_civa") = CDbl(t_iva.Text)
   data_temp.Recordset("reg_cab") = 99
   data_temp.Recordset("servicio") = 0
   data_temp.Recordset("ced_socio") = t_ced.Text
   data_temp.Recordset("fact") = t_codced.Text  'codcedula
   data_temp.Recordset("moneda") = "UYU"
   data_temp.Recordset("nro_flia") = 1
'       data_temp.Recordset("grupo") = data_lineas.Recordset("grupo") 'nro de cobrador
   If cbofpago.ListIndex = 0 Then
      data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cred") 'rubro
   Else
      data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cdo") 'rubro
   End If
   data_temp.Recordset("arancel") = Format(t_imp.Text, "Standard")  't_total
   data_temp.Recordset("nro_med_a") = XnroMed
   data_temp.Recordset("nom_med_a") = Xnommed
   data_temp.Recordset("precio_est") = t_imp.Text 't_total
   data_temp.Recordset("imp_iva") = CDbl(t_iva.Text)
   data_temp.Recordset.Update
   If frm_factedesp.labtrasla.Caption = 1 Or frm_factedesp.labtrasla.Caption = 2 Or _
      frm_factedesp.labtrasla.Caption = 14 Or frm_factedesp.labtrasla.Caption = 10 Or _
      frm_factedesp.labtrasla.Caption = 11 Or frm_factedesp.labtrasla.Caption = 9 Or _
      frm_factedesp.labtrasla.Caption = 15 Then
      If frm_factedesp.labcateg.Caption = "MSP" Then
      Else
        Xcandelin = Xcandelin + 1
        data_temp.Recordset.AddNew
        data_temp.Recordset("linea") = Xcandelin
        data_temp.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
        data_temp.Recordset("in_unid") = "INT1"
        data_temp.Recordset("in_mat") = 5 'gravado a tasa mínima
        data_temp.Recordset("cantidad") = 1
        If t_pie.Text <> "" Then
           data_temp.Recordset("in_obs") = t_pie.Text
        End If
        data_temp.Recordset("tipo") = cbofpago.Text 'contado/crédito
        data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
        data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
        data_temp.Recordset("cod_cli") = t_mat.Text
        data_temp.Recordset("nom_cli") = Mid(t_nombre.Text, 1, 100)
        data_temp.Recordset("convenio") = t_codcnv.Text
        data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
        data_estudios.Refresh
        If data_estudios.Recordset.RecordCount > 0 Then
           data_temp.Recordset("cod_prod") = data_estudios.Recordset("codest")
           data_temp.Recordset("in_usuario") = "10017"
           data_temp.Recordset("nom_prod") = data_estudios.Recordset("descrip")
           data_temp.Recordset("nom_flia") = "SERVICIOS DOMICILIARIOS"
        End If
        data_temp.Recordset("operador") = frm_factedesp.labusuario.Caption
        data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
        data_temp.Recordset("imp_timbre") = 0
        data_temp.Recordset("tot_lin") = 0
        data_temp.Recordset("base") = data_parse.Recordset("base")
        data_temp.Recordset("pre_civa") = 0
        data_temp.Recordset("reg_cab") = 99
        data_temp.Recordset("servicio") = 0
        data_temp.Recordset("ced_socio") = t_ced.Text
        data_temp.Recordset("fact") = t_codced.Text  'codcedula
        data_temp.Recordset("moneda") = "UYU"
        data_temp.Recordset("nro_flia") = 1
        If cbofpago.ListIndex = 0 Then
           data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cred") 'rubro
        Else
           data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cdo") 'rubro
        End If
        data_temp.Recordset("arancel") = 0
        data_temp.Recordset("nro_med_a") = XnroMed
        data_temp.Recordset("nom_med_a") = Mid(Xnommed, 1, 40)
        data_temp.Recordset("precio_est") = 0
        data_temp.Recordset("imp_iva") = 0
        data_temp.Recordset.Update
      End If
   End If
   data_temp.Refresh
   If labtimbre.Caption = "SI" Then
      Xcandelin = Xcandelin + 1
      data_temp.Recordset.AddNew
      data_temp.Recordset("linea") = Xcandelin
      data_temp.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
      data_temp.Recordset("in_unid") = "INT1"
      data_temp.Recordset("in_mat") = 1 'no lleva iva
      data_temp.Recordset("cantidad") = 1
      data_temp.Recordset("tipo") = cbofpago.Text 'contado/crédito
      data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
      data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
      data_temp.Recordset("cod_cli") = t_mat.Text
      data_temp.Recordset("nom_cli") = Mid(t_nombre.Text, 1, 100)
      data_temp.Recordset("convenio") = t_codcnv.Text
      data_temp.Recordset("cod_prod") = 995
      data_temp.Recordset("in_usuario") = "TIMBRE PROFESIONAL"
      data_temp.Recordset("nom_prod") = "TIMBRE PROFESIONAL"
      data_temp.Recordset("nom_flia") = "TIMBRE PROFESIONAL"
      data_temp.Recordset("operador") = frm_factedesp.labusuario.Caption
      data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
      data_temp.Recordset("imp_timbre") = t_imptimbre.Text
      data_temp.Recordset("tot_lin") = t_imptimbre.Text
      data_temp.Recordset("base") = data_par.Recordset("codsuc")
      data_temp.Recordset("pre_civa") = 0
      data_temp.Recordset("reg_cab") = 99
      data_temp.Recordset("servicio") = 0
      data_temp.Recordset("ced_socio") = t_ced.Text
      data_temp.Recordset("fact") = t_codced.Text  'codcedula
      data_temp.Recordset("moneda") = "UYU"
      data_temp.Recordset("nro_flia") = 8
      If cbofpago.ListIndex = 0 Then
         data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cred") 'rubro
      Else
         data_temp.Recordset("rub_cont") = data_par.Recordset("rub_cdo") 'rubro
      End If
      data_temp.Recordset("arancel") = Format(t_total.Text, "Standard")
'     data_temp.Recordset("nro_med_a") = XnroMed
'     data_temp.Recordset("nom_med_a") = Xnommed
      data_temp.Recordset("precio_est") = t_imptimbre.Text
      data_temp.Recordset("imp_iva") = 0
      data_temp.Recordset.Update
   End If
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveLast
      data_temp.Recordset.MoveFirst
   End If
   data_cabeza2.Recordset.AddNew
   data_cabeza2.Recordset("cl_tipcli") = "1.0"
   If Label7.Caption = "E-FACTURA" Then
      data_cabeza2.Recordset("cl_tipocli") = 111
   Else
      If Label7.Caption = "E-TICKET" Then
         data_cabeza2.Recordset("cl_tipocli") = 101
      Else
         data_cabeza2.Recordset("cl_tipocli") = 101
      End If
   End If
   data_cabeza2.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
   data_cabeza2.Recordset("cl_nrovend") = 1
   If cbofpago.Text = "CONTADO" Then
      data_cabeza2.Recordset("cl_forpago") = 1
   Else
      If cbofpago.Text = "CREDITO" Then
         data_cabeza2.Recordset("cl_forpago") = 2
      Else
         data_cabeza2.Recordset("cl_forpago") = 2
     End If
   End If
   data_cabeza2.Recordset("fecha_modi") = Format(labvence.Caption, "dd/mm/yyyy")
   data_cabeza2.Recordset("cl_diacobr") = Trim(Str(data_par.Recordset("ruc")))
   data_cabeza2.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
   data_cabeza2.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
   data_cabeza2.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
   data_cabeza2.Recordset("cl_referen") = data_par.Recordset("domic")
   data_cabeza2.Recordset("tit_tarj") = data_par.Recordset("ciudad")
   data_cabeza2.Recordset("cl_nomconv") = data_par.Recordset("dpto")
    'receptor
   data_cabeza2.Recordset("cl_nro_sup") = Xtipodedocumento
   data_cabeza2.Recordset("hora_baja") = "UY"
   If Xtipodedocumento = 3 Then
      data_cabeza2.Recordset("cl_nom_sup") = t_ced.Text & t_codced.Text
   Else
       data_cabeza2.Recordset("cl_nom_sup") = t_mat.Text
   End If
      'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
   data_cabeza2.Recordset("info_debit") = Mid(t_nombre.Text, 1, 130)
   data_cabeza2.Recordset("cl_direcci") = frm_factedesp.labdire.Caption
   data_cabeza2.Recordset("cl_zona") = frm_factedesp.labzon.Caption
   data_cabeza2.Recordset("cl_localid") = "URUGUAY" 'opcional
   data_cabeza2.Recordset("cl_codigo") = t_mat.Text
   data_cabeza2.Recordset("usu_baja") = "UYU"
   data_cabeza2.Recordset("saldo_doc2") = Format(t_total.Text, "Standard") - Format(t_iva.Text, "Standard")
   data_cabeza2.Recordset("cl_atrasop") = Xlatasa
   data_cabeza2.Recordset("cl_decuota") = Xlatasa22
   data_cabeza2.Recordset("saldo_cc") = Format(t_iva.Text, "Standard")
   data_cabeza2.Recordset("saldo_cc2") = 0 'iva básico
   data_cabeza2.Recordset("saldo_doc") = Format(t_total.Text, "Standard")
   data_cabeza2.Recordset("cl_grupo") = data_temp.Recordset.RecordCount
   data_cabeza2.Recordset("saldo_chc") = Format(t_total.Text, "Standard")
   data_cabeza2.Recordset.Update
   data_cabeza2.Refresh
   
   data_estudios.Recordset.Close
   
'   MsgBox "Termina datos Facturación"
   If data_temp.Recordset.RecordCount > 0 Then
      b_etck_Click
'      Command4_Click
   End If
   'fin de cabezal
Else
   MsgBox "Faltan datos para poder grabar, verifique fecha!", vbInformation
   data_erro.Recordset.AddNew
   data_erro.Recordset("id") = 6
   data_erro.Recordset("fecha") = Date
   data_erro.Recordset("hora") = Format(Time, "HH:mm")
   data_erro.Recordset("nroerr") = Err.Number
   data_erro.Recordset("desc") = "En el command1 faltan datos"
   data_erro.Recordset.Update

End If
        
'Exit Sub


'Cierrosieser2:
'              If Err.Number = 3155 Then
'                 data_erro.Recordset.AddNew
'                 data_erro.Recordset("id") = 6
'                 data_erro.Recordset("fecha") = Date
'                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
'                 data_erro.Recordset("nroerr") = Err.Number
'                 data_erro.Recordset("desc") = "En el command1 de despachofact2"
'                 data_erro.Recordset.Update
'              Else
'                 data_erro.Recordset.AddNew
'                 data_erro.Recordset("id") = 6
'                 data_erro.Recordset("fecha") = Date
'                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
'                 data_erro.Recordset("nroerr") = Err.Number
'                 data_erro.Recordset("desc") = "En el command1 de despachofact2"
'                 data_erro.Recordset.Update
'              End If
'              End

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command4_Click()
Dim Xlatasa, Xlatasa22 As Double
Dim Xcandelin, Xdondeelerr As Integer

On Error GoTo Cierrosieser4
Xdondeelerr = 0

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")
'labserie.Caption = "A"
'labnrofact.Caption = 100011
If data_temp.Recordset.RecordCount > 0 Then
   data_lin.RecordSource = "select * from linmmdd where cod_cli =" & data_temp.Recordset("cod_cli")
   data_lin.Refresh
   data_lin2.RecordSource = "Select * from hc_torax"
   data_lin2.Refresh
   data_lin3.RecordSource = "Select * from indica_enfc"
   data_lin3.Refresh
   
   data_cablocal.Recordset.AddNew
 '           data_cabezal.Recordset("id") = 1
   data_cablocal.Recordset("cl_tipcli") = "1.0"
   data_cablocal.Recordset("cl_telefon") = Label7.Caption
   If Label7.Caption = "E-FACTURA" Then
      data_cablocal.Recordset("cl_tipocli") = 111
   Else
      If Label7.Caption = "E-TICKET" Then
         data_cablocal.Recordset("cl_tipocli") = 101
      Else
         data_cablocal.Recordset("cl_tipocli") = 101
      End If
   End If
   If Label7.Caption = "E-FACTURA" Then
      data_cablocal.Recordset("cl_nombre") = "RUT COMPRADOR"
   Else
      If Label7.Caption = "E-TICKET" Then
         data_cablocal.Recordset("cl_nombre") = "CONSUMO FINAL"
      End If
   End If
   data_cablocal.Recordset("cl_socmnro") = labserie.Caption
   data_cablocal.Recordset("cl_numero") = Val(labnrofact.Caption)
   data_cablocal.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
   data_cablocal.Recordset("cl_nrovend") = 1 'linea de detalle iva incluido
   If cbofpago.Text = "CONTADO" Then
      data_cablocal.Recordset("cl_forpago") = 1
   Else
      If cbofpago.Text = "CREDITO" Then
         data_cablocal.Recordset("cl_forpago") = 2
      Else
         data_cablocal.Recordset("cl_forpago") = 2
      End If
   End If
   data_cablocal.Recordset("cl_diacobr") = Trim(Str(data_par.Recordset("ruc")))
   data_cablocal.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
   data_cablocal.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
   data_cablocal.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
   data_cablocal.Recordset("cl_referen") = data_par.Recordset("domic")
   data_cablocal.Recordset("tit_tarj") = data_par.Recordset("ciudad")
   data_cablocal.Recordset("cl_nomconv") = data_par.Recordset("dpto")
        'receptor
   
   If Label7.Caption = "E-TICKET" Then
      data_cablocal.Recordset("cl_nro_sup") = Xtipodedocumento 'tipo de documento
   Else
      data_cablocal.Recordset("cl_nro_sup") = Xtipodedocumento '2 RUT, 3 CI
   End If
   data_cablocal.Recordset("hora_baja") = "UY" 'codigo del pais del documento
   data_cablocal.Recordset("cl_codigo") = t_mat.Text
   data_cablocal.Recordset("usu_baja") = "UYU" 'tipo de moneda
   data_cablocal.Recordset("saldo_doc2") = data_cabeza2.Recordset("saldo_doc2") 'total monto neto iva mínimo
   data_cablocal.Recordset("cl_atrasop") = Xlatasa
   data_cablocal.Recordset("cl_decuota") = Xlatasa22
   data_cablocal.Recordset("saldo_cc") = Format(t_iva.Text, "Standard")
   data_cablocal.Recordset("saldo_cc2") = 0 'iva básico
   data_cablocal.Recordset("saldo_doc") = Format(t_total.Text, "Standard")
   data_cablocal.Recordset("cl_grupo") = data_temp.Recordset.RecordCount 'nro de líneas
   data_cablocal.Recordset("saldo_chc") = Format(t_total.Text, "Standard")
   data_cablocal.Recordset.Update
   data_cablocal.RecordSource = "Select * from cabezados where cl_codigo =" & t_mat.Text & " and cl_numero =" & Val(labnrofact.Caption)
   data_cablocal.Refresh
   If data_cablocal.Recordset.RecordCount > 0 Then
      data_cablocal.Recordset.MoveFirst
   End If
   Xdondeelerr = 1
   Do While Not data_temp.Recordset.EOF
      Xcandelin = Xcandelin + 1
      data_lin.Recordset.AddNew
      If IsNull(data_temp.Recordset("linea")) = False Then
         data_lin.Recordset("linea") = data_temp.Recordset("linea")
      Else
         data_lin.Recordset("linea") = 1
      End If
      If labnrofact.Caption <> "" Then
         If IsNumeric(labnrofact.Caption) = True Then
            data_lin.Recordset("factura") = Val(labnrofact.Caption)
         Else
            data_lin.Recordset("factura") = Val(labnrofact.Caption)
         End If
      Else
         data_lin.Recordset("factura") = 0
      End If
      data_lin.Recordset("tipo") = data_temp.Recordset("tipo")
      data_lin.Recordset("realizada") = Format(data_temp.Recordset("realizada"), "dd/mm/yyyy")
      data_lin.Recordset("fecha") = Format(data_temp.Recordset("fecha"), "dd/mm/yyyy")
      data_lin.Recordset("cod_cli") = data_temp.Recordset("cod_cli")
      data_lin.Recordset("nom_cli") = Mid(data_temp.Recordset("nom_cli"), 1, 30)
      data_lin.Recordset("convenio") = data_temp.Recordset("convenio")
      data_lin.Recordset("cod_prod") = data_temp.Recordset("cod_prod")
      data_lin.Recordset("nom_prod") = Mid(data_temp.Recordset("nom_prod"), 1, 50)
      data_lin.Recordset("operador") = data_temp.Recordset("operador")
      data_lin.Recordset("hora") = data_temp.Recordset("hora")
      If IsNull(data_temp.Recordset("imp_timbre")) = False Then
         data_lin.Recordset("imp_timbre") = Format(data_temp.Recordset("imp_timbre"), "Standard") ' sub total de la línea
      Else
         data_lin.Recordset("imp_timbre") = 0
      End If
      data_lin.Recordset("tot_lin") = data_temp.Recordset("tot_lin") ' total de la linea de la factura
      data_lin.Recordset("valor_iva") = data_temp.Recordset("pre_civa")
      data_lin.Recordset("base") = data_temp.Recordset("base")
      data_lin.Recordset("nom_med_a") = data_temp.Recordset("nom_med_a")
      data_lin.Recordset("rub_cont") = data_temp.Recordset("rub_cont")
      data_lin.Recordset("nom_flia") = data_temp.Recordset("nom_flia")
      data_lin.Recordset("pre_civa") = data_temp.Recordset("pre_civa")
      data_lin.Recordset("reg_cab") = data_temp.Recordset("reg_cab") '=99
      data_lin.Recordset("servicio") = data_temp.Recordset("servicio")
      If IsNull(data_temp.Recordset("ced_socio")) = False Then
         data_lin.Recordset("ced_socio") = data_temp.Recordset("ced_socio")
      Else
         data_lin.Recordset("ced_socio") = 0
      End If
      data_lin.Recordset("fact") = data_temp.Recordset("fact") 'codced
      data_lin.Recordset("moneda") = data_temp.Recordset("moneda")
      data_lin.Recordset("nro_flia") = data_temp.Recordset("nro_flia")
      data_lin.Recordset("rub_cont") = data_temp.Recordset("rub_cont")
      data_lin.Recordset("arancel") = data_temp.Recordset("arancel")
      data_lin.Recordset("nro_med_a") = data_temp.Recordset("nro_med_a")
      data_lin.Recordset("precio_est") = data_temp.Recordset("precio_est")
      data_lin.Recordset("imp_iva") = data_temp.Recordset("imp_iva")
      If labserie.Caption <> "" Then
         data_lin.Recordset("moneda") = labserie.Caption
      Else
         data_lin.Recordset("moneda") = "A"
      End If
      data_lin.Recordset("tipo_mov") = Trim(Str("2"))
      If IsNull(data_cabeza2.Recordset("cl_tipocli")) = False Then
         If data_cabeza2.Recordset("cl_tipocli") = 111 Then
            data_lin.Recordset("pendiente") = "F"
         Else
            If data_cabeza2.Recordset("cl_tipocli") = 101 Then
               data_lin.Recordset("pendiente") = "T"
            Else
               data_lin.Recordset("pendiente") = "X"
            End If
         End If
      Else
         data_lin.Recordset("pendiente") = "T"
      End If
      data_lin.Recordset.Update
      Xdondeelerr = 2
'      data_lin2.Recordset.AddNew
'      data_lin2.Recordset("hora") = "INT1"
'      data_lin2.Recordset("descrip") = labserie.Caption  'serie del comprobante
'      data_lin2.Recordset("hc_nro") = 2 'tasa minima
'      data_lin2.Recordset("hc_cod") = Val(labnrofact.Caption)
'      data_lin2.Recordset("hc_mat") = data_temp.Recordset("linea")
'      data_lin2.Recordset.Update
'      data_lin2.Refresh
      data_caja.Recordset.AddNew
      data_caja.Recordset("fecha") = data_temp.Recordset("fecha")
      data_caja.Recordset("numero") = data_temp.Recordset("rub_cont")
      data_caja.Recordset("nombre") = Mid(data_temp.Recordset("rub_nomb"), 1, 35)
      data_caja.Recordset("moneda") = "$"
      data_caja.Recordset("movimiento") = "INGRESO"
      data_caja.Recordset("imp_fact") = data_temp.Recordset("tot_lin")
      data_caja.Recordset("documento") = labnrofact.Caption
      If cbofpago.Text = "CREDITO" Then
         data_caja.Recordset("observ") = "CREDITO " + labserie.Caption & " " & labnrofact.Caption
      Else
         data_caja.Recordset("observ") = "CONTADO " + labnrofact.Caption
      End If
      data_caja.Recordset("saldo") = data_temp.Recordset("tot_lin")
      data_caja.Recordset("usuario") = data_temp.Recordset("operador")
      data_caja.Recordset("hora") = data_temp.Recordset("hora")
      data_caja.Recordset("base") = data_temp.Recordset("base")
      data_caja.Recordset("cod_serv") = data_temp.Recordset("cod_prod")
      data_caja.Recordset("nom_serv") = Mid(data_temp.Recordset("nom_prod"), 1, 50)
      data_caja.Recordset("cod_socio") = data_temp.Recordset("cod_cli")
      data_caja.Recordset("nom_socio") = Mid(data_temp.Recordset("nom_cli"), 1, 30)
      If IsNull(data_temp.Recordset("imp_iva")) = False Then
         data_caja.Recordset("imp_iva") = Format(data_temp.Recordset("imp_iva"), "Standard")
      Else
         data_caja.Recordset("imp_iva") = 0
      End If
      data_caja.Recordset("opiva") = 1 ' 10% , 0 NO, 2 22%
      data_caja.Recordset.Update
      Xdondeelerr = 3
    
      data_temp.Recordset.MoveNext
   Loop
   data_lin.Recordset.Close
   data_caja.Recordset.Close
   Xcandelin = 0
   Xconlin = 0
   
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveFirst
      Do While Not data_temp.Recordset.EOF
         If data_temp.Recordset("tot_lin") > 0 Then
            data_deudas.RecordSource = "Select * from deudas where cliente =" & data_temp.Recordset("cod_cli")
            data_deudas.Refresh
            data_deudas.Recordset.AddNew
            data_deudas.Recordset("cod_cnv") = frm_factedesp.labcateg.Caption
            data_deudas.Recordset("nom_cnv") = Mid(frm_factedesp.labnomcat.Caption, 1, 20)
            data_deudas.Recordset("cliente") = data_temp.Recordset("cod_cli")
            data_deudas.Recordset("nombre") = Mid(data_temp.Recordset("nom_cli"), 1, 70)
            data_deudas.Recordset("fecha") = Date
            data_deudas.Recordset("tipodoc") = "CRE"
            data_deudas.Recordset("nro_superv") = 30
            data_deudas.Recordset("documento") = Val(labnrofact.Caption)
            data_deudas.Recordset("tipocta") = labserie.Caption
            data_deudas.Recordset("importe") = data_temp.Recordset("tot_lin")
            data_deudas.Recordset("moneda") = 1
            data_deudas.Recordset("origen") = "E-TICKET NRO." & labserie.Caption & " " & labnrofact.Caption
            data_deudas.Recordset("saldo_cc") = data_temp.Recordset("tot_lin")
            data_deudas.Recordset("mes") = 0
            data_deudas.Recordset("ano") = 0
            data_deudas.Recordset("estado_cta") = 1
            data_deudas.Recordset("tiquet") = 0
            data_deudas.Recordset("deudas") = 0
            data_deudas.Recordset("total") = data_temp.Recordset("tot_lin")
            data_deudas.Recordset("iva") = data_temp.Recordset("imp_iva")
            data_deudas.Recordset("servi") = 0
            data_deudas.Recordset("nro_vende") = data_temp.Recordset("linea")
            data_deudas.Recordset.Update
'            data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_temp.Recordset("cod_cli")
'            data_cli.Refresh
'            If data_cli.Recordset.RecordCount > 0 Then
'               data_cli.Recordset.Edit
'               If IsNull(data_cli.Recordset("saldo_cc")) = False Then
'                  data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") + data_temp.Recordset("tot_lin")
'               Else
'                  data_cli.Recordset("saldo_cc") = data_temp.Recordset("tot_lin")
'               End If
'               data_cli.Recordset.Update
'            End If
         End If
         data_temp.Recordset.MoveNext
      Loop
   End If
   data_deudas.Recordset.Close
   Xdondeelerr = 4
   data_temp.Recordset.MoveFirst
   data_ctrlf.Recordset.Edit
   data_ctrlf.Recordset("fecha") = Date
   data_ctrlf.Recordset("hora") = Format(Time, "HH:mm")
   data_ctrlf.Recordset.Update
'   frm_histdesp.Show vbModal
   List1.Clear
   data_llamado.RecordSource = "Select * from llamado where nrolla =" & labnrollamado.Caption
   data_llamado.Refresh
   If data_llamado.Recordset.RecordCount > 0 Then
      If IsNull(data_llamado.Recordset("totend")) = False Then
         If data_llamado.Recordset("totend") <> "" Then
            If data_llamado.Recordset("totend") = "FACT" Then
            Else
               data_llamado.Recordset.Edit
               data_llamado.Recordset("totend") = "FACT"
               data_llamado.Recordset.Update
            End If
         Else
            data_llamado.Recordset.Edit
            data_llamado.Recordset("totend") = "FACT"
            data_llamado.Recordset.Update
         End If
      Else
         data_llamado.Recordset.Edit
         data_llamado.Recordset("totend") = "FACT"
         data_llamado.Recordset.Update
      End If
      data_llamado.Refresh
   End If
   data_llamado.Recordset.Close
'   MsgBox "Termina de grabar después del envío"
   Xdondeelerr = 5
       
    data_cabezal.Recordset.AddNew
    '           data_cabezal.Recordset("id") = 1
    data_cabezal.Recordset("cl_tipcli") = "1.0"
    data_cabezal.Recordset("cl_tipocli") = data_cablocal.Recordset("cl_tipocli")
    data_cabezal.Recordset("cl_socmnro") = data_cablocal.Recordset("cl_socmnro")
    data_cabezal.Recordset("cl_numero") = data_cablocal.Recordset("cl_numero")
    data_cabezal.Recordset("cl_fnac") = data_cablocal.Recordset("cl_fnac")
    data_cabezal.Recordset("fecha_reac") = data_cablocal.Recordset("fecha_reac")
    data_cabezal.Recordset("cl_tj_venc") = data_cablocal.Recordset("cl_tj_venc")
    data_cabezal.Recordset("cl_nrovend") = data_cablocal.Recordset("cl_nrovend")
    data_cabezal.Recordset("cl_forpago") = data_cablocal.Recordset("cl_forpago")
    data_cabezal.Recordset("cl_celular") = data_cablocal.Recordset("cl_celular") 'descripcion f.pago
    data_cabezal.Recordset("fecha_modi") = data_cablocal.Recordset("fecha_modi")
    data_cabezal.Recordset("cl_diacobr") = data_cablocal.Recordset("cl_diacobr")
    data_cabezal.Recordset("cl_nrotarj") = data_cablocal.Recordset("cl_nrotarj")
    data_cabezal.Recordset("cl_tjemi_n") = data_cablocal.Recordset("cl_tjemi_n")
    data_cabezal.Recordset("cl_tjemi_c") = data_cablocal.Recordset("cl_tjemi_c")
    data_cabezal.Recordset("cl_referen") = data_cablocal.Recordset("cl_referen")
    data_cabezal.Recordset("tit_tarj") = data_cablocal.Recordset("tit_tarj")
    data_cabezal.Recordset("cl_nomconv") = data_cablocal.Recordset("cl_nomconv")
    'receptor
    data_cabezal.Recordset("cl_nro_sup") = data_cablocal.Recordset("cl_nro_sup")
    data_cabezal.Recordset("hora_baja") = data_cablocal.Recordset("hora_baja")
    data_cabezal.Recordset("cl_nom_sup") = data_cablocal.Recordset("cl_nom_sup")
        'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
    data_cabezal.Recordset("info_debit") = data_cablocal.Recordset("info_debit")
    data_cabezal.Recordset("cl_direcci") = data_cablocal.Recordset("cl_direcci")
    data_cabezal.Recordset("cl_zona") = data_cablocal.Recordset("cl_zona")
    'data_cabezal.Recordset("cl_email") = frm_convenios.t_dpto.Text 'opcional
    data_cabezal.Recordset("cl_localid") = data_cablocal.Recordset("cl_localid") 'opcional
    data_cabezal.Recordset("cl_codigo") = data_cablocal.Recordset("cl_codigo")
    data_cabezal.Recordset("usu_baja") = data_cablocal.Recordset("usu_baja") 'moneda
    data_cabezal.Recordset("saldo_chc2") = data_cablocal.Recordset("saldo_chc2") 'valor dolar
    data_cabezal.Recordset("saldo_cc") = data_cablocal.Recordset("saldo_cc")  'iva minimo
    data_cabezal.Recordset("saldo_cc2") = data_cablocal.Recordset("saldo_cc2") 'iva básico
    data_cabezal.Recordset("cl_atrasoa") = data_cablocal.Recordset("cl_atrasoa") 'subtot iva 22
    data_cabezal.Recordset("cl_cedula") = data_cablocal.Recordset("cl_cedula") 'subtot iva cero
    data_cabezal.Recordset("saldo_doc2") = data_cablocal.Recordset("saldo_doc2")
    data_cabezal.Recordset("cl_atrasop") = data_cablocal.Recordset("cl_atrasop")
    data_cabezal.Recordset("cl_decuota") = data_cablocal.Recordset("cl_decuota")
    data_cabezal.Recordset("saldo_doc") = data_cablocal.Recordset("saldo_doc")
    data_cabezal.Recordset("cl_grupo") = data_cablocal.Recordset("cl_grupo")
    data_cabezal.Recordset("saldo_chc") = data_cablocal.Recordset("saldo_chc")
    data_cabezal.Recordset("cl_telefon") = data_cablocal.Recordset("cl_telefon")
    data_cabezal.Recordset("cl_nombre") = data_cablocal.Recordset("cl_nombre")
    data_cabezal.Recordset("cl_cuopaga") = data_cablocal.Recordset("cl_cuopaga")
    data_cabezal.Recordset("codmotbaja") = data_cablocal.Recordset("codmotbaja")
    data_cabezal.Recordset("ultanopmut") = data_cablocal.Recordset("ultanopmut")
    data_cabezal.Recordset("cl_fultvta") = data_cablocal.Recordset("cl_fultvta")
    data_cabezal.Recordset("cl_entre") = data_cablocal.Recordset("cl_entre")
    data_cabezal.Recordset("codmotbaja") = data_cablocal.Recordset("codmotbaja")
    data_cabezal.Recordset("ultanopmut") = data_cablocal.Recordset("ultanopmut")
    data_cabezal.Recordset("cl_fultvta") = data_cablocal.Recordset("cl_fultvta")
    data_cabezal.Recordset("cl_entre") = data_cablocal.Recordset("cl_entre")
    data_cabezal.Recordset("cl_fultpag") = data_cablocal.Recordset("cl_fultpag")
    data_cabezal.Recordset("cl_ultmesp") = data_cablocal.Recordset("cl_ultmesp")
    data_cabezal.Recordset("cl_nomvend") = data_cablocal.Recordset("cl_nomvend")
    data_cabezal.Recordset("cl_fax") = data_cablocal.Recordset("cl_fax")
    data_cabezal.Recordset.Update
    'fin de cabezal
    data_cablocal.Recordset.Edit
    data_cablocal.Recordset("cl_codced") = 1
    data_cablocal.Recordset.Update
    Xdondeelerr = 6
    data_cabezal.Recordset.Close
    data_lin3.Recordset.Close
    data_lin2.Recordset.Close
   
'   MsgBox "Proceso de facturación terminado con éxito.", vbInformation
'   Unload Me
Else
   MsgBox "No hay líneas de facturación"
End If
'terminado


Exit Sub

Cierrosieser4:
             If Err.Number = 3155 Then
                data_erro.Recordset.AddNew
                data_erro.Recordset("id") = 5
                data_erro.Recordset("fecha") = Date
                data_erro.Recordset("hora") = Format(Time, "HH:mm")
                data_erro.Recordset("nroerr") = Err.Number
                data_erro.Recordset("desc") = "En el command4 " & Trim(Str(Xdondeelerr))
                data_erro.Recordset.Update
             Else
                data_erro.Recordset.AddNew
                data_erro.Recordset("id") = 5
                data_erro.Recordset("fecha") = Date
                data_erro.Recordset("hora") = Format(Time, "HH:mm")
                data_erro.Recordset("nroerr") = Err.Number
                data_erro.Recordset("desc") = "En el command4 " & Trim(Str(Xdondeelerr))
                data_erro.Recordset.Update
             End If
             End
             
End Sub

Private Sub Form_Load()
Dim Xmat, Xced, Xcodced As Long
Dim Xmasdiezui, Ximpdomi, Ximptrasla, XImp, Ximpsuma As Double
Dim XsocioOK, Xtrasla, XX, Xdiasfact As Integer
Dim Xclave, Xconvcod, Xclavefin As String
Dim Xfec1, Xfec2 As Date

Dim Xlaui As Double

'''On Error GoTo Cierrosieser

XsocioOK = 0
Ximpdomi = 0
Ximptrasla = 0
data_hist.Connect = "ODBC;DSN=sappnew;"
data_hist.RecordSource = "Select * from abmsocio where fecha >=#" & Format(Date, "yyyy/mm/dd") & "#"
data_hist.Refresh
labtimbre.Caption = "NO"
data_erro.DatabaseName = App.Path & "\errores.mdb"
data_erro.RecordSource = "errores"
data_erro.Refresh

data_cablocal.DatabaseName = App.Path & "\cablocal.mdb"
data_cablocal.RecordSource = "cabezados"
data_cablocal.Refresh

data_eror.DatabaseName = App.Path & "\erores.mdb"
data_eror.RecordSource = "erores"
data_eror.Refresh
data_llamado.Connect = "odbc;dsn=sappnew;"
data_ui.Connect = "ODBC;DSN=sappnew;"
data_ui.RecordSource = "hc_frecresp"
data_ui.Refresh
'''''''''''''''''' CONVENIO PARTICULAR PRECIOS!!!!
''''''''''''''''' altas de ficha con valor cero en la cédula
'''''''''''''''''TRASLADOS COORDINADOS NO LO FACTURA
'''''''''''''''''CATEGORIA MSP
'''''''' si ya esta facturado que no vuelva a facturar
''''1073 en emision y no incluir en la tabla, y los informes de documentos no procesados
'''' CUANDO ES AREA PROTEG Y NO TIENE DATOS DEL SOCIO
'''cuando es valor cero pasar a recibo

data_parse.DatabaseName = App.Path & "\PARSE.mdb"
data_parse.RecordSource = "parsec0"
data_parse.Refresh

data_ctrlf.DatabaseName = App.Path & "\ctrf.mdb"
data_ctrlf.RecordSource = "ctrf"
data_ctrlf.Refresh

data_conslla.Connect = "odbc;dsn=sappnew;"
Xfec1 = data_ctrlf.Recordset("fecha")
Xfec2 = Date
Xdiasfact = DateDiff("d", Xfec1, Xfec2)

data_temp.DatabaseName = App.Path & "\factura.mdb"

If frm_factedesp.labnro.Caption <> "" Then
   labnrollamado.Caption = frm_factedesp.labnro.Caption
Else
   labnrollamado.Caption = 0
End If
data_conve.Connect = "odbc;dsn=sappnew;"

data_cabeza2.DatabaseName = App.Path & "\factura.mdb"
data_cabeza2.RecordSource = "cabezados"
data_cabeza2.Refresh
If data_cabeza2.Recordset.RecordCount > 0 Then
   data_cabeza2.Recordset.MoveFirst
   Do While Not data_cabeza2.Recordset.EOF
      data_cabeza2.Recordset.Delete
      data_cabeza2.Recordset.MoveNext
   Loop
End If
data_cabeza2.Refresh
    
Xcandelin = 0
Xmat = 0
Xced = 0
    
Xtot = 0
Xsubt = 0
Xivvva = 0
Xivauno = 0
    'MsgBox "Verifique los datos de facturación y presione el botón de ACEPTAR", vbInformation
    
mf.Text = Format(frm_factedesp.txt_fecha.Text, "dd/mm/yyyy")
labhora.Caption = Format(frm_factedesp.labhora.Caption, "HH:mm")
    
data_verfac.Connect = "ODBC;DSN=sappnew;"
    
data_verfac.Connect = "ODBC;DSN=sappnew;"
    
data_cabezal.Connect = "ODBC;DSN=sappnew;"
data_cabezal.RecordSource = "Select * from clirespl where cl_codigo =" & 25048
data_cabezal.Refresh
    
data_deudas.Connect = "ODBC;DSN=sappnew;"
    
data_lin2.Connect = "ODBC;DSN=sappnew;"
    
data_caja.Connect = "ODBC;DSN=sappnew;"
data_caja.RecordSource = "select * from caja where fecha >=#" & Format(Date, "yyyy/mm/dd") & "#"
data_caja.Refresh
    
data_estudios.Connect = "ODBC;DSN=sappnew;"
    
data_arancel.Connect = "ODBC;DSN=sappnew;"
   
data_lin3.Connect = "ODBC;DSN=sappnew;"

data_lin.Connect = "ODBC;DSN=sappnew;"
    
data_par.Connect = "ODBC;DSN=sappfact;"
data_par.RecordSource = "paramsapp"
data_par.Refresh
    
t_pie.Text = "Caja de Usuario: " & frm_factedesp.labusuario & " FECHA:" & Format(Date, "dd/mm/yyyy")
t_mat.Text = 0
Xtrasla = 0
If frm_factedesp.labtrasla.Caption <> "" Then
   If Val(frm_factedesp.labtrasla.Caption) >= 0 Then
      Xtrasla = frm_factedesp.labtrasla.Caption
   Else
      Xtrasla = 50
   End If
Else
   Xtrasla = 50
End If
labusuario.Caption = frm_factedesp.labusuario
data_cli.Connect = "ODBC;DSN=sappnew;"
If frm_factedesp.labmatric.Caption <> "" Then
   Xmat = frm_factedesp.labmatric.Caption
Else
   Xmat = 0
End If
If frm_factedesp.labced.Caption <> "" Then
   Xced = frm_factedesp.labced.Caption
   Xcodced = frm_factedesp.labcodced.Caption
Else
   Xced = 0
   Xcodced = 0
End If
Xclavefin = ""
Xclave = frm_factedesp.labclave.Caption
Xclavefin = frm_factedesp.codfinal.Caption
Xconvcod = frm_factedesp.labcateg.Caption
t_imp.Text = 0
labcodest.Caption = ""
    'celestes
If Xclavefin = "V" Then
   If Xtrasla = 1 Or Xtrasla = 2 Or Xtrasla = 14 Then
      labcodtras.Caption = 90006
   Else
      If Xtrasla = 10 Or Xtrasla = 11 Then
         labcodtras.Caption = 90007
      Else
         If Xtrasla = 9 Then
            labcodtras.Caption = 90015
         Else
            If Xtrasla = 15 Then
               labcodtras.Caption = 90003
            End If
         End If
      End If
   End If
Else
   If Xclavefin = "A" Then
      If Xtrasla = 1 Or Xtrasla = 2 Or Xtrasla = 14 Then
         labcodtras.Caption = 90005
      Else
         If Xtrasla = 10 Or Xtrasla = 11 Then
            labcodtras.Caption = 90007
         Else
            If Xtrasla = 9 Then
               labcodtras.Caption = 90015
            Else
               If Xtrasla = 15 Then
                  labcodtras.Caption = 90002
               End If
            End If
         End If
      End If
   Else
      If Xclavefin = "R" Or Xclavefin = "N" Then
         If Xtrasla = 1 Or Xtrasla = 2 Or Xtrasla = 14 Then
            labcodtras.Caption = 90005
         Else
            If Xtrasla = 10 Or Xtrasla = 11 Then
               labcodtras.Caption = 90007
            Else
               If Xtrasla = 9 Then
                  labcodtras.Caption = 90015
               Else
                  If Xtrasla = 15 Then
                     labcodtras.Caption = 90002
                  End If
               End If
            End If
         End If
      Else
         If Xtrasla = 1 Or Xtrasla = 2 Or Xtrasla = 14 Then
            labcodtras.Caption = 90006
         Else
            If Xtrasla = 10 Or Xtrasla = 11 Then
               labcodtras.Caption = 90007
            Else
               If Xtrasla = 9 Then
                  labcodtras.Caption = 90015
               Else
                  If Xtrasla = 15 Then
                     labcodtras.Caption = 90003
                  End If
               End If
            End If
         End If
      End If
   End If
End If
    
If Xclave = "V" Or Xclave = "C" Then
   If Xconvcod = "911" Or Xconvcod = "911B" Then
      labcodest.Caption = 10014
   Else
      labcodest.Caption = 10002
   End If
Else
   If Xclave = "A" Then
      If Xconvcod = "911" Or Xconvcod = "911B" Then
         labcodest.Caption = 10013
      Else
         labcodest.Caption = 10004
      End If
   Else
      If Xclave = "R" Then
         If Xconvcod = "911" Or Xconvcod = "911B" Then
            labcodest.Caption = 10012
         Else
            labcodest.Caption = 10006
         End If
      Else
         If Xclave = "Z" Then
            labcodest.Caption = 14004
         Else
            If Xclave = "N" Then
               If Xconvcod = "911" Or Xconvcod = "911B" Then
                  labcodest.Caption = 10012
               Else
                  labcodest.Caption = 10016
               End If
            Else
               labcodest.Caption = 10016
            End If
         End If
      End If
   End If
End If
If Xconvcod = "MSP" Then
   If Xclave = "R" Then
      labcodest.Caption = 90017
   Else
      If Xclave = "A" Then
         labcodest.Caption = 90018
      Else
         labcodest.Caption = 90019
      End If
   End If
End If
If Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "CERCAS" Then
   labcodest.Caption = 10008
End If
If Val(frm_factedesp.labmovilpas.Caption) = 2015 Then
   labcodest.Caption = 10018
End If
If Xtrasla <= 0 Then
   List1.AddItem "SERVICIOS DOMICILIARIOS"
   labcodsrv.Caption = 10017
Else
    '   List1.AddItem "SERVICIOS DOMICILIARIOS"
   List1.AddItem "SERVICIOS DE TRASLADO"
   labcodsrv.Caption = 90021
End If
    'labcodsrv.Caption = 10017
Dim Xop1 As Integer
Dim Xval_estudio, Xdescuento As Double

List1.ListIndex = 0
cbofpago.ListIndex = 0
If Xdiasfact <= 5 Then
   If Xmat <> 0 Then
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         t_nombre.Text = Mid(frm_factedesp.labnom.Caption, 1, 60)
         t_mat.Text = data_cli.Recordset("cl_codigo")
         t_codcnv.Text = Xconvcod
         t_nomcnv.Text = Mid(frm_factedesp.labnomcat.Caption, 1, 30)
         If IsNull(data_cli.Recordset("cl_cedula")) = False Then
            If IsNull(data_cli.Recordset("cl_codced")) = False Then
               t_ced.Text = data_cli.Recordset("cl_cedula")
               t_codced.Text = data_cli.Recordset("cl_codced")
            Else
               t_codced.Text = ""
            End If
         Else
            t_ced.Text = ""
         End If
         Xtipodedocumento = 4
      Else
         XsocioOK = 0
         data_cli.RecordSource = "clientes"
         data_cli.Refresh
         data_cli.Recordset.AddNew
         data_cli.Recordset("cl_codigo") = data_par.Recordset("nro_socio") + 1
         Xmat = data_par.Recordset("nro_socio") + 1
         data_par.Recordset.Edit
         data_par.Recordset("nro_socio") = data_par.Recordset("nro_socio") + 1
         data_par.Recordset.Update
         data_cli.Recordset("cl_cedula") = Xced
         data_cli.Recordset("cl_codced") = Xcodced
         data_cli.Recordset("cl_apellid") = Mid(frm_factedesp.labnom.Caption, 1, 60)
         data_cli.Recordset("cl_codconv") = frm_factedesp.labcateg.Caption
         data_cli.Recordset("cl_nomconv") = Mid(frm_factedesp.labnomcat.Caption, 1, 30)
         data_cli.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
         If Val(frm_factedesp.labsexo.Caption) = 1 Then
            data_cli.Recordset("cl_sexo") = 2
         Else
            data_cli.Recordset("cl_sexo") = 1
         End If
         data_cli.Recordset("estado") = 1
         data_cli.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
         If Trim(frm_factedesp.labtelef.Caption) <> "" Then
            data_cli.Recordset("cl_telefon") = Mid(frm_factedesp.labtelef.Caption, 1, 20)
         End If
         If Trim(frm_factedesp.labdire.Caption) <> "" Then
            data_cli.Recordset("cl_direcci") = frm_factedesp.labdire.Caption
         End If
         If Trim(frm_factedesp.lablocal.Caption) <> "" Then
            data_cli.Recordset("cl_entre") = frm_factedesp.lablocal.Caption
            data_cli.Recordset("cl_zona") = Mid(frm_factedesp.lablocal.Caption, 1, 25)
         End If
         data_cli.Recordset.Update
         data_cli.Refresh
         t_mat.Text = Xmat
         t_nombre.Text = frm_factedesp.labnom.Caption
         t_codcnv.Text = frm_factedesp.labcateg.Caption
         t_nomcnv.Text = frm_factedesp.labnomcat.Caption
         t_ced.Text = Xced
         t_codced.Text = Xcodced
         Xtipodedocumento = 4
         data_hist.Recordset.AddNew
         data_hist.Recordset("cl_codigo") = t_mat.Text
         data_hist.Recordset("cl_motivo") = "ALTA DESPACHO"
         data_hist.Recordset("desc") = "ALTA"
         data_hist.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
         data_hist.Recordset("hora") = Format(Time, "HH:mm")
         data_hist.Recordset("usuario") = "DESPACHO"
         data_hist.Recordset("convenio") = t_codcnv.Text
         data_hist.Recordset("base") = 19
         data_hist.Recordset.Update
      End If
   Else
      If Xced <> 0 Then
         data_cli.RecordSource = "Select * from clientes where int(cl_cedula) =" & Int(Xced)
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            t_nombre.Text = Mid(frm_factedesp.labnom.Caption, 1, 60)
            t_mat.Text = data_cli.Recordset("cl_codigo")
            t_codcnv.Text = Xconvcod
            t_nomcnv.Text = Mid(frm_factedesp.labnomcat.Caption, 1, 30)
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If IsNull(data_cli.Recordset("cl_codced")) = False Then
                  t_ced.Text = data_cli.Recordset("cl_cedula")
                  t_codced.Text = data_cli.Recordset("cl_codced")
               Else
                  t_codced.Text = ""
               End If
            Else
               t_ced.Text = ""
            End If
            Xmat = data_cli.Recordset("cl_codigo")
            Xtipodedocumento = 4
         Else
            data_cli.RecordSource = "clientes"
            data_cli.Refresh
            data_cli.Recordset.AddNew
            data_cli.Recordset("cl_codigo") = data_par.Recordset("nro_socio") + 1
            Xmat = data_par.Recordset("nro_socio") + 1
            data_par.Recordset.Edit
            data_par.Recordset("nro_socio") = data_par.Recordset("nro_socio") + 1
            data_par.Recordset.Update
            data_cli.Recordset("cl_cedula") = Xced
            data_cli.Recordset("cl_codced") = Xcodced
            data_cli.Recordset("cl_apellid") = Mid(frm_factedesp.labnom.Caption, 1, 60)
            data_cli.Recordset("cl_codconv") = frm_factedesp.labcateg.Caption
            data_cli.Recordset("cl_nomconv") = Mid(frm_factedesp.labnomcat.Caption, 1, 30)
            data_cli.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
            If Val(frm_factedesp.labsexo.Caption) = 1 Then
               data_cli.Recordset("cl_sexo") = 2
            Else
               data_cli.Recordset("cl_sexo") = 1
            End If
            data_cli.Recordset("estado") = 1
            data_cli.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
            If Trim(frm_factedesp.labtelef.Caption) <> "" Then
               data_cli.Recordset("cl_telefon") = Mid(frm_factedesp.labtelef.Caption, 1, 20)
            End If
            If Trim(frm_factedesp.labdire.Caption) <> "" Then
               data_cli.Recordset("cl_direcci") = frm_factedesp.labdire.Caption
            End If
            If Trim(frm_factedesp.lablocal.Caption) <> "" Then
               data_cli.Recordset("cl_entre") = frm_factedesp.lablocal.Caption
               data_cli.Recordset("cl_zona") = Mid(frm_factedesp.lablocal.Caption, 1, 25)
            End If
            data_cli.Recordset.Update
            data_cli.Refresh
            t_mat.Text = Xmat
            t_nombre.Text = frm_factedesp.labnom.Caption
            t_codcnv.Text = frm_factedesp.labcateg.Caption
            t_nomcnv.Text = frm_factedesp.labnomcat.Caption
            t_ced.Text = Xced
            t_codced.Text = Xcodced
            Xtipodedocumento = 4
            data_hist.Recordset.AddNew
            data_hist.Recordset("cl_codigo") = t_mat.Text
            data_hist.Recordset("cl_motivo") = "ALTA DESPACHO"
            data_hist.Recordset("desc") = "ALTA"
            data_hist.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
            data_hist.Recordset("hora") = Format(Time, "HH:mm")
            data_hist.Recordset("usuario") = "DESPACHO"
            data_hist.Recordset("convenio") = t_codcnv.Text
            data_hist.Recordset("base") = 19
            data_hist.Recordset.Update
         End If
      Else
         data_cli.RecordSource = "clientes"
         data_cli.Refresh
         data_cli.Recordset.AddNew
         data_cli.Recordset("cl_codigo") = data_par.Recordset("nro_socio") + 1
         Xmat = data_par.Recordset("nro_socio") + 1
         data_par.Recordset.Edit
         data_par.Recordset("nro_socio") = data_par.Recordset("nro_socio") + 1
         data_par.Recordset.Update
         data_cli.Recordset("cl_cedula") = Xced
         data_cli.Recordset("cl_codced") = 0
         data_cli.Recordset("cl_apellid") = Mid(frm_factedesp.labnom.Caption, 1, 60)
         data_cli.Recordset("cl_codconv") = frm_factedesp.labcateg.Caption
         data_cli.Recordset("cl_nomconv") = Mid(frm_factedesp.labnomcat.Caption, 1, 30)
         data_cli.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
         If Val(frm_factedesp.labsexo.Caption) = 1 Then
            data_cli.Recordset("cl_sexo") = 2
         Else
            data_cli.Recordset("cl_sexo") = 1
         End If
         data_cli.Recordset("estado") = 1
         data_cli.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
         If Trim(frm_factedesp.labtelef.Caption) <> "" Then
            data_cli.Recordset("cl_telefon") = Mid(frm_factedesp.labtelef.Caption, 1, 20)
         End If
         If Trim(frm_factedesp.labdire.Caption) <> "" Then
            data_cli.Recordset("cl_direcci") = frm_factedesp.labdire.Caption
         End If
         If Trim(frm_factedesp.lablocal.Caption) <> "" Then
            data_cli.Recordset("cl_entre") = frm_factedesp.lablocal.Caption
            data_cli.Recordset("cl_zona") = Mid(frm_factedesp.lablocal.Caption, 1, 25)
         End If
         data_cli.Recordset.Update
         data_cli.Refresh
         t_mat.Text = Xmat
         t_nombre.Text = frm_factedesp.labnom.Caption
         t_codcnv.Text = frm_factedesp.labcateg.Caption
         t_nomcnv.Text = frm_factedesp.labnomcat.Caption
         t_ced.Text = Xced
         t_codced.Text = 0
         Xtipodedocumento = 4
         data_hist.Recordset.AddNew
         data_hist.Recordset("cl_codigo") = t_mat.Text
         data_hist.Recordset("cl_motivo") = "ALTA DESPACHO"
         data_hist.Recordset("desc") = "ALTA"
         data_hist.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
         data_hist.Recordset("hora") = Format(Time, "HH:mm")
         data_hist.Recordset("usuario") = "DESPACHO"
         data_hist.Recordset("convenio") = t_codcnv.Text
         data_hist.Recordset("base") = 19
         data_hist.Recordset.Update
      End If
   End If
            
   t_imp.Text = 0
   t_imp.Text = Format(t_imp.Text, "Standard")
   t_total.Text = t_imp.Text
   t_iva.Text = 0
   t_iva.Text = Format(t_iva.Text, "Standard")
   Ximpdomi = 0
   Ximptrasla = 0
   If XsocioOK = 8 Then
'      MsgBox "Hay un error en los datos del paciente, verifique para poder facturar"
      data_erro.Recordset.AddNew
      data_erro.Recordset("id") = 19
      data_erro.Recordset("fecha") = Date
      data_erro.Recordset("hora") = Format(Time, "HH:mm")
      data_erro.Recordset("nroerr") = 8
      data_erro.Recordset("desc") = "Alta de ficha"
      data_erro.Recordset.Update
      End
   Else
      If Xtrasla = 1 Or Xtrasla = 2 Or Xtrasla = 14 Or Xtrasla = 10 Or _
         Xtrasla = 11 Or Xtrasla = 9 Or Xtrasla = 15 Then
         If frm_factedesp.labcateg.Caption = "MSP" Then
            data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
            data_estudios.Refresh
         Else
            data_estudios.RecordSource = "Select * from estudios where codest =" & labcodtras.Caption
            data_estudios.Refresh
         End If
      Else
         data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
         data_estudios.Refresh
      End If
      data_conve.RecordSource = "Select * from convenio where cnv_codigo ='" & Xconvcod & "'"
      data_conve.Refresh
      If data_conve.Recordset.RecordCount > 0 Then
         If IsNull(data_conve.Recordset("cnv_aran")) = False Then
            Xop1 = data_conve.Recordset("cnv_aran")
         Else
            Xop1 = 0
         End If
      Else
         Xop1 = 0
      End If
      If data_estudios.Recordset.RecordCount > 0 Then
         labdescest.Caption = data_estudios.Recordset("descrip")
         If Xconvcod = "PART" Then
            Xval_estudio = data_estudios.Recordset("part")
         Else
            Xval_estudio = data_estudios.Recordset("cons")
         End If
         data_arancel.RecordSource = "Select * from Aran_servicios where id_gpo =" & Xop1 & " and id_serv =" & Val(labcodest.Caption)
'         data_arancel.RecordSource = "Select * from arancel where ara_cnvcod ='" & Xconvcod & "' and ara_famnro =" & data_estudios.Recordset("codest")
         data_arancel.Refresh
         If data_arancel.Recordset.RecordCount > 0 Then
            If data_arancel.Recordset("prec_serv") > 0 Then
               XImp = data_arancel.Recordset("prec_serv")
            Else
               If data_arancel.Recordset("por_serv") = 100 Then
                  XImp = 0
               Else
                  If data_arancel.Recordset("por_serv") = 0 Then
                     XImp = Xval_estudio
                  Else
                     Xdescuento = data_arancel.Recordset("por_serv") * Xval_estudio / 100
                     XImp = Xval_estudio - Xdescuento
                  End If
               End If
            End If
         Else
            XImp = data_estudios.Recordset("part")
         End If
         data_arancel.Recordset.Close
      Else
         XImp = 0
      End If
      t_imp.Text = Format(XImp, "Standard")
      t_total.Text = t_imp.Text
      t_iva.Text = t_imp.Text * 0.1 / 1.1
      t_iva.Text = Format(t_iva.Text, "Standard")
'hasta aqui
'''"CERSEM" Or Xconvcod = "CERCAS"
      If Xconvcod <> "PART" Then
         If data_conve.Recordset.RecordCount > 0 Then
            If IsNull(data_conve.Recordset("cnv_colrec")) = False Then
               If Xconvcod = "SA" Or Xconvcod = "SAF" Or Xconvcod = "CCOMS" Or Xconvcod = "CERCAS" Or Xconvcod = "SEGAM" Or _
                  Xconvcod = "EMERN" Or Xconvcod = "EMERG" Or Xconvcod = "EMERNE" Or Xconvcod = "CAAM" Or Xconvcod = "911" Or _
                  Xconvcod = "SAP" Or Xconvcod = "VIV19" Or Xconvcod = "VIV20" Or Xconvcod = "CAAMEP" Or Xconvcod = "911B" Or _
                  Xconvcod = "MSP" Or Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "APNORE" Or Xconvcod = "CASH" Or _
                  Xconvcod = "SJ01" Or data_conve.Recordset("cnv_colrec") = "M" Or data_conve.Recordset("cnv_colrec") = "R" Or _
                  Mid(Xconvcod, 1, 4) = "TALA" Or frm_factedesp.labcodzon.Caption = 3 Then
                  t_imp.Text = 0
                  t_imp.Text = Format(t_imp.Text, "Standard")
                  t_total.Text = t_imp.Text
                  t_iva.Text = 0
                  t_iva.Text = Format(t_iva.Text, "Standard")
                  Ximpdomi = 0
                  Ximptrasla = 0
               End If
            Else
               If Xconvcod = "SA" Or Xconvcod = "SAF" Or Xconvcod = "CCOMS" Or Xconvcod = "CERCAS" Or Xconvcod = "SEGAM" Or _
                  Xconvcod = "EMERN" Or Xconvcod = "EMERG" Or Xconvcod = "EMERNE" Or Xconvcod = "CAAM" Or Xconvcod = "911" Or _
                  Xconvcod = "SAP" Or Xconvcod = "VIV19" Or Xconvcod = "VIV20" Or Xconvcod = "CAAMEP" Or Xconvcod = "911B" Or _
                  Xconvcod = "MSP" Or Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "APNORE" Or Xconvcod = "CASH" Or _
                  Xconvcod = "SJ01" Or Mid(Xconvcod, 1, 4) = "TALA" Or frm_factedesp.labcodzon.Caption = 3 Then
                  t_imp.Text = 0
                  t_imp.Text = Format(t_imp.Text, "Standard")
                  t_total.Text = t_imp.Text
                  t_iva.Text = 0
                  t_iva.Text = Format(t_iva.Text, "Standard")
                  Ximpdomi = 0
                  Ximptrasla = 0
               End If
            End If
         Else
            If Xconvcod = "SA" Or Xconvcod = "SAF" Or Xconvcod = "CCOMS" Or Xconvcod = "CERCAS" Or Xconvcod = "SEGAM" Or _
               Xconvcod = "EMERN" Or Xconvcod = "EMERG" Or Xconvcod = "EMERNE" Or Xconvcod = "CAAM" Or _
               Xconvcod = "SAP" Or Xconvcod = "VIV19" Or Xconvcod = "VIV20" Or Xconvcod = "CAAMEP" Or _
               Xconvcod = "MSP" Or Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "APNORE" Or _
               Xconvcod = "SJ01" Or Mid(Xconvcod, 1, 4) = "TALA" Or frm_factedesp.labcodzon.Caption = 3 Then
               t_imp.Text = 0
               t_imp.Text = Format(t_imp.Text, "Standard")
               t_total.Text = t_imp.Text
               t_iva.Text = 0
               t_iva.Text = Format(t_iva.Text, "Standard")
               Ximpdomi = 0
               Ximptrasla = 0
            End If
         End If
      End If
      If Xconvcod = "MUCAFL" Or Xconvcod = "MUCAMA" Or Xconvcod = "MUCAMI" Or Xconvcod = "MUCAMM" Or Xconvcod = "MUCAMP" Or _
         Xconvcod = "MUCAMS" Or Xconvcod = "MUCAMT" Or Xconvcod = "MUCATA" Or Xconvcod = "SOLEME" Or Xconvcod = "911B" Or _
         Xconvcod = "CAAMEP" Or Xconvcod = "SOLAF" Or Xconvcod = "SOLAMB" Or Xconvcod = "SOC" Or Xconvcod = "911" Or _
         Xconvcod = "CASH" Or Xconvcod = "911B" Or Xconvcod = "CERSEM" Or Xconvcod = "CPS" Then
         t_imp.Text = 0
         t_imp.Text = Format(t_imp.Text, "Standard")
         t_total.Text = t_imp.Text
         t_iva.Text = 0
         t_iva.Text = Format(t_iva.Text, "Standard")
         Ximpdomi = 0
         Ximptrasla = 0
      End If
      If frm_factedesp.labcosto.Caption <> "" Then
         If frm_factedesp.labcosto.Caption > 0 Then
            t_imp.Text = Format(frm_factedesp.labcosto.Caption, "Standard")
            t_total.Text = t_imp.Text
            t_iva.Text = t_imp.Text * 0.1 / 1.1
            t_iva.Text = Format(t_iva.Text, "Standard")
         Else
            t_imp.Text = 0
            t_imp.Text = Format(t_imp.Text, "Standard")
            t_total.Text = t_imp.Text
            t_iva.Text = 0
            t_iva.Text = Format(t_iva.Text, "Standard")
            Ximpdomi = 0
            Ximptrasla = 0
         End If
      Else
      End If
      If frm_factedesp.labtick.Caption = 1 Then
         t_imp.Text = 0
         t_imp.Text = Format(t_imp.Text, "Standard")
         t_total.Text = t_imp.Text
         t_iva.Text = 0
         t_iva.Text = Format(t_iva.Text, "Standard")
         Ximpdomi = 0
         Ximptrasla = 0
      End If
'      If Val(frm_factedesp.labmovilpas.Caption) = 2015 Then
'         t_imp.Text = 0
'         t_imp.Text = Format(t_imp.Text, "Standard")
'         t_total.Text = t_imp.Text
'         t_iva.Text = 0
'         t_iva.Text = Format(t_iva.Text, "Standard")
'         Ximpdomi = 0
'         Ximptrasla = 0
'      End If
      If t_total.Text > 0 Then
         If data_ui.Recordset.RecordCount > 0 Then
            Xlaui = CDbl(data_ui.Recordset("descrip"))
            Xmasdiezui = CDbl(t_total.Text) / Xlaui
         End If
      End If
      If t_imp.Text > 0 Then
         If Xmasdiezui > 10000 Then
            Label7.Caption = "E-FACTURA"
            labnrofact.Caption = ""
            labserie.Caption = ""
         Else
            Label7.Caption = "E-TICKET"
            labnrofact.Caption = ""
            labserie.Caption = ""
         End If
      End If
      If Trim(frm_factedesp.labtimbre.Caption) <> "" Then
         labtimbre.Caption = frm_factedesp.labtimbre.Caption
      Else
         labtimbre.Caption = ""
      End If
      If labtimbre.Caption = "SI" Then
         data_estudios.RecordSource = "Select * from estudios where codest =" & 995
         data_estudios.Refresh
         If data_estudios.Recordset.RecordCount > 0 Then
            t_total.Text = Val(t_total.Text) + data_estudios.Recordset("cons")
            t_imptimbre.Text = Val(data_estudios.Recordset("cons"))
         End If
      End If
      data_temp.DatabaseName = App.Path & "\factura.mdb"
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
      data_estudios.Recordset.Close
      data_conve.Recordset.Close
      
      If t_total.Text <> "" Then
         If t_total.Text > 0 Then
            If Xmasdiezui > 10000 Then
               If Xtipodedocumento = 4 Then
                  If t_ced.Text <> "" Then
                     Xtipodedocumento = 3
                     Command1_Click
                  Else
                  End If
               Else
                  If t_ced.Text <> "" Then
                     Xtipodedocumento = 3
                     Command1_Click
                  Else
                  End If
               End If
            Else
               Command1_Click
            End If
         Else
            b_findos_Click
         End If
      Else
         b_findos_Click
      End If
   End If
Else
   MsgBox "ERROR EN FECHA DEL SISTEMA COMUNIQUE A ADMINISTRACION!!! FECHA ACTUAL:" & Date
   data_ctrlf.Recordset.Edit
   data_ctrlf.Recordset("fecha") = Date
   data_ctrlf.Recordset("hora") = Format(Time, "HH:mm")
   data_ctrlf.Recordset.Update
    data_erro.Recordset.AddNew
    data_erro.Recordset("id") = 6
    data_erro.Recordset("fecha") = Date
    data_erro.Recordset("hora") = Format(Time, "HH:mm")
    data_erro.Recordset("nroerr") = 19
    data_erro.Recordset("desc") = "Fecha del sistema"
    data_erro.Recordset.Update
   End
End If

'Exit Sub

'Cierrosieser:
'              If Err.Number = 3155 Then
'                 data_erro.Recordset.AddNew
'                 data_erro.Recordset("id") = 6
'                 data_erro.Recordset("fecha") = Date
'                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
'                 data_erro.Recordset("nroerr") = Err.Number
'                 data_erro.Recordset("desc") = "En el load de despachofact2"
'                 data_erro.Recordset.Update
'                 MsgBox "NRO:" & labnrollamado.Caption
'              Else
'                 data_erro.Recordset.AddNew
'                 data_erro.Recordset("id") = 6
'                 data_erro.Recordset("fecha") = Date
'                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
'                 data_erro.Recordset("nroerr") = Err.Number
'                 data_erro.Recordset("desc") = "En el load de despachofact2"
'                 data_erro.Recordset.Update
'                 MsgBox "NRO:" & labnrollamado.Caption
'              End If
'              End
             
             
End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With


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

On Error GoTo Xxquepasaalenv

    If ResultadoCfe Is Nothing Then
        MsgBox Mensaje
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 1
        data_erro.Recordset("desc") = Mid(Trim(Mensaje), 1, 130)
        data_erro.Recordset.Update
        Exit Sub
    End If

    If Not ResultadoCfe.OperacionEjecutada Or ResultadoCfe.EstadoCfe Is Nothing Then
        If ResultadoCfe.Mensaje <> vbNullString Then Mensaje = Mensaje & ": " & ResultadoCfe.Mensaje
        MsgBox Mensaje
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 2
        data_erro.Recordset("desc") = Mid(Trim(Mensaje), 1, 130)
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.Error Then
        Mensaje = Mensaje & ", ocurrió un error"
        If ResultadoCfe.EstadoCfe.Mensaje <> vbNullString Then _
            Mensaje = Mensaje & ": " & ResultadoCfe.EstadoCfe.Mensaje
        MsgBox Mensaje
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 3
        data_erro.Recordset("desc") = Mid(Trim(Mensaje), 1, 130)
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.SerieNumeroCfe Is Nothing Then
        MsgBox "El CFE no trae número de folio, no se puede terminar la factura"
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 4
        data_erro.Recordset("desc") = "El CFE no trae número de folio"
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.DatosCae Is Nothing Then
        MsgBox "El CFE no trae datos del CAE, no se puede terminar la factura"
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 5
        data_erro.Recordset("desc") = "CFE no trae datos del CAE"
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If (CInt(ResultadoCfe.EstadoCfe.SerieNumeroCfe.TipoCFE) < 200) Then
        Dim strFile As String
        strFile = App.Path & "\qr.bmp"
        Dim objresultado As Resultado
        Set objresultado = objPosCfe.GenerarQr(ResultadoCfe.EstadoCfe.DatosQr, 100, strFile)

        Dim strMensaje As String
        strMensaje = "No se pudo generar el QR"

        If objresultado Is Nothing Then
            MsgBox strMensaje
            data_erro.Recordset.AddNew
            data_erro.Recordset("id") = 19
            data_erro.Recordset("fecha") = Date
            data_erro.Recordset("hora") = Format(Time, "HH:mm")
            data_erro.Recordset("nroerr") = 6
            data_erro.Recordset("desc") = Mid(Trim(strMensaje), 1, 130)
            data_erro.Recordset.Update
            
            Exit Sub
        End If

        If Not objresultado.OperacionExitosa Then
            If objresultado.Mensaje <> vbNullString Then strMensaje = strMensaje & ": " & objresultado.Mensaje
            MsgBox strMensaje
            data_erro.Recordset.AddNew
            data_erro.Recordset("id") = 19
            data_erro.Recordset("fecha") = Date
            data_erro.Recordset("hora") = Format(Time, "HH:mm")
            data_erro.Recordset("nroerr") = 7
            data_erro.Recordset("desc") = Mid(Trim(strMensaje), 1, 130)
            data_erro.Recordset.Update
            
            Exit Sub
        End If

'        imgQr.Picture = LoadPicture(strFile)
    End If
    If Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 0 Or _
       Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 11 Then
       labserie.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Serie
       labnrofact.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
    Else
       data_eror.Recordset.AddNew
       data_eror.Recordset("nro") = ResultadoCfe.EstadoCfe.CodigoRespuesta
       data_eror.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
       data_eror.Recordset("hora") = Format(Time, "HH:mm")
       data_eror.Recordset("obs") = "NOMBRE:" & t_nombre.Text & " CAT:" & t_codcnv.Text
       data_eror.Recordset.Update
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
 '       "Código de seguridad: " & ResultadoCfe.EstadoCfe.CodigoSeguridad & vbNewLine & _
'        "Código de respuesta: " & ResultadoCfe.EstadoCfe.CodigoRespuesta & vbNewLine & _
'        "Fecha de firma: " & ResultadoCfe.EstadoCfe.FechaFirma & vbNewLine & _
'        "GUID: " & ResultadoCfe.EstadoCfe.Guid & vbNewLine & _
'        "Mensaje: " & ResultadoCfe.EstadoCfe.Mensaje & vbNewLine & _
'        "Pendiente de envío: " & CStr(ResultadoCfe.EstadoCfe.PendienteDeEnvio) & vbNewLine

    strUltimoGuid = ResultadoCfe.EstadoCfe.Guid
    Set objUltimaSerieNumero = ResultadoCfe.EstadoCfe.SerieNumeroCfe

    'cmdConsultaXguid.Enabled = True
    'cmdConsultaXnumero.Enabled = True
Exit Sub

Xxquepasaalenv:
              If Err.Number = 3155 Then
                 data_erro.Recordset.AddNew
                 data_erro.Recordset("id") = 7
                 data_erro.Recordset("fecha") = Date
                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
                 data_erro.Recordset("nroerr") = Err.Number
                 data_erro.Recordset("desc") = "Al desplegar infor e-tck"
                 data_erro.Recordset.Update
              Else
                 data_erro.Recordset.AddNew
                 data_erro.Recordset("id") = 7
                 data_erro.Recordset("fecha") = Date
                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
                 data_erro.Recordset("nroerr") = Err.Number
                 data_erro.Recordset("desc") = "Al desplegar infor e-tck"
                 data_erro.Recordset.Update
              End If
              End


End Sub

