VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_despachofact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos para la factura"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   Icon            =   "frm_despachofact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
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
      Connect         =   "odbc;dsn=sapp;"
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
      Caption         =   "FIN"
      Height          =   495
      Left            =   3600
      TabIndex        =   36
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
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
      ItemData        =   "frm_despachofact.frx":058A
      Left            =   1800
      List            =   "frm_despachofact.frx":0594
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
      Picture         =   "frm_despachofact.frx":05AA
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
      Picture         =   "frm_despachofact.frx":0B34
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
      Picture         =   "frm_despachofact.frx":10BE
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "frm_despachofact"
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

Private Sub b_efct_Click()
Dim strIdTransac As String

If Label7.Caption = "E-FACTURA" Then

    Set objPosCfe = New PosCfe
    
    Dim objresultado As Resultado
'    Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
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
        MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
            "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
        
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
              .cantidad.FromString Trim(Str(data_temp.Recordset("cantidad")))
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

If Label7.Caption = "E-TICKET" Then

    Set objPosCfe = New PosCfe
    
    Dim objresultado As Resultado
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
          .cantidad.FromString Trim(Str(data_temp.Recordset("cantidad")))
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

End Sub

Private Sub b_findos_Click()
Dim Xivauno As Double

Dim Xlf As Date
Dim Xelano, Xdiasfact, Xcandelin, XX2 As Integer
Dim Xfecvence As Date
Dim Xlatasa, Xlatasa22 As Double

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")

Xcandelin = 0
Xelano = Year(Date) + 1
Xfecvence = Format(mf.Text, "dd/mm/yyyy")
Xfecvence = Xfecvence + 30

XX2 = 0
Dim XnroMed As Integer
Dim Xnommed As String
If frm_largador.txt_codmed.Text <> "" Then
   If frm_largador.txt_codmed.Text > 0 Then
      XnroMed = frm_largador.txt_codmed.Text
      Xnommed = frm_largador.dbcbomed.Text
   Else
      XnroMed = 440
      Xnommed = "OTROS MEDICOS"
   End If
Else
   XnroMed = 440
   Xnommed = "OTROS MEDICOS"
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
   If frm_largador.cbotras.ListIndex = 1 Or frm_largador.cbotras.ListIndex = 2 Or _
      frm_largador.cbotras.ListIndex = 14 Or frm_largador.cbotras.ListIndex = 10 Or _
      frm_largador.cbotras.ListIndex = 11 Or frm_largador.cbotras.ListIndex = 9 Or _
      frm_largador.cbotras.ListIndex = 15 Then
      If frm_largador.txt_cat.Text = "MSP" Then
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
   data_temp.Recordset("operador") = WElusuario
   data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
   data_temp.Recordset("imp_timbre") = t_imp.Text
   data_temp.Recordset("tot_lin") = t_total.Text
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
   data_temp.Recordset("arancel") = Format(t_total.Text, "Standard")
   data_temp.Recordset("nro_med_a") = XnroMed
   data_temp.Recordset("nom_med_a") = Xnommed
   data_temp.Recordset("precio_est") = t_total.Text
   data_temp.Recordset("imp_iva") = CDbl(t_iva.Text)
   data_temp.Recordset.Update
   
   If frm_largador.cbotras.ListIndex = 1 Or frm_largador.cbotras.ListIndex = 2 Or _
      frm_largador.cbotras.ListIndex = 14 Or frm_largador.cbotras.ListIndex = 10 Or _
      frm_largador.cbotras.ListIndex = 11 Or frm_largador.cbotras.ListIndex = 9 Or _
      frm_largador.cbotras.ListIndex = 15 Then
      If frm_largador.txt_cat.Text = "MSP" Then
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
        data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
        data_estudios.Refresh
        If data_estudios.Recordset.RecordCount > 0 Then
           data_temp.Recordset("cod_prod") = data_estudios.Recordset("codest")
           data_temp.Recordset("in_usuario") = "10017"
           data_temp.Recordset("nom_prod") = data_estudios.Recordset("descrip")
           data_temp.Recordset("nom_flia") = "SERVICIOS DOMICILIARIOS"
        End If
        data_temp.Recordset("operador") = WElusuario
        data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
        data_temp.Recordset("imp_timbre") = 0
        data_temp.Recordset("tot_lin") = 0
        data_temp.Recordset("base") = data_par.Recordset("codsuc")
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
        data_temp.Recordset("nom_med_a") = Xnommed
        data_temp.Recordset("precio_est") = 0
        data_temp.Recordset("imp_iva") = 0
        data_temp.Recordset.Update
      End If
   End If
   data_temp.Refresh
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
      data_lin.RecordSource = "linmmdd"
      data_lin.Refresh
      data_lin2.RecordSource = "Select * from hc_torax"
      data_lin2.Refresh
      data_lin3.RecordSource = "Select * from indica_enfc"
      data_lin3.Refresh
      data_cabezal.Recordset.AddNew
     '           data_cabezal.Recordset("id") = 1
      data_cabezal.Recordset("cl_tipcli") = "1.0"
      data_cabezal.Recordset("cl_tipocli") = 9
      data_cabezal.Recordset("cl_socmnro") = labserie.Caption
      data_cabezal.Recordset("cl_numero") = Val(labnrofact.Caption)
      data_cabezal.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
      data_cabezal.Recordset("cl_nrovend") = 1 'linea de detalle iva incluido
      data_cabezal.Recordset("cl_forpago") = 1
      data_cabezal.Recordset("cl_diacobr") = Trim(Str(data_par.Recordset("ruc")))
      data_cabezal.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
      data_cabezal.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
      data_cabezal.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
      data_cabezal.Recordset("cl_referen") = data_par.Recordset("domic")
      data_cabezal.Recordset("tit_tarj") = data_par.Recordset("ciudad")
      data_cabezal.Recordset("cl_nomconv") = data_par.Recordset("dpto")
             'receptor
      data_cabezal.Recordset("cl_nro_sup") = Xtipodedocumento 'tipo de documento
      data_cabezal.Recordset("hora_baja") = "UY" 'codigo del pais del documento
      data_cabezal.Recordset("cl_codigo") = t_mat.Text
      data_cabezal.Recordset("usu_baja") = "UYU" 'tipo de moneda
      data_cabezal.Recordset("saldo_doc2") = Format(t_total.Text, "Standard") 'total monto neto iva mínimo
      data_cabezal.Recordset("cl_atrasop") = Xlatasa
      data_cabezal.Recordset("cl_decuota") = Xlatasa22
      data_cabezal.Recordset("saldo_cc") = Format(t_iva.Text, "Standard")
      data_cabezal.Recordset("saldo_cc2") = 0 'iva básico
      data_cabezal.Recordset("saldo_doc") = Format(t_total.Text, "Standard")
      data_cabezal.Recordset("cl_grupo") = data_temp.Recordset.RecordCount 'nro de líneas
      data_cabezal.Recordset("saldo_chc") = Format(t_total.Text, "Standard")
      data_cabezal.Recordset.Update
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
         data_lin.Recordset.Update
                    
'         data_lin2.Recordset.AddNew
'         data_lin2.Recordset("hora") = "INT1"
'         data_lin2.Recordset("descrip") = labserie.Caption  'serie del comprobante
'         data_lin2.Recordset("hc_nro") = 2 'tasa minima
'         data_lin2.Recordset("hc_cod") = Val(labnrofact.Caption)
'         data_lin2.Recordset("hc_mat") = data_temp.Recordset("linea")
'         data_lin2.Recordset.Update
'         data_lin2.Refresh
                              
'         data_caja.Recordset.AddNew
'         data_caja.Recordset("fecha") = data_temp.Recordset("fecha")
'         data_caja.Recordset("numero") = data_temp.Recordset("rub_cont")
'         data_caja.Recordset("nombre") = Mid(data_temp.Recordset("rub_nomb"), 1, 35)
'         data_caja.Recordset("moneda") = "$"
'         data_caja.Recordset("movimiento") = "INGRESO"
'         data_caja.Recordset("imp_fact") = data_temp.Recordset("tot_lin")
'         data_caja.Recordset("documento") = labnrofact.Caption
'         If cbofpago.Text = "CREDITO" Then
'            data_caja.Recordset("observ") = "CREDITO " + labnrofact.Caption
'         Else
'            data_caja.Recordset("observ") = "CONTADO " + labnrofact.Caption
'         End If
'         data_caja.Recordset("saldo") = data_temp.Recordset("tot_lin")
'         data_caja.Recordset("usuario") = data_temp.Recordset("operador")
'         data_caja.Recordset("hora") = data_temp.Recordset("hora")
'         data_caja.Recordset("base") = data_temp.Recordset("base")
'         data_caja.Recordset("cod_serv") = data_temp.Recordset("cod_prod")
'         data_caja.Recordset("nom_serv") = Mid(data_temp.Recordset("nom_prod"), 1, 50)
'         data_caja.Recordset("cod_socio") = data_temp.Recordset("cod_cli")
'         data_caja.Recordset("nom_socio") = Mid(data_temp.Recordset("nom_cli"), 1, 30)
'         If IsNull(data_temp.Recordset("imp_iva")) = False Then
'            data_caja.Recordset("imp_iva") = Format(data_temp.Recordset("imp_iva"), "Standard")
'         Else
'            data_caja.Recordset("imp_iva") = 0
'         End If
'         data_caja.Recordset("opiva") = 1 ' 10% , 0 NO, 2 22%
'         data_caja.Recordset.Update
         
         data_temp.Recordset.MoveNext
      Loop
      Xcandelin = 0
      Xconlin = 0
      data_ctrlf.Recordset.Edit
      data_ctrlf.Recordset("fecha") = Date
      data_ctrlf.Recordset.Update

      data_llamado.RecordSource = "Select * from llamado where nrolla =" & labnrollamado.Caption
      data_llamado.Refresh
      If data_llamado.Recordset.RecordCount > 0 Then
         If data_llamado.Recordset("pend") <> 2 Then
            data_llamado.Recordset.Edit
            data_llamado.Recordset("pend") = 2
            data_llamado.Recordset("totend") = "FACT"
            data_llamado.Recordset.Update
         End If
      End If
       
'      MsgBox "Proceso de facturación terminado", vbInformation
   
   Else
      MsgBox "No hay líneas de facturación"
   End If
Else
   MsgBox "Hay un error en fecha o datos de la factura, NO SE PUEDE FACTURAR, REINTENTE!", vbCritical
End If
'terminado


End Sub

Private Sub Command1_Click()
Dim Xivauno As Double

Dim Xlf As Date
Dim Xelano, Xcandelin, XX2 As Integer
Dim Xfecvence As Date
Dim Xlatasa, Xlatasa22 As Double

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")

Xcandelin = 0
Xelano = Year(Date) + 1
Xfecvence = Format(mf.Text, "dd/mm/yyyy")
Xfecvence = Xfecvence + 30

XX2 = 0
Dim XnroMed As Integer
Dim Xnommed As String

If frm_largador.txt_codmed.Text <> "" Then
   If frm_largador.txt_codmed.Text > 0 Then
      XnroMed = frm_largador.txt_codmed.Text
      Xnommed = frm_largador.dbcbomed.Text
   Else
      XnroMed = 440
      Xnommed = "OTROS MEDICOS"
   End If
Else
   XnroMed = 440
   Xnommed = "OTROS MEDICOS"
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
   data_temp.Recordset("libro_rub") = Label7.Caption ' tipo de documento (Ej.e-ticket)
'       data_temp.Recordset("unidad") = labserie.Caption
   data_temp.Recordset("in_unid") = "INT1"
   data_temp.Recordset("in_mat") = 2 'gravado a tasa mínima
   data_temp.Recordset("cantidad") = 1
   If t_pie.Text <> "" Then
      data_temp.Recordset("in_obs") = t_pie.Text
   End If
'       data_temp.Recordset("factura") = labnrofact.Caption
   data_temp.Recordset("tipo") = cbofpago.Text 'contado/crédito
   data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
   data_temp.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
   data_temp.Recordset("cod_cli") = t_mat.Text
   data_temp.Recordset("nom_cli") = t_nombre.Text
   data_temp.Recordset("convenio") = t_codcnv.Text
   If frm_largador.cbotras.ListIndex = 1 Or frm_largador.cbotras.ListIndex = 2 Or _
      frm_largador.cbotras.ListIndex = 14 Or frm_largador.cbotras.ListIndex = 10 Or _
      frm_largador.cbotras.ListIndex = 11 Or frm_largador.cbotras.ListIndex = 9 Or _
      frm_largador.cbotras.ListIndex = 15 Then
      If frm_largador.txt_cat.Text = "MSP" Then
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
   data_temp.Recordset("operador") = WElusuario
   data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
   data_temp.Recordset("imp_timbre") = t_imp.Text
   data_temp.Recordset("tot_lin") = t_total.Text
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
   data_temp.Recordset("arancel") = Format(t_total.Text, "Standard")
   data_temp.Recordset("nro_med_a") = XnroMed
   data_temp.Recordset("nom_med_a") = Xnommed
   data_temp.Recordset("precio_est") = t_total.Text
   data_temp.Recordset("imp_iva") = CDbl(t_iva.Text)
   data_temp.Recordset.Update
   If frm_largador.cbotras.ListIndex = 1 Or frm_largador.cbotras.ListIndex = 2 Or _
      frm_largador.cbotras.ListIndex = 14 Or frm_largador.cbotras.ListIndex = 10 Or _
      frm_largador.cbotras.ListIndex = 11 Or frm_largador.cbotras.ListIndex = 9 Or _
      frm_largador.cbotras.ListIndex = 15 Then
      If frm_largador.txt_cat.Text = "MSP" Then
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
        data_temp.Recordset("nom_cli") = t_nombre.Text
        data_temp.Recordset("convenio") = t_codcnv.Text
        data_estudios.RecordSource = "Select * from estudios where codest =" & labcodest.Caption
        data_estudios.Refresh
        If data_estudios.Recordset.RecordCount > 0 Then
           data_temp.Recordset("cod_prod") = data_estudios.Recordset("codest")
           data_temp.Recordset("in_usuario") = "10017"
           data_temp.Recordset("nom_prod") = data_estudios.Recordset("descrip")
           data_temp.Recordset("nom_flia") = "SERVICIOS DOMICILIARIOS"
        End If
        data_temp.Recordset("operador") = WElusuario
        data_temp.Recordset("hora") = Format(labhora.Caption, "HH:mm")
        data_temp.Recordset("imp_timbre") = 0
        data_temp.Recordset("tot_lin") = 0
        data_temp.Recordset("base") = data_par.Recordset("codsuc")
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
        data_temp.Recordset("nom_med_a") = Xnommed
        data_temp.Recordset("precio_est") = 0
        data_temp.Recordset("imp_iva") = 0
        data_temp.Recordset.Update
      End If
   End If
   data_temp.Refresh
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
   If t_pie.Text <> "" Then
      data_cabeza2.Recordset("obsp") = t_pie.Text
   End If
   data_cabeza2.Recordset.Update
   data_cabeza2.Refresh
   If data_temp.Recordset.RecordCount > 0 Then
      b_etck_Click
'      Command4_Click
   End If
   'fin de cabezal
Else
   MsgBox "Faltan datos para poder grabar, verifique fecha!", vbInformation
End If
        
        

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command4_Click()
Dim Xlatasa, Xlatasa22 As Double
Dim Xcandelin As Integer

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")
'labserie.Caption = "A"
'labnrofact.Caption = 100011
If data_temp.Recordset.RecordCount > 0 Then
   data_lin.RecordSource = "linmmdd"
   data_lin.Refresh
   data_lin2.RecordSource = "Select * from hc_torax"
   data_lin2.Refresh
   data_lin3.RecordSource = "Select * from indica_enfc"
   data_lin3.Refresh
   data_cabezal.Recordset.AddNew
 '           data_cabezal.Recordset("id") = 1
   data_cabezal.Recordset("cl_tipcli") = "1.0"
   data_cabezal.Recordset("cl_telefon") = Label7.Caption
   If Label7.Caption = "E-FACTURA" Then
      data_cabezal.Recordset("cl_tipocli") = 111
   Else
      If Label7.Caption = "E-TICKET" Then
         data_cabezal.Recordset("cl_tipocli") = 101
      Else
         data_cabezal.Recordset("cl_tipocli") = 101
      End If
   End If
            
   data_cabezal.Recordset("cl_socmnro") = labserie.Caption
   data_cabezal.Recordset("cl_numero") = Val(labnrofact.Caption)
   data_cabezal.Recordset("cl_fnac") = Format(mf.Text, "dd/mm/yyyy")
'  data_cabezal.Recordset("fecha_reac") = Format(mfd.Text, "dd/mm/yyyy") 'fecha de servicios
'  data_cabezal.Recordset("cl_tj_venc") = Format(mfh.Text, "dd/mm/yyyy") idem
   data_cabezal.Recordset("cl_nrovend") = 1 'linea de detalle iva incluido
   If cbofpago.Text = "CONTADO" Then
      data_cabezal.Recordset("cl_forpago") = 1
   Else
      If cbofpago.Text = "CREDITO" Then
         data_cabezal.Recordset("cl_forpago") = 2
      Else
         data_cabezal.Recordset("cl_forpago") = 2
      End If
   End If
'         If labvence.Text <> "__/__/____" Then
 '           data_cabezal.Recordset("fecha_modi") = Format(labvence.Text, "dd/mm/yyyy")
'         End If
         'datos emisor
   data_cabezal.Recordset("cl_diacobr") = Trim(Str(data_par.Recordset("ruc")))
   data_cabezal.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
   data_cabezal.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
   data_cabezal.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
   data_cabezal.Recordset("cl_referen") = data_par.Recordset("domic")
   data_cabezal.Recordset("tit_tarj") = data_par.Recordset("ciudad")
   data_cabezal.Recordset("cl_nomconv") = data_par.Recordset("dpto")
        'receptor
   
   If Label7.Caption = "E-TICKET" Then
      data_cabezal.Recordset("cl_nro_sup") = Xtipodedocumento 'tipo de documento
   Else
      data_cabezal.Recordset("cl_nro_sup") = Xtipodedocumento '2 RUT, 3 CI
   End If
   data_cabezal.Recordset("hora_baja") = "UY" 'codigo del pais del documento
'         data_cabezal.Recordset("cl_nom_sup") = frm_convenios22.txt_ruc.Text
            'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
'         data_cabezal.Recordset("info_debit") = frm_convenios22.t_razon.Text
'            data_cabezal.Recordset("cl_direcci") = frm_convenios22.txt_direc.Text
'            If frm_convenios22.txt_localid.Text <> "" Then
'               data_cabezal.Recordset("cl_zona") = frm_convenios22.txt_localid.Text
'            End If
'            If frm_convenios22.t_dpto.Text <> "" Then
'               data_cabezal.Recordset("cl_email") = frm_convenios22.t_dpto.Text 'opcional
'            End If
            
'         data_cabezal.Recordset("cl_localid") = "URUGUAY" 'opcional
   data_cabezal.Recordset("cl_codigo") = t_mat.Text
   data_cabezal.Recordset("usu_baja") = "UYU" 'tipo de moneda
   data_cabezal.Recordset("saldo_doc2") = data_cabeza2.Recordset("saldo_doc2") 'total monto neto iva mínimo
   data_cabezal.Recordset("cl_atrasop") = Xlatasa
   data_cabezal.Recordset("cl_decuota") = Xlatasa22
   data_cabezal.Recordset("saldo_cc") = Format(t_iva.Text, "Standard")
   data_cabezal.Recordset("saldo_cc2") = 0 'iva básico
   data_cabezal.Recordset("saldo_doc") = Format(t_total.Text, "Standard")
   data_cabezal.Recordset("cl_grupo") = data_temp.Recordset.RecordCount 'nro de líneas
   data_cabezal.Recordset("saldo_chc") = Format(t_total.Text, "Standard")
   data_cabezal.Recordset.Update
  'fin de cabezal
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
       '      data_lin.Recordset("costo") = CDbl(labstot.Caption) ' sub total de la factura
      data_lin.Recordset("valor_iva") = data_temp.Recordset("pre_civa")
      data_lin.Recordset("base") = data_temp.Recordset("base")
'               data_lin.Recordset("ruc") = data_temp.Recordset("ruc")
'            data_lin.Recordset("tipo_mov") = data_temp.Recordset("tipo_mov")
'               data_lin.Recordset("grupo") = data_temp.Recordset("grupo") 'cobrador
'               data_lin.Recordset("solicitant") = data_temp.Recordset("solicitant")
      data_lin.Recordset("nom_med_a") = data_temp.Recordset("nom_med_a")
'               data_lin.Recordset("vto") = data_temp.Recordset("vto")
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
      data_lin.Recordset.Update
                    
      data_lin2.Recordset.AddNew
      data_lin2.Recordset("hora") = "INT1"
      data_lin2.Recordset("descrip") = labserie.Caption  'serie del comprobante
      data_lin2.Recordset("hc_nro") = 2 'tasa minima
      data_lin2.Recordset("hc_cod") = Val(labnrofact.Caption)
      data_lin2.Recordset("hc_mat") = data_temp.Recordset("linea")
      data_lin2.Recordset.Update
      data_lin2.Refresh
      data_caja.Recordset.AddNew
      data_caja.Recordset("fecha") = data_temp.Recordset("fecha")
      data_caja.Recordset("numero") = data_temp.Recordset("rub_cont")
      data_caja.Recordset("nombre") = Mid(data_temp.Recordset("rub_nomb"), 1, 35)
      data_caja.Recordset("moneda") = "$"
      data_caja.Recordset("movimiento") = "INGRESO"
      data_caja.Recordset("imp_fact") = data_temp.Recordset("tot_lin")
      data_caja.Recordset("documento") = labnrofact.Caption
      If cbofpago.Text = "CREDITO" Then
         data_caja.Recordset("observ") = "CREDITO " + labnrofact.Caption
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
    
      data_temp.Recordset.MoveNext
   Loop
   Xcandelin = 0
   Xconlin = 0
   
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveFirst
      Do While Not data_temp.Recordset.EOF
         If data_temp.Recordset("tot_lin") > 0 Then
            data_deudas.RecordSource = "Select * from deudas where cliente =" & data_temp.Recordset("cod_cli")
            data_deudas.Refresh
            data_deudas.Recordset.AddNew
            data_deudas.Recordset("cod_cnv") = frm_largador.txt_cat.Text
            data_deudas.Recordset("nom_cnv") = Mid(frm_largador.txt_nomcat.Text, 1, 20)
            data_deudas.Recordset("cliente") = data_temp.Recordset("cod_cli")
            data_deudas.Recordset("nombre") = data_temp.Recordset("nom_cli")
            data_deudas.Recordset("fecha") = Date
            data_deudas.Recordset("tipodoc") = "CRE"
            data_deudas.Recordset("nro_superv") = 30
            data_deudas.Recordset("documento") = Val(labnrofact.Caption)
            data_deudas.Recordset("tipocta") = data_temp.Recordset("moneda")
            data_deudas.Recordset("importe") = data_temp.Recordset("tot_lin")
            data_deudas.Recordset("moneda") = 1
            data_deudas.Recordset("origen") = "E-TICKET NRO." & data_temp.Recordset("moneda") & " " & labnrofact.Caption
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
            data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_temp.Recordset("cod_cli")
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_cli.Recordset.Edit
               If IsNull(data_cli.Recordset("saldo_cc")) = False Then
                  data_cli.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc") + data_temp.Recordset("tot_lin")
               Else
                  data_cli.Recordset("saldo_cc") = data_temp.Recordset("tot_lin")
               End If
               data_cli.Recordset.Update
            End If
         End If
         data_temp.Recordset.MoveNext
      Loop
   End If
   data_ctrlf.Recordset.Edit
   data_ctrlf.Recordset("fecha") = Date
   data_ctrlf.Recordset.Update
'   frm_histdesp.Show vbModal
      
   data_llamado.RecordSource = "Select * from llamado where nrolla =" & labnrollamado.Caption
   data_llamado.Refresh
   If data_llamado.Recordset.RecordCount > 0 Then
      If data_llamado.Recordset("pend") <> 2 Then
         data_llamado.Recordset.Edit
         data_llamado.Recordset("pend") = 2
         data_llamado.Recordset("totend") = "FACT"
         data_llamado.Recordset.Update
      End If
   End If
   
'   MsgBox "Proceso de facturación terminado con éxito.", vbInformation
'   Unload Me
Else
   MsgBox "No hay líneas de facturación"
End If
'terminado

End Sub

Private Sub Form_Load()
Dim Xmat, Xced, Xcodced As Long
Dim Xmasdiezui, Ximpdomi, Ximptrasla, XImp, Ximpsuma As Double
Dim XsocioOK, Xtrasla, XX, Xdiasfact As Integer
Dim Xclave, Xconvcod, Xclavefin As String
Dim Xfec1, Xfec2 As Date

Dim Xlaui As Double

XsocioOK = 0
Ximpdomi = 0
Ximptrasla = 0

data_ui.Connect = "ODBC;DSN=sapp;"
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

data_ctrlf.DatabaseName = App.Path & "\ctrf.mdb"
data_ctrlf.RecordSource = "ctrf"
data_ctrlf.Refresh

data_conslla.DatabaseName = App.Path & "\sapp.mdb"

Xfec1 = data_ctrlf.Recordset("fecha")
Xfec2 = Date
Xdiasfact = DateDiff("d", Xfec1, Xfec2)

data_temp.DatabaseName = App.Path & "\factura.mdb"

If frm_largador.txt_nro.Text <> "" Then
   labnrollamado.Caption = frm_largador.txt_nro.Text
Else
   labnrollamado.Caption = 0
End If
data_conve.Connect = "odbc;dsn=sapp;"

data_conslla.RecordSource = "Select * from llamado where nrolla =" & labnrollamado.Caption & " and totend in ('FACT')"
data_conslla.Refresh
If data_conslla.Recordset.RecordCount > 0 Then
Else
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
    
    mf.Text = Format(Date, "dd/mm/yyyy")
    labhora.Caption = Format(Time, "HH:mm")
    
    data_verfac.Connect = "ODBC;DSN=sapp;"
    
    data_verfac.Connect = "ODBC;DSN=sapp;"
    
    data_cabezal.Connect = "ODBC;DSN=sapp;"
    data_cabezal.RecordSource = "clirespl"
    data_cabezal.Refresh
    
    data_deudas.Connect = "ODBC;DSN=sapp;"
    
    data_lin2.Connect = "ODBC;DSN=sapp;"
    
    data_caja.Connect = "ODBC;DSN=sapp;"
    data_caja.RecordSource = "caja"
    data_caja.Refresh
    
    data_estudios.Connect = "ODBC;DSN=sapp;"
    
    data_arancel.Connect = "ODBC;DSN=sapp;"
    
    data_lin3.Connect = "ODBC;DSN=sapp;"
    
    data_lin.Connect = "ODBC;DSN=sapp;"
    
    data_par.Connect = "ODBC;DSN=sappfact;"
    data_par.RecordSource = "paramsapp"
    data_par.Refresh
    
    t_pie.Text = "Caja de Usuario: " & WElusuario & " FECHA:" & Format(Date, "dd/mm/yyyy")
    t_mat.Text = 0
    Xtrasla = 0
    Xtrasla = frm_largador.cbotras.ListIndex
    labusuario.Caption = WElusuario
    data_cli.Connect = "ODBC;DSN=sapp;"
    If frm_largador.txt_mat.Text <> "" Then
       Xmat = frm_largador.txt_mat.Text
    Else
       Xmat = 0
    End If
    If frm_largador.txt_ced.Text <> "" Then
       Xced = frm_largador.txt_ced.Text
       Xcodced = frm_largador.t_codced.Text
    Else
       Xced = 0
       Xcodced = 0
    End If
    Xclavefin = ""
    If frm_largador.cbocolor.Text = "VERDE" Or frm_largador.cbocolor.Text = "CELESTE" Then
       Xclave = "V"
    Else
       If frm_largador.cbocolor.Text = "AMARILLO" Then
          Xclave = "A"
       Else
          If frm_largador.cbocolor.Text = "ROJO" Then
             Xclave = "R"
          Else
             If frm_largador.cbocolor.Text = "AZUL" Then
                Xclave = "Z"
             Else
                If frm_largador.cbocolor.Text = "NEGRO" Then
                   Xclave = "N"
                Else
                   Xclave = "V"
                End If
             End If
          End If
       End If
    End If
    If frm_largador.cbocolfin.Text = "VERDE" Then
       Xclavefin = "V"
    Else
       If frm_largador.cbocolfin.Text = "AMARILLO" Then
          Xclavefin = "A"
       Else
          If frm_largador.cbocolfin.Text = "ROJO" Then
             Xclavefin = "R"
          Else
             If frm_largador.cbocolfin.Text = "AZUL" Then
                Xclavefin = "V"
             Else
                If frm_largador.cbocolfin.Text = "NEGRO" Then
                   Xclavefin = "R"
                Else
                   Xclavefin = "V"
                End If
             End If
          End If
       End If
    End If
    
    If frm_largador.txt_cat.Text <> "" Then
       Xconvcod = frm_largador.txt_cat.Text
    Else
       Xconvcod = "PART"
    End If
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
          If Xclavefin = "R" Then
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
    
    If Xclave = "V" Then
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
    
    If Xtrasla <= 0 Then
       List1.AddItem "SERVICIOS DOMICILIARIOS"
       labcodsrv.Caption = 10017
    Else
    '   List1.AddItem "SERVICIOS DOMICILIARIOS"
       List1.AddItem "SERVICIOS DE TRASLADO"
       labcodsrv.Caption = 90021
    End If
    'labcodsrv.Caption = 10017
    List1.ListIndex = 0
    cbofpago.ListIndex = 0
    
    If Xced > 0 Then
        If Xdiasfact <= 5 Then
            If Xmat <> 0 Then
               data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
               data_cli.Refresh
               If data_cli.Recordset.RecordCount > 0 Then
                  t_nombre.Text = data_cli.Recordset("cl_apellid")
                  t_mat.Text = data_cli.Recordset("cl_codigo")
                  t_codcnv.Text = Xconvcod
                  t_nomcnv.Text = data_cli.Recordset("cl_nomconv")
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
                  MsgBox "Hay un error en los datos de la matrícula del socio, verifique para poder facturar el servicio", vbExclamation
                  XsocioOK = 8
               End If
            Else
               If Xced <> 0 Then
                  data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xced
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     t_nombre.Text = data_cli.Recordset("cl_apellid")
                     t_mat.Text = data_cli.Recordset("cl_codigo")
                     t_codcnv.Text = Xconvcod
                     t_nomcnv.Text = data_cli.Recordset("cl_nomconv")
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
                     Xtipodedocumento = 3
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
                     data_cli.Recordset("cl_apellid") = frm_largador.txt_nomb.Text
                     data_cli.Recordset("cl_codconv") = frm_largador.txt_cat.Text
                     data_cli.Recordset("cl_nomconv") = frm_largador.txt_nomcat.Text
                     data_cli.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
                     If frm_largador.Combo3.ListIndex = 0 Then
                        data_cli.Recordset("cl_sexo") = 1
                     Else
                        data_cli.Recordset("cl_sexo") = 2
                     End If
                     data_cli.Recordset("estado") = 1
                     data_cli.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
                     If frm_largador.txt_tel.Text <> "" Then
                        data_cli.Recordset("cl_telefon") = frm_largador.txt_tel.Text
                     End If
                     If frm_largador.txt_locali.Text <> "" Then
                        data_cli.Recordset("cl_direcci") = frm_largador.txt_locali.Text
                     End If
                     data_cli.Recordset.Update
                     t_nombre.Text = frm_largador.txt_nomb.Text
                     t_codcnv.Text = frm_largador.txt_cat.Text
                     t_nomcnv.Text = frm_largador.txt_nomcat.Text
                     t_ced.Text = Xced
                     t_codced.Text = Xcodced
                     Xtipodedocumento = 3
                  End If
               Else
                  MsgBox "No registró la cédula, debe ingresar documento para poder facturar servicio", vbInformation
                  XsocioOK = 8
               End If
            End If
            '       If lperno.List(lperno.ListIndex) = lper.List(lper.ListIndex) Then
            
            If XsocioOK = 8 Then
               MsgBox "Hay un error en los datos del paciente, verifique para poder facturar"
            Else
               If Xtrasla = 1 Or Xtrasla = 2 Or Xtrasla = 14 Or Xtrasla = 10 Or _
                  Xtrasla = 11 Or Xtrasla = 9 Or Xtrasla = 15 Then
                  If frm_largador.txt_cat.Text = "MSP" Then
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
                  labdescest.Caption = data_estudios.Recordset("descrip")
                  data_arancel.RecordSource = "Select * from arancel where ara_cnvcod ='" & Xconvcod & "' and ara_famnro =" & data_estudios.Recordset("codest")
                  data_arancel.Refresh
                  If data_arancel.Recordset.RecordCount > 0 Then
                     If IsNull(data_arancel.Recordset("ara_precio")) = False Then
                        If data_arancel.Recordset("ara_precio") > 0 Then
                           XImp = data_arancel.Recordset("ara_precio")
                        Else
                           If IsNull(data_arancel.Recordset("ara_porcen")) = False Then
                              XImp = data_estudios.Recordset("ucfh") - (data_estudios.Recordset("ucfh") * (data_arancel.Recordset("ara_porcen") / 100))
    '                          XImp = data_estudios.Recordset("cons") - (data_estudios.Recordset("cons") * (data_arancel.Recordset("ara_porcen") / 100))
                           Else
                              XImp = 0
                           End If
                        End If
                     Else
                        XImp = 0
                     End If
                  Else
                     If IsNull(data_estudios.Recordset("ucfh")) = False Then
                        XImp = data_estudios.Recordset("ucfh")
                     Else
                        XImp = 0
                     End If
                  End If
               Else
                  XImp = 0
               End If
               t_imp.Text = Format(XImp, "Standard")
               t_total.Text = t_imp.Text
               t_iva.Text = t_imp.Text * 0.1 / 1.1
               t_iva.Text = Format(t_iva.Text, "Standard")
               If frm_largador.chtmut.value = 1 Then
                  t_imp.Text = 0
                  t_imp.Text = Format(t_imp.Text, "Standard")
                  t_total.Text = t_imp.Text
                  t_iva.Text = 0
                  t_iva.Text = Format(t_iva.Text, "Standard")
                  Ximpdomi = 0
                  Ximptrasla = 0
               End If
    'hasta aqui
    '''"CERSEM" Or Xconvcod = "CERCAS"
               If Xconvcod <> "PART" Then
                  data_conve.RecordSource = "Select * from convenio where cnv_codigo ='" & Xconvcod & "'"
                  data_conve.Refresh
                  If data_conve.Recordset.RecordCount > 0 Then
                     If IsNull(data_conve.Recordset("cnv_colrec")) = False Then
                        If Xconvcod = "SA" Or Xconvcod = "SAF" Or Xconvcod = "CCOMS" Or Xconvcod = "CERCAS" Or Xconvcod = "SEGAM" Or _
                           Xconvcod = "EMERN" Or Xconvcod = "EMERG" Or Xconvcod = "EMERNE" Or Xconvcod = "CAAM" Or _
                           Xconvcod = "SAP" Or Xconvcod = "VIV19" Or Xconvcod = "VIV20" Or Xconvcod = "CAAMEP" Or _
                           Xconvcod = "MSP" Or Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "APNORE" Or _
                           Xconvcod = "SJ01" Or data_conve.Recordset("cnv_colrec") = "M" Or data_conve.Recordset("cnv_colrec") = "R" Or _
                           Mid(Xconvcod, 1, 4) = "TALA" Or frm_largador.cbozona.Text = 3 Then
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
                           Xconvcod = "EMERN" Or Xconvcod = "EMERG" Or Xconvcod = "EMERNE" Or Xconvcod = "CAAM" Or _
                           Xconvcod = "SAP" Or Xconvcod = "VIV19" Or Xconvcod = "VIV20" Or Xconvcod = "CAAMEP" Or _
                           Xconvcod = "MSP" Or Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "APNORE" Or _
                           Xconvcod = "SJ01" Or Mid(Xconvcod, 1, 4) = "TALA" Or frm_largador.cbozona.Text = 3 Then
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
                        Xconvcod = "SJ01" Or Mid(Xconvcod, 1, 4) = "TALA" Or frm_largador.cbozona.Text = 3 Then
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
                  Xconvcod = "MUCAMS" Or Xconvcod = "MUCAMT" Or Xconvcod = "MUCATA" Or Xconvcod = "SOLEME" Or _
                  Xconvcod = "CAAMEP" Or Xconvcod = "SOLAF" Or Xconvcod = "SOLAMB" Or Xconvcod = "SOC" Then
                  t_imp.Text = 0
                  t_imp.Text = Format(t_imp.Text, "Standard")
                  t_total.Text = t_imp.Text
                  t_iva.Text = 0
                  t_iva.Text = Format(t_iva.Text, "Standard")
                  Ximpdomi = 0
                  Ximptrasla = 0
               End If
               If frm_largador.txt_costo.Text <> "" Then
                  If frm_largador.txt_costo.Text > 0 Then
                     t_imp.Text = Format(frm_largador.txt_costo.Text, "Standard")
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
                  t_imp.Text = 0
                  t_imp.Text = Format(t_imp.Text, "Standard")
                  t_total.Text = t_imp.Text
                  t_iva.Text = 0
                  t_iva.Text = Format(t_iva.Text, "Standard")
                  Ximpdomi = 0
                  Ximptrasla = 0
               End If
               If data_ui.Recordset.RecordCount > 0 Then
                  Xlaui = CDbl(data_ui.Recordset("descrip"))
                  Xmasdiezui = CDbl(t_total.Text) / Xlaui
               End If
               If Label7.Caption = "E-TICKET" Then
                  labnrofact.Caption = ""
                  labserie.Caption = ""
               Else
                  If Label7.Caption = "E-FACTURA" Then
                     labnrofact.Caption = ""
                     labserie.Caption = ""
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
                  If t_imp.Text <> "" Then
                     If t_imp.Text > 0 Then
                        If Xmasdiezui > 10000 Then
                           If Xtipodedocumento = 4 Then
                              If t_ced.Text <> "" Then
                                 Xtipodedocumento = 3
                                 Command1_Click
                              Else
                                 MsgBox "Atención! no se pudo facturar. Faltan datos en el documento. Comunique a administración!"
                                 Unload Me
                              End If
                           Else
                              If t_ced.Text <> "" Then
                                 Xtipodedocumento = 3
                                 Command1_Click
                              Else
                                 MsgBox "Atención! no se pudo facturar. Faltan datos en el documento. Comunique a administración!"
                                 Unload Me
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
            MsgBox "Verifique FECHA del sistema, no se puede facturar. COMUNIQUE A ADMINISTRACION!!"
        End If
    Else
        If Xconvcod = "911" Or Xconvcod = "911B" Then
           data_cli.RecordSource = "clientes"
           data_cli.Refresh
           data_cli.Recordset.AddNew
           data_cli.Recordset("cl_codigo") = data_par.Recordset("nro_socio") + 1
           Xmat = data_par.Recordset("nro_socio") + 1
           data_par.Recordset.Edit
           data_par.Recordset("nro_socio") = data_par.Recordset("nro_socio") + 1
           data_par.Recordset.Update
           data_cli.Recordset("cl_cedula") = 0
           data_cli.Recordset("cl_codced") = 0
           data_cli.Recordset("cl_apellid") = frm_largador.txt_nomb.Text
           data_cli.Recordset("cl_codconv") = frm_largador.txt_cat.Text
           data_cli.Recordset("cl_nomconv") = frm_largador.txt_nomcat.Text
           data_cli.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
           If frm_largador.Combo3.ListIndex = 0 Then
              data_cli.Recordset("cl_sexo") = 1
           Else
              data_cli.Recordset("cl_sexo") = 2
           End If
           data_cli.Recordset("estado") = 1
           data_cli.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
           If frm_largador.txt_tel.Text <> "" Then
              data_cli.Recordset("cl_telefon") = frm_largador.txt_tel.Text
           End If
           If frm_largador.txt_locali.Text <> "" Then
              data_cli.Recordset("cl_direcci") = frm_largador.txt_locali.Text
           End If
           data_cli.Recordset.Update
           t_nombre.Text = frm_largador.txt_nomb.Text
           t_codcnv.Text = frm_largador.txt_cat.Text
           t_nomcnv.Text = frm_largador.txt_nomcat.Text
           t_ced.Text = 0
           t_codced.Text = 0
           Xtipodedocumento = 4
           t_imp.Text = 0
           t_imp.Text = Format(t_imp.Text, "Standard")
           t_total.Text = t_imp.Text
           t_iva.Text = 0
           t_iva.Text = Format(t_iva.Text, "Standard")
           Ximpdomi = 0
           Ximptrasla = 0
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
           b_findos_Click
        End If
    End If
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
        MsgBox "El CFE no trae número de folio, no se puede terminar la factura"
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.DatosCae Is Nothing Then
        MsgBox "El CFE no trae datos del CAE, no se puede terminar la factura"
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
            Exit Sub
        End If

        If Not objresultado.OperacionExitosa Then
            If objresultado.Mensaje <> vbNullString Then strMensaje = strMensaje & ": " & objresultado.Mensaje
            MsgBox strMensaje
            Exit Sub
        End If

'        imgQr.Picture = LoadPicture(strFile)
    End If
    labserie.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.serie
    labnrofact.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
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
End Sub

