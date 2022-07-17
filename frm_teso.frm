VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_teso 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja de tesorería"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   Icon            =   "frm_teso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7080
      TabIndex        =   50
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc data_cl 
      Height          =   375
      Left            =   4320
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_cl"
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
   Begin MSAdodcLib.Adodc data_cajtes 
      Height          =   375
      Left            =   2520
      Top             =   6120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_cajtes"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8520
      TabIndex        =   49
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b_cam 
      Caption         =   "Cambiar Usuario"
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
      Left            =   6720
      TabIndex        =   44
      Top             =   120
      Width           =   1935
   End
   Begin VB.Data data_val 
      Caption         =   "data_val"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4800
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton bctrol 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Controlar Saldos..."
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
      Left            =   6720
      MouseIcon       =   "frm_teso.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton bant 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Anterior <---"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton bsig 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Siguiente --->"
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
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Data data_rub 
      Caption         =   "data_rub"
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
      RecordSource    =   "rubteso"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton bimp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_teso.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Imprimir Caja"
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_teso.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Buscar datos"
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      Picture         =   "frm_teso.frx":1260
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Cancelar la acción"
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton bborra 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      Picture         =   "frm_teso.frx":17EA
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Eliminar el registro actual"
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "frm_teso.frx":1D74
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Modificar registro actual"
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_teso.frx":22FE
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Grabar datos"
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton balta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_teso.frx":2888
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Dar alta un nuevo registro"
      Top             =   5760
      Width           =   495
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
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   6
      Top             =   720
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
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de caja"
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
      ForeColor       =   &H00FFFFC0&
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8655
      Begin MSMask.MaskEdBox mfvencec 
         Height          =   375
         Left            =   6480
         TabIndex        =   52
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   255
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
      Begin VB.Data data_part 
         Caption         =   "data_part"
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
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton b_buscl 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Buscar..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         Picture         =   "frm_teso.frx":2E12
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txt_tcam 
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
         Left            =   6240
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_impiva 
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
         Left            =   6360
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox cboiva 
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
         ItemData        =   "frm_teso.frx":339C
         Left            =   5160
         List            =   "frm_teso.frx":33A9
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_con 
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
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txt_haber 
         Alignment       =   1  'Right Justify
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
         Left            =   6600
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txt_debe 
         Alignment       =   1  'Right Justify
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
         Left            =   6600
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin MSDBCtls.DBCombo dbcborub 
         Bindings        =   "frm_teso.frx":33BB
         DataSource      =   "data_rub"
         Height          =   660
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1164
         _Version        =   393216
         Style           =   1
         ListField       =   ""
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
      Begin VB.TextBox txt_codrub 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cbomon 
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
         ItemData        =   "frm_teso.frx":33D2
         Left            =   1920
         List            =   "frm_teso.frx":33DC
         TabIndex        =   16
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txt_obs 
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   22
         Top             =   3480
         Width           =   6015
      End
      Begin VB.TextBox txt_imp 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txt_imp2 
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
         Left            =   1920
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox t_cli 
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
         Left            =   1920
         TabIndex        =   46
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Vencimiento de Cheque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   51
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label labnomcl 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1920
         TabIndex        =   48
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "COMERCIO:"
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
         TabIndex        =   45
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         Caption         =   "T.Cambio:"
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
         Left            =   4920
         TabIndex        =   43
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         Caption         =   "U$s."
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
         Left            =   720
         TabIndex        =   42
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IVA:"
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
         Left            =   3960
         TabIndex        =   41
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         Caption         =   "S=Salida"
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
         Left            =   7320
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "E=Entrada"
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
         Left            =   7320
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label labsaldop 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3960
         TabIndex        =   27
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "SALDO $."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   26
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   8640
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "OBSERVACIONES"
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
         TabIndex        =   25
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CONCEPTO:"
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
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IMPORTE:"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "MONEDA:"
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
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "AL HABER:"
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
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "AL DEBE:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CÓDIGO:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox txt_base 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "HORA:"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FECHA....:"
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
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Usuario:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   5880
      Picture         =   "frm_teso.frx":3405
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_teso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_buscl_Click()
Xestaok = 7
frm_labo.Show vbModal

End Sub

Private Sub b_cam_Click()
'Dim Xcambiau, Xcambiac As String
'Dim Xlafetes2 As Date
'Xlafetes2 = CDate("01/01/2010")
'Xcambiau = InputBox("Ingrese usuario:", "Nombre de usuario")
'Xcambiac = InputBox("Ingrese contraseña:", "Contraseña de usuario")
'If Xcambiau <> "" Then
'   If Xcambiac <> "" Then
'      frm_usuario.data_usuario.Recordset.FindFirst "usuario = '" & Xcambiau & "'"
'      If Not frm_usuario.data_usuario.Recordset.NoMatch Then
'         If frm_usuario.data_usuario.Recordset("clave") = UCase(Xcambiac) Then
'            WElusuario = UCase(Trim(Xcambiau))
'            WNombase = UCase(WElusuario)
'            data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' order by nromov"
'            data_cajtes.Refresh
'            If data_cajtes.Recordset.RecordCount > 0 Then
'               data_cajtes.Recordset.MoveLast
''               igualacaj
'            End If
'         Else
'            MsgBox "Clave incorrecta", vbInformation, "Seguridad"
'         End If
'      Else
'         MsgBox "Usuario no registrado", vbInformation, "Seguridad"
'      End If
'   End If
'End If
MsgBox "OPCION NO HABILITADA DESDE LA EDICION!"

End Sub

Private Sub balta_Click()
Dim Xlafetes3 As Date
Xlafetes3 = CDate("01/01/2010")
Frame1.Enabled = True
limpiar
txt_hora.Enabled = True
mfecha.Enabled = True
mfecha.Text = Format(Date, "dd/mm/yyyy")
txt_hora.Text = Format(Time, "HH:mm")
txt_base.Text = WNombase
txt_hora.Enabled = False
mfecha.Enabled = False
txt_codrub.SetFocus
'''mfecha.Text = "__/__/____"
'''mfecha.SetFocus
'''data_cajtes.RecordSource = "Select * from tesorero order by nromov DESC"
' data_cajtes.RecordSource = "Select * from tesorero where fecha >=#" &  order by nromov"
'''data_cajtes.Refresh
'data_cajtes.Recordset.MoveLast
Text1.Text = data_part.Recordset("nrocaja") + 1
data_part.Recordset.Edit
data_part.Recordset("nrocaja") = data_part.Recordset("nrocaja") + 1
data_part.Recordset.Update
data_part.Refresh
data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' order by nromov DESC"
'data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' And fecha >=#" & Format(Xlafetes3, "yyyy/mm/dd") & "# order by nromov"
data_cajtes.Refresh
If data_cajtes.Recordset.RecordCount > 0 Then
'   data_cajtes.Recordset.MoveLast
   Wsaldos = data_cajtes.Recordset("saldos")
Else
   Wsaldos = 0
End If
data_cajtes.Recordset.AddNew
labsaldop.Caption = Format(Wsaldos, "Standard")
XAcnv = 1
balta.Enabled = False
bmodif.Enabled = False
bgraba.Enabled = True
bcance.Enabled = True
bbusca.Enabled = False
bborra.Enabled = False
bimp.Enabled = False
bctrol.Enabled = False
bsig.Enabled = False
bant.Enabled = False
b_cam.Enabled = False


End Sub

Private Sub bant_Click()
If data_cajtes.Recordset.BOF = True Then
   MsgBox "Primer registro", vbInformation, "Mensaje"
Else
   data_cajtes.Recordset.MovePrevious
   If data_cajtes.Recordset.BOF = True Then
      MsgBox "Primer registro", vbInformation, "Mensaje"
   Else
      igualacaj
   End If
End If

End Sub

Private Sub bborra_Click()
If data_cajtes.Recordset.EOF = False Then
   If data_cajtes.Recordset("fecha") = Date Then
        Dim Xress As String
        Dim Xlafetes4 As Date
        Xlafetes4 = CDate("01/01/2010")
        Xress = MsgBox("Desea eliminar el registro seleccionado?", vbYesNo + vbInformation, "Mensaje")
        If Xress = vbYes Then
            data_cajtes.Recordset.Delete
            data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' order by nromov"
        '    data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' And fecha >=#" & Format(Xlafetes4, "yyyy/mm/dd") & "# order by nromov"
            data_cajtes.Refresh
            bctrol_Click
            data_cajtes.Recordset.MoveLast
            igualacaj
        '    MsgBox "Recuerde CONTROLAR SALDOS luego de BORRAR", vbCritical, "Mensaje"
        End If
   Else
      MsgBox "No es un registro del día", vbCritical
   End If
Else
   MsgBox "Atención!!!: último registro, regrese al anterior", vbCritical
End If

End Sub

Private Sub bbusca_Click()
frm_busctes.Show vbModal

End Sub

Private Sub bcance_Click()
If XAcnv = 1 Then
   data_cajtes.Recordset.CancelUpdate
   If data_cajtes.Recordset.EOF = False Then
      data_cajtes.Recordset.MoveLast
   End If
   balta.Enabled = True
   bmodif.Enabled = True
   bgraba.Enabled = False
   bcance.Enabled = False
   bbusca.Enabled = True
   bborra.Enabled = True
   bimp.Enabled = True
   bctrol.Enabled = True
   bsig.Enabled = True
   bant.Enabled = True
   b_cam.Enabled = True
   txt_codrub.Enabled = True
   dbcborub.Enabled = True
   igualacaj
   Frame1.Enabled = False
   XAcnv = 0
   oculta
Else
'   data_cajtes.Recordset.MoveLast
'   data_cajtes.Recordset.MoveFirst
   balta.Enabled = True
   bmodif.Enabled = True
   bgraba.Enabled = False
   bcance.Enabled = False
   bbusca.Enabled = True
   bborra.Enabled = True
   bimp.Enabled = True
   bctrol.Enabled = True
   bsig.Enabled = True
   bant.Enabled = True
   b_cam.Enabled = True
   txt_codrub.Enabled = True
   dbcborub.Enabled = True
   igualacaj
   Frame1.Enabled = False
   XAcnv = 0
   oculta
End If
txt_hora.Enabled = False
mfecha.Enabled = False

End Sub

Private Sub bctrol_Click()
Dim Quesaldo As Double
On Error GoTo Quepasatesor

frm_teso.MousePointer = 11
labsaldop.Caption = 0
Quesaldo = 0
If UCase(WElusuario) = "RREGUEIRA" Then
   Quesaldo = 0
Else
   If UCase(WElusuario) = "MCOSTA" Then
      Quesaldo = 2535701.45
   Else
      If UCase(WElusuario) = "MPEREZ" Then
         Quesaldo = 3.48
      Else
         If UCase(WElusuario) = "RMBROU" Then
            Quesaldo = -186730611.25
         Else
            If UCase(WElusuario) = "RMSANT" Then
               Quesaldo = -35354862
            Else
               If UCase(WElusuario) = "ROXANA" Then
                  Quesaldo = 1763.5
               End If
            End If
         End If
      End If
   End If
End If
data_cajtes.Refresh

data_cajtes.Recordset.MoveFirst
'data_cajtes.Recordset.MoveLast
'''data_cajtes.Recordset.MovePrevious

'''Quesaldo = data_cajtes.Recordset("saldos")
'''data_cajtes.Recordset.MoveNext

Do While Not data_cajtes.Recordset.EOF
   If data_cajtes.Recordset("concep") = "E" Then
      Quesaldo = Quesaldo + data_cajtes.Recordset("monto")
   Else
      Quesaldo = Quesaldo - data_cajtes.Recordset("monto")
   End If
   If IsNull(data_cajtes.Recordset("saldos")) = True Then
'      data_cajtes.Recordset.Edit
      data_cajtes.Recordset("saldos") = Quesaldo
      data_cajtes.Recordset.Update
   Else
      If Format(data_cajtes.Recordset("saldos"), "Standard") = Format(Quesaldo, "Standard") Then
'      If data_cajtes.Recordset("saldos") = Quesaldo Then
      Else
'         data_cajtes.Recordset.Edit
         data_cajtes.Recordset("saldos") = Quesaldo
         data_cajtes.Recordset.Update
      End If
   End If
   labsaldop.Caption = Format(data_cajtes.Recordset("saldos"), "Standard")
   data_cajtes.Recordset.MoveNext
Loop
frm_teso.MousePointer = 0

Exit Sub

Quepasatesor:
             If Err.Number = 3155 Then
                MsgBox "Error al grabar, verifique"
             Else
                MsgBox "No se pudo terminar de generar el saldo ERROR: " & Err.Number & " " & Err.Description
             End If
'igualacaj

End Sub

Private Sub bgraba_Click()
If XAcnv = 1 Then
   If IsDate(mfecha.Text) = True Then
      If txt_codrub.Text <> "" Then
         If txt_codrub.Text <> 0 Then
            If Len(txt_obs.Text) > 5 Then
                If txt_con.Text = "E" Then
                   Wsaldos = Wsaldos + txt_imp.Text
                Else
                   If txt_con.Text = "S" Then
                      Wsaldos = Wsaldos - txt_imp.Text
                   Else
                      Wsaldos = Wsaldos + 0
                   End If
                End If
                Fungraba
                data_cajtes.Recordset("saldos") = Wsaldos
                data_cajtes.Recordset("nromov") = Text1.Text
                data_cajtes.Recordset.Update
                XAcnv = 0
                balta.Enabled = True
                bmodif.Enabled = True
                bgraba.Enabled = False
                bcance.Enabled = False
                bbusca.Enabled = True
                bborra.Enabled = True
                bimp.Enabled = True
                bctrol.Enabled = True
                bsig.Enabled = True
                bant.Enabled = True
                b_cam.Enabled = True
                Frame1.Enabled = False
                txt_hora.Enabled = False
                mfecha.Enabled = False
'                bctrol_Click
'                labsaldop.Caption = Format(data_cajtes.Recordset("saldos"), "Standard")
            Else
                MsgBox "Debe ingresar en observación más de 5 letras", vbInformation, "Mensaje"
                txt_obs.SetFocus
            End If
         Else
            MsgBox "VERIFIQUE el RUBRO", vbCritical, "Mensaje"
            txt_codrub.SetFocus
         End If
      Else
         MsgBox "VERIFIQUE el RUBRO", vbCritical, "Mensaje"
         txt_codrub.SetFocus
      End If
   Else
      MsgBox "VERIFIQUE LA FECHA", vbCritical, "Mensaje"
      mfecha.SetFocus
   End If
Else
   If IsDate(mfecha.Text) = True Then
      If txt_codrub.Text <> "" Then
         If txt_codrub.Text <> 0 Then
            If Len(txt_obs.Text) > 5 Then
                If txt_con.Text = "E" Then
                   Wsaldos = Wsaldos + txt_imp.Text
                Else
                   If txt_con.Text = "S" Then
                      Wsaldos = Wsaldos - txt_imp.Text
                   Else
                      Wsaldos = Wsaldos + 0
                   End If
                End If
'                data_cajtes.Recordset.Edit
                data_cajtes.Recordset("saldos") = Wsaldos
                Fungraba
                data_cajtes.Recordset.Update
                data_cajtes.Recordset.MoveFirst
                XAcnv = 0
                balta.Enabled = True
                bmodif.Enabled = True
                bgraba.Enabled = False
                bcance.Enabled = False
                bbusca.Enabled = True
                bborra.Enabled = True
                bimp.Enabled = True
                bctrol.Enabled = True
                bsig.Enabled = True
                bant.Enabled = True
                b_cam.Enabled = True
                txt_codrub.Enabled = True
                dbcborub.Enabled = True
                Frame1.Enabled = False
                txt_hora.Enabled = False
                mfecha.Enabled = False
'                bctrol_Click
            Else
                MsgBox "Debe ingresar en observación más de 5 letras", vbInformation, "Mensaje"
                txt_obs.SetFocus
            End If
         Else
            MsgBox "VERIFIQUE el RUBRO", vbCritical, "Mensaje"
            txt_codrub.SetFocus
         End If
      Else
         MsgBox "VERIFIQUE el RUBRO", vbCritical, "Mensaje"
         txt_codrub.SetFocus
      End If
   Else
      MsgBox "VERIFIQUE LA FECHA", vbCritical, "Mensaje"
      mfecha.SetFocus
   End If
End If
End Sub

Private Sub bimp_Click()
frm_impteso.Show vbModal

End Sub

Private Sub bmodif_Click()

If data_cajtes.Recordset.EOF = False Then
   If data_cajtes.Recordset("fecha") = Date Then
     If data_cajtes.Recordset.EOF = False Then
        balta.Enabled = False
        bmodif.Enabled = False
        bgraba.Enabled = True
        bcance.Enabled = True
        bbusca.Enabled = False
        bborra.Enabled = False
        bimp.Enabled = False
        bctrol.Enabled = False
        bsig.Enabled = False
        bant.Enabled = False
        b_cam.Enabled = False
        txt_hora.Enabled = True
        mfecha.Enabled = True

        Frame1.Enabled = True
        XAcnv = 0
        oculta
        igualacaj
        
        If IsNull(data_cajtes.Recordset("moneda")) = False Then
           If data_cajtes.Recordset("moneda") = 2 Then
              txt_tcam.Visible = True
              txt_imp2.Visible = True
              Label13.Visible = True
              Label14.Visible = True
           Else
              txt_tcam.Visible = False
              txt_imp2.Visible = False
              Label13.Visible = False
              Label14.Visible = False
           End If
        Else
           txt_tcam.Visible = False
           txt_imp2.Visible = False
           Label13.Visible = False
           Label14.Visible = False
        End If
        txt_hora.Enabled = False
        mfecha.Enabled = False
        If txt_imp.Visible = True Then
           If txt_imp.Enabled = True Then
              txt_imp.SetFocus
           Else
              If txt_imp2.Visible = True Then
                 If txt_imp2.Enabled = True Then
                    txt_imp2.SetFocus
                 End If
              End If
           End If
        End If
     Else
        MsgBox "Atención!!!: último registro, regrese al anterior", vbCritical, "Mensaje"
   End If
 Else
    MsgBox "No es un registro del día", vbCritical, "Tesorería"
 End If
Else
    MsgBox "Atención!!!: último registro, regrese al anterior", vbCritical
End If

End Sub

Private Sub bsig_Click()
If data_cajtes.Recordset.EOF = True Then
   MsgBox "Ultimo registro", vbInformation, "Mensaje"
Else
   data_cajtes.Recordset.MoveNext
   If data_cajtes.Recordset.EOF = True Then
      MsgBox "Ultimo registro", vbInformation, "Mensaje"
   Else
      igualacaj
   End If
End If

End Sub

Private Sub cboiva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_obs.Enabled = True Then
      txt_obs.SetFocus
   End If
End If

End Sub

Private Sub cboiva_LostFocus()
Dim Xvaliva As Double
If txt_imp.Text <> "" Then
    If cboiva.ListIndex = 0 Then
       txt_impiva.Text = Format(0, "Standard")
    Else
       If cboiva.ListIndex = 1 Then
          Xvaliva = txt_imp.Text / 1.1
          Xvaliva = Xvaliva * 0.1
          txt_impiva.Text = Format(Xvaliva, "Standard")
       Else
          If cboiva.ListIndex = 2 Then
             Xvaliva = txt_imp.Text / 1.22
             Xvaliva = Xvaliva * 0.22
             txt_impiva.Text = Format(Xvaliva, "Standard")
          Else
             txt_impiva.Text = Format(0, "Standard")
             cboiva.ListIndex = 0
          End If
       End If
    End If
End If
End Sub

Private Sub cbomon_Click()
If cbomon.Text = "DOLARES AMERICANOS" Then
   Label13.Visible = True
   Label14.Visible = True
   txt_tcam.Visible = True
   txt_imp2.Visible = True
   txt_imp.Enabled = False
Else
   Label13.Visible = False
   txt_tcam.Visible = False
   txt_imp2.Visible = False
   Label14.Visible = False
   txt_imp.Enabled = True
End If

End Sub

Private Sub cbomon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cbomon.Text = "DOLARES AMERICANOS" Then
      Label13.Visible = True
      Label14.Visible = True
      txt_tcam.Visible = True
      txt_imp2.Visible = True
      txt_imp.Enabled = False
   Else
      Label13.Visible = False
      txt_tcam.Visible = False
      txt_imp2.Visible = False
      Label14.Visible = False
      txt_imp.Enabled = True
   End If
   If txt_tcam.Visible = True Then
      txt_imp2.SetFocus
   Else
      txt_imp.SetFocus
   End If
   
End If


End Sub

Private Sub cbomon_LostFocus()
If cbomon.Text = "DOLARES AMERICANOS" Then
   Label13.Visible = True
   Label14.Visible = True
   txt_tcam.Visible = True
   txt_tcam.Text = Format(data_val.Recordset("valor"), "Standard")
   txt_imp.Enabled = False
   txt_imp2.Enabled = True
   txt_imp2.SetFocus
Else
   Label13.Visible = False
   Label14.Visible = False
   txt_tcam.Visible = False
   txt_imp.Enabled = True
   txt_imp.SetFocus
End If

End Sub

Private Sub Command1_Click()
frm_teso.MousePointer = 11
labsaldop.Caption = 0
Quesaldo = 0
If UCase(WElusuario) = "RREGUEIRA" Then
   Quesaldo = 1466269
Else
   If UCase(WElusuario) = "MCOSTA" Then
      Quesaldo = 2535701
   Else
      If UCase(WElusuario) = "MPEREZ" Then
         Quesaldo = 3
      Else
         If UCase(WElusuario) = "RMBROU" Then
            Quesaldo = -186730611
         Else
            If UCase(WElusuario) = "RMSANT" Then
               Quesaldo = -35354862
            End If
         End If
      End If
   End If
End If

'''data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' order by nromov"

'data_cajtes.Refresh
data_cajtes.Recordset.MoveFirst
'If data_cajtes.Recordset("concep") = "E" Then
'   Quesaldo = data_cajtes.Recordset("saldos") - data_cajtes.Recordset("monto")
'Else
'   Quesaldo = data_cajtes.Recordset("saldos") + data_cajtes.Recordset("monto")
'End If
Do While Not data_cajtes.Recordset.EOF
   If data_cajtes.Recordset("concep") = "E" Then
      Quesaldo = Quesaldo + data_cajtes.Recordset("monto")
   Else
      Quesaldo = Quesaldo - data_cajtes.Recordset("monto")
   End If
   If IsNull(data_cajtes.Recordset("saldos")) = True Then
'      data_cajtes.Recordset.Edit
      data_cajtes.Recordset("saldos") = Quesaldo
      data_cajtes.Recordset.Update
   Else
      If Format(data_cajtes.Recordset("saldos"), "Standard") = Format(Quesaldo, "Standard") Then
'      If data_cajtes.Recordset("saldos") = Quesaldo Then
      Else
'         data_cajtes.Recordset.Edit
         data_cajtes.Recordset("saldos") = Quesaldo
         data_cajtes.Recordset.Update
      End If
   End If
   labsaldop.Caption = Format(data_cajtes.Recordset("saldos"), "Standard")
   data_cajtes.Recordset.MoveNext
Loop
frm_teso.MousePointer = 0

'igualacaj

End Sub

Private Sub Command2_Click()
Dim Saldin As Double
Dim Nro As Double
Nro = 40389
Saldin = 111
data_cajtes.RecordSource = "Select * from tesorero where fecha >='" & Format("07/05/2018", "yyyy-mm-dd") & "' and usuario ='" & "FOSORIO" & "' and nromov >=" & 389091 & " and nromov <=" & 389122 & " order by nromov"
data_cajtes.Refresh
If data_cajtes.Recordset.RecordCount > 0 Then
   data_cajtes.Recordset.MoveFirst
   Do While Not data_cajtes.Recordset.EOF
      data_cajtes.Recordset("nromov") = Nro
      data_cajtes.Recordset.Update
      Nro = Nro + 1
      data_cajtes.Recordset.MoveNext
   Loop
End If
MsgBox "Terminado"


End Sub

Private Sub dbcborub_DblClick(Area As Integer)

    If txt_codrub.Text <> "" Then
    Else
       dbcborub.ListField = "nombre"
       dbcborub.BoundColumn = "nombre"
       data_rub.Recordset.FindFirst "nombre ='" & dbcborub.Text & "'"
       If Not data_rub.Recordset.NoMatch Then
          dbcborub.Text = data_rub.Recordset("nombre")
          txt_codrub.Text = data_rub.Recordset("codigo")
          dbcborub.Height = 500
          dbcborub.ListField = ""
          dbcborub.BoundColumn = ""
          If WNombase = "BANCOITAU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111018
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111018
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "CHEQUES" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 114012
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 114012
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCOSTA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111119
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111119
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "SANTDOL" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111020
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111020
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          If WNombase = "MPEREZ" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111121
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111121
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROUD" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111006
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111006
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "FOSORIO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111128
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111128
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "ROXANA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111015
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111015
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PAOLA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111125
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111125
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PPONS" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111127
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111127
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "GUSTAVO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111123
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111123
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          dbcborub.Text = data_rub.Recordset("nombre")
          txt_con.Text = data_rub.Recordset("es")
          If txt_codrub.Text > 400000 And txt_codrub.Text <= 499999 Then
             If txt_codrub.Text = 421013 Then
                cboiva.ListIndex = 0
                txt_impiva.Text = 0
                cboiva.Enabled = False
                txt_impiva.Enabled = False
              Else
                cboiva.ListIndex = 0
                txt_impiva.Text = 0
                cboiva.Enabled = True
                txt_impiva.Enabled = True
              End If
          Else
              If txt_codrub.Text = 523122 Then
                 cboiva.ListIndex = 0
                 txt_impiva.Text = 0
                 cboiva.Enabled = False
                 txt_impiva.Enabled = False
              Else
                 If txt_codrub.Text >= 115000 And txt_codrub.Text <= 115999 Then
                    cboiva.ListIndex = 0
                    txt_impiva.Text = 0
                    cboiva.Enabled = True
                    txt_impiva.Enabled = True
                 Else
                    If txt_codrub.Text >= 120000 And txt_codrub.Text <= 129999 Then
                       cboiva.ListIndex = 0
                       txt_impiva.Text = 0
                       cboiva.Enabled = True
                       txt_impiva.Enabled = True
                    Else
                       If txt_codrub.Text >= 1200 And txt_codrub.Text <= 1299 Then
                          cboiva.ListIndex = 0
                          txt_impiva.Text = 0
                          cboiva.Enabled = True
                          txt_impiva.Enabled = True
                       Else
                          cboiva.ListIndex = 0
                          txt_impiva.Text = 0
                          cboiva.Enabled = False
                          txt_impiva.Enabled = False
                       End If
                    End If
                 End If
              End If
          End If
          Xconce = data_rub.Recordset("es")
          txt_con.SetFocus
       Else
          data_rub.RecordSource = "select * from rubteso where nombre >='" & dbcborub.Text & "' order by nombre"
          data_rub.Refresh
          dbcborub.Height = 2350
          If WNombase = "BANCOITAU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111018
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111018
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "CHEQUES" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 114012
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 114012
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCOSTA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111119
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111119
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "SANTDOL" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111020
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111020
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          If WNombase = "MPEREZ" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111121
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111121
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROUD" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111006
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111006
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "FOSORIO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111128
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111128
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "ROXANA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111015
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111015
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PAOLA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111125
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111125
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PPONS" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111127
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111127
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "GUSTAVO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111123
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 111123
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          dbcborub.Text = data_rub.Recordset("nombre")
'          txt_codrub.Text = data_rub.Recordset("codigo")
          txt_con.Text = data_rub.Recordset("es")
          If data_rub.Recordset("codigo") > 400000 And data_rub.Recordset("codigo") <= 499999 Then
             If data_rub.Recordset("codigo") = 421013 Then
                cboiva.ListIndex = 0
                txt_impiva.Text = 0
                cboiva.Enabled = False
                txt_impiva.Enabled = False
             Else
                cboiva.ListIndex = 0
                txt_impiva.Text = 0
                cboiva.Enabled = True
                txt_impiva.Enabled = True
             End If
          Else
             If data_rub.Recordset("codigo") = 523122 Then
                cboiva.ListIndex = 0
                txt_impiva.Text = 0
                cboiva.Enabled = False
                txt_impiva.Enabled = False
             Else
                If data_rub.Recordset("codigo") >= 115000 And data_rub.Recordset("codigo") <= 115999 Then
                   cboiva.ListIndex = 0
                   txt_impiva.Text = 0
                   cboiva.Enabled = True
                   txt_impiva.Enabled = True
                Else
                   If data_rub.Recordset("codigo") >= 120000 And data_rub.Recordset("codigo") <= 129999 Then
                      cboiva.ListIndex = 0
                      txt_impiva.Text = 0
                      cboiva.Enabled = True
                      txt_impiva.Enabled = True
                   Else
                      If data_rub.Recordset("codigo") >= 1200 And data_rub.Recordset("codigo") <= 1299 Then
                         cboiva.ListIndex = 0
                         txt_impiva.Text = 0
                         cboiva.Enabled = True
                         txt_impiva.Enabled = True
                      Else
                         cboiva.ListIndex = 0
                         txt_impiva.Text = 0
                         cboiva.Enabled = False
                         txt_impiva.Enabled = False
                      End If
                   End If
                End If
             End If
          End If

'          txt_imp.SetFocus
          Xconce = data_rub.Recordset("es")
       End If
    End If
If txt_debe.Text <> "" Then
   Xdeb = txt_debe.Text
   Xhab = txt_haber.Text
Else
   Xdeb = 9999999
   Xhab = 9999999
   
End If

End Sub

Private Sub dbcborub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txt_codrub.Text <> "" Then
    Else
       dbcborub.ListField = "nombre"
       dbcborub.BoundColumn = "nombre"
       data_rub.Recordset.FindFirst "nombre ='" & dbcborub.Text & "'"
       If Not data_rub.Recordset.NoMatch Then
          dbcborub.Text = data_rub.Recordset("nombre")
          txt_codrub.Text = data_rub.Recordset("codigo")
          dbcborub.Height = 500
          dbcborub.ListField = ""
          dbcborub.BoundColumn = ""
          If WNombase = "BANCOITAU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111018
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111018
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "CHEQUES" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 114012
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 114012
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          If WNombase = "MCOSTA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111119
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111119
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "SANTDOL" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111020
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111020
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          If WNombase = "MPEREZ" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111121
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111121
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROUD" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111006
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111006
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "FOSORIO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111128
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111128
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "ROXANA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111015
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111015
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PAOLA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111125
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111125
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PPONS" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111127
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111127
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "GUSTAVO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111123
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111123
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          dbcborub.Text = data_rub.Recordset("nombre")
          txt_con.Text = data_rub.Recordset("es")
          If data_rub.Recordset("codigo") > 400000 And data_rub.Recordset("codigo") <= 499999 Then
             If data_rub.Recordset("codigo") = 412013 Then
                cboiva.ListIndex = 0
                txt_impiva.Text = 0
                cboiva.Enabled = False
                txt_impiva.Enabled = False
              Else
                cboiva.ListIndex = 0
                txt_impiva.Text = 0
                cboiva.Enabled = True
                txt_impiva.Enabled = True
              End If
          Else
              If data_rub.Recordset("codigo") = 523122 Then
                 cboiva.ListIndex = 0
                 txt_impiva.Text = 0
                 cboiva.Enabled = False
                 txt_impiva.Enabled = False
              Else
                 If data_rub.Recordset("codigo") >= 115000 And data_rub.Recordset("codigo") <= 115999 Then
                    cboiva.ListIndex = 0
                    txt_impiva.Text = 0
                    cboiva.Enabled = True
                    txt_impiva.Enabled = True
                 Else
                    If data_rub.Recordset("codigo") >= 120000 And data_rub.Recordset("codigo") <= 129999 Then
                       cboiva.ListIndex = 0
                       txt_impiva.Text = 0
                       cboiva.Enabled = True
                       txt_impiva.Enabled = True
                    Else
                       If data_rub.Recordset("codigo") >= 1200 And data_rub.Recordset("codigo") <= 1299 Then
                          cboiva.ListIndex = 0
                          txt_impiva.Text = 0
                          cboiva.Enabled = True
                          txt_impiva.Enabled = True
                       Else
                          cboiva.ListIndex = 0
                          txt_impiva.Text = 0
                          cboiva.Enabled = False
                          txt_impiva.Enabled = False
                       End If
                    End If
                 End If
              End If
          End If
          Xconce = data_rub.Recordset("es")
          txt_con.SetFocus
       Else
          data_rub.RecordSource = "select * from rubteso where nombre >='" & dbcborub.Text & "' order by nombre"
          data_rub.Refresh
          dbcborub.Height = 2350
          If WNombase = "BANCOITAU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111018
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111018
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "CHEQUES" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 114012
                txt_haber.Text = data_rub.Recordset("codigo")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("codigo")
                   txt_haber.Text = 114012
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          If WNombase = "MCOSTA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111119
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111119
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "SANTDOL" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111020
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111020
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          If WNombase = "MPEREZ" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111121
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111121
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROUD" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111006
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111006
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "FOSORIO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111128
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111128
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "ROXANA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111015
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111015
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "MCCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMCREDIT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111005
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111005
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMSANT" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111007
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111007
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "RMBROU" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111001
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111001
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PAOLA" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111125
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111125
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "PPONS" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111127
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111127
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          If WNombase = "GUSTAVO" Then
             If data_rub.Recordset("es") = "E" Then
                txt_debe.Text = 111123
                txt_haber.Text = data_rub.Recordset("haber")
             Else
                If data_rub.Recordset("es") = "S" Then
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = 111123
                Else
                   txt_debe.Text = data_rub.Recordset("debe")
                   txt_haber.Text = data_rub.Recordset("haber")
                End If
             End If
          End If
          
          dbcborub.Text = data_rub.Recordset("nombre")
'          txt_codrub.Text = data_rub.Recordset("codigo")
          txt_con.Text = data_rub.Recordset("es")
          If txt_codrub.Text <> "" Then
            If data_rub.Recordset("codigo") > 400000 And data_rub.Recordset("codigo") <= 499999 Then
               If txt_codrub.Text = 412013 Then
                  cboiva.ListIndex = 0
                  txt_impiva.Text = 0
                  cboiva.Enabled = False
                  txt_impiva.Enabled = False
               Else
                  cboiva.ListIndex = 0
                  txt_impiva.Text = 0
                  cboiva.Enabled = True
                  txt_impiva.Enabled = True
               End If
            Else
               If data_rub.Recordset("codigo") = 523122 Then
                  cboiva.ListIndex = 0
                  txt_impiva.Text = 0
                  cboiva.Enabled = False
                  txt_impiva.Enabled = False
               Else
                  If data_rub.Recordset("codigo") >= 115000 And data_rub.Recordset("codigo") <= 115999 Then
                     cboiva.ListIndex = 0
                     txt_impiva.Text = 0
                     cboiva.Enabled = True
                     txt_impiva.Enabled = True
                  Else
                     If data_rub.Recordset("codigo") >= 120000 And data_rub.Recordset("codigo") <= 129999 Then
                        cboiva.ListIndex = 0
                        txt_impiva.Text = 0
                        cboiva.Enabled = True
                        txt_impiva.Enabled = True
                     Else
                        If data_rub.Recordset("codigo") >= 1200 And data_rub.Recordset("codigo") <= 1299 Then
                           cboiva.ListIndex = 0
                           txt_impiva.Text = 0
                           cboiva.Enabled = True
                           txt_impiva.Enabled = True
                        Else
                           cboiva.ListIndex = 0
                           txt_impiva.Text = 0
                           cboiva.Enabled = False
                           txt_impiva.Enabled = False
                        End If
                     End If
                  End If
               End If
            End If
         End If
'          txt_imp.SetFocus
          Xconce = data_rub.Recordset("es")
       End If
    End If
    If txt_debe.Text <> "" Then
       Xdeb = txt_debe.Text
       Xhab = txt_haber.Text
    Else
       Xdeb = 999999
       Xhab = 999999
    End If
End If

End Sub

Private Sub dbcborub_LostFocus()
If txt_codrub.Text <> "" Then
    If Xdeb = txt_codrub.Text Then
    Else
       If Xhab = txt_codrub.Text Then
       Else
          Xdeb = 999999
          Xhab = 999999
          MsgBox "ATENCION: Verifique el rubro", vbCritical, "Tesorería"
          dbcborub.SetFocus
       End If
    End If
End If

If cbomon.Enabled = True Then
   cbomon.SetFocus
   If data_rub.Recordset("moneda") = 2 Then
      cbomon.ListIndex = 1
   Else
      cbomon.ListIndex = 0
   End If
End If

End Sub

Private Sub Form_Load()
Dim Xresp As String
Dim Xsiono As Integer
Dim Xlafetes As Date
Xlafetes = CDate("01/07/2015")
Xsiono = 0
Control_vence
data_cl.ConnectionString = "dsn=" & Xconexrmt
'data_cajtes.DatabaseName = App.Path & "\sapp.mdb"
data_cajtes.ConnectionString = "dsn=" & Xconexrmt
data_part.DatabaseName = App.path & "\parteso.mdb"
data_part.RecordSource = "parteso"
data_part.Refresh

'data_cajtes.Refresh
data_rub.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_val.DatabaseName = App.path & "\tcambio.mdb"
data_val.RecordSource = "tvalor"
data_val.Refresh
WBase = frm_menu.data_parse.Recordset("base")
If UCase(WElusuario) = "SANTDOL" Then
   WNombase = UCase(WElusuario)
   Xsiono = 1
Else
   If UCase(WElusuario) = "MCOSTA" Or UCase(WElusuario) = "CHEQUES" Then
      WNombase = UCase(WElusuario)
      Xsiono = 1
   Else
      If UCase(WElusuario) = "BANCOITAU" Then
         WNombase = UCase(WElusuario)
         Xsiono = 1
      Else
         If UCase(WElusuario) = "ROXANA" Then
            WNombase = UCase(WElusuario)
            Xsiono = 1
         Else
            If UCase(WElusuario) = "MCBROU" Then
               WNombase = UCase(WElusuario)
               Xsiono = 1
            Else
               If UCase(WElusuario) = "MCSANT" Then
                  WNombase = UCase(WElusuario)
                  Xsiono = 1
               Else
                  If UCase(WElusuario) = "MCCREDIT" Then
                     WNombase = UCase(WElusuario)
                     Xsiono = 1
                  Else
                     If UCase(WElusuario) = "RMBROU" Then
                        WNombase = UCase(WElusuario)
                        Xsiono = 1
                     Else
                        If UCase(WElusuario) = "RMSANT" Then
                           WNombase = UCase(WElusuario)
                           Xsiono = 1
                        Else
                           If UCase(WElusuario) = "FOSORIO" Then
                              WNombase = UCase(WElusuario)
                              Xsiono = 1
                           Else
                              If UCase(WElusuario) = "PAOLA" Then
                                 WNombase = UCase(WElusuario)
                                 Xsiono = 1
                              Else
                                 If UCase(WElusuario) = "MPEREZ" Then
                                    WNombase = UCase(WElusuario)
                                    Xsiono = 1
                                 Else
                                    If UCase(WElusuario) = "RMBROUD" Then
                                       WNombase = UCase(WElusuario)
                                       Xsiono = 1
                                    Else
                                      If UCase(WElusuario) = "PPONS" Then
                                         WNombase = UCase(WElusuario)
                                         Xsiono = 1
                                      Else
                                        If UCase(WElusuario) = "JFERNAN" Then
                                           WNombase = InputBox("Ingrese USUARIO de tesorería: ", "Autenticación")
                                           WNombase = UCase(WNombase)
                                           If WNombase = "SANTDOL" Then
                                              Xsiono = 1
                                           Else
                                              If WNombase = "MCOSTA" Then
                                                 Xsiono = 1
                                              Else
                                                 If WNombase = "ROXANA" Then
                                                    Xsiono = 1
                                                 Else
                                                    If WNombase = "BANCOITAU" Then
                                                       Xsiono = 1
                                                    Else
                                                       If WNombase = "MCBROU" Then
                                                          Xsiono = 1
                                                       Else
                                                          If WNombase = "MCSANT" Then
                                                             Xsiono = 1
                                                          Else
                                                             If WNombase = "MCCREDIT" Then
                                                                Xsiono = 1
                                                             Else
                                                                If WNombase = "RMBROU" Then
                                                                   Xsiono = 1
                                                                Else
                                                                   If WNombase = "FOSORIO" Then
                                                                      Xsiono = 1
                                                                   Else
                                                                      If WNombase = "RMSANT" Then
                                                                         Xsiono = 1
                                                                      Else
                                                                         If WNombase = "PAOLA" Then
                                                                            Xsiono = 1
                                                                         Else
                                                                            If WNombase = "MPEREZ" Then
                                                                               Xsiono = 1
                                                                            Else
                                                                               If WNombase = "RMBROUD" Then
                                                                                  Xsiono = 1
                                                                               Else
                                                                                  If WNombase = "PPONS" Then
                                                                                     Xsiono = 1
                                                                                  Else
                                                                                     MsgBox "Usuario no autorizado", vbCritical, "Mensaje"
                                                                                     Xsiono = 0
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
                                        Else
                                           MsgBox "Usuario no autorizado", vbCritical, "Mensaje"
                                          Xsiono = 0
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
Dim Xfecht As Date
Xfecht = CDate("01/01/2018")
If Xsiono = 1 Then
   If WElusuario = "RREGUEIRA" Then
      data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' and fecha >='" & Format(Xfecht, "yyyy-mm-dd") & "' order by nromov"
   Else
      data_cajtes.RecordSource = "Select * from tesorero where usuario ='" & WNombase & "' order by fecha,hora"
   End If
   data_cajtes.Refresh
   If data_cajtes.Recordset.RecordCount > 0 Then
      data_cajtes.Recordset.MoveFirst
      bctrol_Click
      data_cajtes.Recordset.MoveLast
      igualacaj
   End If
Else

   End
End If

Xsiono = 0

End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfecha_GotFocus()
''''i = Len(oMyTextBox.Text)
''''oMyTextBox.SelStart = 0
''''oMyTextBox.SelLength = i

End Sub

Private Sub mfecha_LostFocus()
If mfecha.Text <> "__/__/____" Then
Else
   MsgBox "Ingrese fecha"
End If

End Sub

Private Sub t_cli_KeyPress(KeyAscii As Integer)
If t_cli.Text = "" Then
   t_cli.Text = 0
End If
If KeyAscii = 13 Then
   cbomon.SetFocus
End If

End Sub

Private Sub t_cli_LostFocus()
If t_cli.Text <> "" Then
   data_cl.RecordSource = "Select * from abmdesp where nro =" & t_cli.Text & " and base <>" & 99
   data_cl.Refresh
   If data_cl.Recordset.RecordCount > 0 Then
      labnomcl.Caption = data_cl.Recordset("obsmot")
   Else
      labnomcl.Caption = ""
   End If
Else
   labnomcl.Caption = ""
End If

End Sub

Private Sub txt_codrub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_codrub.Text <> "" Then
      If IsNumeric(txt_codrub.Text) = True Then
        If txt_codrub.Text > 400000 And txt_codrub.Text <= 499999 Then
           If txt_codrub.Text = 421013 Then
              cboiva.ListIndex = 0
              txt_impiva.Text = 0
              cboiva.Enabled = False
              txt_impiva.Enabled = False
           Else
              cboiva.ListIndex = 0
              txt_impiva.Text = 0
              cboiva.Enabled = True
              txt_impiva.Enabled = True
           End If
        Else
           If txt_codrub.Text = 523122 Then
              cboiva.ListIndex = 0
              txt_impiva.Text = 0
              cboiva.Enabled = False
              txt_impiva.Enabled = False
           Else
              If txt_codrub.Text >= 115000 And txt_codrub.Text <= 115999 Then
                 cboiva.ListIndex = 0
                 txt_impiva.Text = 0
                 cboiva.Enabled = True
                 txt_impiva.Enabled = True
              Else
                 If txt_codrub.Text >= 120000 And txt_codrub.Text <= 129999 Then
                    cboiva.ListIndex = 0
                    txt_impiva.Text = 0
                    cboiva.Enabled = True
                    txt_impiva.Enabled = True
                 Else
                    If txt_codrub.Text >= 1200 And txt_codrub.Text <= 1299 Then
                       cboiva.ListIndex = 0
                       txt_impiva.Text = 0
                       cboiva.Enabled = True
                       txt_impiva.Enabled = True
                    Else
                       cboiva.ListIndex = 0
                       txt_impiva.Text = 0
                       cboiva.Enabled = False
                       txt_impiva.Enabled = False
                    End If
                 End If
              End If
           End If
        End If
      End If
   End If
   If dbcborub.Enabled = True Then
      dbcborub.SetFocus
   End If
End If

End Sub

Private Sub txt_codrub_LostFocus()
If IsNumeric(txt_codrub.Text) = True Then
   data_rub.RecordSource = "Select * from rubteso"
   data_rub.Refresh
   data_rub.Recordset.FindFirst "codigo =" & txt_codrub.Text
   If Not data_rub.Recordset.NoMatch Then
      If WNombase = "BANCOITAU" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111018
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111018
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "CHEQUES" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 114012
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 114012
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      
      If WNombase = "MCOSTA" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111119
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111119
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "SANTDOL" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111020
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111020
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      
      If WNombase = "MPEREZ" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111121
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111121
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "RMBROUD" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111006
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111006
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
            
      If WNombase = "FOSORIO" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111128
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111128
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "ROXANA" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111015
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111015
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "MCBROU" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111001
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111001
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "MCSANT" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111007
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111007
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "MCCREDIT" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111005
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111005
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "RMCREDIT" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111005
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111005
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "RMSANT" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111007
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111007
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "RMBROU" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111001
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111001
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "PAOLA" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111125
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111125
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "PPONS" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111127
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111127
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      If WNombase = "GUSTAVO" Then
         If data_rub.Recordset("es") = "E" Then
            txt_debe.Text = 111123
            txt_haber.Text = data_rub.Recordset("codigo")
         Else
            If data_rub.Recordset("es") = "S" Then
               txt_debe.Text = data_rub.Recordset("codigo")
               txt_haber.Text = 111123
            Else
               txt_debe.Text = data_rub.Recordset("debe")
               txt_haber.Text = data_rub.Recordset("haber")
            End If
         End If
      End If
      
      dbcborub.Text = data_rub.Recordset("nombre")
      txt_con.Text = data_rub.Recordset("es")
      If cbomon.Enabled = True Then
         cbomon.SetFocus
      End If
      If data_rub.Recordset("moneda") = 2 Then
         cbomon.ListIndex = 1
      Else
         cbomon.ListIndex = 0
      End If
      If txt_codrub.Text > 400000 And txt_codrub.Text <= 499999 Then
         If txt_codrub.Text = 421013 Then
            cboiva.ListIndex = 0
            txt_impiva.Text = 0
            cboiva.Enabled = False
            txt_impiva.Enabled = False
         Else
            cboiva.ListIndex = 0
            txt_impiva.Text = 0
            cboiva.Enabled = True
            txt_impiva.Enabled = True
         End If
      Else
         If txt_codrub.Text = 523122 Then
            cboiva.ListIndex = 0
            txt_impiva.Text = 0
            cboiva.Enabled = False
            txt_impiva.Enabled = False
         Else
            If txt_codrub.Text >= 115000 And txt_codrub.Text <= 115999 Then
               cboiva.ListIndex = 0
               txt_impiva.Text = 0
               cboiva.Enabled = True
               txt_impiva.Enabled = True
            Else
               If txt_codrub.Text >= 120000 And txt_codrub.Text <= 129999 Then
                  cboiva.ListIndex = 0
                  txt_impiva.Text = 0
                  cboiva.Enabled = True
                  txt_impiva.Enabled = True
               Else
                  If txt_codrub.Text >= 1200 And txt_codrub.Text <= 1299 Then
                     cboiva.ListIndex = 0
                     txt_impiva.Text = 0
                     cboiva.Enabled = True
                     txt_impiva.Enabled = True
                  Else
                     cboiva.ListIndex = 0
                     txt_impiva.Text = 0
                     cboiva.Enabled = False
                     txt_impiva.Enabled = False
                  End If
               End If
            End If
         End If
      End If
      Xconce = UCase(data_rub.Recordset("es"))
   Else
      dbcborub.Text = ""
      txt_debe.Text = ""
      txt_haber.Text = ""
      txt_con.Text = ""
      txt_codrub.Text = ""
      Xconce = ""
      MsgBox "No encontrado", vbInformation, "Mensaje"
      If dbcborub.Enabled = True Then
         dbcborub.SetFocus
      End If
   End If
Else
   dbcborub.Text = ""
   txt_debe.Text = ""
   txt_haber.Text = ""
   txt_con.Text = ""
   txt_codrub.Text = ""
   Xconce = ""
   If dbcborub.Enabled = True Then
      dbcborub.SetFocus
   End If
End If
If txt_debe.Text <> "" Then
   Xdeb = txt_debe.Text
   Xhab = txt_haber.Text
Else
   Xdeb = 9999999
   Xhab = 9999999
End If


End Sub

Private Sub txt_con_Change()
   KeyAscii = Asc(UCase(chr(KeyAscii)))

End Sub

Private Sub txt_con_KeyPress(KeyAscii As Integer)
If txt_con.Text <> "" Then
   KeyAscii = Asc(UCase(chr(KeyAscii)))
   If KeyAscii = 13 Then
'      If txt_imp.Enabled = True Then
'         txt_imp.SetFocus
'      Else
'         txt_imp2.SetFocus
'      End If
      t_cli.SetFocus
   End If
End If

End Sub

Private Sub txt_con_LostFocus()
'Dim Xdeb, Xhab As Long
If UCase(txt_con.Text) = "E" Then
'   Xdeb = txt_debe.Text
'   Xhab = txt_haber.Text
   If UCase(txt_con.Text) = Trim(UCase(Xconce)) Then
      txt_debe.Text = Xdeb
      txt_haber.Text = Xhab
   Else
      txt_debe.Text = Xhab
      txt_haber.Text = Xdeb
   End If
Else
   If UCase(txt_con.Text) = "S" Then
'      Xdeb = txt_debe.Text
'      Xhab = txt_haber.Text
      If UCase(txt_con.Text) = Trim(UCase(Xconce)) Then
         txt_debe.Text = Xdeb
         txt_haber.Text = Xhab
      Else
         txt_debe.Text = Xhab
         txt_haber.Text = Xdeb
      End If
   Else
      MsgBox "Ingrese un concepto válido", vbCritical, "Mensaje"
      txt_con.SetFocus
   End If
End If

End Sub

Private Sub txt_imp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cboiva.Enabled = True Then
      cboiva.SetFocus
   Else
      txt_obs.SetFocus
   End If
End If

End Sub

Private Sub txt_imp_LostFocus()
If txt_imp.Text <> "" Then
   If cbomon.Text = "DOLARES AMERICANOS" Then
      If txt_imp2.Text <> "" Then
'         txt_imp2.Text = txt_imp.Text * txt_tcam.Text
         txt_imp2.Text = Format(txt_imp2.Text, "Standard")
         txt_imp.Text = Format(txt_imp.Text, "Standard")
      Else
         MsgBox "Ingrese tipo de cambio"
         txt_tcam.SetFocus
      End If
   Else
      txt_imp.Text = Format(txt_imp.Text, "Standard")
   End If
End If

End Sub

Public Function igualacaj()
If data_cajtes.Recordset.EOF = False Then
    If IsNull(data_cajtes.Recordset("fecha")) = False Then
       mfecha.Text = Format(data_cajtes.Recordset("fecha"), "dd/mm/yyyy")
    Else
       mfecha.Text = "__/__/____"
    End If
    If IsNull(data_cajtes.Recordset("hora")) = False Then
       txt_hora.Text = Format(data_cajtes.Recordset("hora"), "HH:mm")
    Else
       txt_hora.Text = ""
    End If
    If IsNull(data_cajtes.Recordset("vence_cheq")) = False Then
       mfvencec.Text = Format(data_cajtes.Recordset("vence_cheq"), "dd/mm/yyyy")
    Else
       mfvencec.Text = "__/__/____"
    End If
    If IsNull(data_cajtes.Recordset("bandera")) = False Then
       t_cli.Text = data_cajtes.Recordset("bandera")
       If t_cli.Text <> "" Then
          data_cl.RecordSource = "Select * from abmdesp where nro =" & t_cli.Text & " and base <>" & 99
          data_cl.Refresh
          If data_cl.Recordset.RecordCount > 0 Then
             labnomcl.Caption = data_cl.Recordset("obsmot")
          Else
             labnomcl.Caption = ""
          End If
       Else
          labnomcl.Caption = ""
       End If
    Else
       t_cli.Text = 0
       labnomcl.Caption = ""
    End If
    
    If IsNull(data_cajtes.Recordset("usuario")) = False Then
       txt_base.Text = data_cajtes.Recordset("usuario")
    Else
       txt_base.Text = WNombase
    End If
    If IsNull(data_cajtes.Recordset("cod_rub")) = False Then
       txt_codrub.Text = data_cajtes.Recordset("cod_rub")
    Else
       txt_codrub.Text = ""
    End If
    
    If IsNull(data_cajtes.Recordset("nom_rub")) = False Then
       dbcborub.Text = data_cajtes.Recordset("nom_rub")
    Else
       dbcborub.Text = ""
    End If
    If IsNull(data_cajtes.Recordset("cod_debe")) = False Then
       txt_debe.Text = data_cajtes.Recordset("cod_debe")
    Else
       txt_debe.Text = ""
    End If
    If IsNull(data_cajtes.Recordset("cod_haber")) = False Then
       txt_haber.Text = data_cajtes.Recordset("cod_haber")
    Else
       txt_haber.Text = ""
    End If
    If IsNull(data_cajtes.Recordset("moneda")) = False Then
       If data_cajtes.Recordset("moneda") = 1 Then
          cbomon.ListIndex = 0
       Else
          If data_cajtes.Recordset("moneda") = 2 Then
             cbomon.ListIndex = 1
          Else
             cbomon.ListIndex = 0
          End If
       End If
    Else
       cbomon.ListIndex = 0
    End If
    If IsNull(data_cajtes.Recordset("concep")) = False Then
       txt_con.Text = data_cajtes.Recordset("concep")
    Else
       txt_con.Text = ""
    End If
    If IsNull(data_cajtes.Recordset("monto")) = False Then
       txt_imp.Text = Format(data_cajtes.Recordset("monto"), "Standard")
    Else
       txt_imp.Text = ""
    End If
    If IsNull(data_cajtes.Recordset("iva")) = False Then
       If data_cajtes.Recordset("iva") = 1 Then
          cboiva.ListIndex = 1
          txt_impiva.Text = Format(data_cajtes.Recordset("impiva"), "Standard")
       Else
          If data_cajtes.Recordset("iva") = 2 Then
             cboiva.ListIndex = 2
             txt_impiva.Text = Format(data_cajtes.Recordset("impiva"), "Standard")
          Else
             cboiva.ListIndex = 0
             txt_impiva.Text = Format(0, "Standard")
          End If
       End If
    Else
       cboiva.ListIndex = 0
       txt_impiva.Text = Format(0, "Standard")
    End If
    
    If IsNull(data_cajtes.Recordset("obs")) = False Then
       txt_obs.Text = data_cajtes.Recordset("obs")
    Else
       txt_obs.Text = ""
    End If
    If IsNull(data_cajtes.Recordset("moneda")) = False Then
       If data_cajtes.Recordset("moneda") = 2 Then
          If IsNull(data_cajtes.Recordset("saldou")) = False Then
             txt_imp2.Visible = True
             txt_tcam.Visible = True
             txt_imp2.Text = Format(data_cajtes.Recordset("saldou"), "Standard")
             txt_tcam.Text = Format(data_cajtes.Recordset("tcam"), "Standard")
          Else
             txt_imp2.Text = 0
             txt_tcam.Text = 0
             txt_imp2.Visible = False
             txt_tcam.Visible = False
          End If
       End If
    End If
    If IsNull(data_cajtes.Recordset("saldos")) = False Then
       labsaldop.Caption = Format(data_cajtes.Recordset("saldos"), "Standard")
    Else
       labsaldop.Caption = ""
    End If
End If

End Function

Public Function Fungraba()
data_cajtes.Recordset("base") = WBase
data_cajtes.Recordset("fecha") = Format(mfecha.Text, "dd/mm/yyyy")
data_cajtes.Recordset("hora") = Format(txt_hora.Text, "HH:mm")
data_cajtes.Recordset("usuario") = Mid(txt_base.Text, 1, 10)
data_cajtes.Recordset("cod_rub") = txt_codrub.Text
data_cajtes.Recordset("nom_rub") = Mid(dbcborub.Text, 1, 40)
data_cajtes.Recordset("cod_debe") = txt_debe.Text
data_cajtes.Recordset("cod_haber") = txt_haber.Text
If mfvencec.Text <> "__/__/____" Then
   data_cajtes.Recordset("vence_cheq") = CDate(mfvencec.Text)
Else
   If IsNull(data_cajtes.Recordset("vence_cheq")) = False Then
      data_cajtes.Recordset("vence_cheq") = Null
   End If
End If
If t_cli.Text <> "" Then
   data_cajtes.Recordset("bandera") = t_cli.Text
Else
   data_cajtes.Recordset("bandera") = 0
End If
If cbomon.ListIndex = 0 Then
   data_cajtes.Recordset("moneda") = 1
Else
   If cbomon.ListIndex = 1 Then
      data_cajtes.Recordset("moneda") = 2
   Else
      data_cajtes.Recordset("moneda") = 1
   End If
End If
If UCase(txt_con.Text) = "E" Then
   data_cajtes.Recordset("concep") = "E"
   data_cajtes.Recordset("descon") = "ENTRADA"
Else
   If UCase(txt_con.Text) = "S" Then
      data_cajtes.Recordset("concep") = "S"
      data_cajtes.Recordset("descon") = "SALIDA"
   Else
      If UCase(txt_con.Text) = "A" Then
         data_cajtes.Recordset("concep") = "A"
         data_cajtes.Recordset("descon") = "SIN M"
      Else
         data_cajtes.Recordset("concep") = "A"
         data_cajtes.Recordset("descon") = "SIN MOV"
      End If
   End If
End If
If cbomon.Text = "DOLARES AMERICANOS" Then
   If txt_imp.Text <> "" Then
      If txt_imp2.Text <> "" Then
         data_cajtes.Recordset("monto") = Format(txt_imp.Text, "Standard")
         data_cajtes.Recordset("saldou") = Format(txt_imp2.Text, "Standard")
         data_cajtes.Recordset("tcam") = Format(txt_tcam.Text, "Standard")
      Else
         data_cajtes.Recordset("monto") = Format(txt_imp.Text, "Standard")
         data_cajtes.Recordset("saldou") = 0
         data_cajtes.Recordset("tcam") = 0
      End If
   Else
      data_cajtes.Recordset("monto") = 0
      data_cajtes.Recordset("saldou") = 0
      data_cajtes.Recordset("tcam") = 0
   End If
Else
   If txt_imp.Text <> "" Then
      data_cajtes.Recordset("monto") = Format(txt_imp.Text, "Standard")
   Else
      data_cajtes.Recordset("monto") = 0
   End If
End If
If cboiva.ListIndex = 0 Then
   data_cajtes.Recordset("impiva") = 0
   data_cajtes.Recordset("iva") = 0
Else
   If cboiva.ListIndex = 1 Then
      If txt_impiva.Text <> "" Then
         data_cajtes.Recordset("impiva") = Format(txt_impiva.Text, "Standard")
         data_cajtes.Recordset("iva") = 1
      Else
         data_cajtes.Recordset("impiva") = 0
         data_cajtes.Recordset("iva") = 0
      End If
   Else
      If cboiva.ListIndex = 2 Then
         If txt_impiva.Text <> "" Then
            data_cajtes.Recordset("impiva") = Format(txt_impiva.Text, "Standard")
            data_cajtes.Recordset("iva") = 2
         Else
            data_cajtes.Recordset("impiva") = 0
            data_cajtes.Recordset("iva") = 0
         End If
      Else
         data_cajtes.Recordset("impiva") = 0
         data_cajtes.Recordset("iva") = 0
      End If
   End If
End If

If txt_obs.Text <> "" Then
   data_cajtes.Recordset("obs") = Mid(txt_obs.Text, 1, 50)
Else
   data_cajtes.Recordset("obs") = ""
End If

End Function

Public Function limpiar()
txt_base.Text = WNombase
mfecha.Text = "__/__/____"
txt_hora.Text = ""
txt_codrub.Text = ""
dbcborub.Text = ""
txt_debe.Text = ""
txt_haber.Text = ""
cbomon.ListIndex = 0
txt_con.Text = ""
txt_imp.Text = ""
txt_obs.Text = ""
txt_impiva.Text = ""
cboiva.ListIndex = 0
txt_imp2.Text = 0
txt_tcam.Text = 0
labsaldop.Caption = ""
t_cli.Text = ""
labnomcl.Caption = ""
mfvencec.Text = "__/__/____"

End Function

Private Sub txt_imp2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txt_imp.Enabled = True Then
      txt_imp.SetFocus
   Else
      txt_obs.SetFocus
   End If
End If

End Sub

Private Sub txt_imp2_LostFocus()
If txt_imp2.Text <> "" Then
   If cbomon.Text = "DOLARES AMERICANOS" Then
      If txt_tcam.Text <> "" Then
         txt_imp.Text = txt_imp2.Text * txt_tcam.Text
         txt_imp.Text = Format(txt_imp.Text, "Standard")
         txt_imp2.Text = Format(txt_imp2.Text, "Standard")
      Else
         MsgBox "Ingrese tipo de cambio"
         txt_tcam.SetFocus
      End If
   Else
'      txt_imp.Text = Format(txt_imp.Text, "Standard")
   End If
End If

End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub txt_tcam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_imp2.SetFocus
End If

End Sub

Private Sub txt_tcam_LostFocus()
If txt_tcam.Text <> data_val.Recordset("valor") Then
   data_val.Recordset.Edit
   data_val.Recordset("valor") = txt_tcam.Text
   data_val.Recordset("fecha") = Date
   data_val.Recordset.Update
End If

End Sub

Public Function oculta()
txt_tcam.Visible = False
txt_imp2.Visible = False
Label13.Visible = False
Label14.Visible = False

End Function


Public Sub Control_vence()
Dim Xsqlpromo, XsqlCons, XelGrupo As String
Dim Xrecclii As New ADODB.Recordset
Dim XrecGraba As New ADODB.Recordset
Dim XfechaChe As Date
XfechaChe = Date + 10

XelGrupo = ""

ConectarBD
ConbdSapp.Open

Xsqlpromo = "Select * from tesorero where vence_cheq >='" & Format(Date, "yyyy-mm-dd") & "' and vence_cheq <='" & Format(XfechaChe, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      MsgBox "Cheque vence el " & Format(Xrecclii("vence_cheq"), "dd/mm/yyyy") & " Obs: " & Xrecclii("obs")
      Xrecclii.MoveNext
   Loop
Else
   MsgBox "No hay cheques a vencer en los 10 días.", vbInformation
End If
Xrecclii.Close

ConbdSapp.Close

End Sub
