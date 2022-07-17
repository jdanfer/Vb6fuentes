VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_nuevas 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar Nuevas  entregas"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7410
   Icon            =   "frm_nuevas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_promos 
      Caption         =   "data_promos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc data_deu 
      Height          =   375
      Left            =   5160
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "data_deu"
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
   Begin MSAdodcLib.Adodc data_final 
      Height          =   330
      Left            =   5400
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Caption         =   "data_final"
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
   Begin VB.Data data_par 
      Caption         =   "data_par"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_ui 
      Caption         =   "data_ui"
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
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton b_fin 
      Caption         =   "fin"
      Height          =   495
      Left            =   3360
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton b_efct 
      Caption         =   "efct"
      Height          =   495
      Left            =   2400
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton b_etck 
      Caption         =   "etck"
      Height          =   495
      Left            =   3240
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Data data_ctrl 
      Caption         =   "data_ctrl"
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
      RecordSource    =   "ULTNRO"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton btn_sale 
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
      Left            =   6600
      Picture         =   "frm_nuevas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Salir"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton btn_can 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_nuevas.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cancelar acción"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton btn_gra 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_nuevas.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Guardar datos"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton btn_nue 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_nuevas.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo registro"
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos del recibo"
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
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.TextBox t_descu 
         Height          =   285
         Left            =   4080
         TabIndex        =   46
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox t_cantt 
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
         Left            =   3360
         TabIndex        =   43
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox t_deuda 
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
         Left            =   5520
         TabIndex        =   42
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox t_timb 
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
         Left            =   2040
         TabIndex        =   40
         Top             =   3120
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc data_cnv 
         Height          =   330
         Left            =   240
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
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
         Caption         =   "data_cnv"
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
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   4440
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "data_cli"
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
      Begin VB.TextBox t_serie 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   38
         Top             =   3600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         DataField       =   "qr"
         DataSource      =   "Data1"
         Height          =   1575
         Left            =   4920
         ScaleHeight     =   1515
         ScaleWidth      =   1635
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Data data_cablocal 
         Caption         =   "data_cablocal"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_emilocal 
         Caption         =   "data_emilocal"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_cabemi 
         Caption         =   "data_cabemi"
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
         Top             =   4320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt_ruc 
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
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   25
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txt_a 
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
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   23
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txt_m 
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
         MaxLength       =   2
         TabIndex        =   22
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txt_nrorec 
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
         Left            =   2400
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txt_nomcob 
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
         Left            =   2880
         TabIndex        =   14
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txt_cob 
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
         Left            =   2040
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txt_imp 
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
         Left            =   4320
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txt_col 
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
         Left            =   2040
         TabIndex        =   9
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txt_nomcat 
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
         Left            =   3240
         TabIndex        =   7
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txt_cat 
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
         Left            =   2040
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_nom 
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
         Left            =   2040
         TabIndex        =   5
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox txt_mat 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label labdescpor 
         Height          =   255
         Left            =   1800
         TabIndex        =   48
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label labcodpromo 
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label labdespromo 
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
         Left            =   2520
         TabIndex        =   44
         Top             =   4080
         Width           =   4095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "Deudas $."
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
         TabIndex        =   41
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Timbres $."
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
         TabIndex        =   39
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label labnrofact 
         Height          =   375
         Left            =   4920
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label labserie 
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label labtipof 
         Height          =   375
         Left            =   4800
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "RUT:"
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
         TabIndex        =   24
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "MES:"
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
         TabIndex        =   21
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Número de recibo:"
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
         TabIndex        =   15
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Cobrador:"
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
         TabIndex        =   12
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Importe:"
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
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Color de recibo:"
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
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Categoría:"
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
         TabIndex        =   4
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
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
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "MATRICULA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label labvenceok 
      Height          =   255
      Left            =   1440
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labcodseg 
      Height          =   255
      Left            =   5400
      TabIndex        =   35
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labcae 
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label labautoriza 
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labvence 
      Height          =   255
      Left            =   3240
      TabIndex        =   32
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Picture         =   "frm_nuevas.frx":1A6A
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1695
   End
End
Attribute VB_Name = "frm_nuevas"
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




Private Sub b_efct_Click()
Dim strIdTransac As String

Dim Xnograva2 As Double
Dim Xindi, Xlalinea As Integer

Dim Ximpposi As Double
Ximpposi = 0

Xnograva2 = 0

Set objPosCfe = New PosCfe
Dim objresultado As Resultado
'Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)

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
'  MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'       "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
    
'Enviando
  If Not EstaInicializado() Then Exit Sub
     Dim objCfe As CFE
     Set objCfe = New CFE
     Dim objCf As ClassFactory
     Set objCf = New ClassFactory
     Set objCfe.EFact = New EFact
    With objCfe.EFact.Encabezado.IdDoc
       .TipoCFE = objCf.EnumConverter.IdDoc_Fact_TipoCFEFromString(Trim(str(Data1.Recordset("debe_haber"))))
       .FchEmis.SetDate Year(Data1.Recordset("fecha")), Month(Data1.Recordset("fecha")), Day(Data1.Recordset("fecha"))
       .IsValidMntBruto = True
       .MntBruto = IdDoc_Tck_MntBruto_1
       .FmaPago = IdDoc_Fact_FmaPago_2
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
        .TipoDocRecep = DocType_2
        .CodPaisRecep = CodPaisType_UY
        .DocRecep = Data1.Recordset("ruc")
        .RznSocRecep = Data1.Recordset("apellidos")
        .DirRecep = Data1.Recordset("dir_cli")
        .CiudadRecep = Data1.Recordset("zona")
     End With
     With objCfe.EFact.Encabezado.Totales
        If IsNull(Data1.Recordset("tiquet")) = False Then
           Xnograva2 = Data1.Recordset("tiquet")
        End If
        If IsNull(Data1.Recordset("deudas")) = False Then
           Xnograva2 = Xnograva2 + Data1.Recordset("deudas")
        End If
        .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(Data1.Recordset("tipodoc"))
        .IsValidTpoCambio = True
        .TpoCambio.FromString "1"
        .IsValidMntNetoIvaTasaMin = True
        .IsValidMntIVATasaMin = True
        .IsValidMntNoGrv = True
        .MntNoGrv.FromString Format(Xnograva2, "0.00")
        .MntNetoIvaTasaMin.FromString Format(Data1.Recordset("servi"), "0.00")
        .IVATasaMin = TasaIVAType_10FullStop000
        .MntIVATasaMin.FromString Format(Data1.Recordset("iva"), "0.00")
        .CantLinDet.FromString Data1.Recordset("numero")
        .MntTotal.FromString Format(Data1.Recordset("total"), "0.00")
        .MntPagar.FromString Format(Data1.Recordset("total"), "0.00")
    End With
    data_cabemi.RecordSource = "Select * from cabemi where cliente2 =" & Data1.Recordset("cliente") & " and nro_linea not in (12) order by nro_linea"
    data_cabemi.Refresh
    If data_cabemi.Recordset.RecordCount > 0 Then
       data_cabemi.Recordset.MoveFirst
       Do While Not data_cabemi.Recordset.EOF
         With objCfe.EFact.Detalle.Item.AddNew
             If data_cabemi.Recordset("serie") = "DS" Then
                Ximpposi = data_cabemi.Recordset("monto") - data_cabemi.Recordset("nro_doc")
               .NroLinDet.FromString Trim(str(data_cabemi.Recordset("nro_linea")))
               .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_cabemi.Recordset("indic_fact"))))
               .NomItem = data_cabemi.Recordset("descrip")
               .cantidad.FromString Trim(str(data_cabemi.Recordset("cantidad")))
               .UniMed = "N/A"
               .PrecioUnitario.FromString Format(data_cabemi.Recordset("imp_srv"), "0.00")
               .IsValidDescuentoMonto = True
               .IsValidDescuentoPct = True
               .DescuentoPct.FromString Format(Data1.Recordset("descpor"), "0")
               .DescuentoMonto.FromString Format(data_cabemi.Recordset("nro_doc"), "0.00")
               .MontoItem.FromString Format(Ximpposi, "0.00")
             Else
               .NroLinDet.FromString Trim(str(data_cabemi.Recordset("nro_linea")))
               .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_cabemi.Recordset("indic_fact"))))
               .NomItem = data_cabemi.Recordset("descrip")
               .cantidad.FromString Trim(str(data_cabemi.Recordset("cantidad")))
               .UniMed = "N/A"
               .PrecioUnitario.FromString Format(data_cabemi.Recordset("imp_srv"), "0.00")
               .MontoItem.FromString Format(data_cabemi.Recordset("monto"), "0.00")
             End If
         End With
         data_cabemi.Recordset.MoveNext
       Loop
    End If
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
    Data1.Recordset.Edit
    Data1.Recordset("tipocta") = labserie.Caption
    Data1.Recordset("documento") = labnrofact.Caption
    Data1.Recordset.Update
    t_serie.Text = labserie.Caption
    txt_nrorec.Text = labnrofact.Caption
    
    labserie.Caption = ""
    labnrofact.Caption = ""
    b_fin_Click

End Sub

Private Sub b_etck_Click()
Dim strIdTransac As String
Dim Xnograva, XimpCuota As Double
Dim Xindi, Xlalinea As Integer
Dim Ximpposi As Double
Ximpposi = 0

Xindi = 2

Xnograva = 0
XimpCuota = 0
Set objPosCfe = New PosCfe
Dim objresultado As Resultado
'Set objresultado = objPosCfe.Inicializar("SAPP-105", "FD-105", vbNullString)
Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-206", vbNullString)

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
'   MsgBox "Conectado: " & CStr(objresultado22.Conectado) & vbNewLine & _
'         "Requiere clave del certificado: " & CStr(objresultado22.RequiereClaveCertificado)
 
   If Not EstaInicializado() Then Exit Sub
      Dim objCfe As CFE
      Set objCfe = New CFE
      Dim objCf As ClassFactory
      Set objCf = New ClassFactory
      Set objCfe.ETck = New ETck

     'Enviando
    With objCfe.ETck.Encabezado.IdDoc
      .TipoCFE = IdDoc_Tck_TipoCFE_101
      .FchEmis.SetDate Year(Data1.Recordset("fecha")), Month(Data1.Recordset("fecha")), Day(Data1.Recordset("fecha"))
      .IsValidMntBruto = True
      .MntBruto = IdDoc_Tck_MntBruto_1
      .FmaPago = IdDoc_Tck_FmaPago_2
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
      If Xtipodedocumento = 3 Then
         .TipoDocRecep = DocType_3
      Else
         .TipoDocRecep = DocType_4
      End If
      .CodPaisRecep = CodPaisType_UY
      .Receptor_Tck_Choice.DocRecepExt = Trim(str(Data1.Recordset("cliente")))
      If Xtipodedocumento = 3 Then
         .Receptor_Tck_Choice.DocRecep = Trim(str(Data1.Recordset("cedula"))) & Trim(str(Data1.Recordset("cod")))
      End If
      .RznSocRecep = Data1.Recordset("apellidos")
      .DirRecep = Data1.Recordset("dir_cli")
      .CiudadRecep = Data1.Recordset("zona")
    End With
          
    With objCfe.ETck.Encabezado.Totales
      If IsNull(Data1.Recordset("tiquet")) = False Then
         Xnograva = Data1.Recordset("tiquet")
      End If
      If IsNull(Data1.Recordset("deudas")) = False Then
         Xnograva = Xnograva + Data1.Recordset("deudas")
      End If
      .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(Data1.Recordset("tipodoc"))
      .IsValidTpoCambio = True
      .TpoCambio.FromString "1"
      .IsValidMntNetoIvaTasaMin = True
      .IsValidMntIVATasaMin = True
      .IsValidMntNoGrv = True
      .MntNoGrv.FromString Format(Xnograva, "0.00")
      .MntNetoIvaTasaMin.FromString Format(Data1.Recordset("servi"), "0.00")
      .IVATasaMin = TasaIVAType_10FullStop000
      .MntIVATasaMin.FromString Format(Data1.Recordset("iva"), "0.00")
      .CantLinDet.FromString Trim(str(Data1.Recordset("numero")))
      .MntTotal.FromString Format(Data1.Recordset("total"), "0.00")
      .MntPagar.FromString Format(Data1.Recordset("total"), "0.00")
    End With
    
    data_cabemi.RecordSource = "Select * from cabemi where cliente2 =" & Data1.Recordset("cliente") & " and nro_linea not in (12) order by nro_linea"
    data_cabemi.Refresh
    If data_cabemi.Recordset.RecordCount > 0 Then
       data_cabemi.Recordset.MoveFirst
       Xlalinea = 1
       Do While Not data_cabemi.Recordset.EOF
          With objCfe.ETck.Detalle.Item.AddNew
            If Val(data_cabemi.Recordset("monto")) < 0 Then
            Else
               If data_cabemi.Recordset("serie") = "DS" Then
                  Ximpposi = data_cabemi.Recordset("monto") - data_cabemi.Recordset("nro_doc")
                  .NroLinDet.FromString Trim(str(data_cabemi.Recordset("nro_linea")))
                  .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_cabemi.Recordset("indic_fact"))))
                  .NomItem = data_cabemi.Recordset("descrip")
                  .cantidad.FromString Trim(str(data_cabemi.Recordset("cantidad")))
                  .UniMed = "N/A"
                  .PrecioUnitario.FromString Format(data_cabemi.Recordset("imp_srv"), "0.00")
                  .IsValidDescuentoMonto = True
                  .IsValidDescuentoPct = True
                  .DescuentoPct.FromString Format(Data1.Recordset("descpor"), "0")
                  .DescuentoMonto.FromString Format(data_cabemi.Recordset("nro_doc"), "0.00")
                  .MontoItem.FromString Format(Ximpposi, "0.00")
                  Xlalinea = Xlalinea + 1
               Else
                  .NroLinDet.FromString Trim(str(data_cabemi.Recordset("nro_linea")))
                  .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_cabemi.Recordset("indic_fact"))))
                  .NomItem = data_cabemi.Recordset("descrip")
                  .cantidad.FromString Trim(str(data_cabemi.Recordset("cantidad")))
                  .UniMed = "N/A"
                  .PrecioUnitario.FromString Format(data_cabemi.Recordset("imp_srv"), "0.00")
                  .MontoItem.FromString Format(data_cabemi.Recordset("monto"), "0.00")
                  Xlalinea = Xlalinea + 1
               End If
            End If
          End With
          data_cabemi.Recordset.MoveNext
       Loop
    End If
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
    Data1.Recordset.Edit
    Data1.Recordset("tipocta") = labserie.Caption
    Data1.Recordset("documento") = labnrofact.Caption
    Data1.Recordset.Update
    t_serie.Text = labserie.Caption
    txt_nrorec.Text = labnrofact.Caption
    labserie.Caption = ""
    labnrofact.Caption = ""
    b_fin_Click
'   b_efct_Click


End Sub

Private Sub b_fin_Click()
borranue
Frame1.Enabled = False
btn_gra.Enabled = False
btn_can.Enabled = False
btn_nue.Enabled = True

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      data_cabemi.RecordSource = "Select * from cabemi where cliente2 =" & Data1.Recordset("cliente")
      data_cabemi.Refresh
      If data_cabemi.Recordset.RecordCount > 0 Then
         data_cabemi.Recordset.MoveFirst
         Do While Not data_cabemi.Recordset.EOF
            data_cabemi.Recordset.Edit
            data_cabemi.Recordset("serie") = Data1.Recordset("tipocta")
            data_cabemi.Recordset("nro_doc") = Data1.Recordset("documento")
            data_cabemi.Recordset.Update
            data_cabemi.Recordset.MoveNext
         Loop
         data_cabemi.Refresh
         data_cabemi.Recordset.MoveFirst
         Do While Not data_cabemi.Recordset.EOF
            data_cablocal.Recordset.AddNew
            data_cablocal.Recordset("serie") = data_cabemi.Recordset("serie")
            data_cablocal.Recordset("nro_doc") = data_cabemi.Recordset("nro_doc")
            data_cablocal.Recordset("cod_srv") = data_cabemi.Recordset("cod_srv")
            data_cablocal.Recordset("descrip") = data_cabemi.Recordset("descrip")
            data_cablocal.Recordset("imp_srv") = data_cabemi.Recordset("imp_srv")
            data_cablocal.Recordset("nro_linea") = data_cabemi.Recordset("nro_linea")
            data_cablocal.Recordset("fecha2") = data_cabemi.Recordset("fecha2")
            data_cablocal.Recordset("tipo_cod") = data_cabemi.Recordset("tipo_cod")
            data_cablocal.Recordset("indic_fact") = data_cabemi.Recordset("indic_fact")
            data_cablocal.Recordset("cantidad") = data_cabemi.Recordset("cantidad")
            data_cablocal.Recordset("monto") = data_cabemi.Recordset("monto")
            data_cablocal.Recordset("cliente2") = data_cabemi.Recordset("cliente2")
            data_cablocal.Recordset.Update
            data_cabemi.Recordset.MoveNext
         Loop
      End If
      Data1.Recordset.MoveNext
   Loop
'   MsgBox "Se continúa con la carga de las facturas a la deuda del socio.", vbInformation
   Data1.Recordset.MoveFirst
   data_emilocal.Recordset.AddNew
   data_emilocal.Recordset("cliente") = Data1.Recordset("cliente")
   data_emilocal.Recordset("apellidos") = Data1.Recordset("apellidos")
   data_emilocal.Recordset("cod_cnv") = Data1.Recordset("cod_cnv")
   If IsNull(Data1.Recordset("ruc")) = False Then
      If Data1.Recordset("ruc") <> "" Then
         data_emilocal.Recordset("ruc") = Data1.Recordset("ruc")
      End If
   End If
   data_emilocal.Recordset("nom_cnv") = Data1.Recordset("nom_cnv")
   data_emilocal.Recordset("nro_cobr") = Data1.Recordset("nro_cobr")
   data_emilocal.Recordset("color_rec") = Data1.Recordset("color_rec")
   data_emilocal.Recordset("importe") = Data1.Recordset("importe")
   data_emilocal.Recordset("servi") = Data1.Recordset("servi")
   data_emilocal.Recordset("total") = Data1.Recordset("total")
   data_emilocal.Recordset("mes") = Data1.Recordset("mes")
   data_emilocal.Recordset("ano") = Data1.Recordset("ano")
   data_emilocal.Recordset("dir_cli") = Data1.Recordset("dir_cli")
   data_emilocal.Recordset("fecha") = Data1.Recordset("fecha")
   data_emilocal.Recordset("grupo") = Data1.Recordset("grupo")
   data_emilocal.Recordset("nom_cobr") = Data1.Recordset("nom_cobr")
   data_emilocal.Recordset("iva") = Data1.Recordset("iva")
   data_emilocal.Recordset("debe_haber") = Data1.Recordset("debe_haber")
   data_emilocal.Recordset("tipocta") = Data1.Recordset("tipocta")
   data_emilocal.Recordset("cedula") = Data1.Recordset("cedula")
   data_emilocal.Recordset("cod") = Data1.Recordset("cod")
   data_emilocal.Recordset("tipodoc") = Data1.Recordset("tipodoc")
   data_emilocal.Recordset("tipo") = Data1.Recordset("tipo")
   data_emilocal.Recordset("moneda") = Data1.Recordset("moneda")
   data_emilocal.Recordset("origen") = Data1.Recordset("origen")
   data_emilocal.Recordset("operador") = Data1.Recordset("operador")
   data_emilocal.Recordset("fecha_cobr") = Data1.Recordset("fecha_cobr")
   data_emilocal.Recordset("fvence") = Data1.Recordset("fvence")
   data_emilocal.Recordset("Autoriza") = Data1.Recordset("Autoriza")
   data_emilocal.Recordset("RangoCAE") = Data1.Recordset("RangoCAE")
   data_emilocal.Recordset("CodSeg") = Data1.Recordset("CodSeg")
   data_emilocal.Recordset("qr") = Data1.Recordset("qr")
   data_emilocal.Recordset("zona") = Data1.Recordset("zona")
   data_emilocal.Recordset("tiquet") = Data1.Recordset("tiquet")
   data_emilocal.Recordset("deudas") = Data1.Recordset("deudas")
   data_emilocal.Recordset("numero") = Data1.Recordset("numero")
   data_emilocal.Recordset("documento") = Data1.Recordset("documento")
   data_emilocal.Recordset("promo") = Data1.Recordset("promo")
   data_emilocal.Recordset("descimp") = Data1.Recordset("descimp")
   data_emilocal.Recordset("descpor") = Data1.Recordset("descpor")
   
   data_emilocal.Recordset.Update

' agregar mysql
   data_final.Recordset.AddNew
   data_final.Recordset("cliente") = Data1.Recordset("cliente")
   data_final.Recordset("cod_cnv") = Data1.Recordset("cod_cnv")
   data_final.Recordset("nom_cnv") = Data1.Recordset("nom_cnv")
   data_final.Recordset("ruc") = Data1.Recordset("ruc")
   data_final.Recordset("tipocta") = Data1.Recordset("tipocta")
   data_final.Recordset("apellidos") = Data1.Recordset("apellidos")
   data_final.Recordset("cedula") = Data1.Recordset("cedula")
   data_final.Recordset("cod") = Data1.Recordset("cod")
   data_final.Recordset("fecha") = Data1.Recordset("fecha")
   data_final.Recordset("tipodoc") = Data1.Recordset("tipodoc")
   data_final.Recordset("documento") = Data1.Recordset("documento")
   data_final.Recordset("tipo") = Data1.Recordset("tipo")
   data_final.Recordset("importe") = Data1.Recordset("importe")
   data_final.Recordset("debe_haber") = Data1.Recordset("debe_haber")
   data_final.Recordset("moneda") = Data1.Recordset("moneda")
   data_final.Recordset("origen") = Data1.Recordset("origen")
   data_final.Recordset("operador") = Data1.Recordset("operador")
   data_final.Recordset("hora") = Data1.Recordset("hora")
   data_final.Recordset("dir_cli") = Data1.Recordset("dir_cli")
   data_final.Recordset("loc_cli") = Data1.Recordset("loc_cli")
   data_final.Recordset("tel_cli") = Data1.Recordset("tel_cli")
   data_final.Recordset("nro_superv") = Data1.Recordset("nro_superv")
   data_final.Recordset("nom_superv") = Data1.Recordset("nom_superv")
   data_final.Recordset("nro_vende") = Data1.Recordset("nro_vende")
   data_final.Recordset("nom_vende") = Data1.Recordset("nom_vende")
   data_final.Recordset("grupo") = Data1.Recordset("grupo")
   data_final.Recordset("numero") = Data1.Recordset("numero")
   data_final.Recordset("zona") = Data1.Recordset("zona")
   data_final.Recordset("nro_cobr") = Data1.Recordset("nro_cobr")
   data_final.Recordset("nom_cobr") = Data1.Recordset("nom_cobr")
   data_final.Recordset("mes") = Data1.Recordset("mes")
   data_final.Recordset("ano") = Data1.Recordset("ano")
   data_final.Recordset("color_rec") = Data1.Recordset("color_rec")
   data_final.Recordset("fecha_ing") = Data1.Recordset("fecha_ing")
   data_final.Recordset("fecha_nac") = Data1.Recordset("fecha_nac")
   data_final.Recordset("tiquet") = Data1.Recordset("tiquet")
   data_final.Recordset("deudas") = Data1.Recordset("deudas")
   data_final.Recordset("servi") = Data1.Recordset("servi")
   data_final.Recordset("iva") = Data1.Recordset("iva")
   data_final.Recordset("total") = Data1.Recordset("total")
   data_final.Recordset("promo") = Data1.Recordset("promo")
   data_final.Recordset("descimp") = Data1.Recordset("descimp")
   data_final.Recordset("descpor") = Data1.Recordset("descpor")
   data_final.Recordset.Update
   DoEvents
      
   data_deu.Recordset.AddNew
    data_deu.Recordset("cod_cnv") = Data1.Recordset("cod_cnv")
    data_deu.Recordset("nom_cnv") = Mid(Data1.Recordset("nom_cnv"), 1, 20)
    data_deu.Recordset("tipocta") = Data1.Recordset("tipocta")
    data_deu.Recordset("cliente") = Data1.Recordset("cliente")
    data_deu.Recordset("nombre") = Data1.Recordset("apellidos")
    data_deu.Recordset("fecha") = Data1.Recordset("fecha")
    data_deu.Recordset("tipodoc") = Data1.Recordset("tipodoc")
    data_deu.Recordset("documento") = Data1.Recordset("documento")
    data_deu.Recordset("importe") = Data1.Recordset("importe")
    data_deu.Recordset("moneda") = Data1.Recordset("moneda")
    data_deu.Recordset("origen") = Data1.Recordset("origen")
    data_deu.Recordset("nro_vende") = Data1.Recordset("nro_vende")
    data_deu.Recordset("grupo") = Data1.Recordset("grupo")
    data_deu.Recordset("saldo_cc") = 0
    data_deu.Recordset("mes") = Data1.Recordset("mes")
    data_deu.Recordset("ano") = Data1.Recordset("ano")
    data_deu.Recordset("nro_cobr") = Data1.Recordset("nro_cobr")
    data_deu.Recordset("nom_cobr") = Data1.Recordset("nom_cobr")
    data_deu.Recordset("estado_cta") = 1
    data_deu.Recordset("tiquet") = 0
    data_deu.Recordset("deudas") = 0
    data_deu.Recordset("total") = Data1.Recordset("total")
    data_deu.Recordset("servi") = Data1.Recordset("servi")
    data_deu.Recordset("iva") = Data1.Recordset("iva")
    data_deu.Recordset("nro_superv") = 0
    data_deu.Recordset("promo") = Data1.Recordset("promo")
    data_deu.Recordset("descimp") = Data1.Recordset("descimp")
    data_deu.Recordset("descpor") = Data1.Recordset("descpor")
    data_deu.Recordset.Update


    MsgBox "Proceso terminado.", vbInformation
End If



End Sub

Private Sub btn_can_Click()
Data1.Recordset.CancelUpdate
borranue
Frame1.Enabled = False
btn_gra.Enabled = False
btn_can.Enabled = False
btn_nue.Enabled = True

End Sub

Private Sub btn_gra_Click()
Dim Xiv As Double
Dim XImp As Double
Dim Xivanuevo As Double
Dim Xelimps, Xtots, XnuevoImp As Double
Dim Lineas As Integer
Lineas = 0

XnuevoImp = 0

Xivanuevo = 0
Dim Xfecsrv As Date
Xfecsrv = Date + 30
Dim Xlaui, Xmasdiezui As Double

Dim Xdig, Xrut, Xtot, Xtot2, Xfactor, i As Integer
Dim Xbandrut As Integer
Xtipodedocumento = 4
Xbandrut = 0
Xtot2 = 0
If Trim(t_deuda.Text) = "" Then
   t_deuda.Text = 0
End If
If Trim(t_timb.Text) = "" Then
   t_timb.Text = 0
End If
If Trim(t_cantt.Text) = "" Then
   t_cantt.Text = 0
End If

labtipof.Caption = "E-TICKET"
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If
If data_cabemi.Recordset.RecordCount > 0 Then
   data_cabemi.Recordset.MoveFirst
   Do While Not data_cabemi.Recordset.EOF
      data_cabemi.Recordset.Delete
      data_cabemi.Recordset.MoveNext
   Loop
End If

If data_ui.Recordset.RecordCount > 0 Then
   Xlaui = CDbl(data_ui.Recordset("descrip"))
   Xmasdiezui = CDbl(txt_imp.Text) / Xlaui
End If
If txt_imp.Text <> "" Then
   If txt_imp.Text > 0 Then
      If Xmasdiezui > 10000 Then
         If Xtipodedocumento = 4 Then
            Xtipodedocumento = 3
         End If
      End If
   End If
End If

If txt_ruc.Text <> "" Then
   i = 0
   If Len(Trim(txt_ruc.Text)) = 12 Then
      If IsNumeric(txt_ruc.Text) Then
         Xdig = Val(Mid(txt_ruc.Text, 12, 1))
         Xrut = Val(Mid(txt_ruc.Text, 1, 12))
         Xtot = 0
         Xfactor = 2
         For i = 1 To 11
             If i = 1 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 4
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 2 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 3
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 3 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 2
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 4 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 9
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 5 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 8
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 6 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 7
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 7 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 6
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 8 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 5
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 9 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 4
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 10 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 3
                Xtot2 = Xtot2 + Xtot
             End If
             If i = 11 Then
                Xtot = Val(Mid(txt_ruc.Text, i, 1)) * 2
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
         If Xdig = Val(Mid(txt_ruc.Text, 12, 1)) Then
            labtipof.Caption = "E-FACTURA"
         Else
            MsgBox "El RUT ingresado no es correcto.", vbCritical
               '  Command1.Enabled = False
            Xbandrut = 9
         End If
      Else
         MsgBox "El RUT ingresado no es numérico.", vbCritical
              'Command1.Enabled = False
         Xbandrut = 9
      End If
   Else
      MsgBox "El RUT no tiene la cantidad de dígitos correctamente.", vbCritical
           'Command1.Enabled = False
      Xbandrut = 9
   End If
End If


If Xbandrut = 9 Then
   MsgBox "Hay un error en el RUT, verifique!", vbCritical
Else
   Data1.Recordset.AddNew
   Data1.Recordset("cliente") = txt_mat.Text
   Data1.Recordset("apellidos") = txt_nom.Text
   Data1.Recordset("cod_cnv") = txt_cat.Text
   If txt_ruc.Text <> "" Then
      Data1.Recordset("ruc") = txt_ruc.Text
   End If
   Data1.Recordset("nom_cnv") = txt_nomcat.Text
   Data1.Recordset("nro_cobr") = txt_cob.Text
   Data1.Recordset("color_rec") = txt_col.Text
   Data1.Recordset("importe") = Format(txt_imp.Text, "0.00")
   If t_descu.Text <> "" Then
      If Val(t_descu.Text) > 0 Then
         txt_imp.Text = Val(txt_imp.Text) - Val(t_descu.Text)
         Xivanuevo = txt_imp.Text / 1.1 * 0.1
         Data1.Recordset("servi") = Val(txt_imp.Text) - Xivanuevo
         txt_imp.Text = Val(txt_imp.Text) + Val(t_descu.Text)
         Data1.Recordset("descimp") = -Val(t_descu.Text)
         Data1.Recordset("descpor") = Val(labdescpor.Caption)
         Data1.Recordset("promo") = labdespromo.Caption
      Else
         Xivanuevo = txt_imp.Text / 1.1 * 0.1
         Data1.Recordset("servi") = Val(txt_imp.Text) - Xivanuevo
      End If
   Else
      Xivanuevo = txt_imp.Text / 1.1 * 0.1
      Data1.Recordset("servi") = Val(txt_imp.Text) - Xivanuevo
   End If
   Data1.Recordset("servi") = Format(Data1.Recordset("servi"), "0.00")
   Xtots = Val(txt_imp.Text) + Val(t_deuda.Text) + Val(t_timb.Text)
   If t_descu.Text <> "" Then
      If Val(t_descu.Text) > 0 Then
         Data1.Recordset("total") = Xtots - Val(t_descu.Text)
'         Data1.Recordset("total") = Format(Xtots, "0.00")
      Else
         Data1.Recordset("total") = Format(Xtots, "0.00")
      End If
   Else
      Data1.Recordset("total") = Format(Xtots, "0.00")
   End If
   Data1.Recordset("mes") = txt_m.Text
   Data1.Recordset("ano") = txt_a.Text
   Data1.Recordset("dir_cli") = data_cli.Recordset("cl_direcci")
   Data1.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
   Data1.Recordset("grupo") = 0
   Data1.Recordset("nom_cobr") = txt_nomcob.Text
   Data1.Recordset("iva") = Xivanuevo
   Data1.Recordset("debe_haber") = 101
   If txt_ruc.Text <> "" Then
      Data1.Recordset("debe_haber") = 111
   End If
   Data1.Recordset("tipocta") = "SR"
   Data1.Recordset("cedula") = Int(data_cli.Recordset("cl_cedula"))
   Data1.Recordset("cod") = data_cli.Recordset("cl_codced")
   Data1.Recordset("tipodoc") = "UYU" 'moneda
   Data1.Recordset("tipo") = "N.EMISION"
   Data1.Recordset("moneda") = 2 'fpago crédito
   Data1.Recordset("origen") = "Cuota " + Trim(str(txt_m.Text)) + "/" + Trim(str(txt_a.Text))
   Data1.Recordset("operador") = WElusuario
   Data1.Recordset("fecha_cobr") = Xfecsrv 'servicio hasta
   Data1.Recordset("zona") = "Sin Datos"
   Data1.Recordset("tiquet") = Format(Val(t_timb.Text), "0.00")
   Data1.Recordset("deudas") = Format(Val(t_deuda.Text), "0.00")
   If t_timb.Text > 0 Then
      If t_deuda.Text > 0 Then
         Data1.Recordset("numero") = 3
      Else
         Data1.Recordset("numero") = 2
      End If
   Else
      If t_deuda.Text > 0 Then
         Data1.Recordset("numero") = 2
      Else
         Data1.Recordset("numero") = 1
      End If
   End If
   Data1.Recordset.Update
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
   End If

   Dim Xeltex As String
   Xeltex = Data1.Recordset("apellidos")
   labnrofact.Caption = Xeltex
   data_cabemi.Recordset.AddNew
   If t_descu.Text <> "" Then
      If Val(t_descu.Text) > 0 Then
         data_cabemi.Recordset("serie") = "DS"
         data_cabemi.Recordset("nro_doc") = Val(t_descu.Text)
      Else
         data_cabemi.Recordset("serie") = "SR"
         data_cabemi.Recordset("nro_doc") = 0
      End If
   Else
      data_cabemi.Recordset("serie") = "SR"
      data_cabemi.Recordset("nro_doc") = 0
   End If
   data_cabemi.Recordset("cod_srv") = "881"
   data_cabemi.Recordset("descrip") = "CUOTA MENSUAL"
   data_cabemi.Recordset("imp_srv") = Format(Val(txt_imp.Text), "0.00")
   data_cabemi.Recordset("nro_linea") = 1
   data_cabemi.Recordset("fecha2") = Date
   data_cabemi.Recordset("tipo_cod") = "INT1"
   If txt_imp.Text > 0 Then
      data_cabemi.Recordset("indic_fact") = 2
   Else
      data_cabemi.Recordset("indic_fact") = 5
   End If
   data_cabemi.Recordset("cantidad") = 1
   data_cabemi.Recordset("monto") = Format(Val(txt_imp.Text), "0.00")
   data_cabemi.Recordset("cliente2") = data_cli.Recordset("cl_codigo")
   data_cabemi.Recordset("mesc") = txt_m.Text
   data_cabemi.Recordset("anioc") = txt_a.Text
   data_cabemi.Recordset.Update
   data_cabemi.Refresh
   Lineas = Lineas + 1
   If data_cabemi.Recordset.RecordCount > 0 Then
      data_cabemi.Recordset.MoveFirst
   End If
   
   If t_deuda.Text > 0 Then
      data_cabemi.Recordset.AddNew
      data_cabemi.Recordset("cantidad") = 1
      data_cabemi.Recordset("serie") = "SR"
      data_cabemi.Recordset("nro_doc") = 0
      data_cabemi.Recordset("cod_srv") = "882"
      data_cabemi.Recordset("descrip") = "DEUDAS POR SERVICIOS"
      data_cabemi.Recordset("imp_srv") = Format(Val(t_deuda.Text), "0.00")
      data_cabemi.Recordset("nro_linea") = 2
      data_cabemi.Recordset("fecha2") = Date
      data_cabemi.Recordset("tipo_cod") = "INT1"
      data_cabemi.Recordset("indic_fact") = 1
      data_cabemi.Recordset("monto") = Format(Val(t_deuda.Text), "0.00")
      data_cabemi.Recordset("cliente2") = data_cli.Recordset("cl_codigo")
      data_cabemi.Recordset("mesc") = txt_m.Text
      data_cabemi.Recordset("anioc") = txt_a.Text
      data_cabemi.Recordset.Update
      Lineas = Lineas + 1
   End If
    
'timbres
   If t_timb.Text > 0 Then
      data_cabemi.Recordset.AddNew
      data_cabemi.Recordset("cantidad") = t_cantt.Text
      data_cabemi.Recordset("serie") = "SR"
      data_cabemi.Recordset("nro_doc") = 0
      data_cabemi.Recordset("cod_srv") = "995"
      data_cabemi.Recordset("descrip") = "TIMBRE PROFESIONAL"
      Xelimps = t_timb.Text / t_cantt.Text
      data_cabemi.Recordset("imp_srv") = Format(Xelimps, "0.00")
      If t_deuda.Text > 0 Then
         data_cabemi.Recordset("nro_linea") = 3
      Else
         data_cabemi.Recordset("nro_linea") = 2
      End If
      data_cabemi.Recordset("fecha2") = Date
      data_cabemi.Recordset("tipo_cod") = "INT1"
      data_cabemi.Recordset("indic_fact") = 1
      data_cabemi.Recordset("monto") = Format(Val(t_timb.Text), "0.00")
      data_cabemi.Recordset("cliente2") = data_cli.Recordset("cl_codigo")
      data_cabemi.Recordset("mesc") = txt_m.Text
      data_cabemi.Recordset("anioc") = txt_a.Text
      data_cabemi.Recordset.Update
      Lineas = Lineas + 1
   End If
   
   If t_descu.Text <> "" Then
      If Val(t_descu.Text) > 0 Then
         data_cabemi.Recordset.AddNew
         data_cabemi.Recordset("serie") = "SR"
         data_cabemi.Recordset("nro_doc") = 0
         data_cabemi.Recordset("cod_srv") = "883"
         data_cabemi.Recordset("descrip") = "PROMOCION " & labdespromo.Caption
         data_cabemi.Recordset("imp_srv") = -Val(t_descu.Text)
         data_cabemi.Recordset("nro_linea") = 12
         data_cabemi.Recordset("fecha2") = Date
         data_cabemi.Recordset("tipo_cod") = "INT1"
         data_cabemi.Recordset("indic_fact") = 2 '7
         data_cabemi.Recordset("cantidad") = 1
         data_cabemi.Recordset("monto") = -Val(t_descu.Text)
         data_cabemi.Recordset("cliente2") = data_cli.Recordset("cl_codigo")
         data_cabemi.Recordset("mesc") = txt_m.Text
         data_cabemi.Recordset("anioc") = txt_a.Text
         data_cabemi.Recordset.Update
      End If
   End If
   
   If txt_ruc.Text = "" Then
      b_etck_Click
   Else
      b_efct_Click
   End If

End If




End Sub

Private Sub btn_nue_Click()
Frame1.Enabled = True
borranue
t_serie.Text = ""
txt_nrorec.Text = ""

txt_mat.SetFocus
Data1.Recordset.AddNew
txt_m.Text = Month(Date)
txt_a.Text = Year(Date)
btn_gra.Enabled = True
btn_can.Enabled = True
btn_nue.Enabled = False

End Sub

Private Sub btn_sale_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim Nombre, NomCab, Dbnombre As String
Dim Xmes, Xano As Long

MsgBox "RECUERDE: las nuevas entregas solo se pueden realizar desde el 1° hasta el último día del mes!", vbExclamation

Xmes = Month(Date)
Xano = Year(Date)
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cnv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cnv.ConnectionString = "dsn=" & Xconexrmt
'data_final.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_final.ConnectionString = "dsn=" & Xconexrmt
Data1.DatabaseName = App.path & "\emitemp.mdb"

data_cabemi.DatabaseName = App.path & "\emitemp.mdb"
data_promos.Connect = "odbc;dsn=" & Xconexrmt & ";"

'Data1.Connect = "ODBC;DSN=sappfact;"
Nombre = "emi"
NomCab = "CAB"
Dbnombre = "DB"
If Xmes > 9 Then
   Nombre = Nombre + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   NomCab = NomCab + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   Dbnombre = Dbnombre + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
Else
   Nombre = Nombre + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   NomCab = NomCab + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   Dbnombre = Dbnombre + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
End If
data_emilocal.DatabaseName = App.path & "\" & Dbnombre & ".mdb"
data_cablocal.DatabaseName = App.path & "\" & Dbnombre & ".mdb"

Data1.RecordSource = "emision"
Data1.Refresh
''data_final.RecordSource = Nombre
''data_final.Refresh
data_cabemi.RecordSource = "cabemi"
data_cabemi.Refresh

data_ui.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ui.RecordSource = "hc_frecresp"
data_ui.Refresh

'data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deu.ConnectionString = "dsn=" & Xconexrmt
data_deu.RecordSource = "Select * from deudas limit 5"
data_deu.Refresh

data_emilocal.RecordSource = "select * from " & Nombre & " where cliente >=" & 1000 & " and cliente <=" & 30000
data_emilocal.Refresh

data_cablocal.RecordSource = "select * from " & NomCab & " where cliente2 >=" & 1000 & " and cliente2 <=" & 30000
data_cablocal.Refresh

data_final.RecordSource = "Select * from " & Nombre & " where cliente >=" & 1000 & " and cliente <=" & 30000
data_final.Refresh

data_par.Connect = "odbc;dsn=sappfact;"
data_par.RecordSource = "paramsapp"
data_par.Refresh



End Sub

Public Function borranue()
txt_mat.Text = ""
txt_nom.Text = ""
txt_ruc.Text = ""
txt_cat.Text = ""
txt_nomcat.Text = ""
txt_col.Text = ""
txt_imp.Text = ""
txt_cob.Text = ""
txt_nomcob.Text = ""
txt_nrorec.Text = ""
txt_m.Text = ""
txt_a.Text = ""


End Function

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_deuda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nrorec.SetFocus
End If

End Sub

Private Sub t_timb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_deuda.SetFocus
End If

End Sub

Private Sub txt_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nom.SetFocus
End If

End Sub

Private Sub txt_mat_LostFocus()
'data_cli.Recordset.FindFirst "cl_codigo =" & txt_mat.Text
Dim Descu As String
Descu = ""
labcodpromo.Caption = ""
labdescpor.Caption = ""
data_cli.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Text
data_cli.Refresh
t_descu.Text = ""
If data_cli.Recordset.RecordCount > 0 Then
   txt_cat.Text = data_cli.Recordset("cl_codconv")
   
   If data_cli.Recordset("estado") = 2 Then
      MsgBox "Socio figura de BAJA, verifique!!", vbCritical
   End If
   If IsNull(data_cli.Recordset("idpromos")) = False Then
      labcodpromo.Caption = data_cli.Recordset("idpromos")
      If Val(labcodpromo.Caption) > 0 Then
         data_promos.RecordSource = "select * from promocion_gpo where id=" & Val(labcodpromo.Caption)
         data_promos.Refresh
         If data_promos.Recordset.RecordCount > 0 Then
            data_promos.Recordset.MoveFirst
            If data_promos.Recordset("descu_imp") > 0 Then
               t_descu.Text = data_promos.Recordset("descu_imp")
               labdescpor.Caption = 0
            Else
               Descu = "0." & data_promos.Recordset("descu_por")
               labdescpor.Caption = data_promos.Recordset("descu_por")
               data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
               data_cnv.Refresh
               If data_cnv.Recordset.RecordCount > 0 Then
                  t_descu.Text = data_cnv.Recordset("cnv_precio") * CDbl(Descu)
               Else
                  t_descu.Text = 0
               End If
            End If
'1000483
            labdespromo.Caption = data_promos.Recordset("descrip")
         Else
            labdespromo.Caption = ""
         End If
      End If
   End If
   
   txt_nom.Text = data_cli.Recordset("cl_apellid")
   txt_nomcat.Text = data_cli.Recordset("cl_nomconv")
   txt_cob.Text = data_cli.Recordset("cl_nrocobr")
   txt_nomcob.Text = data_cli.Recordset("cl_nomcobr")
   t_timb.Text = 0
   t_deuda.Text = 0
'   data_cnv.Recordset.FindFirst "cnv_codigo ='" & txt_cat.Text & "'"
   data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & txt_cat.Text & "'"
   data_cnv.Refresh
   If data_cnv.Recordset.RecordCount > 0 Then
      txt_imp.Text = data_cnv.Recordset("cnv_precio")
      txt_col.Text = data_cnv.Recordset("cnv_colrec")
      If IsNull(data_cnv.Recordset("cnv_ruc")) = False Then
         txt_ruc.Text = Trim(data_cnv.Recordset("cnv_ruc"))
         If Trim(txt_ruc.Text) <> "" Then
            txt_nom.Text = data_cnv.Recordset("cnv_entre")
         End If
      Else
         txt_ruc.Text = ""
      End If
   Else
      txt_imp.Text = 0
   End If
   txt_nrorec.Text = 0
Else
   txt_imp.Text = 0
   t_timb.Text = 0
   t_deuda.Text = 0
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
        labnrofact.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
        Data1.Recordset.Edit
        labvence.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Vencimiento)
        labautoriza.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Autorizacion)
        labcae.Caption = labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroDesde) & " - " & labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroHasta)
        labcodseg.Caption = CStr(ResultadoCfe.EstadoCfe.CodigoSeguridad)
        If Len(labvence.Caption) = 8 Then
           labvenceok.Caption = Mid(labvence.Caption, 7, 2) & "/" & Mid(labvence.Caption, 5, 2) & "/" & Mid(labvence.Caption, 1, 4)
        Else
           labvenceok.Caption = "31/12/2016"
        End If
        Data1.Recordset("fvence") = CDate(labvenceok.Caption)
        Data1.Recordset("Autoriza") = Val(labautoriza.Caption)
        Data1.Recordset("RangoCAE") = Trim(labcae.Caption)
        Data1.Recordset("CodSeg") = Trim(labcodseg.Caption)
        Picture1.Picture = LoadPicture(App.path & "\qr.bmp")
        Data1.Recordset.Update


        strUltimoGuid = ResultadoCfe.EstadoCfe.Guid
        Set objUltimaSerieNumero = ResultadoCfe.EstadoCfe.SerieNumeroCfe
    Else
        MsgBox "Error en el documento", vbCritical
        labnrofact.Caption = ""
        Unload Me
        
    End If
    'cmdConsultaXguid.Enabled = True
    'cmdConsultaXnumero.Enabled = True
End Sub


