VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_acondmov 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Acondicionamiento de móviles"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_acondmov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ver TODO"
      Height          =   255
      Left            =   3720
      TabIndex        =   36
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Data data_enf 
      Caption         =   "data_enf"
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_med 
      Caption         =   "data_med"
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
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tabla Chóferes"
      Height          =   375
      Left            =   5760
      TabIndex        =   27
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3240
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf2 
      Caption         =   "data_inf2"
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
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_chof 
      Caption         =   "data_chof"
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
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_busca 
      Caption         =   "data_busca"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_graba 
      Caption         =   "data_graba"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_acondmov.frx":0442
      Height          =   1575
      Left            =   240
      OleObjectBlob   =   "frm_acondmov.frx":045B
      TabIndex        =   23
      Top             =   6360
      Width           =   7455
   End
   Begin VB.TextBox t_movb 
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   6240
      Picture         =   "frm_acondmov.frx":0FD2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton b_eli 
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   5040
      Picture         =   "frm_acondmov.frx":1414
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00C0C0FF&
      Enabled         =   0   'False
      Height          =   735
      Left            =   3840
      Picture         =   "frm_acondmov.frx":1856
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00C0C0FF&
      Enabled         =   0   'False
      Height          =   735
      Left            =   2640
      Picture         =   "frm_acondmov.frx":1C98
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   1440
      Picture         =   "frm_acondmov.frx":20DA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   240
      Picture         =   "frm_acondmov.frx":251C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Datos de registro del acondicionamiento"
      Enabled         =   0   'False
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox t_enf 
         Height          =   360
         Left            =   2280
         TabIndex        =   35
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox t_med 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2280
         TabIndex        =   33
         Top             =   1800
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_acondmov.frx":295E
         Left            =   2280
         List            =   "frm_acondmov.frx":296B
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   6120
         TabIndex        =   24
         ToolTipText     =   "Ver tabla de choferes..."
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox t_obs 
         Height          =   735
         Left            =   2280
         MaxLength       =   35
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3960
         Width           =   5175
      End
      Begin MSMask.MaskEdBox mhorh 
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         Format          =   "HH:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mhord 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         Format          =   "HH:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_chof 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   4
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox t_mov 
         Height          =   375
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Enfermería:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Médico:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Servicio:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label labmed 
         Height          =   375
         Left            =   6000
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label laben 
         Height          =   375
         Left            =   4800
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label labcod 
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label9 
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Demora total:"
         Height          =   255
         Left            =   5880
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label6 
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Observaciones:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Fecha/Hora de finalización:"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Fecha/Hora de comienzo:"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Chofer:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Número de móvil:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Buscar por móvil:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5880
      Width           =   1935
   End
End
Attribute VB_Name = "frm_acondmov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()

End Sub
