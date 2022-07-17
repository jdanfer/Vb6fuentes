VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_seguicovid 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento CODIV-19"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12330
   Icon            =   "frm_seguicovid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12330
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Datos de Seguimiento"
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
      Height          =   8415
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   11895
      Begin VB.Data data_consultar 
         Caption         =   "data_consultar"
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
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton b_addrea 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7080
         Picture         =   "frm_seguicovid.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Agregar a la lista"
         Top             =   3120
         Width           =   615
      End
      Begin VB.ListBox List3 
         Height          =   1035
         Left            =   7680
         TabIndex        =   56
         Top             =   3120
         Width           =   3975
      End
      Begin VB.ComboBox cbotestrea 
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
         ItemData        =   "frm_seguicovid.frx":0B14
         Left            =   2040
         List            =   "frm_seguicovid.frx":0B1E
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton b_addsol 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7080
         Picture         =   "frm_seguicovid.frx":0B36
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Agregar a la lista"
         Top             =   2160
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   1035
         Left            =   7680
         TabIndex        =   53
         Top             =   2160
         Width           =   3975
      End
      Begin VB.ComboBox cbotest 
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
         ItemData        =   "frm_seguicovid.frx":10C0
         Left            =   2040
         List            =   "frm_seguicovid.frx":10CA
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Data data_medsapp 
         Caption         =   "data_medsapp"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_lin 
         Caption         =   "data_lin"
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
         Top             =   4920
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Data adoaltas 
         Caption         =   "adoaltas"
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
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Alta_ficha"
         Height          =   495
         Left            =   3600
         TabIndex        =   51
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Elegir otro paciente y liberar este registro"
         Height          =   495
         Left            =   240
         Picture         =   "frm_seguicovid.frx":10E2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Deja el registro pendiente en seguimiento para que pueda tomarlo otro médico"
         Top             =   4200
         Width           =   3975
      End
      Begin VB.CheckBox chi 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Insuf.Resp."
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
         Left            =   10320
         TabIndex        =   49
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Diarrea"
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
         Left            =   8520
         TabIndex        =   48
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resfrío"
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
         Left            =   8520
         TabIndex        =   47
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox cht 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tos"
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
         Left            =   7200
         TabIndex        =   46
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chf 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fiebre"
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
         Left            =   7200
         TabIndex        =   45
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chestuvo 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cuarentena antes del inicio de síntomas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4920
         TabIndex        =   44
         Top             =   960
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mftras 
         Height          =   375
         Left            =   10320
         TabIndex        =   43
         Top             =   5640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chtrasl 
         BackColor       =   &H0000C000&
         Caption         =   "Traslado y alta de seguimiento"
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
         Left            =   7200
         TabIndex        =   42
         Top             =   5640
         Width           =   3135
      End
      Begin VB.ComboBox t_resultiso 
         Height          =   315
         ItemData        =   "frm_seguicovid.frx":166C
         Left            =   5640
         List            =   "frm_seguicovid.frx":1676
         TabIndex        =   41
         Text            =   "t_resultiso"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frm_seguicovid.frx":168E
         Left            =   4800
         List            =   "frm_seguicovid.frx":16A1
         TabIndex        =   40
         Top             =   6240
         Width           =   6855
      End
      Begin MSMask.MaskEdBox mfcomun 
         Height          =   375
         Left            =   3720
         TabIndex        =   38
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfsolici 
         Height          =   375
         Left            =   5640
         TabIndex        =   36
         Top             =   2160
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
      Begin VB.CheckBox chisosol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hisopado solicitado"
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
         Left            =   240
         TabIndex        =   35
         Top             =   2160
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mfreaiso 
         Height          =   375
         Left            =   5640
         TabIndex        =   33
         Top             =   2760
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
      Begin VB.CheckBox chisorea 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hisopado realizado"
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
         Left            =   240
         TabIndex        =   31
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_seguicovid.frx":17DB
         Left            =   2040
         List            =   "frm_seguicovid.frx":17F7
         TabIndex        =   30
         Top             =   1560
         Width           =   3015
      End
      Begin MSMask.MaskEdBox mfprox 
         Height          =   375
         Left            =   4800
         TabIndex        =   28
         Top             =   5640
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
      Begin MSMask.MaskEdBox mfsint 
         Height          =   375
         Left            =   10200
         TabIndex        =   26
         Top             =   480
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
      Begin MSMask.MaskEdBox mfContact 
         Height          =   375
         Left            =   5160
         TabIndex        =   25
         Top             =   480
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
      Begin VB.Data data_lla2 
         Caption         =   "data_lla2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   -120
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
         Height          =   495
         Left            =   1680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data data_conshc 
         Caption         =   "data_conshc"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_hce 
         Caption         =   "data_hce"
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
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   9000
         TabIndex        =   23
         Top             =   4320
         Width           =   735
         _ExtentX        =   1296
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
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   7680
         TabIndex        =   22
         Top             =   4320
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
      Begin VB.TextBox t_nro 
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
         Left            =   11040
         TabIndex        =   20
         ToolTipText     =   "Ingrese aquí el número de control a grabar"
         Top             =   6720
         Width           =   615
      End
      Begin VB.CommandButton b_graba 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   10200
         Picture         =   "frm_seguicovid.frx":186B
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Graba los datos (menos los controles de abajo)"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton b_graba2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9600
         Picture         =   "frm_seguicovid.frx":1DF5
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Graba los datos del control"
         Top             =   7920
         Width           =   375
      End
      Begin VB.TextBox t_ensuma 
         Height          =   1215
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   7080
         Width           =   7095
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   240
         TabIndex        =   15
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CheckBox chcmed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Control médico"
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
         Left            =   9720
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox chctel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Control telefónico"
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
         Left            =   7200
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox chcomun 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comunicación a Epidemiología"
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
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   3495
      End
      Begin VB.TextBox t_sintomas 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox t_inicio 
         Height          =   375
         Left            =   7200
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frm_seguicovid.frx":237F
         Left            =   3000
         List            =   "frm_seguicovid.frx":2395
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox chviaje 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Viaje"
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro. de controles"
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
         Left            =   240
         TabIndex        =   59
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro.de control:"
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
         Left            =   9720
         TabIndex        =   58
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Diagnóstico:"
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
         Left            =   2520
         TabIndex        =   39
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
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
         Left            =   4680
         TabIndex        =   37
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resultado"
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
         Left            =   4680
         TabIndex        =   34
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
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
         Left            =   4680
         TabIndex        =   32
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grupo:"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H000080FF&
         Caption         =   "Fecha próximo control:"
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
         Left            =   2520
         TabIndex        =   27
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de Contacto"
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
         Left            =   5160
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000C000&
         Caption         =   "Fecha y hora de ALTA"
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
         Left            =   4440
         TabIndex        =   21
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "En suma del seguimiento:"
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
         Left            =   2520
         TabIndex        =   16
         Top             =   6720
         Width           =   7095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Controles:(Doble click para ver control)"
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
         Left            =   240
         TabIndex        =   14
         Top             =   5160
         Width           =   11415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Síntomas:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Inicio de síntomas:"
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
         Left            =   7200
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contacto:"
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
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label labllam 
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labnom 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Paciente:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frm_seguicovid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_addrea_Click()
Dim XX As Integer
XX = 0
If chisorea.Value = 1 Then
   If cbotestrea.ListIndex >= 0 Then
      If mfreaiso.Text <> "__/__/____" Then
         data_consultar.RecordSource = "select * from seguimiento_tests where seguim_id =" & Data1.Recordset("nrolla") & " and realiza in (1) order by fecha"
         data_consultar.Refresh
         data_consultar.Recordset.AddNew
         data_consultar.Recordset("seguim_id") = Data1.Recordset("nrolla")
         data_consultar.Recordset("fecha") = CDate(mfreaiso.Text)
         data_consultar.Recordset("tipo") = cbotestrea.Text
         data_consultar.Recordset("solicita") = 0
         data_consultar.Recordset("realiza") = 1
         If Trim(t_resultiso.Text) <> "" Then
            data_consultar.Recordset("resultado") = t_resultiso.Text
         End If
         data_consultar.Recordset.Update
         List3.AddItem Format(mfreaiso.Text, "dd/mm/yyyy") & " -->" & cbotestrea.Text
         mfreaiso.Text = "__/__/____"
         cbotestrea.ListIndex = -1
         chisorea.Value = 0
      Else
         MsgBox "Falta dato."
      End If
   Else
      MsgBox "Falta dato."
   End If
Else
   MsgBox "Falta dato."
End If


End Sub

Private Sub b_addsol_Click()
Dim XX As Integer
XX = 0
If chisosol.Value = 1 Then
   If cbotest.ListIndex >= 0 Then
      If mfsolici.Text <> "__/__/____" Then
         data_consultar.RecordSource = "select * from seguimiento_tests where seguim_id =" & Data1.Recordset("nrolla") & " and solicita in (1) order by fecha"
         data_consultar.Refresh
         data_consultar.Recordset.AddNew
         data_consultar.Recordset("seguim_id") = Data1.Recordset("nrolla")
         data_consultar.Recordset("fecha") = CDate(mfsolici.Text)
         data_consultar.Recordset("tipo") = cbotest.Text
         data_consultar.Recordset("solicita") = 1
         data_consultar.Recordset("realiza") = 0
         data_consultar.Recordset.Update
         List2.AddItem Format(mfsolici.Text, "dd/mm/yyyy") & " -->" & cbotest.Text
         mfsolici.Text = "__/__/____"
         cbotest.ListIndex = -1
         chisosol.Value = 0
      Else
         MsgBox "Falta dato."
      End If
   Else
      MsgBox "Falta dato."
   End If
Else
   MsgBox "Falta dato."
End If


End Sub

Private Sub b_graba_Click()
'On Error GoTo Quees
'8/4 10.07

If IsNull(Data1.Recordset("viaje")) = False Then
   If Data1.Recordset("viaje") <> chviaje.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("viaje") = chviaje.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("viaje") = chviaje.Value
   Data1.Recordset.Update
End If

If mfContact.Text <> "__/__/____" Then
   If IsNull(Data1.Recordset("fec_contac")) = False Then
      If Format(Data1.Recordset("fec_contac"), "dd/mm/yyyy") <> Format(mfContact.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("fec_contac") = Format(mfContact.Text, "dd/mm/yyyy")
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("fec_contac") = Format(mfContact.Text, "dd/mm/yyyy")
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("fec_contac")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("fec_contac") = Null
      Data1.Recordset.Update
   End If
End If
If mfcomun.Text <> "__/__/____" Then
   If IsNull(Data1.Recordset("fec_comunica")) = False Then
      If Format(Data1.Recordset("fec_comunica"), "dd/mm/yyyy") <> Format(mfcomun.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("fec_comunica") = Format(mfcomun.Text, "dd/mm/yyyy")
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("fec_comunica") = Format(mfcomun.Text, "dd/mm/yyyy")
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("fec_comunica")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("fec_comunica") = Null
      Data1.Recordset.Update
   End If
End If

If IsNull(Data1.Recordset("cuarent_ant")) = False Then
   If Data1.Recordset("cuarent_ant") <> chestuvo.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("cuarent_ant") = chestuvo.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("cuarent_ant") = chestuvo.Value
   Data1.Recordset.Update
End If
If IsNull(Data1.Recordset("fiebre")) = False Then
   If Data1.Recordset("fiebre") <> chf.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("fiebre") = chf.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("fiebre") = chf.Value
   Data1.Recordset.Update
End If
If IsNull(Data1.Recordset("tos")) = False Then
   If Data1.Recordset("tos") <> cht.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("tos") = cht.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("tos") = cht.Value
   Data1.Recordset.Update
End If
If IsNull(Data1.Recordset("resfrio")) = False Then
   If Data1.Recordset("resfrio") <> chr.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("resfrio") = chr.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("resfrio") = chr.Value
   Data1.Recordset.Update
End If
If IsNull(Data1.Recordset("diarrea")) = False Then
   If Data1.Recordset("diarrea") <> chd.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("diarrea") = chd.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("diarrea") = chd.Value
   Data1.Recordset.Update
End If
If IsNull(Data1.Recordset("insuf")) = False Then
   If Data1.Recordset("insuf") <> chi.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("insuf") = chi.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("insuf") = chi.Value
   Data1.Recordset.Update
End If

If mfsint.Text <> "__/__/____" Then
   If IsNull(Data1.Recordset("fec_sint")) = False Then
      If Format(Data1.Recordset("fec_sint"), "dd/mm/yyyy") <> Format(mfsint.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("fec_sint") = Format(mfsint.Text, "dd/mm/yyyy")
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("fec_sint") = Format(mfsint.Text, "dd/mm/yyyy")
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("fec_sint")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("fec_sint") = Null
      Data1.Recordset.Update
   End If
End If

If mfsolici.Text <> "__/__/____" Then
   If IsNull(Data1.Recordset("isopa_fecsol")) = False Then
      If Format(Data1.Recordset("isopa_fecsol"), "dd/mm/yyyy") <> Format(mfsolici.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("isopa_fecsol") = Format(mfsolici.Text, "dd/mm/yyyy")
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("isopa_fecsol") = Format(mfsolici.Text, "dd/mm/yyyy")
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("isopa_fecsol")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("isopa_fecsol") = Null
      Data1.Recordset.Update
   End If
End If

If mfsolici2.Text <> "__/__/____" Then
   If IsNull(Data1.Recordset("fecisosol2")) = False Then
      If Format(Data1.Recordset("fecisosol2"), "dd/mm/yyyy") <> Format(mfsolici2.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("fecisosol2") = Format(mfsolici2.Text, "dd/mm/yyyy")
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("fecisosol2") = Format(mfsolici2.Text, "dd/mm/yyyy")
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("fecisosol2")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("fecisosol2") = Null
      Data1.Recordset.Update
   End If
End If

If mfreaiso.Text <> "__/__/____" Then
   If IsNull(Data1.Recordset("isopa_fecrea")) = False Then
      If Format(Data1.Recordset("isopa_fecrea"), "dd/mm/yyyy") <> Format(mfreaiso.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("isopa_fecrea") = Format(mfreaiso.Text, "dd/mm/yyyy")
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("isopa_fecrea") = Format(mfreaiso.Text, "dd/mm/yyyy")
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("isopa_fecrea")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("isopa_fecrea") = Null
      Data1.Recordset.Update
   End If
End If

If mfreaiso2.Text <> "__/__/____" Then
   If IsNull(Data1.Recordset("fecisorea2")) = False Then
      If Format(Data1.Recordset("fecisorea2"), "dd/mm/yyyy") <> Format(mfreaiso2.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("fecisorea2") = Format(mfreaiso2.Text, "dd/mm/yyyy")
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("fecisorea2") = Format(mfreaiso2.Text, "dd/mm/yyyy")
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("fecisorea2")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("fecisorea2") = Null
      Data1.Recordset.Update
   End If
End If

If Combo1.Text <> "" Then
    If IsNull(Data1.Recordset("contacto")) = False Then
       If Data1.Recordset("contacto") <> Combo1.Text Then
          Data1.Recordset.Edit
          Data1.Recordset("contacto") = Combo1.Text
          Data1.Recordset.Update
       End If
    Else
       Data1.Recordset.Edit
       Data1.Recordset("contacto") = Combo1.Text
       Data1.Recordset.Update
    End If
Else
    If IsNull(Data1.Recordset("contacto")) = False Then
       Data1.Recordset.Edit
       Data1.Recordset("contacto") = Null
       Data1.Recordset.Update
    End If
End If

If Combo2.Text <> "" Then
    If IsNull(Data1.Recordset("grupo_covid")) = False Then
       If Data1.Recordset("grupo_covid") <> Combo2.Text Then
          Data1.Recordset.Edit
          Data1.Recordset("grupo_covid") = Combo2.Text
          Data1.Recordset.Update
       End If
    Else
       Data1.Recordset.Edit
       Data1.Recordset("grupo_covid") = Combo2.Text
       Data1.Recordset.Update
    End If
Else
    If IsNull(Data1.Recordset("grupo_covid")) = False Then
       Data1.Recordset.Edit
       Data1.Recordset("grupo_covid") = Null
       Data1.Recordset.Update
    End If
End If

If t_inicio.Text <> "" Then
    If IsNull(Data1.Recordset("inicio_sint")) = False Then
       If Data1.Recordset("inicio_sint") <> t_inicio.Text Then
          Data1.Recordset.Edit
          Data1.Recordset("inicio_sint") = t_inicio.Text
          Data1.Recordset.Update
       End If
    Else
       Data1.Recordset.Edit
       Data1.Recordset("inicio_sint") = t_inicio.Text
       Data1.Recordset.Update
    End If
Else
    If IsNull(Data1.Recordset("inicio_sint")) = False Then
       Data1.Recordset.Edit
       Data1.Recordset("inicio_sint") = Null
       Data1.Recordset.Update
    End If
End If

If t_sintomas.Text <> "" Then
    If IsNull(Data1.Recordset("sintomas")) = False Then
       If Data1.Recordset("sintomas") <> t_sintomas.Text Then
          Data1.Recordset.Edit
          Data1.Recordset("sintomas") = t_sintomas.Text
          Data1.Recordset.Update
       End If
    Else
       Data1.Recordset.Edit
       Data1.Recordset("sintomas") = t_sintomas.Text
       Data1.Recordset.Update
    End If
Else
    If IsNull(Data1.Recordset("sintomas")) = False Then
       Data1.Recordset.Edit
       Data1.Recordset("sintomas") = Null
       Data1.Recordset.Update
    End If
End If
If IsNull(Data1.Recordset("comunic_epi")) = False Then
   If Data1.Recordset("comunic_epi") <> chcomun.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("comunic_epi") = chcomun.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("comunic_epi") = chcomun.Value
   Data1.Recordset.Update
End If

If IsNull(Data1.Recordset("isopa_sol")) = False Then
   If Data1.Recordset("isopa_sol") <> chisosol.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("isopa_sol") = chisosol.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("isopa_sol") = chisosol.Value
   Data1.Recordset.Update
End If

If IsNull(Data1.Recordset("isopa_rea")) = False Then
   If Data1.Recordset("isopa_rea") <> chisorea.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("isopa_rea") = chisorea.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("isopa_rea") = chisorea.Value
   Data1.Recordset.Update
End If

If t_resultiso.Text <> "" Then
   If IsNull(Data1.Recordset("isopa_result")) = False Then
      If Data1.Recordset("isopa_result") <> t_resultiso.Text Then
         Data1.Recordset.Edit
         Data1.Recordset("isopa_result") = t_resultiso.Text
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("isopa_result") = t_resultiso.Text
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("isopa_result")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("isopa_result") = Null
      Data1.Recordset.Update
   End If
End If

If t_result2.Text <> "" Then
   If IsNull(Data1.Recordset("resuliso2")) = False Then
      If Data1.Recordset("resuliso2") <> t_result2.Text Then
         Data1.Recordset.Edit
         Data1.Recordset("resuliso2") = t_result2.Text
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("resuliso2") = t_result2.Text
      Data1.Recordset.Update
   End If
Else
   If IsNull(Data1.Recordset("resuliso2")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("resuliso2") = Null
      Data1.Recordset.Update
   End If
End If

If IsNull(Data1.Recordset("ctrol_telef")) = False Then
   If Data1.Recordset("ctrol_telef") <> chctel.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("ctrol_telef") = chctel.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("ctrol_telef") = chctel.Value
   Data1.Recordset.Update
End If

If IsNull(Data1.Recordset("ctrol_medic")) = False Then
   If Data1.Recordset("ctrol_medic") <> chcmed.Value Then
      Data1.Recordset.Edit
      Data1.Recordset("ctrol_medic") = chcmed.Value
      Data1.Recordset.Update
   End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("ctrol_medic") = chcmed.Value
   Data1.Recordset.Update
End If

If mf.Enabled = True Then
   If mf.Text <> "__/__/____" Then
      If IsNull(Data1.Recordset("cierre_fec")) = False Then
         If Data1.Recordset("cierre_fec") <> mf.Text Then
            Data1.Recordset.Edit
            Data1.Recordset("cierre_fec") = mf.Text
            Data1.Recordset.Update
         End If
      Else
         Data1.Recordset.Edit
         Data1.Recordset("cierre_fec") = mf.Text
         Data1.Recordset.Update
      End If
   Else
      If IsNull(Data1.Recordset("cierre_fec")) = False Then
         Data1.Recordset.Edit
         Data1.Recordset("cierre_fec") = Null
         Data1.Recordset.Update
      End If
   End If

   If mh.Enabled = True Then
      If mh.Text <> "__:__" Then
         If IsNull(Data1.Recordset("cierre_hora")) = False Then
            If Data1.Recordset("cierre_hora") <> mh.Text Then
               Data1.Recordset.Edit
               Data1.Recordset("cierre_hora") = mh.Text
               Data1.Recordset.Update
            End If
         Else
            Data1.Recordset.Edit
            Data1.Recordset("cierre_hora") = mh.Text
            Data1.Recordset.Update
         End If
      Else
         If IsNull(Data1.Recordset("cierre_hora")) = False Then
            Data1.Recordset.Edit
            Data1.Recordset("cierre_hora") = Null
            Data1.Recordset.Update
         End If
      End If
   End If
   If t_resultiso.Text = "Negativo" Then
      If IsNull(Data1.Recordset("prox_control")) = False Then
         If Format(Data1.Recordset("prox_control"), "yyyy/mm/dd") <> Format(Date, "yyyy/mm/dd") Then
            Data1.Recordset.Edit
            Data1.Recordset("prox_control") = Format(Date, "dd/mm/yyyy")
            Data1.Recordset("cmt_enproceso") = 2
            Data1.Recordset.Update
         End If
      Else
         Data1.Recordset.Edit
         Data1.Recordset("prox_control") = Format(Date, "dd/mm/yyyy")
         Data1.Recordset("cmt_enproceso") = 2
         Data1.Recordset.Update
      End If
   Else
      If t_resultiso.Text = "Positivo" Then
         If IsNull(Data1.Recordset("prox_control")) = False Then
            If Format(Data1.Recordset("prox_control"), "dd/mm/yyyy") <> Format(Date, "dd/mm/yyyy") Then
               Data1.Recordset.Edit
               Data1.Recordset("prox_control") = Format(Date, "dd/mm/yyyy")
               Data1.Recordset("cmt_enproceso") = 2
               Data1.Recordset.Update
            End If
         Else
            Data1.Recordset.Edit
            Data1.Recordset("prox_control") = Format(Date, "dd/mm/yyyy")
            Data1.Recordset("cmt_enproceso") = 2
            Data1.Recordset.Update
         End If
      Else
      
      End If
   End If
   If mf.Text <> "__/__/____" And mh.Text <> "__:__" Then
      MsgBox "Se cerrará el llamado y no estará más en los pendientes de seguimiento", vbInformation
      data_lla2.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
      data_lla2.Refresh
      If data_lla2.Recordset.RecordCount > 0 Then
         If data_lla2.Recordset("pend") = 4 Then
            data_lla2.Recordset.Edit
            data_lla2.Recordset("pend") = 2
            data_lla2.Recordset("movilpas") = 2015
            data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
            data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
            data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
            data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
            data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
            data_lla2.Recordset("hor_llega") = Format(mh.Text, "HH:mm")
            data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
            data_lla2.Recordset("hor_rea") = Format(mh.Text, "HH:mm")
            data_lla2.Recordset("diag") = "CMT COVID-19"
            data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
            If IsNull(data_lla2.Recordset("mes")) = False Then
               If data_lla2.Recordset("mes") > 30 Then
                  data_lla2.Recordset("mes") = 0
               End If
            End If
             'data_lla.Recordset("codmed") = txt_codmed.Text
             'data_lla.Recordset("nommed") = dbcbomed.Text
            data_lla2.Recordset.Update
         Else
            If data_lla2.Recordset("pend") <> 2 Then
                data_lla2.Recordset.Edit
                data_lla2.Recordset("pend") = 2
                data_lla2.Recordset("movilpas") = 2015
                data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
                data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
                data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                data_lla2.Recordset("hor_llega") = Format(mh.Text, "HH:mm")
                data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                data_lla2.Recordset("hor_rea") = Format(mh.Text, "HH:mm")
                data_lla2.Recordset("diag") = "COVID-19 DERIVADO A BASE"
                data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
                If IsNull(data_lla2.Recordset("mes")) = False Then
                   If data_lla2.Recordset("mes") > 30 Then
                      data_lla2.Recordset("mes") = 0
                   End If
                End If
                 'data_lla.Recordset("codmed") = txt_codmed.Text
                 'data_lla.Recordset("nommed") = dbcbomed.Text
                data_lla2.Recordset.Update
                 
                 MsgBox "ATENCION!! El llamado ya NO figura cómo pasado a COVID! Se pasará como terminado.", vbInformation
            End If
         End If
      End If
   Else
       MsgBox "El llamado continuará cómo PASADO A Seguimiento Covid"
   End If
   
End If
MsgBox "Datos grabados correctamente"

'Exit Sub

'Quees:
'        If Err.Number = 3155 Then
'           MsgBox "ERROR: " & Err.Number & " " & Err.Description
'        Else
'           MsgBox "ERROR: " & Err.Number & " " & Err.Description
'        End If

End Sub

Private Sub b_graba2_Click()
On Error GoTo Quepasaalg
Dim Xhcesi, XcedDoc, XTextoparaHC As String
Dim Xcrear As Integer
Dim Xnrofactura As Long
Dim Xnromedico As Integer
Dim Xnommedico As String
Xnromedico = 0
Xnommedico = ""

Xnrofactura = 0
Xhcesi = ""
XcedDoc = ""


If t_ensuma.Text <> "" And t_nro.Text <> "" And mfprox.Text <> "__/__/____" Then


   If chtrasl.Value = 1 Then
      If mftras.Text <> "__/__/____" Then
         Data1.Recordset.Edit
         Data1.Recordset("cierre_fec") = Date
         Data1.Recordset("cierre_hora") = Format(Time, "HH:mm")
         Data1.Recordset.Update
        MsgBox "Se cerrará el llamado y no estará más en los pendientes de seguimiento", vbInformation
        data_lla2.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
        data_lla2.Refresh
        If data_lla2.Recordset.RecordCount > 0 Then
           If data_lla2.Recordset("pend") = 4 Then
              data_lla2.Recordset.Edit
              data_lla2.Recordset("pend") = 2
              data_lla2.Recordset("cmt_enproceso") = 2
              data_lla2.Recordset("movilpas") = 2015
              data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
              data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
              data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
              data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
              data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
              data_lla2.Recordset("hor_llega") = Format(Time, "HH:mm")
              data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
              data_lla2.Recordset("hor_rea") = Format(Time, "HH:mm")
              data_lla2.Recordset("diag") = "CMT COVID-19-TRASLADO"
              data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
              If IsNull(data_lla2.Recordset("mes")) = False Then
                 If data_lla2.Recordset("mes") > 30 Then
                    data_lla2.Recordset("mes") = 0
                 End If
              End If
               'data_lla.Recordset("codmed") = txt_codmed.Text
               'data_lla.Recordset("nommed") = dbcbomed.Text
              data_lla2.Recordset.Update
           Else
              If data_lla2.Recordset("pend") <> 2 Then
                  data_lla2.Recordset.Edit
                  data_lla2.Recordset("pend") = 2
                  data_lla2.Recordset("cmt_enproceso") = 2
                  data_lla2.Recordset("movilpas") = 2015
                  data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
                  data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
                  data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("hor_llega") = Format(Time, "HH:mm")
                  data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("hor_rea") = Format(Time, "HH:mm")
                  data_lla2.Recordset("diag") = "COVID-19 DERIVADO A BASE"
                  data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
                  If IsNull(data_lla2.Recordset("mes")) = False Then
                     If data_lla2.Recordset("mes") > 30 Then
                        data_lla2.Recordset("mes") = 0
                     End If
                  End If
                   'data_lla.Recordset("codmed") = txt_codmed.Text
                   'data_lla.Recordset("nommed") = dbcbomed.Text
                  data_lla2.Recordset.Update
                   
                   MsgBox "ATENCION!! El llamado ya NO figura cómo pasado a seguimiento! Se pasará como terminado.", vbInformation
              End If
           End If
        End If
      End If
   End If
   If IsNull(Data1.Recordset("prox_control")) = False Then
      If Format(Data1.Recordset("prox_control"), "dd/mm/yyyy") <> Format(mfprox.Text, "dd/mm/yyyy") Then
         Data1.Recordset.Edit
         Data1.Recordset("prox_control") = Format(mfprox.Text, "dd/mm/yyyy")
         Data1.Recordset("cmt_enproceso") = 2
         Data1.Recordset.Update
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("prox_control") = Format(mfprox.Text, "dd/mm/yyyy")
      Data1.Recordset("cmt_enproceso") = 2
      Data1.Recordset.Update
   End If
'    If t_resultiso.Text = "Negativo" Then
'       Data1.Recordset.Edit
'       Data1.Recordset("cierre_fec") = Date
'       Data1.Recordset("cierre_hora") = Format(Time, "HH:mm")
       
'       Data1.Recordset("prox_control") = Format(Date, "dd/mm/yyyy")
'       Data1.Recordset("cmt_enproceso") = 2
'       Data1.Recordset.Update
'    Else
'       If t_resultiso.Text = "Positivo" Then
       
'       Else
'          If Format(Data1.Recordset("prox_control"), "yyyy/mm/dd") <> Format("01/12/2021", "dd/mm/yyyy") Then
'             Data1.Recordset.Edit
'             Data1.Recordset("prox_control") = Format("31/12/2021", "dd/mm/yyyy")
'             Data1.Recordset("cmt_enproceso") = 2
'             Data1.Recordset.Update
'          End If
'       End If
'    End If
      
   Data2.RecordSource = "select * from seguimiento_covid where id_llamado =" & Val(labllam.Caption) & " and dia =" & t_nro.Text
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      
      Data2.Recordset("trasl") = chtrasl.Value
      If mftras.Text <> "__/__/____" Then
         Data2.Recordset("fecha_trasl") = Format(mftras.Text, "dd/mm/yyyy")
      Else
         If IsNull(Data2.Recordset("fecha_trasl")) = False Then
            Data2.Recordset("fecha_trasl") = Null
         End If
      End If
      
      If Combo3.ListIndex >= 0 Then
         If IsNull(Data2.Recordset("diagnost")) = False Then
            If Data2.Recordset("diagnost") <> Combo3.Text Then
               Data2.Recordset.Edit
               Data2.Recordset("diagnost") = Combo3.Text
               Data2.Recordset.Update
            End If
         End If
      Else
         MsgBox "No ingresó diagnóstico.", vbCritical
      End If
      If IsNull(Data2.Recordset("texto")) = False Then
         If Data2.Recordset("texto") <> t_ensuma.Text Then
            Data2.Recordset.Edit
            Data2.Recordset("texto") = Format(Date, "dd/mm/yyyy") & "-->" & t_ensuma.Text
            Data2.Recordset.Update
         End If
      Else
         Data2.Recordset.Edit
         Data2.Recordset("texto") = Format(Date, "dd/mm/yyyy") & "-->" & t_ensuma.Text
         Data2.Recordset.Update
      End If
      If IsNull(Data2.Recordset("fecha_control")) = False Then
         If Format(Data2.Recordset("fecha_control"), "dd/mm/yyyy") <> Format(mfprox.Text, "dd/mm/yyyy") Then
            Data2.Recordset.Edit
            Data2.Recordset("fecha_control") = Format(mfprox.Text, "dd/mm/yyyy")
            Data2.Recordset.Update
         End If
      End If
   Else
      Data2.Recordset.AddNew
      Data2.Recordset("trasl") = chtrasl.Value
      If mftras.Text <> "__/__/____" Then
         Data2.Recordset("fecha_trasl") = Format(mftras.Text, "dd/mm/yyyy")
      End If
      Data2.Recordset("id_llamado") = Val(labllam.Caption)
      If Combo3.ListIndex >= 0 Then
         Data2.Recordset("diagnost") = Combo3.Text
      End If
      Data2.Recordset("texto") = Format(Date, "dd/mm/yyyy") & "-->" & t_ensuma.Text
      Data2.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
      Data2.Recordset("hora") = Format(Time, "HH:mm")
      Data2.Recordset("dia") = t_nro.Text
      If frm_largador.txt_ced.Text <> "" Then
         Data2.Recordset("ced") = frm_largador.txt_ced.Text
         Data2.Recordset("codced") = frm_largador.t_codced.Text
      End If
      If frm_largador.txt_mat.Text <> "" Then
         Data2.Recordset("matricula") = frm_largador.txt_mat.Text
      End If
      Data2.Recordset("fecha_control") = Format(mfprox.Text, "dd/mm/yyyy")
      Data2.Recordset("nom_usu") = WElusuario
      Data2.Recordset.Update
      MsgBox "Registro grabado correctamente"
    'alta
   End If
   List1.Clear
   Data2.RecordSource = "select * from seguimiento_covid where id_llamado =" & frm_largador.txt_nro.Text & " order by dia"
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      Data2.Recordset.MoveFirst
      Do While Not Data2.Recordset.EOF
         List1.AddItem Data2.Recordset("dia")
         Data2.Recordset.MoveNext
      Loop
   End If
   
   If frm_largador.txt_ced.Text <> "" Then
       If frm_largador.txt_ced.Text > 0 Then
          If frm_largador.txt_mat.Text <> "" Then
             If frm_largador.txt_mat.Text > 0 Then
                Xcrear = 8
             Else
                Xcrear = 0
             End If
          Else
             Xcrear = 0
          End If
       Else
          Xcrear = 0
       End If
   Else
       Xcrear = 0
   End If
   Xhcesi = MsgBox("Desea crear Historia Clínica para este paciente con estos datos?", vbInformation + vbYesNo)
   XcedDoc = InputBox("Ingrese su número de cédula completo (Ejemplo: para CI:1234567-8, ingresar: 12345678)")

   If Xhcesi = vbYes And XcedDoc <> "" Then
      data_conshc.RecordSource = "select * from us where documento ='" & Trim(XcedDoc) & "'"
      data_conshc.Refresh
      
      If Xcrear = 8 And data_conshc.Recordset.RecordCount > 0 Then
         If IsNull(data_conshc.Recordset("us_desc")) = False Then
            Xnromedico = Val(data_conshc.Recordset("us_desc"))
            data_medsapp.RecordSource = "select * from medicos where med_cod =" & Xnromedico
            data_medsapp.Refresh
            If data_medsapp.Recordset.RecordCount > 0 Then
               If IsNull(data_medsapp.Recordset("med_nombre")) = False Then
                  Xnommedico = data_medsapp.Recordset("med_nombre")
               Else
                  Xnromedico = 440
                  Xnommedico = "OTROS MEDICOS"
               End If
            Else
               Xnromedico = 440
               Xnommedico = "OTROS MEDICOS"
            End If
         Else
            Xnromedico = 440
            Xnommedico = "OTROS MEDICOS"
         End If
         
         If Combo1.Text <> "" Then
            If mfContact.Text <> "__/__/____" Then
               XTextoparaHC = "Contacto: " & Combo1.Text & " FECHA:" & Format(mfContact.Text, "dd/mm/yyyy") & vbCrLf
            Else
               XTextoparaHC = "Contacto: " & Combo1.Text
            End If
         End If
         If t_inicio.Text <> "" Then
            If Trim(XTextoparaHC) = "" Then
               If mfsint.Text <> "__/__/____" Then
                  XTextoparaHC = "Inicio de síntomas: " & t_inicio.Text & " FECHA:" & Format(mfsint.Text, "dd/mm/yyyy") & vbCrLf
               Else
                  XTextoparaHC = "Inicio de síntomas: " & t_inicio.Text & vbCrLf
               End If
            Else
               If mfsint.Text <> "__/__/____" Then
                  XTextoparaHC = XTextoparaHC & "Inicio de síntomas: " & t_inicio.Text & " FECHA:" & Format(mfsint.Text, "dd/mm/yyyy") & vbCrLf
               Else
                  XTextoparaHC = XTextoparaHC & "Inicio de síntomas: " & t_inicio.Text & vbCrLf
               End If
            End If
         End If
         If t_sintomas.Text <> "" Then
            If Trim(XTextoparaHC) = "" Then
               XTextoparaHC = "Síntomas: " & t_sintomas.Text & vbCrLf
            Else
               XTextoparaHC = XTextoparaHC & "Síntomas: " & t_sintomas.Text & vbCrLf
            End If
         End If
         If chisosol.Value = 1 Then
            If mfsolici.Text <> "__/__/____" Then
               If Trim(XTextoparaHC) = "" Then
                  XTextoparaHC = "Isopado solicitado día: " & Format(mfsolici.Text, "dd/mm/yyyy") & vbCrLf
               Else
                  XTextoparaHC = XTextoparaHC & "Isopado solicitado día: " & Format(mfsolici.Text, "dd/mm/yyyy") & vbCrLf
               End If
            Else
               If Trim(XTextoparaHC) = "" Then
                  XTextoparaHC = "Isopado solicitado sin fecha" & vbCrLf
               Else
                  XTextoparaHC = XTextoparaHC & "Isopado solicitado sin fecha" & vbCrLf
               End If
            End If
         End If
         
         If chisorea.Value = 1 Then
            If mfreaiso.Text <> "__/__/____" Then
               If t_resultiso.Text <> "" Then
                  If Trim(XTextoparaHC) = "" Then
                     XTextoparaHC = "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & " Resultado: " & t_resultiso.Text & vbCrLf
                  Else
                     XTextoparaHC = XTextoparaHC & "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & " Resultado: " & t_resultiso.Text & vbCrLf
                  End If
               Else
                  If Trim(XTextoparaHC) = "" Then
                     XTextoparaHC = "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & vbCrLf
                  Else
                     XTextoparaHC = XTextoparaHC & "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & vbCrLf
                  End If
               End If
            Else
               If Trim(XTextoparaHC) = "" Then
                  XTextoparaHC = "Isopado solicitado sin fecha" & vbCrLf
               Else
                  XTextoparaHC = XTextoparaHC & "Isopado solicitado sin fecha" & vbCrLf
               End If
            End If
         End If
         If chcomun.Value = 1 Then
            If mfcomun.Text <> "__/__/____" Then
               If Trim(XTextoparaHC) = "" Then
                  XTextoparaHC = "Comunicado a Epidemiología el " & Format(mfcomun.Text, "dd/mm/yyyy") & vbCrLf
               Else
                  XTextoparaHC = XTextoparaHC & "Comunicado a Epidemiología el " & Format(mfcomun.Text, "dd/mm/yyyy") & vbCrLf
               End If
            Else
               If Trim(XTextoparaHC) = "" Then
                  XTextoparaHC = "Comunicado a Epidemiología " & vbCrLf
               Else
                  XTextoparaHC = XTextoparaHC & "Comunicado a Epidemiología " & vbCrLf
               End If
            End If
         End If
         If mf.Text <> "__/__/____" Then
            If Trim(XTextoparaHC) = "" Then
               XTextoparaHC = "FECHA DE ALTA: " & Format(mf.Text, "dd/mm/yyyy")
            Else
               XTextoparaHC = XTextoparaHC & "FECHA DE ALTA: " & Format(mf.Text, "dd/mm/yyyy")
            End If
         End If
         data_hce.RecordSource = "select * from cabezal_hc where cb_mat =" & frm_largador.txt_mat.Text
         data_hce.Refresh
         If data_hce.Recordset.RecordCount > 0 Then
            data_hce.RecordSource = "select * from cabezal_hcdig where mat =" & frm_largador.txt_mat.Text
            data_hce.Refresh
            data_par.Recordset.Edit
            data_par.Recordset("p_hc") = data_par.Recordset("p_hc") + 1
            data_par.Recordset.Update
            data_hce.Recordset.AddNew
            data_hce.Recordset("id") = data_par.Recordset("p_hc")
            data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
            data_hce.Recordset("mat") = frm_largador.txt_mat.Text
            data_hce.Recordset("cednum") = frm_largador.txt_ced.Text
            If frm_largador.t_codced.Text <> "" Then
               data_hce.Recordset("cedtext") = frm_largador.txt_ced.Text & frm_largador.t_codced.Text
               data_hce.Recordset("codced") = frm_largador.t_codced.Text
            Else
               data_hce.Recordset("cedtext") = frm_largador.txt_ced.Text
               data_hce.Recordset("codced") = 0
            End If
            data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
            data_hce.Recordset("hora") = Format(Time, "HH:mm:ss")
            data_hce.Recordset("codigo") = 3
            data_hce.Recordset("tipo_cons") = 9
            data_hce.Recordset("tipo_consd") = "Orientación Telefónica"
            data_hce.Recordset("hc_base") = 18
            data_hce.Recordset("hc_codmed") = data_conshc.Recordset("id")
            data_hce.Recordset("hc_nommed") = data_conshc.Recordset("nombre") & " " & data_conshc.Recordset("apellidos")
            data_hce.Recordset("hc_cpmed") = data_conshc.Recordset("cp")
            If frm_largador.txt_edad.Text <> "" Then
               data_hce.Recordset("hc_naca") = frm_largador.txt_edad.Text
            End If
'                  adohc1.Recordset("hc_nacm") = Xwedm
'                  adohc1.Recordset("hc_nacd") = Xwedd
            data_hce.Recordset.Update

            data_hce.RecordSource = "Select * from hc_mcyotro where id =" & 529
            data_hce.Refresh
            data_hce.Recordset.AddNew
            data_hce.Recordset("id") = data_par.Recordset("p_hc")
            data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
            data_hce.Recordset("hc_mat") = frm_largador.txt_mat.Text
            data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
            data_hce.Recordset("hora") = Format(Time, "HH:mm")
            data_hce.Recordset("hc_mc") = "Orientación telefónica"
            If t_ensuma.Text <> "" Then
               If t_nro.Text <> "" Then
                  data_hce.Recordset("hc_otros") = XTextoparaHC & "Consulta Nro." & t_nro.Text & vbCrLf & "EN SUMA:" & t_ensuma.Text
               Else
                  data_hce.Recordset("hc_otros") = XTextoparaHC & "Consulta Nro.1" & vbCrLf & "EN SUMA:" & t_ensuma.Text
               End If
            Else
               data_hce.Recordset("hc_otros") = XTextoparaHC & "Sin Datos"
            End If
            data_hce.Recordset.Update

            data_hce.RecordSource = "Select * from cli_crmdeudas where nrofact =" & data_par.Recordset("p_hc")
            data_hce.Refresh
            data_hce.Recordset.AddNew
            data_hce.Recordset("id") = data_par.Recordset("p_hc")
            data_hce.Recordset("base") = frm_largador.txt_mat.Text
            data_hce.Recordset("nrofact") = data_par.Recordset("p_hc")
            If Combo3.ListIndex >= 0 Then
               data_hce.Recordset("obs") = Combo3.Text
            Else
               data_hce.Recordset("obs") = "registro de orientación clínica por vía telefónica"
            End If
            data_hce.Recordset("usuario") = "Z719"
            data_hce.Recordset("forma_pago") = 1
            data_hce.Recordset("var1n") = 3
            data_hce.Recordset.Update

            data_hce.RecordSource = "Select * from cabezal_hcdig where id =" & data_par.Recordset("p_hc") & " and mat =" & frm_largador.txt_mat.Text
            data_hce.Refresh
            If data_hce.Recordset.RecordCount > 0 Then
               If IsNull(data_hce.Recordset("hc_fin")) = True Then
                  data_hce.Recordset.Edit
                  data_hce.Recordset("hc_fin") = 5
                  data_hce.Recordset.Update
               End If
            End If
    
            data_lin.RecordSource = "select * from param_gral"
            data_lin.Refresh
            Xnrofactura = data_lin.Recordset("p_linmmdd") + 1
            data_lin.Recordset.Edit
            data_lin.Recordset("p_linmmdd") = data_lin.Recordset("p_linmmdd") + 1
            data_lin.Recordset.Update
            '   labcodest.Caption = 10018

            data_lin.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
            data_lin.Refresh
            data_lin.Recordset.AddNew
            data_lin.Recordset("linea") = 1
            data_lin.Recordset("factura") = Xnrofactura
            data_lin.Recordset("tipo") = "REG."
            data_lin.Recordset("realizada") = Date
            data_lin.Recordset("fecha") = Date
            data_lin.Recordset("cod_cli") = frm_largador.txt_mat.Text
            data_lin.Recordset("nom_cli") = Mid(frm_largador.txt_nomb.Text, 1, 30)
            data_lin.Recordset("convenio") = frm_largador.txt_cat.Text
            data_lin.Recordset("cod_prod") = 10018
            data_lin.Recordset("nom_prod") = "CONSULTA TELEFONICA"
            data_lin.Recordset("operador") = WElusuario
            data_lin.Recordset("hora") = Format(Time, "HH:mm")
            data_lin.Recordset("imp_timbre") = 0
            data_lin.Recordset("tot_lin") = 0
            data_lin.Recordset("valor_iva") = 0
            data_lin.Recordset("base") = frm_menu.data_parse.Recordset("base")
            data_lin.Recordset("nom_med_a") = Mid(Xnommedico, 1, 40)
            data_lin.Recordset("pre_civa") = 0
            data_lin.Recordset("reg_cab") = 99
            If frm_largador.txt_ced.Text <> "" Then
               data_lin.Recordset("ced_socio") = frm_largador.txt_ced.Text
            End If
            If frm_largador.t_codced.Text <> "" Then
               data_lin.Recordset("fact") = frm_largador.t_codced.Text
            End If
            data_lin.Recordset("moneda") = "A"
            data_lin.Recordset("nro_flia") = 1
            data_lin.Recordset("nom_flia") = "MEDICINA GENERAL"
            data_lin.Recordset("rub_cont") = frm_menu.data_parse.Recordset("srvcnt")
            data_lin.Recordset("arancel") = 0
            data_lin.Recordset("nro_med_a") = Xnromedico
            data_lin.Recordset("precio_est") = 0
            data_lin.Recordset("imp_iva") = 0
            data_lin.Recordset("tipo_mov") = "2"
            data_lin.Recordset("pendiente") = "X"
            data_lin.Recordset.Update
    
            MsgBox "HC creada correctamente", vbInformation
         Else
            If frm_largador.txt_mat.Text <> "" Then
               If frm_largador.txt_mat.Text <> 0 Then
                  
                  Command2_Click
                  
                  data_conshc.RecordSource = "select * from us where documento ='" & Trim(XcedDoc) & "'"
                  data_conshc.Refresh
                  
                  If data_conshc.Recordset.RecordCount > 0 Then
                     If Combo1.Text <> "" Then
                        If mfContact.Text <> "__/__/____" Then
                           XTextoparaHC = "Contacto: " & Combo1.Text & " FECHA:" & Format(mfContact.Text, "dd/mm/yyyy") & vbCrLf
                        Else
                           XTextoparaHC = "Contacto: " & Combo1.Text
                        End If
                     End If
                     If t_inicio.Text <> "" Then
                        If Trim(XTextoparaHC) = "" Then
                           If mfsint.Text <> "__/__/____" Then
                              XTextoparaHC = "Inicio de síntomas: " & t_inicio.Text & " FECHA:" & Format(mfsint.Text, "dd/mm/yyyy") & vbCrLf
                           Else
                              XTextoparaHC = "Inicio de síntomas: " & t_inicio.Text & vbCrLf
                           End If
                        Else
                           If mfsint.Text <> "__/__/____" Then
                              XTextoparaHC = XTextoparaHC & "Inicio de síntomas: " & t_inicio.Text & " FECHA:" & Format(mfsint.Text, "dd/mm/yyyy") & vbCrLf
                           Else
                              XTextoparaHC = XTextoparaHC & "Inicio de síntomas: " & t_inicio.Text & vbCrLf
                           End If
                        End If
                     End If
                     If t_sintomas.Text <> "" Then
                        If Trim(XTextoparaHC) = "" Then
                           XTextoparaHC = "Síntomas: " & t_sintomas.Text & vbCrLf
                        Else
                           XTextoparaHC = XTextoparaHC & "Síntomas: " & t_sintomas.Text & vbCrLf
                        End If
                     End If
                     If chisosol.Value = 1 Then
                        If mfsolici.Text <> "__/__/____" Then
                           If Trim(XTextoparaHC) = "" Then
                              XTextoparaHC = "Isopado solicitado día: " & Format(mfsolici.Text, "dd/mm/yyyy") & vbCrLf
                           Else
                              XTextoparaHC = XTextoparaHC & "Isopado solicitado día: " & Format(mfsolici.Text, "dd/mm/yyyy") & vbCrLf
                           End If
                        Else
                           If Trim(XTextoparaHC) = "" Then
                              XTextoparaHC = "Isopado solicitado sin fecha" & vbCrLf
                           Else
                              XTextoparaHC = XTextoparaHC & "Isopado solicitado sin fecha" & vbCrLf
                           End If
                        End If
                     End If
                     
                     If chisorea.Value = 1 Then
                        If mfreaiso.Text <> "__/__/____" Then
                           If t_resultiso.Text <> "" Then
                              If Trim(XTextoparaHC) = "" Then
                                 XTextoparaHC = "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & " Resultado: " & t_resultiso.Text & vbCrLf
                              Else
                                 XTextoparaHC = XTextoparaHC & "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & " Resultado: " & t_resultiso.Text & vbCrLf
                              End If
                           Else
                              If Trim(XTextoparaHC) = "" Then
                                 XTextoparaHC = "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & vbCrLf
                              Else
                                 XTextoparaHC = XTextoparaHC & "Isopado realizado día: " & Format(mfreaiso.Text, "dd/mm/yyyy") & vbCrLf
                              End If
                           End If
                        Else
                           If Trim(XTextoparaHC) = "" Then
                              XTextoparaHC = "Isopado solicitado sin fecha" & vbCrLf
                           Else
                              XTextoparaHC = XTextoparaHC & "Isopado solicitado sin fecha" & vbCrLf
                           End If
                        End If
                     End If
                     If chcomun.Value = 1 Then
                        If mfcomun.Text <> "__/__/____" Then
                           If Trim(XTextoparaHC) = "" Then
                              XTextoparaHC = "Comunicado a Epidemiología el " & Format(mfcomun.Text, "dd/mm/yyyy") & vbCrLf
                           Else
                              XTextoparaHC = XTextoparaHC & "Comunicado a Epidemiología el " & Format(mfcomun.Text, "dd/mm/yyyy") & vbCrLf
                           End If
                        Else
                           If Trim(XTextoparaHC) = "" Then
                              XTextoparaHC = "Comunicado a Epidemiología " & vbCrLf
                           Else
                              XTextoparaHC = XTextoparaHC & "Comunicado a Epidemiología " & vbCrLf
                           End If
                        End If
                     End If
                     If mf.Text <> "__/__/____" Then
                        If Trim(XTextoparaHC) = "" Then
                           XTextoparaHC = "FECHA DE ALTA: " & Format(mf.Text, "dd/mm/yyyy")
                        Else
                           XTextoparaHC = XTextoparaHC & "FECHA DE ALTA: " & Format(mf.Text, "dd/mm/yyyy")
                        End If
                     End If
                     data_hce.RecordSource = "select * from cabezal_hc where cb_mat =" & frm_largador.txt_mat.Text
                     data_hce.Refresh
                     If data_hce.Recordset.RecordCount > 0 Then
                        data_hce.RecordSource = "select * from cabezal_hcdig where mat =" & frm_largador.txt_mat.Text
                        data_hce.Refresh
                        data_par.Recordset.Edit
                        data_par.Recordset("p_hc") = data_par.Recordset("p_hc") + 1
                        data_par.Recordset.Update
                        data_hce.Recordset.AddNew
                        data_hce.Recordset("id") = data_par.Recordset("p_hc")
                        data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
                        data_hce.Recordset("mat") = frm_largador.txt_mat.Text
                        data_hce.Recordset("cednum") = frm_largador.txt_ced.Text
                        If frm_largador.t_codced.Text <> "" Then
                           data_hce.Recordset("cedtext") = frm_largador.txt_ced.Text & frm_largador.t_codced.Text
                           data_hce.Recordset("codced") = frm_largador.t_codced.Text
                        Else
                           data_hce.Recordset("cedtext") = frm_largador.txt_ced.Text
                           data_hce.Recordset("codced") = 0
                        End If
                        data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
                        data_hce.Recordset("hora") = Format(Time, "HH:mm:ss")
                        data_hce.Recordset("codigo") = 3
                        data_hce.Recordset("tipo_cons") = 9
                        data_hce.Recordset("tipo_consd") = "Orientación Telefónica"
                        data_hce.Recordset("hc_base") = frm_menu.data_parse.Recordset("base")
                        data_hce.Recordset("hc_codmed") = data_conshc.Recordset("id")
                        data_hce.Recordset("hc_nommed") = data_conshc.Recordset("nombre") & " " & data_conshc.Recordset("apellidos")
                        data_hce.Recordset("hc_cpmed") = data_conshc.Recordset("cp")
                        If frm_largador.txt_edad.Text <> "" Then
                           data_hce.Recordset("hc_naca") = frm_largador.txt_edad.Text
                        End If
            '                  adohc1.Recordset("hc_nacm") = Xwedm
            '                  adohc1.Recordset("hc_nacd") = Xwedd
                        data_hce.Recordset.Update
            
                        data_hce.RecordSource = "Select * from hc_mcyotro where id =" & 529
                        data_hce.Refresh
                        data_hce.Recordset.AddNew
                        data_hce.Recordset("id") = data_par.Recordset("p_hc")
                        data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
                        data_hce.Recordset("hc_mat") = frm_largador.txt_mat.Text
                        data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
                        data_hce.Recordset("hora") = Format(Time, "HH:mm")
                        data_hce.Recordset("hc_mc") = "Orientación telefónica"
                        If t_ensuma.Text <> "" Then
                           If t_nro.Text <> "" Then
                              data_hce.Recordset("hc_otros") = XTextoparaHC & "Consulta Nro." & t_nro.Text & vbCrLf & "EN SUMA:" & t_ensuma.Text
                           Else
                              data_hce.Recordset("hc_otros") = XTextoparaHC & "Consulta Nro.1" & vbCrLf & "EN SUMA:" & t_ensuma.Text
                           End If
                        Else
                           data_hce.Recordset("hc_otros") = XTextoparaHC & "Sin Datos"
                        End If
                        data_hce.Recordset.Update
            
                        data_hce.RecordSource = "Select * from cli_crmdeudas where nrofact =" & data_par.Recordset("p_hc")
                        data_hce.Refresh
                        data_hce.Recordset.AddNew
                        data_hce.Recordset("id") = data_par.Recordset("p_hc")
                        data_hce.Recordset("base") = frm_largador.txt_mat.Text
                        data_hce.Recordset("nrofact") = data_par.Recordset("p_hc")
                        If Combo3.ListIndex >= 0 Then
                           data_hce.Recordset("obs") = Combo3.Text
                        Else
                           data_hce.Recordset("obs") = "registro de orientación clínica por vía telefónica"
                        End If
                        data_hce.Recordset("usuario") = "Z719"
                        data_hce.Recordset("forma_pago") = 1
                        data_hce.Recordset("var1n") = 3
                        data_hce.Recordset.Update
            
                        data_hce.RecordSource = "Select * from cabezal_hcdig where id =" & data_par.Recordset("p_hc") & " and mat =" & frm_largador.txt_mat.Text
                        data_hce.Refresh
                        If data_hce.Recordset.RecordCount > 0 Then
                           If IsNull(data_hce.Recordset("hc_fin")) = True Then
                              data_hce.Recordset.Edit
                              data_hce.Recordset("hc_fin") = 5
                              data_hce.Recordset.Update
                           End If
                        End If
                        data_lin.RecordSource = "select * from param_gral"
                        data_lin.Refresh
                        Xnrofactura = data_lin.Recordset("p_linmmdd") + 1
                        data_lin.Recordset.Edit
                        data_lin.Recordset("p_linmmdd") = data_lin.Recordset("p_linmmdd") + 1
                        data_lin.Recordset.Update
                        '   labcodest.Caption = 10018

                        data_lin.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
                        data_lin.Refresh
                        data_lin.Recordset.AddNew
                        data_lin.Recordset("linea") = 1
                        data_lin.Recordset("factura") = Xnrofactura
                        data_lin.Recordset("tipo") = "REG."
                        data_lin.Recordset("realizada") = Date
                        data_lin.Recordset("fecha") = Date
                        data_lin.Recordset("cod_cli") = frm_largador.txt_mat.Text
                        data_lin.Recordset("nom_cli") = Mid(frm_largador.txt_nomb.Text, 1, 30)
                        data_lin.Recordset("convenio") = frm_largador.txt_cat.Text
                        data_lin.Recordset("cod_prod") = 10018
                        data_lin.Recordset("nom_prod") = "CONSULTA TELEFONICA"
                        data_lin.Recordset("operador") = WElusuario
                        data_lin.Recordset("hora") = Format(Time, "HH:mm")
                        data_lin.Recordset("imp_timbre") = 0
                        data_lin.Recordset("tot_lin") = 0
                        data_lin.Recordset("valor_iva") = 0
                        data_lin.Recordset("base") = frm_menu.data_parse.Recordset("base")
                        data_lin.Recordset("nom_med_a") = Mid(Xnommedico, 1, 40)
                             
                        data_lin.Recordset("pre_civa") = 0
                        data_lin.Recordset("reg_cab") = 99
                        If frm_largador.txt_ced.Text <> "" Then
                           data_lin.Recordset("ced_socio") = frm_largador.txt_ced.Text
                        End If
                        If frm_largador.t_codced.Text <> "" Then
                           data_lin.Recordset("fact") = frm_largador.t_codced.Text
                        End If
                        data_lin.Recordset("moneda") = "A"
                        data_lin.Recordset("nro_flia") = 1
                        data_lin.Recordset("nom_flia") = "MEDICINA GENERAL"
                        data_lin.Recordset("rub_cont") = frm_menu.data_parse.Recordset("srvcnt")
                        data_lin.Recordset("arancel") = 0
                        data_lin.Recordset("nro_med_a") = Xnromedico
                        data_lin.Recordset("precio_est") = 0
                        data_lin.Recordset("imp_iva") = 0
                        data_lin.Recordset("tipo_mov") = "2"
                        data_lin.Recordset("pendiente") = "X"
                        data_lin.Recordset.Update
                                
                        MsgBox "HC creada correctamente", vbInformation
                     Else
                        MsgBox "No se pudo crear, verifique datos de cédula y matrícula del paciente", vbInformation
                     End If
                  Else
                      MsgBox "No se puede crear porque no se encuentra CI del médico", vbInformation
                  End If
                  
               Else
                  MsgBox "No se encuentra registro de ficha en HCE, deberá crearla manualmente.", vbInformation
               End If
            Else
               MsgBox "No se encuentra registro de ficha en HCE, deberá crearla manualmente.", vbInformation
            End If
         End If
      Else
         MsgBox "No se encuentra número de cédula del médico o faltan datos del paciente, no se puede crear HCE", vbInformation
      End If
   End If

Else
   MsgBox "No hay datos a grabar"
   
End If

Exit Sub

Quepasaalg:
            If Err.Number = 3155 Then
               MsgBox "Error: " & Err.Number & " " & Err.Description
            Else
               MsgBox "Error: " & Err.Number & " " & Err.Description
            End If
            
End Sub

Private Sub Command1_Click()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("cmt_enproceso")) = False Then
      Data1.Recordset.Edit
      Data1.Recordset("cmt_enproceso") = Null
      Data1.Recordset("cmt_usproc") = Null
      Data1.Recordset.Update
      MsgBox "El llamado ha quedado habilitado para otro médico", vbInformation
   End If
End If
Unload Me

End Sub

Private Sub Command2_Click()
   adoaltas.Connect = "ODBC;DSN=sappnew;"
   adoaltas.RecordSource = "Select * from cabezal_hc where cb_mat =" & frm_largador.txt_mat.Text
   adoaltas.Refresh
   If adoaltas.Recordset.RecordCount > 0 Then
   Else
     data_par.Recordset.Edit
     data_par.Recordset("p_hc") = data_par.Recordset("p_hc") + 1
     data_par.Recordset.Update
      
     adoaltas.Recordset.AddNew
     adoaltas.Recordset("id") = data_par.Recordset("p_hc")
     adoaltas.Recordset("cb_mat") = frm_largador.txt_mat.Text
     adoaltas.Recordset("cb_mattext") = Trim(str(frm_largador.txt_mat.Text))
     If frm_largador.txt_ced.Text <> "" And frm_largador.t_codced.Text <> "" Then
        adoaltas.Recordset("cb_ced") = frm_largador.txt_ced.Text & frm_largador.t_codced.Text
      Else
         adoaltas.Recordset("cb_ced") = "XX"
      End If
      adoaltas.Recordset("cb_tipdoc") = "DNI"
      adoaltas.Recordset("cb_paisdoc") = "Uruguay"
      adoaltas.Recordset("cb_codpaisced") = "UY"
      If frm_largador.txt_nomb.Text <> "" Then
         adoaltas.Recordset("cb_nom1") = Mid(frm_largador.txt_nomb.Text, 11, 40)
         adoaltas.Recordset("cb_nom2") = Mid(frm_largador.txt_nomb.Text, 11, 40)
         adoaltas.Recordset("cb_ape1") = Mid(frm_largador.txt_nomb.Text, 1, 40)
         adoaltas.Recordset("cb_ape2") = Mid(frm_largador.txt_nomb.Text, 1, 40)
      End If
      adoaltas.Recordset("cb_estado") = 1
      If frm_largador.txt_cat.Text <> "" Then
         adoaltas.Recordset("cb_codconv") = frm_largador.txt_cat.Text
      End If
      If frm_largador.Combo3.Text = "MASC" Then
         adoaltas.Recordset("cb_sexo") = "M"
      Else
         adoaltas.Recordset("cb_sexo") = "F"
      End If
      adoaltas.Recordset("cb_indmult") = 0
      adoaltas.Recordset("cb_indmultc") = 0
      adoaltas.Recordset("cb_indmultd") = "F"
      adoaltas.Recordset("cb_indfall") = 0
      adoaltas.Recordset("cb_tipoviad") = "CALLE"
      If frm_largador.txt_direc.Text <> "" Then
         adoaltas.Recordset("cb_nomvia") = Mid(frm_largador.txt_direc.Text, 1, 20)
      End If
      adoaltas.Recordset("cb_nropta") = "S/N"
      adoaltas.Recordset("cb_dptod") = "CANELONES"
      adoaltas.Recordset("cb_dptoc") = "3"
      adoaltas.Recordset("cb_paisd") = "Uruguay"
      adoaltas.Recordset("cb_paisc") = "UY"
      adoaltas.Recordset.Update
      adoaltas.Refresh
   End If
   
End Sub

Private Sub Form_Load()

data_hce.Connect = "odbc;dsn=sappnew;"

data_lla2.Connect = "odbc;dsn=sappnew;"
data_lin.Connect = "odbc;dsn=sappnew;"
data_medsapp.Connect = "odbc;dsn=sappnew;"
data_consultar.Connect = "odbc;dsn=sappnew;"

data_conshc.Connect = "odbc;dsn=sappnew;"
Data2.Connect = "odbc;dsn=sappnew;"

data_par.Connect = "odbc;dsn=sappnew;"
data_par.RecordSource = "param_gral"
data_par.Refresh

If frm_largador.chcovid.Value = 1 Then
   mf.Enabled = True
   mh.Enabled = True
Else
   mf.Enabled = False
   mh.Enabled = False
End If
   Data1.Connect = "odbc;dsn=sappnew;"
   Data1.RecordSource = "select * from llamado where nrolla =" & frm_largador.txt_nro.Text
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      If IsNull(Data1.Recordset("nombre")) = False Then
         labnom.Caption = Data1.Recordset("nombre")
      End If
      labllam.Caption = Data1.Recordset("nrolla")
      If IsNull(Data1.Recordset("viaje")) = False Then
         chviaje.Value = Data1.Recordset("viaje")
      Else
         chviaje.Value = 0
      End If
      If IsNull(Data1.Recordset("contacto")) = False Then
         Combo1.Text = Data1.Recordset("contacto")
      Else
         Combo1.Text = ""
      End If
      If IsNull(Data1.Recordset("fec_contac")) = False Then
         mfContact.Text = Format(Data1.Recordset("fec_contac"), "dd/mm/yyyy")
      Else
         mfContact.Text = "__/__/____"
      End If
      If IsNull(Data1.Recordset("fec_sint")) = False Then
         mfsint.Text = Format(Data1.Recordset("fec_sint"), "dd/mm/yyyy")
      Else
         mfsint.Text = "__/__/____"
      End If
      If IsNull(Data1.Recordset("grupo_covid")) = False Then
         Combo2.Text = Data1.Recordset("grupo_covid")
      Else
         Combo2.Text = ""
      End If
      If IsNull(Data1.Recordset("cuarent_ant")) = False Then
         chestuvo.Value = Data1.Recordset("cuarent_ant")
      Else
         chestuvo.Value = 0
      End If
      If IsNull(Data1.Recordset("fiebre")) = False Then
         chf.Value = Data1.Recordset("fiebre")
      Else
         chf.Value = 0
      End If
      If IsNull(Data1.Recordset("tos")) = False Then
         cht.Value = Data1.Recordset("tos")
      Else
         cht.Value = 0
      End If
      If IsNull(Data1.Recordset("resfrio")) = False Then
         chr.Value = Data1.Recordset("resfrio")
      Else
         chr.Value = 0
      End If
      If IsNull(Data1.Recordset("diarrea")) = False Then
         chd.Value = Data1.Recordset("diarrea")
      Else
         chd.Value = 0
      End If
      If IsNull(Data1.Recordset("insuf")) = False Then
         chi.Value = Data1.Recordset("insuf")
      Else
         chi.Value = 0
      End If
      If IsNull(Data1.Recordset("isopa_sol")) = False Then
         chisosol.Value = Data1.Recordset("isopa_sol")
      Else
         chisosol.Value = 0
      End If
      If IsNull(Data1.Recordset("isopa_rea")) = False Then
         chisorea.Value = Data1.Recordset("isopa_rea")
      Else
         chisorea.Value = 0
      End If
      If IsNull(Data1.Recordset("isopa_fecsol")) = False Then
         mfsolici.Text = Format(Data1.Recordset("isopa_fecsol"), "dd/mm/yyyy")
      Else
         mfsolici.Text = "__/__/____"
      End If
      If IsNull(Data1.Recordset("isopa_fecrea")) = False Then
         mfreaiso.Text = Format(Data1.Recordset("isopa_fecrea"), "dd/mm/yyyy")
      Else
         mfreaiso.Text = "__/__/____"
      End If
      If IsNull(Data1.Recordset("isopa_result")) = False Then
         t_resultiso.Text = Data1.Recordset("isopa_result")
      Else
         t_resultiso.Text = ""
      End If
      If IsNull(Data1.Recordset("fec_comunica")) = False Then
         mfcomun.Text = Format(Data1.Recordset("fec_comunica"), "dd/mm/yyyy")
      Else
         mfcomun.Text = "__/__/____"
      End If
'      If IsNull(Data1.Recordset("fecisosol2")) = False Then
'         mfsolici2.Text = Format(Data1.Recordset("fecisosol2"), "dd/mm/yyyy")
'      Else
'         mfsolici2.Text = "__/__/____"
'      End If
'      If IsNull(Data1.Recordset("fecisorea2")) = False Then
'         mfreaiso2.Text = Format(Data1.Recordset("fecisorea2"), "dd/mm/yyyy")
'      Else
'         mfreaiso2.Text = "__/__/____"
'      End If
'      If IsNull(Data1.Recordset("resuliso2")) = False Then
'         t_result2.Text = Data1.Recordset("resuliso2")
'      Else
'         t_result2.Text = ""
'      End If
              
      If IsNull(Data1.Recordset("inicio_sint")) = False Then
         t_inicio.Text = Data1.Recordset("inicio_sint")
      Else
         t_inicio.Text = ""
      End If
      If IsNull(Data1.Recordset("sintomas")) = False Then
         t_sintomas.Text = Data1.Recordset("sintomas")
      Else
         t_sintomas.Text = ""
      End If
      If IsNull(Data1.Recordset("comunic_epi")) = False Then
         chcomun.Value = Data1.Recordset("comunic_epi")
      Else
         chcomun.Value = 0
      End If
      If IsNull(Data1.Recordset("ctrol_telef")) = False Then
         chctel.Value = Data1.Recordset("ctrol_telef")
      Else
         chctel.Value = 0
      End If
      If IsNull(Data1.Recordset("ctrol_medic")) = False Then
         chcmed.Value = Data1.Recordset("ctrol_medic")
      Else
         chcmed.Value = 0
      End If
      If IsNull(Data1.Recordset("cierre_fec")) = False Then
         mf.Text = Format(Data1.Recordset("cierre_fec"), "dd/mm/yyyy")
      Else
         mf.Text = "__/__/____"
      End If
      If IsNull(Data1.Recordset("cierre_hora")) = False Then
         mh.Text = Format(Data1.Recordset("cierre_hora"), "HH:mm")
      Else
         mh.Text = "__:__"
      End If
      
      List1.Clear
      Data2.RecordSource = "select * from seguimiento_covid where id_llamado =" & frm_largador.txt_nro.Text & " order by dia"
      Data2.Refresh
      If Data2.Recordset.RecordCount > 0 Then
         Data2.Recordset.MoveFirst
         Do While Not Data2.Recordset.EOF
            List1.AddItem Data2.Recordset("dia")
            Data2.Recordset.MoveNext
         Loop
      Else
         If t_resultiso.Text = "Positivo" Then
            mfprox.Enabled = True
            mfprox.Text = "__/__/____"
         Else
            If IsNull(Data1.Recordset("prox_control")) = False Then
               mfprox.Text = Data1.Recordset("prox_control")
            Else
               mfprox.Text = "__/__/____"
            End If
            mfprox.Enabled = True
        End If
      End If
      List2.Clear
      data_consultar.RecordSource = "select * from seguimiento_tests where seguim_id =" & Data1.Recordset("nrolla") & " and solicita in (1) order by fecha"
      data_consultar.Refresh
      If data_consultar.Recordset.RecordCount > 0 Then
         data_consultar.Recordset.MoveFirst
         Do While Not data_consultar.Recordset.EOF
            List2.AddItem Format(data_consultar.Recordset("fecha"), "dd/mm/yyyy") & "--->" & data_consultar.Recordset("tipo")
            data_consultar.Recordset.MoveNext
         Loop
      End If
      List3.Clear
      data_consultar.RecordSource = "select * from seguimiento_tests where seguim_id =" & Data1.Recordset("nrolla") & " and realiza in (1) order by fecha"
      data_consultar.Refresh
      If data_consultar.Recordset.RecordCount > 0 Then
         data_consultar.Recordset.MoveFirst
         Do While Not data_consultar.Recordset.EOF
            If IsNull(data_consultar.Recordset("resultado")) = False Then
               List3.AddItem Format(data_consultar.Recordset("fecha"), "dd/mm/yyyy") & "--->" & data_consultar.Recordset("tipo") & " -->" & data_consultar.Recordset("resultado")
            Else
               List3.AddItem Format(data_consultar.Recordset("fecha"), "dd/mm/yyyy") & "--->" & data_consultar.Recordset("tipo") & " -->Sin resultado"
            End If
            data_consultar.Recordset.MoveNext
         Loop
      End If
      
   
   Else
      MsgBox "No se encuentra llamado, verifique si grabó. Cierre la ventana de información COVID y vuelva a ingresar.", vbCritical
      
   End If



End Sub

Private Sub List1_DblClick()
t_nro.Text = ""
t_ensuma.Text = ""
t_nro.Text = ""
If List1.ListCount > 0 Then
   Data2.RecordSource = "select * from seguimiento_covid where id_llamado =" & Val(labllam.Caption) & " and dia =" & Val(List1.List(List1.ListIndex))
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      t_nro.Text = List1.List(List1.ListIndex)
      t_ensuma.Text = "Usuario: " & Data2.Recordset("nom_usu") & "-->" & Data2.Recordset("texto") & " Próximo Control: " & Format(Data2.Recordset("fecha_control"), "dd/mm/yyyy")
      If IsNull(Data2.Recordset("diagnost")) = False Then
         Combo3.Text = Data2.Recordset("diagnost")
      Else
         Combo3.Text = ""
      End If
   End If
End If

End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub mf_GotFocus()
mf.Text = Format(Date, "dd/mm/yyyy")
mh.Text = Format(Time, "HH:mm")

End Sub

Private Sub mfprox_LostFocus()
Dim DifDias As Long

If mfprox.Text <> "__/__/____" Then
   DifDias = DateDiff("d", Date, mfprox.Text)
   
   If Format(mfprox.Text, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
      MsgBox "La fecha de próximo control es menor a la fecha actual, verifique!", vbCritical
      mfprox.SetFocus
   Else
      If DifDias > 15 Then
         MsgBox "La fecha de próximo control no puede ser por más de 15 días", vbCritical
         mfprox.SetFocus
      End If
   End If
End If

End Sub

Private Sub mftras_GotFocus()
mftras.Text = Date

End Sub

Private Sub t_resultiso_LostFocus()
If t_resultiso.Text = "Negativo" Or t_resultiso.Text = "Positivo" Then
Else
   t_resultiso.Text = ""
   MsgBox "Texto incorrecto"
End If

End Sub
