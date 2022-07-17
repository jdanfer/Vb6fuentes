VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_metas 
   Caption         =   "Registro de METAS"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11400
   Icon            =   "frm_metas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11400
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_cedmed 
      Height          =   375
      Left            =   9720
      TabIndex        =   98
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid flex1 
      Height          =   1935
      Left            =   240
      TabIndex        =   96
      Top             =   6360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3413
      _Version        =   393216
      BackColorBkg    =   16744576
      FocusRect       =   2
      SelectionMode   =   1
   End
   Begin VB.TextBox t_cedbusca 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   95
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9600
      TabIndex        =   88
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de la meta"
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
      Height          =   2175
      Left            =   240
      TabIndex        =   72
      Top             =   2280
      Visible         =   0   'False
      Width           =   10935
      Begin VB.TextBox t_proxcon3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   87
         Top             =   1410
         Width           =   2415
      End
      Begin VB.ComboBox cboregadul 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":058A
         Left            =   3000
         List            =   "frm_metas.frx":0594
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox t_feca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8040
         MaxLength       =   50
         TabIndex        =   83
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cboscree 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":05A0
         Left            =   5280
         List            =   "frm_metas.frx":05AA
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cbosia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":05B6
         Left            =   1680
         List            =   "frm_metas.frx":05C0
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox t_apemed 
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
         Left            =   8040
         MaxLength       =   50
         TabIndex        =   77
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox t_nommed 
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
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   76
         Top             =   360
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mfctradult 
         Height          =   375
         Left            =   1680
         TabIndex        =   74
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label labcedmed 
         BackColor       =   &H00FF8080&
         Height          =   255
         Left            =   3720
         TabIndex        =   89
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Próxima Consulta:"
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
         Left            =   5400
         TabIndex        =   86
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hoja Registro Adulto Mayor:"
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
         TabIndex        =   84
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecatest:"
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
         Left            =   6720
         TabIndex        =   82
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hoja Screening:"
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
         Left            =   3720
         TabIndex        =   80
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hoja SIA:"
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
         TabIndex        =   78
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MÉDICO REFERENCIA:"
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
         Height          =   375
         Left            =   3720
         TabIndex        =   75
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHA CTROL:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de la meta"
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
      Height          =   2535
      Left            =   360
      TabIndex        =   53
      Top             =   2160
      Visible         =   0   'False
      Width           =   10695
      Begin MSMask.MaskEdBox mfanti 
         Height          =   375
         Left            =   3000
         TabIndex        =   108
         Top             =   1920
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
      Begin VB.TextBox t_semgon 
         Height          =   285
         Left            =   6240
         TabIndex        =   107
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox t_nrocons 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   105
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cbomamo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":05CC
         Left            =   4080
         List            =   "frm_metas.frx":05D6
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cbopap 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":05E2
         Left            =   5520
         List            =   "frm_metas.frx":05EC
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox mfvdrl 
         Height          =   375
         Left            =   8880
         TabIndex        =   99
         Top             =   960
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
      Begin VB.ComboBox cbohcpb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":05F8
         Left            =   1680
         List            =   "frm_metas.frx":0602
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cboprot 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":060E
         Left            =   9360
         List            =   "frm_metas.frx":0618
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cbovdrl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":0624
         Left            =   7920
         List            =   "frm_metas.frx":062E
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox t_cedrn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         MaxLength       =   45
         TabIndex        =   66
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox t_anticon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   64
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox cboodont 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":063A
         Left            =   3960
         List            =   "frm_metas.frx":0644
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboobste 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_metas.frx":0650
         Left            =   1680
         List            =   "frm_metas.frx":065A
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox t_lochc 
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
         Left            =   8640
         MaxLength       =   45
         TabIndex        =   58
         Top             =   360
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mfctremb 
         Height          =   375
         Left            =   1680
         TabIndex        =   55
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label40 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sem.Gonorrea:"
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
         Left            =   4680
         TabIndex        =   106
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NRO.CONS:"
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
         Height          =   375
         Left            =   3480
         TabIndex        =   104
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mamografía:"
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
         Left            =   2880
         TabIndex        =   102
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PAP"
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
         Left            =   4920
         TabIndex        =   100
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consentimiento informado"
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
         Left            =   6960
         TabIndex        =   69
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HCPB al Sistema?"
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
         TabIndex        =   68
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cédula R.N:"
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
         Left            =   7800
         TabIndex        =   65
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Anticoncep."
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
         TabIndex        =   63
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VDRL y HIV:"
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
         Left            =   6600
         TabIndex        =   62
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ctrol.Odont:"
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
         Left            =   2880
         TabIndex        =   60
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Localización de la HC:"
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
         Height          =   375
         Left            =   5880
         TabIndex        =   57
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ctrol.Obstétrico:"
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
         TabIndex        =   56
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHA CTROL:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8760
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Caption         =   "Adodc1"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   31
      Top             =   5160
      Width           =   11055
      Begin VB.CommandButton b_imp 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         Picture         =   "frm_metas.frx":0666
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Informes del sistema"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton b_elim 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         Picture         =   "frm_metas.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Eliminar registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton b_canc 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         Picture         =   "frm_metas.frx":117A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cancelar acciòn"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton b_grab 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         Picture         =   "frm_metas.frx":1704
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Graba los datos"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton b_mod 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         Picture         =   "frm_metas.frx":1C8E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Modificar registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton b_alta 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         Picture         =   "frm_metas.frx":2218
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Crear nuevo registro"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame framnino1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de la META"
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
      Height          =   3135
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox t_dd 
         Height          =   285
         Left            =   6480
         TabIndex        =   92
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox t_mm 
         Height          =   285
         Left            =   5880
         TabIndex        =   91
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox t_aa 
         Height          =   285
         Left            =   5280
         TabIndex        =   90
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox t_ctroft 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         MaxLength       =   80
         TabIndex        =   52
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox t_fdesar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         MaxLength       =   80
         TabIndex        =   50
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox t_odont 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   48
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox t_hemog 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         MaxLength       =   80
         TabIndex        =   46
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox t_proxctr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   44
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox t_obs 
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
         Left            =   1920
         MaxLength       =   120
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   2640
         Width           =   8895
      End
      Begin VB.TextBox t_medic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         MaxLength       =   80
         TabIndex        =   29
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox t_eco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   27
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox t_vacun 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8040
         MaxLength       =   80
         TabIndex        =   25
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox t_lacta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   23
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox t_peso 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9720
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mfctr 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label35 
         BackColor       =   &H00FF8080&
         Caption         =   "Años   Meses   Días"
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
         Left            =   5280
         TabIndex        =   93
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Control Oftalmológico:"
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
         Left            =   4920
         TabIndex        =   51
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Form.Ctrol.del Desarrollo:"
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
         Left            =   4920
         TabIndex        =   49
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CTROL.ODONT:"
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
         TabIndex        =   47
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HEMOGLOBINA:"
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
         Left            =   6600
         TabIndex        =   45
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prox.Control:"
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
         TabIndex        =   30
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Médico:"
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
         Left            =   6120
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ECO Cadera:"
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
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vacunas:"
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
         Left            =   6600
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lactancia:"
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
         TabIndex        =   22
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PESO:"
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
         Left            =   8760
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edad Ctrol:"
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
         Left            =   3600
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHA CTROL:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos del paciente"
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
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.TextBox t_fram 
         Height          =   375
         Left            =   10560
         TabIndex        =   97
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton b_ok 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10560
         Picture         =   "frm_metas.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1320
         Width           =   375
      End
      Begin VB.ComboBox cbomet2 
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
         ItemData        =   "frm_metas.frx":2D2C
         Left            =   3720
         List            =   "frm_metas.frx":2D2E
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   6615
      End
      Begin VB.ComboBox cbomet 
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
         ItemData        =   "frm_metas.frx":2D30
         Left            =   1680
         List            =   "frm_metas.frx":2D32
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox t_base 
         Alignment       =   1  'Right Justify
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
         Left            =   7200
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox t_codc 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label labfnac 
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
         Left            =   6120
         TabIndex        =   43
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label labd 
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
         Left            =   9720
         TabIndex        =   41
         Top             =   840
         Width           =   615
      End
      Begin VB.Label labm 
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
         Left            =   9000
         TabIndex        =   40
         Top             =   840
         Width           =   615
      End
      Begin VB.Label labmat 
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
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "META:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label laba 
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
         Left            =   8280
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7440
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label labcnv 
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
         Left            =   4800
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label labnom 
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
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CÉDULA:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar por cédula:"
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
      TabIndex        =   94
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2955
      Left            =   6600
      Picture         =   "frm_metas.frx":2D34
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   3825
   End
End
Attribute VB_Name = "frm_metas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Regcli As New ADODB.Recordset
Public Regmeta As New ADODB.Recordset
Public Reconsmeta As New ADODB.Recordset
Public Sqlcli, Sqlmeta, Sqlconsmeta As String


Private Sub b_alta_Click()

Frame1.Enabled = True
framnino1.Enabled = True
Frame2.Enabled = True
Frame4.Enabled = True

borrar_camp
t_ced.SetFocus
XAlta = 1
mf.Text = Format(Date, "dd/mm/yyyy")
t_base.Text = frm_menu.data_parse.Recordset("base")
b_alta.Enabled = False
b_grab.Enabled = True
b_mod.Enabled = False
b_elim.Enabled = False
b_imp.Enabled = False
b_canc.Enabled = True
flex1.Enabled = False

End Sub

Private Sub b_canc_Click()
borrar_camp
limpiar
XAlta = 0
b_alta.Enabled = True
b_grab.Enabled = False
b_mod.Enabled = True
b_elim.Enabled = True
b_imp.Enabled = True
b_canc.Enabled = False
Frame1.Enabled = False
framnino1.Visible = False
Frame4.Visible = False
Frame2.Visible = False
flex1.Enabled = True
cargargrid

End Sub

Private Sub b_elim_Click()
Dim Xdesea As String
Xdesea = MsgBox("Desea borrar el registro seleccionado " & Adodc1.Recordset("nombre") & " ??", vbExclamation + vbYesNo)
If Xdesea = vbYes Then
   Adodc1.Recordset.Delete
   cargargrid
   MsgBox "Registro Borrado"

End If

End Sub

Private Sub b_grab_Click()

'On Error GoTo Xerralgrabmet

If t_ced.Text = "" And t_codc.Text = "" Then
   MsgBox "No ingresó documento del paciente", vbExclamation
Else
   If XAlta = 1 Then
      Adodc1.Recordset.AddNew
      If t_fram.Text <> "" Then
         Adodc1.Recordset("frame") = t_fram.Text
      Else
         Adodc1.Recordset("frame") = 0
      End If
      Adodc1.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
      Adodc1.Recordset("cedula") = t_ced.Text
      Adodc1.Recordset("codced") = t_codc.Text
      Adodc1.Recordset("base") = t_base.Text
      Adodc1.Recordset("matric") = Val(labmat.Caption)
      Adodc1.Recordset("nombre") = Mid(labnom.Caption, 1, 100)
      Adodc1.Recordset("cnvcod") = labcnv.Caption
      Adodc1.Recordset("fecnac") = Format(labfnac.Caption, "dd/mm/yyyy")
      If t_cedmed.Text <> "" Then
         Adodc1.Recordset("cedmed") = t_cedmed.Text
      Else
         Adodc1.Recordset("cedmed") = "0"
      End If
      If laba.Caption <> "" Then
         If Val(laba.Caption) > 0 Then
            Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
         Else
            Adodc1.Recordset("edadtex") = labm.Caption & " MESES " & labd.Caption & " DIAS"
         End If
         Adodc1.Recordset("eda") = Val(laba.Caption)
         Adodc1.Recordset("edm") = Val(labm.Caption)
         Adodc1.Recordset("edd") = Val(labd.Caption)
      Else
         Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
         Adodc1.Recordset("eda") = 0
         Adodc1.Recordset("edm") = Val(labm.Caption)
         Adodc1.Recordset("edd") = Val(labd.Caption)
      End If
      Adodc1.Recordset("meta") = cbomet.ListIndex
      Adodc1.Recordset("metadesc") = cbomet.Text
      Adodc1.Recordset("meta2") = cbomet2.ListIndex
      Adodc1.Recordset("meta2desc") = cbomet2.Text
      If mfctr.Text <> "__/__/____" Then
         Adodc1.Recordset("fecctrl") = Format(mfctr.Text, "dd/mm/yyyy")
         If laba.Caption <> "" Then
            If Val(laba.Caption) > 0 Then
               Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
            Else
               Adodc1.Recordset("edadtex") = labm.Caption & " MESES " & labd.Caption & " DIAS"
            End If
            Adodc1.Recordset("edca") = Val(t_aa.Text)
            Adodc1.Recordset("edcm") = Val(t_mm.Text)
            Adodc1.Recordset("edcd") = Val(t_dd.Text)
         Else
            Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
            Adodc1.Recordset("edca") = 0
            Adodc1.Recordset("edcm") = Val(t_mm.Text)
            Adodc1.Recordset("edcd") = Val(t_dd.Text)
         End If
         If t_aa.Text <> "" Then
            If Val(t_aa.Text) > 0 Then
               Adodc1.Recordset("edadtex2") = t_aa.Text & " AÑOS " & t_mm.Text & " MESES " & t_dd.Text & " DIAS"
            Else
               Adodc1.Recordset("edadtex2") = t_mm.Text & " MESES " & t_dd.Text & " DIAS"
            End If
         Else
            Adodc1.Recordset("edadtex2") = t_mm.Text & " MESES " & t_dd.Text & " DIAS"
         End If
         If t_peso.Text <> "" Then
            Adodc1.Recordset("peso") = t_peso.Text
         End If
         If t_lacta.Text <> "" Then
            Adodc1.Recordset("lactan") = t_lacta.Text
         End If
         If t_vacun.Text <> "" Then
            Adodc1.Recordset("vacuna") = t_vacun.Text
         End If
         If t_eco.Text <> "" Then
            Adodc1.Recordset("ecocad") = t_eco.Text
         End If
         If t_hemog.Text <> "" Then
            Adodc1.Recordset("hemog") = t_hemog.Text
         End If
         If t_odont.Text <> "" Then
            Adodc1.Recordset("odont") = t_odont.Text
         End If
         If t_medic.Text <> "" Then
            Adodc1.Recordset("medico") = t_medic.Text
         End If
         If t_proxctr.Text <> "" Then
            Adodc1.Recordset("fecprox") = t_proxctr.Text
         End If
         If t_fdesar.Text <> "" Then
            Adodc1.Recordset("fdesar") = t_fdesar.Text
         End If
         If t_ctroft.Text <> "" Then
            Adodc1.Recordset("oft") = t_ctroft.Text
         End If
         If t_obs.Text <> "" Then
            Adodc1.Recordset("obs") = t_obs.Text
         End If
         Adodc1.Recordset.Update
         Adodc1.Refresh
         borrar_camp
         limpiar
         XAlta = 0
         b_alta.Enabled = True
         b_grab.Enabled = False
         b_mod.Enabled = True
         b_elim.Enabled = True
         b_imp.Enabled = True
         b_canc.Enabled = False
         Frame1.Enabled = False
         framnino1.Visible = False
         Frame4.Visible = False
         Frame2.Visible = False
         flex1.Enabled = True
         cargargrid
      Else
         If mfctremb.Text <> "__/__/____" Then
            Adodc1.Recordset("fecctrl") = Format(mfctremb.Text, "dd/mm/yyyy")
            If t_lochc.Text <> "" Then
               Adodc1.Recordset("lochc") = t_lochc.Text
            End If
            Adodc1.Recordset("obstetr") = cboobste.ListIndex
            Adodc1.Recordset("odont2") = cboodont.ListIndex
            Adodc1.Recordset("vdrl") = cbovdrl.ListIndex
            Adodc1.Recordset("hcpb") = cbohcpb.ListIndex
            Adodc1.Recordset("prot") = cboprot.ListIndex
            If t_anticon.Text <> "" Then
               Adodc1.Recordset("obsemb") = t_anticon.Text
            End If
            If t_cedrn.Text <> "" Then
               Adodc1.Recordset("cedrn") = t_cedrn.Text
            End If
            If t_nrocons.Text <> "" Then
               Adodc1.Recordset("nroconse") = t_nrocons.Text
            End If
            Adodc1.Recordset("cbopap") = cbopap.ListIndex
            If mfvdrl.Text <> "__/__/____" Then
               Adodc1.Recordset("fecvdrl") = Format(mfvdrl.Text, "dd/mm/yyyy")
            End If
            Adodc1.Recordset("cbomamo") = cbomamo.ListIndex
            If mfanti.Text <> "__/__/____" Then
               Adodc1.Recordset("fecanti") = Format(mfanti.Text, "dd/mm/yyyy")
            End If
            If t_semgon.Text <> "" Then
               Adodc1.Recordset("semgono") = t_semgon.Text
            End If
            Adodc1.Recordset.Update
            Adodc1.Refresh
            borrar_camp
            limpiar
            XAlta = 0
            b_alta.Enabled = True
            b_grab.Enabled = False
            b_mod.Enabled = True
            b_elim.Enabled = True
            b_imp.Enabled = True
            b_canc.Enabled = False
            Frame1.Enabled = False
            framnino1.Visible = False
            Frame4.Visible = False
            Frame2.Visible = False
            flex1.Enabled = True
            cargargrid
         Else
            If mfctradult.Text <> "__/__/____" Then
               Adodc1.Recordset("fecctrl") = Format(mfctradult.Text, "dd/mm/yyyy")
               If t_nommed.Text <> "" Then
                  Adodc1.Recordset("nommed") = t_nommed.Text
               End If
               If t_apemed.Text <> "" Then
                  Adodc1.Recordset("apemed") = t_apemed.Text
               End If
               Adodc1.Recordset("sia") = cbosia.ListIndex
               Adodc1.Recordset("screen") = cboscree.ListIndex
               Adodc1.Recordset("adulto") = cboregadul.ListIndex
               If t_feca.Text <> "" Then
                  Adodc1.Recordset("fecatest") = t_feca.Text
               End If
               If t_proxcon3.Text <> "" Then
                  Adodc1.Recordset("fecprox") = t_proxcon3.Text
               End If
               Adodc1.Recordset.Update
               Adodc1.Refresh
               borrar_camp
               limpiar
               XAlta = 0
               b_alta.Enabled = True
               b_grab.Enabled = False
               b_mod.Enabled = True
               b_elim.Enabled = True
               b_imp.Enabled = True
               b_canc.Enabled = False
               Frame1.Enabled = False
               framnino1.Visible = False
               Frame4.Visible = False
               Frame2.Visible = False
               flex1.Enabled = True
               cargargrid
            Else
               MsgBox "Hay un error en la fecha, verifique! No se puede grabar", vbExclamation
               Adodc1.Recordset.CancelUpdate
               Exit Sub
            End If
         End If
      End If
   Else
      If t_fram.Text <> "" Then
         Adodc1.Recordset("frame") = t_fram.Text
      Else
         Adodc1.Recordset("frame") = 0
      End If
      Adodc1.Recordset("fecha") = Format(mf.Text, "dd/mm/yyyy")
      Adodc1.Recordset("cedula") = t_ced.Text
      Adodc1.Recordset("codced") = t_codc.Text
      Adodc1.Recordset("base") = t_base.Text
      Adodc1.Recordset("matric") = Val(labmat.Caption)
      Adodc1.Recordset("nombre") = Mid(labnom.Caption, 1, 100)
      Adodc1.Recordset("cnvcod") = labcnv.Caption
      Adodc1.Recordset("fecnac") = Format(labfnac.Caption, "dd/mm/yyyy")
      If t_cedmed.Text <> "" Then
         Adodc1.Recordset("cedmed") = t_cedmed.Text
      Else
         Adodc1.Recordset("cedmed") = "0"
      End If
      If laba.Caption <> "" Then
         If Val(laba.Caption) > 0 Then
            Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
         Else
            Adodc1.Recordset("edadtex") = labm.Caption & " MESES " & labd.Caption & " DIAS"
         End If
         Adodc1.Recordset("eda") = Val(laba.Caption)
         Adodc1.Recordset("edm") = Val(labm.Caption)
         Adodc1.Recordset("edd") = Val(labd.Caption)
      Else
         Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
         Adodc1.Recordset("eda") = 0
         Adodc1.Recordset("edm") = Val(labm.Caption)
         Adodc1.Recordset("edd") = Val(labd.Caption)
      End If
      Adodc1.Recordset("meta") = cbomet.ListIndex
      Adodc1.Recordset("metadesc") = cbomet.Text
      Adodc1.Recordset("meta2") = cbomet2.ListIndex
      Adodc1.Recordset("meta2desc") = cbomet2.Text
      If mfctr.Text <> "__/__/____" Then
         Adodc1.Recordset("fecctrl") = Format(mfctr.Text, "dd/mm/yyyy")
         If laba.Caption <> "" Then
            If Val(laba.Caption) > 0 Then
               Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
            Else
               Adodc1.Recordset("edadtex") = labm.Caption & " MESES " & labd.Caption & " DIAS"
            End If
         Else
            Adodc1.Recordset("edadtex") = laba.Caption & " AÑOS " & labm.Caption & " MESES " & labd.Caption & " DIAS"
         End If
         If t_aa.Text <> "" Then
            If Val(t_aa.Text) > 0 Then
               Adodc1.Recordset("edadtex2") = t_aa.Text & " AÑOS " & t_mm.Text & " MESES " & t_dd.Text & " DIAS"
            Else
               Adodc1.Recordset("edadtex2") = t_mm.Text & " MESES " & t_dd.Text & " DIAS"
            End If
            Adodc1.Recordset("edca") = Val(t_aa.Text)
            Adodc1.Recordset("edcm") = Val(t_mm.Text)
            Adodc1.Recordset("edcd") = Val(t_dd.Text)
         Else
            Adodc1.Recordset("edadtex2") = t_mm.Text & " MESES " & t_dd.Text & " DIAS"
            Adodc1.Recordset("edca") = 0
            Adodc1.Recordset("edcm") = Val(t_mm.Text)
            Adodc1.Recordset("edcd") = Val(t_dd.Text)
         End If
         If t_peso.Text <> "" Then
            Adodc1.Recordset("peso") = t_peso.Text
         End If
         If t_lacta.Text <> "" Then
            Adodc1.Recordset("lactan") = t_lacta.Text
         End If
         If t_vacun.Text <> "" Then
            Adodc1.Recordset("vacuna") = t_vacun.Text
         End If
         If t_eco.Text <> "" Then
            Adodc1.Recordset("ecocad") = t_eco.Text
         End If
         If t_hemog.Text <> "" Then
            Adodc1.Recordset("hemog") = t_hemog.Text
         End If
         If t_odont.Text <> "" Then
            Adodc1.Recordset("odont") = t_odont.Text
         End If
         If t_medic.Text <> "" Then
            Adodc1.Recordset("medico") = t_medic.Text
         End If
         If t_proxctr.Text <> "" Then
            Adodc1.Recordset("fecprox") = t_proxctr.Text
         End If
         If t_fdesar.Text <> "" Then
            Adodc1.Recordset("fdesar") = t_fdesar.Text
         End If
         If t_ctroft.Text <> "" Then
            Adodc1.Recordset("oft") = t_ctroft.Text
         End If
         If t_obs.Text <> "" Then
            Adodc1.Recordset("obs") = t_obs.Text
         End If
         Adodc1.Recordset.Update
         Adodc1.Refresh
         flex1.Enabled = True
         cargargrid
      Else
         If mfctremb.Text <> "__/__/____" Then
            Adodc1.Recordset("fecctrl") = Format(mfctremb.Text, "dd/mm/yyyy")
            If t_lochc.Text <> "" Then
               Adodc1.Recordset("lochc") = t_lochc.Text
            End If
            Adodc1.Recordset("obstetr") = cboobste.ListIndex
            Adodc1.Recordset("odont2") = cboodont.ListIndex
            Adodc1.Recordset("vdrl") = cbovdrl.ListIndex
            Adodc1.Recordset("hcpb") = cbohcpb.ListIndex
            Adodc1.Recordset("prot") = cboprot.ListIndex
            If t_anticon.Text <> "" Then
               Adodc1.Recordset("obsemb") = t_anticon.Text
            End If
            If t_cedrn.Text <> "" Then
               Adodc1.Recordset("cedrn") = t_cedrn.Text
            End If
            If t_nrocons.Text <> "" Then
               Adodc1.Recordset("nroconse") = t_nrocons.Text
            End If
            Adodc1.Recordset("cbopap") = cbopap.ListIndex
            If mfvdrl.Text <> "__/__/____" Then
               Adodc1.Recordset("fecvdrl") = Format(mfvdrl.Text, "dd/mm/yyyy")
            End If
            Adodc1.Recordset("cbomamo") = cbomamo.ListIndex
            If mfanti.Text <> "__/__/____" Then
               Adodc1.Recordset("fecanti") = Format(mfanti.Text, "dd/mm/yyyy")
            End If
            If t_semgon.Text <> "" Then
               Adodc1.Recordset("semgono") = t_semgon.Text
            End If
            
            Adodc1.Recordset.Update
            Adodc1.Refresh
            flex1.Enabled = True
            cargargrid
         Else
            If mfctradult.Text <> "__/__/____" Then
               Adodc1.Recordset("fecctrl") = Format(mfctradult.Text, "dd/mm/yyyy")
               If t_nommed.Text <> "" Then
                  Adodc1.Recordset("nommed") = t_nommed.Text
               End If
               If t_apemed.Text <> "" Then
                  Adodc1.Recordset("apemed") = t_apemed.Text
               End If
               Adodc1.Recordset("sia") = cbosia.ListIndex
               Adodc1.Recordset("screen") = cboscree.ListIndex
               Adodc1.Recordset("adulto") = cboregadul.ListIndex
               If t_feca.Text <> "" Then
                  Adodc1.Recordset("fecatest") = t_feca.Text
               End If
               If t_proxcon3.Text <> "" Then
                  Adodc1.Recordset("fecprox") = t_proxcon3.Text
               End If
               Adodc1.Recordset.Update
               Adodc1.Refresh
               flex1.Enabled = True
               cargargrid
            Else
               MsgBox "Hay un error en la fecha, verifique! No se puede grabar", vbExclamation
               Adodc1.Recordset.CancelUpdate
               Exit Sub
            End If
         End If
      End If
      borrar_camp
      limpiar
      XAlta = 0
      b_alta.Enabled = True
      b_grab.Enabled = False
      b_mod.Enabled = True
      b_elim.Enabled = True
      b_imp.Enabled = True
      b_canc.Enabled = False
      Frame1.Enabled = False
      framnino1.Visible = False
      Frame4.Visible = False
      Frame2.Visible = False
   End If
End If
   
'Exit Sub

'Xerralgrabmet:
'              If Err.Number = 3157 Then
'                 MsgBox "Error al grabar, verifique la numeración"
'              Else
'                 MsgBox "Error al grabar, verifique los datos"
'              End If
              
End Sub

Private Sub b_imp_Click()
frm_infmetasnew.Show vbModal

End Sub

Private Sub b_mod_Click()
Frame1.Enabled = True
framnino1.Enabled = True
Frame2.Enabled = True
Frame4.Enabled = True

'borrar_camp
't_ced.SetFocus
XAlta = 0
b_alta.Enabled = False
b_grab.Enabled = True
b_mod.Enabled = False
b_elim.Enabled = False
b_imp.Enabled = False
b_canc.Enabled = True
flex1.Enabled = False

End Sub

Private Sub b_ok_Click()
Dim Xnrohc As Double
Dim Xcodmedic As Integer
limpiar
If cbomet2.Text = "CTROL.RECIEN NACIDO" Then
   framnino1.Visible = True
   Frame2.Visible = False
   Frame4.Visible = False
   t_fram.Text = 1
   framnino1.Caption = cbomet2.Text
   ConectarBD
   ConbdSapp.Open
   Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod =" & 190001
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      mfctr.Text = Regcli("fecha")
      CalculaEdad2 (mfctr.Text)
   End If
   Regcli.Close
   Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      Xnrohc = Regcli("hc_nro")
      t_medic.Text = Regcli("hc_nommed")
      Xcodmedic = Regcli("hc_codmed")
   Else
      MsgBox "No se encontrò HC creada para Control META", vbInformation
      Xnrohc = 0
      t_medic.Text = ""
      Xcodmedic = 0
   End If
   Regcli.Close
   Sqlcli = "Select * from us where id =" & Xcodmedic
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      t_cedmed.Text = Regcli("documento")
   Else
      t_cedmed.Text = 0
   End If
   Regcli.Close
   Sqlcli = "Select * from hc_examen where sl_hcnro =" & Xnrohc
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      If IsNull(Regcli("sc_ritmod")) = False Then
         t_peso.Text = Regcli("sc_ritmod")
      Else
         t_peso.Text = ""
      End If
   Else
      t_peso.Text = ""
   End If
   Regcli.Close
   Sqlcli = "Select * from hc_antinmu where hc_nro =" & Xnrohc
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      If IsNull(Regcli("cevfec")) = False Then
         If Format(Regcli("cevfec"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
            t_vacun.Text = "CEV ATRASADO: " & Regcli("cevfec")
         Else
            t_vacun.Text = "CEV VIGENTE: " & Regcli("cevfec")
         End If
      Else
         t_vacun.Text = ""
      End If
   Else
      t_vacun.Text = ""
   End If
   Regcli.Close
   Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      If IsNull(Regcli("hc_descrip")) = False Then
         t_proxctr.Text = Regcli("hc_descrip")
      Else
         t_proxctr.Text = ""
      End If
   Else
      t_proxctr.Text = ""
   End If
   Regcli.Close
   Sqlcli = "Select * from hc_antalim where hc_nro =" & Xnrohc
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      If IsNull(Regcli("descant")) = False Then
         t_lacta.Text = Regcli("descant")
      Else
         t_lacta.Text = ""
      End If
   Else
      t_lacta.Text = ""
   End If
   Regcli.Close
'1603564
   
   ConbdSapp.Close
Else
   If cbomet2.Text = "CTROL.1ER.AÑO DE VIDA" Then
      framnino1.Visible = True
      Frame2.Visible = False
      Frame4.Visible = False
      framnino1.Caption = cbomet2.Text
      t_fram.Text = 1
      ConectarBD
      ConbdSapp.Open
      Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod =" & 190003
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         mfctr.Text = Regcli("fecha")
         CalculaEdad2 (mfctr.Text)
      End If
      Regcli.Close
      Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         Xnrohc = Regcli("hc_nro")
         t_medic.Text = Regcli("hc_nommed")
      Else
         MsgBox "No se encontrò HC creada por Consulta META", vbInformation
         Xnrohc = 0
         t_medic.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_examen where sl_hcnro =" & Xnrohc
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         If IsNull(Regcli("sc_ritmod")) = False Then
            t_peso.Text = Regcli("sc_ritmod")
         Else
            t_peso.Text = ""
         End If
      Else
         t_peso.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_antinmu where hc_nro =" & Xnrohc
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         If IsNull(Regcli("cevfec")) = False Then
            If Format(Regcli("cevfec"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
               t_vacun.Text = "CEV ATRASADO: " & Regcli("cevfec")
            Else
               t_vacun.Text = "CEV VIGENTE: " & Regcli("cevfec")
            End If
         Else
            t_vacun.Text = ""
         End If
      Else
         t_vacun.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         If IsNull(Regcli("hc_descrip")) = False Then
            t_proxctr.Text = Regcli("hc_descrip")
         Else
            t_proxctr.Text = ""
         End If
      Else
         t_proxctr.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "ECO-CADERA" & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         If IsNull(Regcli("ecocadsn")) = False Then
            If Regcli("ecocadsn") = "SI" Then
               t_eco.Text = "ECO-CADERA NORMAL FECHA:" & Regcli("fecha")
            Else
               t_eco.Text = "ECO-CADERA NO"
            End If
         Else
            t_eco.Text = ""
         End If
      Else
         t_eco.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "HEMOGLOBINA" & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         If IsNull(Regcli("hemoghor")) = False Then
            t_hemog.Text = "HEMOGLOBINA: " & Regcli("hemoghor") & " Fecha:" & Regcli("fecha")
         Else
            t_hemog.Text = ""
         End If
      Else
         t_hemog.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "CONTROL ODONTOLOGICO" & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         t_odont.Text = "Ctrol.ODONTOLOGICO Fecha:" & Regcli("fecha")
      Else
         t_odont.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_antalim where hc_nro =" & Xnrohc
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         If IsNull(Regcli("descant")) = False Then
            t_lacta.Text = Regcli("descant")
         Else
            t_lacta.Text = ""
         End If
      Else
         t_lacta.Text = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_antmad2 where hc_nro =" & Xnrohc & " and deschoja ='" & "A los 4 meses" & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         t_fdesar.Text = "Form.Ctrol.Desarrollo " & Regcli("deschoja")
      Else
         t_fdesar.Text = ""
      End If
      
      Regcli.Close
    '1603564
       
      ConbdSapp.Close
   Else
      If cbomet2.Text = "CTROL.2DO.AÑO DE VIDA" Then
         framnino1.Visible = True
         framnino1.Caption = cbomet2.Text
         Frame2.Visible = False
         Frame4.Visible = False
         t_fram.Text = 1
         ConectarBD
         ConbdSapp.Open
         Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod =" & 190004
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            mfctr.Text = Regcli("fecha")
            CalculaEdad2 (mfctr.Text)
         End If
         Regcli.Close
         Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            Xnrohc = Regcli("hc_nro")
            t_medic.Text = Regcli("hc_nommed")
         Else
            MsgBox "No se encontrò HC creada para Control META", vbInformation
            Xnrohc = 0
            t_medic.Text = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_examen where sl_hcnro =" & Xnrohc
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            If IsNull(Regcli("sc_ritmod")) = False Then
               t_peso.Text = Regcli("sc_ritmod")
            Else
               t_peso.Text = ""
            End If
         Else
            t_peso.Text = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_antinmu where hc_nro =" & Xnrohc
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            If IsNull(Regcli("cevfec")) = False Then
               If Format(Regcli("cevfec"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                  t_vacun.Text = "CEV ATRASADO: " & Regcli("cevfec")
               Else
                  t_vacun.Text = "CEV VIGENTE: " & Regcli("cevfec")
               End If
            Else
               t_vacun.Text = ""
            End If
         Else
            t_vacun.Text = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            If IsNull(Regcli("hc_descrip")) = False Then
               t_proxctr.Text = Regcli("hc_descrip")
            Else
               t_proxctr.Text = ""
            End If
         Else
            t_proxctr.Text = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_antalim where hc_nro =" & Xnrohc
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            If IsNull(Regcli("descant")) = False Then
               t_lacta.Text = Regcli("descant")
            Else
               t_lacta.Text = ""
            End If
         Else
            t_lacta.Text = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "CONTROL ODONTOLOGICO" & "'"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            t_odont.Text = "Ctrol.ODONTOLOGICO Fecha:" & Regcli("fecha")
         Else
            t_odont.Text = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_antmad2 where hc_nro =" & Xnrohc & " and deschoja ='" & "A los 18 meses" & "'"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            t_fdesar.Text = "Form.Ctrol.Desarrollo " & Regcli("deschoja")
         Else
            t_fdesar.Text = ""
         End If
         Regcli.Close
         ConbdSapp.Close
      Else
         If cbomet2.Text = "CTROL.3ER.AÑO DE VIDA" Then
            framnino1.Visible = True
            framnino1.Caption = cbomet2.Text
            Frame2.Visible = False
            Frame4.Visible = False
            t_fram.Text = 1
            ConectarBD
            ConbdSapp.Open
            Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod =" & 190005
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               mfctr.Text = Regcli("fecha")
               CalculaEdad2 (mfctr.Text)
            End If
            Regcli.Close
            Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
            With Regcli
                 .CursorLocation = adUseClient
                 .CursorType = adOpenKeyset
                 .LockType = adLockOptimistic
                 .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               Xnrohc = Regcli("hc_nro")
               t_medic.Text = Regcli("hc_nommed")
            Else
               MsgBox "No se encontrò HC creada para Control META", vbInformation
               Xnrohc = 0
               t_medic.Text = ""
            End If
            Regcli.Close
            Sqlcli = "Select * from hc_examen where sl_hcnro =" & Xnrohc
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               If IsNull(Regcli("sc_ritmod")) = False Then
                  t_peso.Text = Regcli("sc_ritmod")
               Else
                  t_peso.Text = ""
               End If
            Else
               t_peso.Text = ""
            End If
            Regcli.Close
            Sqlcli = "Select * from hc_antinmu where hc_nro =" & Xnrohc
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               If IsNull(Regcli("cevfec")) = False Then
                  If Format(Regcli("cevfec"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                     t_vacun.Text = "CEV ATRASADO: " & Regcli("cevfec")
                  Else
                     t_vacun.Text = "CEV VIGENTE: " & Regcli("cevfec")
                  End If
               Else
                  t_vacun.Text = ""
               End If
            Else
               t_vacun.Text = ""
            End If
            Regcli.Close
            Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               If IsNull(Regcli("hc_descrip")) = False Then
                  t_proxctr.Text = Regcli("hc_descrip")
               Else
                  t_proxctr.Text = ""
               End If
            Else
               t_proxctr.Text = ""
            End If
            Regcli.Close
            Sqlcli = "Select * from hc_antalim where hc_nro =" & Xnrohc
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               If IsNull(Regcli("descant")) = False Then
                  t_lacta.Text = Regcli("descant")
               Else
                  t_lacta.Text = ""
               End If
            Else
               t_lacta.Text = ""
            End If
            Regcli.Close
            Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "CONTROL ODONTOLOGICO" & "'"
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               t_odont.Text = "Ctrol.ODONTOLOGICO Fecha:" & Regcli("fecha")
            Else
               t_odont.Text = ""
            End If
            Regcli.Close
            ConbdSapp.Close
         Else
            If cbomet2.Text = "CTROL.4TO.AÑO DE VIDA" Then
               framnino1.Visible = True
               framnino1.Caption = cbomet2.Text
               Frame2.Visible = False
               Frame4.Visible = False
               t_fram.Text = 1
               ConectarBD
               ConbdSapp.Open
               Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod =" & 190030
               With Regcli
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  mfctr.Text = Regcli("fecha")
                  CalculaEdad2 (mfctr.Text)
               End If
               Regcli.Close
               Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
               With Regcli
                    .CursorLocation = adUseClient
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  Xnrohc = Regcli("hc_nro")
                  t_medic.Text = Regcli("hc_nommed")
               Else
                  MsgBox "No se encontrò HC creada para Control META", vbInformation
                  Xnrohc = 0
                  t_medic.Text = ""
               End If
               Regcli.Close
               Sqlcli = "Select * from hc_examen where sl_hcnro =" & Xnrohc
               With Regcli
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  If IsNull(Regcli("sc_ritmod")) = False Then
                     t_peso.Text = Regcli("sc_ritmod")
                  Else
                     t_peso.Text = ""
                  End If
               Else
                  t_peso.Text = ""
               End If
               Regcli.Close
               Sqlcli = "Select * from hc_antinmu where hc_nro =" & Xnrohc
               With Regcli
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  If IsNull(Regcli("cevfec")) = False Then
                     If Format(Regcli("cevfec"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                        t_vacun.Text = "CEV ATRASADO: " & Regcli("cevfec")
                     Else
                        t_vacun.Text = "CEV VIGENTE: " & Regcli("cevfec")
                     End If
                  Else
                     t_vacun.Text = ""
                  End If
               Else
                  t_vacun.Text = ""
               End If
               Regcli.Close
               Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
               With Regcli
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  If IsNull(Regcli("hc_descrip")) = False Then
                     t_proxctr.Text = Regcli("hc_descrip")
                  Else
                     t_proxctr.Text = ""
                  End If
               Else
                  t_proxctr.Text = ""
               End If
               Regcli.Close
               Sqlcli = "Select * from hc_antalim where hc_nro =" & Xnrohc
               With Regcli
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  If IsNull(Regcli("descant")) = False Then
                     t_lacta.Text = Regcli("descant")
                  Else
                     t_lacta.Text = ""
                  End If
               Else
                  t_lacta.Text = ""
               End If
               Regcli.Close
               Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "CONTROL ODONTOLOGICO" & "'"
               With Regcli
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  t_odont.Text = "Ctrol.ODONTOLOGICO Fecha:" & Regcli("fecha")
               Else
                  t_odont.Text = ""
               End If
               Regcli.Close
               Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "CONTROL OFTALMOLOGICO" & "'"
               With Regcli
                   .CursorLocation = adUseClient
                   .CursorType = adOpenKeyset
                   .LockType = adLockOptimistic
                   .Open Sqlcli, ConbdSapp, , , adCmdText
               End With
               If Regcli.RecordCount > 0 Then
                  t_ctroft.Text = "Ctrol.OFTALMOLÓGICO Fecha:" & Regcli("fecha")
               Else
                  t_ctroft.Text = ""
               End If
               
               Regcli.Close
               ConbdSapp.Close
            Else
               If cbomet2.Text = "CTROL.5TO.AÑO DE VIDA" Then
                  framnino1.Visible = True
                  framnino1.Caption = cbomet2.Text
                  Frame2.Visible = False
                  Frame4.Visible = False
                  t_fram.Text = 1
                  ConectarBD
                  ConbdSapp.Open
                  Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod =" & 190031
                  With Regcli
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockOptimistic
                      .Open Sqlcli, ConbdSapp, , , adCmdText
                  End With
                  If Regcli.RecordCount > 0 Then
                     mfctr.Text = Regcli("fecha")
                     CalculaEdad2 (mfctr.Text)
                  End If
                  Regcli.Close
                  Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
                  With Regcli
                       .CursorLocation = adUseClient
                       .CursorType = adOpenKeyset
                       .LockType = adLockOptimistic
                       .Open Sqlcli, ConbdSapp, , , adCmdText
                  End With
                  If Regcli.RecordCount > 0 Then
                     Xnrohc = Regcli("hc_nro")
                     t_medic.Text = Regcli("hc_nommed")
                  Else
                     MsgBox "No se encontrò HC creada para Control META", vbInformation
                     Xnrohc = 0
                     t_medic.Text = ""
                  End If
                  Regcli.Close
                  Sqlcli = "Select * from hc_examen where sl_hcnro =" & Xnrohc
                  With Regcli
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockOptimistic
                      .Open Sqlcli, ConbdSapp, , , adCmdText
                  End With
                  If Regcli.RecordCount > 0 Then
                     If IsNull(Regcli("sc_ritmod")) = False Then
                        t_peso.Text = Regcli("sc_ritmod")
                     Else
                        t_peso.Text = ""
                     End If
                  Else
                     t_peso.Text = ""
                  End If
                  Regcli.Close
                  Sqlcli = "Select * from hc_antinmu where hc_nro =" & Xnrohc
                  With Regcli
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockOptimistic
                      .Open Sqlcli, ConbdSapp, , , adCmdText
                  End With
                  If Regcli.RecordCount > 0 Then
                     If IsNull(Regcli("cevfec")) = False Then
                        If Format(Regcli("cevfec"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                           t_vacun.Text = "CEV ATRASADO: " & Regcli("cevfec")
                        Else
                           t_vacun.Text = "CEV VIGENTE: " & Regcli("cevfec")
                        End If
                     Else
                        t_vacun.Text = ""
                     End If
                  Else
                     t_vacun.Text = ""
                  End If
                  Regcli.Close
                  Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
                  With Regcli
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockOptimistic
                      .Open Sqlcli, ConbdSapp, , , adCmdText
                  End With
                  If Regcli.RecordCount > 0 Then
                     If IsNull(Regcli("hc_descrip")) = False Then
                        t_proxctr.Text = Regcli("hc_descrip")
                     Else
                        t_proxctr.Text = ""
                     End If
                  Else
                     t_proxctr.Text = ""
                  End If
                  Regcli.Close
                  Sqlcli = "Select * from hc_antalim where hc_nro =" & Xnrohc
                  With Regcli
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockOptimistic
                      .Open Sqlcli, ConbdSapp, , , adCmdText
                  End With
                  If Regcli.RecordCount > 0 Then
                     If IsNull(Regcli("descant")) = False Then
                        t_lacta.Text = Regcli("descant")
                     Else
                        t_lacta.Text = ""
                     End If
                  Else
                     t_lacta.Text = ""
                  End If
                  Regcli.Close
                  Sqlcli = "Select * from hc_ctroles where hc_nro =" & Xnrohc & " and descrip ='" & "CONTROL ODONTOLOGICO" & "'"
                  With Regcli
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockOptimistic
                      .Open Sqlcli, ConbdSapp, , , adCmdText
                  End With
                  If Regcli.RecordCount > 0 Then
                     t_odont.Text = "Ctrol.ODONTOLOGICO Fecha:" & Regcli("fecha")
                  Else
                     t_odont.Text = ""
                  End If
                  Regcli.Close
                  ConbdSapp.Close
               Else
                  If cbomet2.Text = "CTROL.EMBARAZADAS" Then
                     Frame2.Visible = True
                     framnino1.Visible = False
                     Frame2.Caption = cbomet2.Text
                     Frame4.Visible = False
                     t_fram.Text = 2
                     ConectarBD
                     ConbdSapp.Open
                     Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod in (190015,190016,190025) order by fecha DESC"
                     With Regcli
                         .CursorLocation = adUseClient
                         .CursorType = adOpenKeyset
                         .LockType = adLockOptimistic
                         .Open Sqlcli, ConbdSapp, , , adCmdText
                     End With
                     If Regcli.RecordCount > 0 Then
                        mfctremb.Text = Regcli("fecha")
                        cbovdrl.ListIndex = 0
                     End If
                     Regcli.Close
                     Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
                     With Regcli
                          .CursorLocation = adUseClient
                          .CursorType = adOpenKeyset
                          .LockType = adLockOptimistic
                          .Open Sqlcli, ConbdSapp, , , adCmdText
                     End With
                     If Regcli.RecordCount > 0 Then
                        Xnrohc = Regcli("hc_nro")
'                        t_medic.Text = Regcli("hc_nommed")
                     Else
                        MsgBox "No se encontrò HC creada para Control META", vbInformation
                        Xnrohc = 0
'                        t_medic.Text = ""
                     End If
                     Regcli.Close
                     ConbdSapp.Close
                  Else
                     If cbomet2.Text = "PESQUISA V.DOMESTICA" Then
                        Frame2.Visible = True
                        framnino1.Visible = False
                        Frame4.Visible = False
                        t_fram.Text = 0
'                       Frame2.Caption = cbomet2.Text
                        ConectarBD
                        ConbdSapp.Open
                        Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod in (190010) order by fecha DESC"
                        With Regcli
                            .CursorLocation = adUseClient
                            .CursorType = adOpenKeyset
                            .LockType = adLockOptimistic
                            .Open Sqlcli, ConbdSapp, , , adCmdText
                        End With
                        If Regcli.RecordCount > 0 Then
                           Dim Xregviolencia As String
                           Xregviolencia = MsgBox("Se realizó Pesquisa de V.Doméstica el:" & Regcli("fecha") & " Desea registrar?", vbYesNo + vbInformation)
                           If Xregviolencia = vbYes Then
                           
                           End If
                        Else
                           MsgBox "No se encontró registro en HCE de V.Doméstica"
                        End If
                        Regcli.Close
                        ConbdSapp.Close
                     Else
                        Command1_Click
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If

End Sub

Private Sub cbomet2_Click()
If cbomet2.Text = "CTROL.12 A 19AÑOS" Or _
   cbomet2.Text = "CTROL.45 A 64AÑOS" Then
   cbomet.ListIndex = 1
Else
   If cbomet2.Text = "CTROL.65 A 74AÑOS" Or cbomet2.Text = "CTROL. >75 AÑOS" Then
      cbomet.ListIndex = 2
   Else
      cbomet.ListIndex = 0
   End If
End If


End Sub

Private Sub Command1_Click()
Dim Xcodmedhc As Integer

If cbomet2.Text = "CTROL.12 A 19AÑOS" Then
   framnino1.Visible = False
   Frame2.Visible = False
   Frame4.Visible = True
   Label31.Visible = False
   Label32.Visible = False
   Label33.Visible = False
   cboscree.Visible = False
   t_feca.Visible = False
   cboregadul.Visible = False
   Label30.Visible = True
   cbosia.Visible = True
   Frame4.Caption = cbomet2.Text
   t_fram.Text = 3
   ConectarBD
   ConbdSapp.Open
   Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod in (190011,190012) order by fecha DESC"
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      mfctr.Text = Regcli("fecha")
   End If
   Regcli.Close
   Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
   With Regcli
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      Xnrohc = Regcli("hc_nro")
      Xcodmedhc = Regcli("hc_codmed")
      t_medic.Text = Regcli("hc_nommed")
      Regcli.Close
      Sqlcli = "Select * from us where id =" & Xcodmedhc
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         t_nommed.Text = Regcli("nombre")
         t_apemed.Text = Regcli("apellidos")
         labcedmed.Caption = Regcli("documento")
      Else
         t_nommed.Text = ""
         t_apemed.Text = ""
         labcedmed.Caption = ""
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_metasr where hc_nro =" & Xnrohc & " and hc_metadesc ='" & "HOJA SIA" & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         cbosia.ListIndex = 0
      Else
         cbosia.ListIndex = 1
      End If
      Regcli.Close
      Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         If IsNull(Regcli("hc_descrip")) = False Then
            t_proxcon3.Text = Regcli("hc_descrip")
         Else
            t_proxcon3.Text = ""
         End If
      Else
         t_proxcon3.Text = ""
      End If
      Regcli.Close
   Else
      MsgBox "No se encontrò HC creada para Control META", vbInformation
      Xnrohc = 0
      Xcodmedhc = 0
      t_medic.Text = ""
   End If
   ConbdSapp.Close
Else
   If cbomet2.Text = "CTROL.45 A 64AÑOS" Then
      framnino1.Visible = False
      Frame2.Visible = False
      Frame4.Visible = True
      Label31.Visible = True
      Label32.Visible = True
      Label33.Visible = False
      cboscree.Visible = True
      t_feca.Visible = True
      cboregadul.Visible = False
      t_fram.Text = 3
      Label30.Visible = False
      cbosia.Visible = False
      Frame4.Caption = cbomet2.Text
      ConectarBD
      ConbdSapp.Open
      Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod in (190013,190014) order by fecha DESC"
      With Regcli
          .CursorLocation = adUseClient
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         mfctr.Text = Regcli("fecha")
      End If
      Regcli.Close
      Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
      With Regcli
           .CursorLocation = adUseClient
           .CursorType = adOpenKeyset
           .LockType = adLockOptimistic
           .Open Sqlcli, ConbdSapp, , , adCmdText
      End With
      If Regcli.RecordCount > 0 Then
         Xnrohc = Regcli("hc_nro")
         Xcodmedhc = Regcli("hc_codmed")
         t_medic.Text = Regcli("hc_nommed")
         Regcli.Close
         Sqlcli = "Select * from us where id =" & Xcodmedhc
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            t_nommed.Text = Regcli("nombre")
            t_apemed.Text = Regcli("apellidos")
            labcedmed.Caption = Regcli("documento")
         Else
            t_nommed.Text = ""
            t_apemed.Text = ""
            labcedmed.Caption = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_metasr where hc_nro =" & Xnrohc & " and hc_metadesc ='" & "CONSULTA-SCREENING PREVENTIVO" & "'"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            cboscree.ListIndex = 0
         Else
            cbosia.ListIndex = 1
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_metascr where hc_nro =" & Xnrohc & " and hc_andesc ='" & "FECATEST PARTICULAR" & "'"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            If IsNull(Regcli("hc_hora")) = False Then
               t_feca.Text = Regcli("hc_hora")
            Else
               t_feca.Text = ""
            End If
         Else
            t_feca.Text = ""
         End If
         Regcli.Close
         Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            If IsNull(Regcli("hc_descrip")) = False Then
               t_proxcon3.Text = Regcli("hc_descrip")
            Else
               t_proxcon3.Text = ""
            End If
         Else
            t_proxcon3.Text = ""
         End If
         Regcli.Close
      Else
         MsgBox "No se encontrò HC creada para Control META", vbInformation
         Xnrohc = 0
         Xcodmedhc = 0
         t_medic.Text = ""
      End If
      ConbdSapp.Close
   Else
      If cbomet2.Text = "CTROL.65 A 74AÑOS" Or cbomet2.Text = "CTROL. >75 AÑOS" Then
         framnino1.Visible = False
         Frame2.Visible = False
         Frame4.Visible = True
         Label31.Visible = False
         Label32.Visible = False
         Label33.Visible = True
         cboscree.Visible = False
         t_feca.Visible = False
         cboregadul.Visible = True
         Label30.Visible = False
         cbosia.Visible = False
         Frame4.Caption = cbomet2.Text
         t_fram.Text = 3
         ConectarBD
         ConbdSapp.Open
         Sqlcli = "Select * from linmmdd where cod_cli =" & labmat.Caption & " and cod_prod in (190018,190019) order by fecha DESC"
         With Regcli
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            mfctr.Text = Regcli("fecha")
         End If
         Regcli.Close
         Sqlcli = "Select * from cabezal_hcdig where cednum =" & t_ced.Text & " and tipo_consd ='" & "Control Meta" & "' and fecha ='" & Format(mfctr.Text, "yyyy/mm/dd") & "'"
         With Regcli
              .CursorLocation = adUseClient
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open Sqlcli, ConbdSapp, , , adCmdText
         End With
         If Regcli.RecordCount > 0 Then
            Xnrohc = Regcli("hc_nro")
            Xcodmedhc = Regcli("hc_codmed")
            t_medic.Text = Regcli("hc_nommed")
            Regcli.Close
            Sqlcli = "Select * from us where id =" & Xcodmedhc
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               t_nommed.Text = Regcli("nombre")
               t_apemed.Text = Regcli("apellidos")
               labcedmed.Caption = Regcli("documento")
            Else
               t_nommed.Text = ""
               t_apemed.Text = ""
               labcedmed.Caption = ""
            End If
            Regcli.Close
            Sqlcli = "Select * from hc_metasr where hc_nro =" & Xnrohc & " and hc_metadesc ='" & "REGISTRO CONS.ADULTO MAYOR" & "'"
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               cboregadul.ListIndex = 0
            Else
               cbosia.ListIndex = 1
            End If
            Regcli.Close
            Sqlcli = "Select * from hc_prescrip where hc_nro =" & Xnrohc & " and hc_tippresd ='" & "CONTROL" & "'"
            With Regcli
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open Sqlcli, ConbdSapp, , , adCmdText
            End With
            If Regcli.RecordCount > 0 Then
               If IsNull(Regcli("hc_descrip")) = False Then
                  t_proxcon3.Text = Regcli("hc_descrip")
               Else
                  t_proxcon3.Text = ""
               End If
            Else
               t_proxcon3.Text = ""
            End If
            Regcli.Close
         Else
            MsgBox "No se encontrò HC creada para Control META", vbInformation
            Xnrohc = 0
            Xcodmedhc = 0
            t_medic.Text = ""
         End If
         ConbdSapp.Close
      Else
   
   
      End If
   End If
End If
End Sub

Private Sub flex1_DblClick()
Dim Xnroreg As Integer

Xnroreg = Val(flex1.TextMatrix(flex1.RowSel, 5))
Adodc1.RecordSource = "Select * from t_meta1 where id =" & Xnroreg
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
   framnino1.Enabled = False
   Frame2.Enabled = False
   Frame4.Enabled = False
   muestrodatos
   
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=192.168.10.24;PORT=3306;DATABASE=sappbd;USER=root;PASSWORD=sapp1987;OPTION=3;"
'Adodc1.RecordSource = "t_meta1"
'Adodc1.Refresh
Adodc1.RecordSource = "Select * from t_descmeta1"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
   Adodc1.Recordset.MoveFirst
   Do While Not Adodc1.Recordset.EOF
      cbomet.AddItem Adodc1.Recordset("descrip")
      Adodc1.Recordset.MoveNext
   Loop
End If

Adodc1.RecordSource = "Select * from t_descmeta2"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
   Adodc1.Recordset.MoveFirst
   Do While Not Adodc1.Recordset.EOF
      cbomet2.AddItem Adodc1.Recordset("descrip")
      Adodc1.Recordset.MoveNext
   Loop
End If
Adodc1.RecordSource = "t_meta1"
Adodc1.Refresh
ConectarBDM
ConbdSappM.Open
Sqlcli = "Select * from t_meta1 order by fecha DESC"
With Regcli
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open Sqlcli, ConbdSappM, , , adCmdText
End With
Dim Xc As Integer
If Regcli.RecordCount > 0 Then
   Regcli.MoveLast
   Regcli.MoveFirst
   Xc = 0
   flex1.Rows = Regcli.RecordCount + 1
   flex1.Cols = 6
   flex1.TextMatrix(Xc, 0) = "FECHA"
   flex1.TextMatrix(Xc, 1) = "CEDULA"
   flex1.TextMatrix(Xc, 2) = "NOMBRE"
   flex1.ColWidth(2) = 2500
   flex1.TextMatrix(Xc, 3) = "META"
   flex1.ColWidth(3) = 2500
   flex1.TextMatrix(Xc, 4) = "EDAD"
   flex1.TextMatrix(Xc, 5) = "NRO.REG."
   Xc = Xc + 1
   Do While Not Regcli.EOF
      flex1.TextMatrix(Xc, 0) = Regcli("fecha")
      flex1.TextMatrix(Xc, 1) = Regcli("cedula")
      flex1.TextMatrix(Xc, 2) = Regcli("nombre")
      flex1.TextMatrix(Xc, 3) = Regcli("meta2desc")
      flex1.TextMatrix(Xc, 4) = Regcli("edadtex")
      flex1.TextMatrix(Xc, 5) = Regcli("id")
      Regcli.MoveNext
      Xc = Xc + 1
   Loop
End If
Regcli.Close
ConbdSappM.Close


End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Public Sub borrar_camp()
t_ced.Text = ""
t_codc.Text = ""
mf.Text = "__/__/____"
t_base.Text = ""
labmat.Caption = ""
labnom.Caption = ""
labcnv.Caption = ""
laba.Caption = ""
labm.Caption = ""
labd.Caption = ""
cbomet.ListIndex = -1
cbomet2.ListIndex = -1

End Sub


Private Sub mf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_base.SetFocus
End If

End Sub

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomet.SetFocus
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codc.SetFocus
End If

End Sub

Private Sub t_ced_LostFocus()
If Trim(t_ced.Text) <> "" Then
   ConectarBD
   ConbdSapp.Open
   Sqlcli = "Select * from clientes where cl_cedula =" & t_ced.Text
   With Regcli
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Sqlcli, ConbdSapp, , , adCmdText
   End With
   If Regcli.RecordCount > 0 Then
      t_codc.Text = Regcli("cl_codced")
      labnom.Caption = Regcli("cl_apellid")
      labmat.Caption = Regcli("cl_codigo")
      labcnv.Caption = Regcli("cl_codconv")
      If IsNull(Regcli("cl_fnac")) = False Then
         labfnac.Caption = Regcli("cl_fnac")
      Else
         labfnac.Caption = Date
      End If
      If IsNull(Regcli("cl_fnac")) = False Then
         CalculaEdad (Regcli("cl_fnac"))
         If Val(laba.Caption) = 0 Then
            If Val(labm.Caption) = 0 Then
               If Val(labd.Caption) > 0 And Val(labd.Caption) <= 10 Then
                  cbomet.ListIndex = 0
                  cbomet2.ListIndex = 1
               End If
            Else
               If Val(labm.Caption) <= 11 Then
'                  Xfecvolver = data_lin.Recordset("fecha") + 36
                  cbomet.ListIndex = 0
                  cbomet2.ListIndex = 2
               End If
            End If
         Else
            If Val(laba.Caption) = 1 And Val(labm.Caption) >= 0 Then
               cbomet.ListIndex = 0
               cbomet2.ListIndex = 2
            Else
               If Val(laba.Caption) = 2 And Val(labm.Caption) >= 0 Then
                  cbomet.ListIndex = 0
                  cbomet2.ListIndex = 3
               Else
                      'cambiar el codigo de facturación
                  If Val(laba.Caption) = 3 And Val(labm.Caption) >= 0 Then
                     cbomet.ListIndex = 0
                     cbomet2.ListIndex = 4
                  Else
                    'cambiar el codigo de facturación
                     If Val(laba.Caption) = 4 And Val(labm.Caption) >= 0 Then
                        cbomet.ListIndex = 0
                        cbomet2.ListIndex = 5
                     Else
                        'cambiar el codigo de facturación
                        If Val(laba.Caption) = 5 And Val(labm.Caption) >= 0 Then
                           cbomet.ListIndex = 0
                           cbomet2.ListIndex = 6
                        Else
                           If Val(laba.Caption) >= 15 And Val(laba.Caption) <= 100 Then
                              If Regcli("cl_sexo") = "FEMENINO" Then
                                 cbomet.ListIndex = 0
                                 cbomet2.ListIndex = 7
                              End If
                           Else
                              If Val(laba.Caption) >= 12 And Val(labm.Caption) <= 19 Then
                                 cbomet.ListIndex = 1
                                 cbomet2.ListIndex = 0
                              Else
                                 If Val(laba.Caption) >= 45 And Val(laba.Caption) <= 64 Then
                                    cbomet.ListIndex = 1
                                    cbomet2.ListIndex = 1
                                 Else
                                    If Val(laba.Caption) >= 65 And Val(laba.Caption) <= 74 Then
                                       cbomet.ListIndex = 2
                                       cbomet2.ListIndex = 0
                                    Else
                                       If Val(laba.Caption) >= 75 And Val(laba.Caption) <= 115 Then
                                          cbomet.ListIndex = 2
                                          cbomet2.ListIndex = 1
                                       Else
                                          cbomet.ListIndex = -1
                                          cbomet2.ListIndex = -1
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
         cbomet.ListIndex = -1
         cbomet2.ListIndex = -1
      End If
   Else
      cbomet.ListIndex = -1
      cbomet2.ListIndex = -1
   End If
   ConbdSapp.Close
Else
   cbomet.ListIndex = -1
   cbomet2.ListIndex = -1
End If


End Sub

Private Sub t_codc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mf.SetFocus
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
'''   labedad.Caption = Anios
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
   laba.Caption = Anios
   labm.Caption = Meses
   labd.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   laba.Caption = 0
   labm.Caption = 0
   labd.Caption = 0

End If

End Sub

Private Sub CalculaEdad2(ByVal Ffechaayer As Date)
Dim FAct As String
Dim Anios As String
Dim Meses As String
Dim Dias As String
Dim newday As String
Dim newmonth As String
Dim newyear As String
Dim FNaci As Date
FNaci = CDate(labfnac.Caption)

FAct = Format(Ffechaayer, "dd/MM/yyyy")

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
'''   labedad.Caption = Anios
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
   t_aa.Text = Anios
   t_mm.Text = Meses
   t_dd.Text = Dias
Else
   MsgBox "Fecha Inválida"
   t_aa.Text = 0
   t_mm.Text = 0
   t_dd.Text = 0

End If

End Sub


Public Function ConectarBDM()
ConbdSappM.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=sappbd;USER=root;PASSWORD=sapp1987;OPTION=3;"

End Function


Public Sub limpiar()
mfctr.Text = "__/__/____"
t_aa.Text = 0
t_mm.Text = 0
t_dd.Text = 0
t_peso.Text = ""
t_lacta.Text = ""
t_vacun.Text = ""
t_eco.Text = ""
t_hemog.Text = ""
t_odont.Text = ""
t_medic.Text = ""
t_proxctr.Text = ""
t_fdesar.Text = ""
t_obs.Text = ""
mfctremb.Text = "__/__/____"
t_lochc.Text = ""
cboobste.ListIndex = -1
cbovdrl.ListIndex = -1
cboodont.ListIndex = -1
cbohcpb.ListIndex = -1
cboprot.ListIndex = -1
t_cedrn.Text = ""
mfctradult.Text = "__/__/____"
t_nommed.Text = ""
t_apemed.Text = ""
cbosia.ListIndex = -1
cboscree.ListIndex = -1
cboregadul.ListIndex = -1
t_feca.Text = ""
t_proxcon3.Text = ""
t_fram.Text = ""
t_nrocons.Text = ""
cbopap.ListIndex = -1
mfvdrl.Text = "__/__/____"
cbomamo.ListIndex = -1
t_anticon.Text = ""
mfanti.Text = "__/__/____"
t_semgon.Text = ""

End Sub

Public Sub muestrodatos()
If IsNull(Adodc1.Recordset("frame")) = False Then
   t_fram.Text = Adodc1.Recordset("frame")
   If t_fram.Text = 1 Then
      framnino1.Visible = True
      Frame2.Visible = False
      Frame4.Visible = False
      framnino1.Enabled = True
      Frame2.Enabled = False
      Frame4.Enabled = False
   Else
      If t_fram.Text = 2 Then
         framnino1.Visible = False
         Frame2.Visible = True
         Frame4.Visible = False
         framnino1.Enabled = False
         Frame2.Enabled = True
         Frame4.Enabled = False
      Else
         If t_fram.Text = 3 Then
            framnino1.Visible = False
            Frame2.Visible = False
            Frame4.Visible = True
            framnino1.Enabled = False
            Frame2.Enabled = False
            Frame4.Enabled = True
         End If
      End If
   End If
End If
If IsNull(Adodc1.Recordset("fecha")) = False Then
   mf.Text = Format(Adodc1.Recordset("fecha"), "dd/mm/yyyy")
Else
   mf.Text = "__/__/____"
End If
If IsNull(Adodc1.Recordset("cedula")) = False Then
   t_ced.Text = Adodc1.Recordset("cedula")
Else
   t_ced.Text = ""
End If
If IsNull(Adodc1.Recordset("codced")) = False Then
   t_codc.Text = Adodc1.Recordset("codced")
Else
   t_codc.Text = ""
End If
If IsNull(Adodc1.Recordset("base")) = False Then
   t_base.Text = Adodc1.Recordset("base")
Else
   t_base.Text = ""
End If
If IsNull(Adodc1.Recordset("matric")) = False Then
   labmat.Caption = Adodc1.Recordset("matric")
Else
   labmat.Caption = 0
End If
If IsNull(Adodc1.Recordset("nombre")) = False Then
   labnom.Caption = Adodc1.Recordset("nombre")
Else
   labnom.Caption = ""
End If
If IsNull(Adodc1.Recordset("cnvcod")) = False Then
   labcnv.Caption = Adodc1.Recordset("cnvcod")
Else
   labcnv.Caption = ""
End If
If IsNull(Adodc1.Recordset("fecnac")) = False Then
   labfnac.Caption = Format(Adodc1.Recordset("fecnac"), "dd/mm/yyyy")
Else
   labfnac.Caption = ""
End If
If IsNull(Adodc1.Recordset("eda")) = False Then
   laba.Caption = Adodc1.Recordset("eda")
Else
   laba.Caption = 0
End If
If IsNull(Adodc1.Recordset("edm")) = False Then
   labm.Caption = Adodc1.Recordset("edm")
Else
   labm.Caption = 0
End If
If IsNull(Adodc1.Recordset("edd")) = False Then
   labd.Caption = Adodc1.Recordset("edd")
Else
   labd.Caption = 0
End If
If IsNull(Adodc1.Recordset("edca")) = False Then
   t_aa.Text = Adodc1.Recordset("edca")
Else
   t_aa.Text = 0
End If
If IsNull(Adodc1.Recordset("edcm")) = False Then
   t_mm.Text = Adodc1.Recordset("edcm")
Else
   t_mm.Text = 0
End If
If IsNull(Adodc1.Recordset("edcd")) = False Then
   t_dd.Text = Adodc1.Recordset("edcd")
Else
   t_dd.Text = 0
End If
If IsNull(Adodc1.Recordset("meta")) = False Then
   cbomet.ListIndex = Adodc1.Recordset("meta")
Else
   cbomet.ListIndex = -1
End If
'      Adodc1.Recordset("metadesc") = cbomet.Text
If IsNull(Adodc1.Recordset("meta2")) = False Then
   cbomet2.ListIndex = Adodc1.Recordset("meta2")
Else
   cbomet2.ListIndex = -1
End If
'      Adodc1.Recordset("meta2desc") = cbomet2.Text
If IsNull(Adodc1.Recordset("fecctrl")) = False Then
   If IsNull(Adodc1.Recordset("frame")) = False Then
      If t_fram.Text = 1 Then
         mfctr.Text = Format(Adodc1.Recordset("fecctrl"), "dd/mm/yyyy")
      Else
         If t_fram.Text = 2 Then
            mfctremb.Text = Format(Adodc1.Recordset("fecctrl"), "dd/mm/yyyy")
         Else
            If t_fram.Text = 3 Then
               mfctradult.Text = Format(Adodc1.Recordset("fecctrl"), "dd/mm/yyyy")
            End If
         End If
      End If
   End If
Else
   mfctr.Text = "__/__/____"
End If
If IsNull(Adodc1.Recordset("peso")) = False Then
   t_peso.Text = Adodc1.Recordset("peso")
Else
   t_peso.Text = 0
End If
If IsNull(Adodc1.Recordset("lactan")) = False Then
   t_lacta.Text = Adodc1.Recordset("lactan")
Else
   t_lacta.Text = ""
End If
If IsNull(Adodc1.Recordset("vacuna")) = False Then
   t_vacun.Text = Adodc1.Recordset("vacuna")
Else
   t_vacun.Text = ""
End If
If IsNull(Adodc1.Recordset("ecocad")) = False Then
   t_eco.Text = Adodc1.Recordset("ecocad")
Else
   t_eco.Text = ""
End If
If IsNull(Adodc1.Recordset("hemog")) = False Then
   t_hemog.Text = Adodc1.Recordset("hemog")
Else
   t_hemog.Text = ""
End If
If IsNull(Adodc1.Recordset("odont")) = False Then
   t_odont.Text = Adodc1.Recordset("odont")
Else
   t_odont.Text = ""
End If
If IsNull(Adodc1.Recordset("medico")) = False Then
   t_medic.Text = Adodc1.Recordset("medico")
Else
   t_medic.Text = ""
End If
If IsNull(Adodc1.Recordset("fecprox")) = False Then
   If IsNull(Adodc1.Recordset("frame")) = False Then
      If t_fram.Text = 1 Then
         t_proxctr.Text = Adodc1.Recordset("fecprox")
      Else
         If t_fram.Text = 3 Then
            t_proxcon3.Text = Adodc1.Recordset("fecprox")
         Else
            t_proxctr.Text = Adodc1.Recordset("fecprox")
         End If
      End If
   End If
Else
   t_proxctr.Text = ""
   t_proxcon3.Text = ""
End If
If IsNull(Adodc1.Recordset("fdesar")) = False Then
   t_fdesar.Text = Adodc1.Recordset("fdesar")
Else
   t_fdesar.Text = ""
End If
If IsNull(Adodc1.Recordset("oft")) = False Then
   t_ctroft.Text = Adodc1.Recordset("oft")
Else
   t_ctroft.Text = ""
End If
If IsNull(Adodc1.Recordset("obs")) = False Then
   t_obs.Text = Adodc1.Recordset("obs")
Else
   t_obs.Text = ""
End If
If IsNull(Adodc1.Recordset("lochc")) = False Then
   t_lochc.Text = Adodc1.Recordset("lochc")
Else
   t_lochc.Text = ""
End If
If IsNull(Adodc1.Recordset("obstetr")) = False Then
   cboobste.ListIndex = Adodc1.Recordset("obstetr")
Else
   cboobste.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("odont2")) = False Then
   cboodont.ListIndex = Adodc1.Recordset("odont2")
Else
   cboodont.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("vdrl")) = False Then
   cbovdrl.ListIndex = Adodc1.Recordset("vdrl")
Else
   cbovdrl.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("hcpb")) = False Then
   cbohcpb.ListIndex = Adodc1.Recordset("hcpb")
Else
   cbohcpb.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("prot")) = False Then
   cboprot.ListIndex = Adodc1.Recordset("prot")
Else
   cboprot.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("cedrn")) = False Then
   t_cedrn.Text = Adodc1.Recordset("cedrn")
Else
   t_cedrn.Text = ""
End If
'               Adodc1.Recordset("fecctrl") = Format(mfctradult.Text, "dd/mm/yyyy")
If IsNull(Adodc1.Recordset("nommed")) = False Then
   t_nommed.Text = Adodc1.Recordset("nommed")
Else
   t_nommed.Text = ""
End If
If IsNull(Adodc1.Recordset("apemed")) = False Then
   t_apemed.Text = Adodc1.Recordset("apemed")
Else
   t_apemed.Text = ""
End If
If IsNull(Adodc1.Recordset("sia")) = False Then
   cbosia.ListIndex = Adodc1.Recordset("sia")
Else
   cbosia.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("screen")) = False Then
   cboscree.ListIndex = Adodc1.Recordset("screen")
Else
   cboscree.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("adulto")) = False Then
   cboregadul.ListIndex = Adodc1.Recordset("adulto")
Else
   cboregadul.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("fecatest")) = False Then
   t_feca.Text = Adodc1.Recordset("fecatest")
Else
   t_feca.Text = ""
End If
If IsNull(Adodc1.Recordset("nroconse")) = False Then
   t_nrocons.Text = Adodc1.Recordset("nroconse")
Else
   t_nrocons.Text = ""
End If
If IsNull(Adodc1.Recordset("nroconse")) = False Then
   t_nrocons.Text = Adodc1.Recordset("nroconse")
Else
   t_nrocons.Text = ""
End If
If IsNull(Adodc1.Recordset("cbopap")) = False Then
   cbopap.ListIndex = Adodc1.Recordset("cbopap")
Else
   cbopap.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("fecvdrl")) = False Then
   mfvdrl.Text = Format(Adodc1.Recordset("fecvdrl"), "dd/mm/yyyy")
Else
   mfvdrl.Text = "__/__/____"
End If
If IsNull(Adodc1.Recordset("cbomamo")) = False Then
   cbomamo.ListIndex = Adodc1.Recordset("cbomamo")
Else
   cbomamo.ListIndex = -1
End If
If IsNull(Adodc1.Recordset("obsemb")) = False Then
   t_anticon.Text = Adodc1.Recordset("obsemb")
Else
   t_anticon.Text = ""
End If
If IsNull(Adodc1.Recordset("fecanti")) = False Then
   mfanti.Text = Format(Adodc1.Recordset("fecanti"), "dd/mm/yyyy")
Else
   mfanti.Text = "__/__/____"
End If
If IsNull(Adodc1.Recordset("semgono")) = False Then
   t_semgon.Text = Adodc1.Recordset("semgono")
Else
   t_semgon.Text = ""
End If

If cbomet2.Text = "CTROL.45 A 64AÑOS" Then
   cboscree.Visible = True
   t_feca.Visible = True
   cboregadul.Visible = False
   Label30.Visible = False
   cbosia.Visible = False
Else
   If cbomet2.Text = "CTROL.65 A 74AÑOS" Or cbomet2.Text = "CTROL. >75 AÑOS" Then
      Label31.Visible = False
      Label32.Visible = False
      Label33.Visible = True
      cboscree.Visible = False
      t_feca.Visible = False
      cboregadul.Visible = True
      Label30.Visible = False
      cbosia.Visible = False
   Else
      Label30.Visible = True
      cbosia.Visible = True
      Label31.Visible = False
      Label32.Visible = False
      Label33.Visible = False
      cboscree.Visible = False
      t_feca.Visible = False
      cboregadul.Visible = False
   End If
End If
      

End Sub

Public Sub cargargrid()
ConectarBDM
ConbdSappM.Open

Sqlcli = "Select * from t_meta1 order by fecha DESC"
With Regcli
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open Sqlcli, ConbdSappM, , , adCmdText
End With
Dim Xc As Integer
If Regcli.RecordCount > 0 Then
   Regcli.MoveLast
   Regcli.MoveFirst
   Xc = 0
   flex1.Rows = Regcli.RecordCount + 1
   flex1.Cols = 6
   flex1.TextMatrix(Xc, 0) = "FECHA"
   flex1.TextMatrix(Xc, 1) = "CEDULA"
   flex1.TextMatrix(Xc, 2) = "NOMBRE"
   flex1.ColWidth(2) = 2500
   flex1.TextMatrix(Xc, 3) = "META"
   flex1.ColWidth(3) = 2500
   flex1.TextMatrix(Xc, 4) = "EDAD"
   flex1.TextMatrix(Xc, 5) = "NRO.REG."
   Xc = Xc + 1
   Do While Not Regcli.EOF
      flex1.TextMatrix(Xc, 0) = Regcli("fecha")
      flex1.TextMatrix(Xc, 1) = Regcli("cedula")
      flex1.TextMatrix(Xc, 2) = Regcli("nombre")
      flex1.TextMatrix(Xc, 3) = Regcli("meta2desc")
      flex1.TextMatrix(Xc, 4) = Regcli("edadtex")
      flex1.TextMatrix(Xc, 5) = Regcli("id")
      Regcli.MoveNext
      Xc = Xc + 1
   Loop
End If
Regcli.Close
ConbdSappM.Close

End Sub
