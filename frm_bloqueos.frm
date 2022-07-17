VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_bloqueos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de captación de afiliados"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10890
   Icon            =   "frm_bloqueos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   10890
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_buscados 
      Caption         =   "data_buscados"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
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
      Height          =   405
      Left            =   5520
      TabIndex        =   50
      Top             =   8400
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por promotor:"
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
      Left            =   2880
      TabIndex        =   49
      Top             =   8400
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por Cédula"
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
      Left            =   240
      TabIndex        =   48
      Top             =   8400
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_bloqueos.frx":058A
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "frm_bloqueos.frx":05A6
      TabIndex        =   47
      Top             =   5880
      Width           =   10335
   End
   Begin VB.Data data_busca 
      Caption         =   "data_busca"
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
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_orig 
      Caption         =   "data_orig"
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
      Top             =   2040
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CommandButton b_infos 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      Picture         =   "frm_bloqueos.frx":2849
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Informes"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_cancelar 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_bloqueos.frx":2DD3
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cancelar acción"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_bloqueos.frx":335D
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Guardar datos ingresados"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_bloqueos.frx":38E7
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Ingresar Fec. Afiliación"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_bloqueos.frx":3E71
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Crear nuevo registro"
      Top             =   5280
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la CAPTACIÓN"
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
      ForeColor       =   &H00C00000&
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.ComboBox cbofunc 
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
         Left            =   1560
         TabIndex        =   45
         Top             =   4200
         Width           =   3975
      End
      Begin MSMask.MaskEdBox mfing 
         Height          =   375
         Left            =   8640
         TabIndex        =   43
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.ComboBox cbocontact 
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
         ItemData        =   "frm_bloqueos.frx":43FB
         Left            =   4200
         List            =   "frm_bloqueos.frx":4411
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   3720
         Width           =   2415
      End
      Begin VB.ComboBox cbogenera 
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
         ItemData        =   "frm_bloqueos.frx":445E
         Left            =   1560
         List            =   "frm_bloqueos.frx":4468
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox cbomut 
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
         ItemData        =   "frm_bloqueos.frx":4474
         Left            =   6960
         List            =   "frm_bloqueos.frx":4487
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   3120
         Width           =   3255
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
         Left            =   1560
         TabIndex        =   34
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox t_dir2 
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
         Left            =   6360
         MaxLength       =   100
         TabIndex        =   32
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox t_dir1 
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   31
         Top             =   2640
         Width           =   4815
      End
      Begin VB.ComboBox cboactual 
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
         ItemData        =   "frm_bloqueos.frx":44C0
         Left            =   7320
         List            =   "frm_bloqueos.frx":44C2
         TabIndex        =   29
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Data data_buscasi 
         Caption         =   "data_buscasi"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   6120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5520
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox t_otro 
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
         Height          =   405
         Left            =   3240
         MaxLength       =   45
         TabIndex        =   28
         ToolTipText     =   "Teléfono alternativo (campo no obligatorio)"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Data data_par 
         Caption         =   "data_par"
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
         Top             =   360
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
         Height          =   375
         Left            =   8160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox t_cel 
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
         Height          =   405
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   27
         ToolTipText     =   "Digite número de celular, campo obligatorio"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox t_tel 
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
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   26
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox t_cced 
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
         Left            =   6360
         MaxLength       =   1
         TabIndex        =   25
         Top             =   1200
         Width           =   375
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
         Left            =   5040
         MaxLength       =   8
         TabIndex        =   24
         ToolTipText     =   "Digite cédula, campo obligatorio"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox t_mat 
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   23
         ToolTipText     =   "Digite matrícula, campo obligatorio"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox t_nom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         MaxLength       =   60
         TabIndex        =   14
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox t_conv 
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
         Left            =   8520
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox t_base 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin MSMask.MaskEdBox mhora 
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.Label labno 
         Height          =   255
         Left            =   7440
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label labcodfunc 
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Captado por:"
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
         Left            =   240
         TabIndex        =   44
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Afiliación:"
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
         Left            =   6840
         TabIndex        =   42
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C00000&
         Caption         =   "Contactado:"
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
         Left            =   2880
         TabIndex        =   40
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C00000&
         Caption         =   "Generante?"
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
         Left            =   240
         TabIndex        =   38
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Mutualista a ingresar:"
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
         Left            =   5040
         TabIndex        =   36
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label labcodzon 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Zona"
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
         Left            =   240
         TabIndex        =   33
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C00000&
         Caption         =   "Mutualista actual:"
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
         Left            =   5640
         TabIndex        =   17
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C00000&
         Caption         =   "Celular:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C00000&
         Caption         =   "Teléfono:"
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
         Left            =   6720
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C00000&
         Caption         =   "Nombre:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Cédula:"
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
         Left            =   3720
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "Matrícula:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label labusu 
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
         Height          =   375
         Left            =   7800
         TabIndex        =   8
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Base:"
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
         Left            =   5040
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Hora:"
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
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   5655
         Left            =   0
         Picture         =   "frm_bloqueos.frx":44C4
         Stretch         =   -1  'True
         Top             =   240
         Width           =   10335
      End
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C00000&
      Caption         =   "Doble click para editar registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8520
      TabIndex        =   52
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   8640
      Picture         =   "frm_bloqueos.frx":4B88
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1575
   End
End
Attribute VB_Name = "frm_bloqueos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acciones_Click()
If t_cod.Text <> "" Then
   frm_solbacc.Show vbModal
Else
   MsgBox "No seleccionó registro", vbInformation
   
End If
End Sub


Private Sub b_cancelar_Click()
Borra_campos

XAlta = 0
labno.Caption = ""

b_nuevo.Enabled = True
b_edita.Enabled = True
b_cancelar.Enabled = False
b_graba.Enabled = False
b_infos.Enabled = True
Frame1.Enabled = False

End Sub


Private Sub b_edita_Click()
Dim DeseaMod As String
Dim Xlafe As String
DeseaMod = MsgBox("Desea ingresar la fecha de afiliación para:" & data_buscados.Recordset("nombre") & " ?", vbInformation + vbYesNo, "BLOQUEOS")
If DeseaMod = vbYes Then
   Xlafe = InputBox("Ingrese fecha de afiliación:", "Bloqueos", Format(Date, "dd/mm/yyyy"))
   If Xlafe <> "" Then
      If Format(CDate(Xlafe), "yyyy/mm/dd") <= Format(Date, "yyyy/mm/dd") Then
         If IsNull(data_buscados.Recordset("fec_afil")) = True Then
            data_buscados.Recordset.Edit
            data_buscados.Recordset("fec_afil") = CDate(Xlafe)
            data_buscados.Recordset.Update
         Else
            MsgBox "Ya tiene una fecha ingresada, no se puede modificar.", vbCritical
         End If
      Else
         MsgBox "La fecha no puede ser mayor a la fecha actual.", vbInformation
      End If
   Else
      MsgBox "Presionó Cancelar, no se grabó.", vbInformation
   End If
End If


End Sub

Private Sub b_graba_Click()
Dim Xelerror As Integer

On Error GoTo Querr

Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo, Xantnro As Long

Xn1 = 2
Xn2 = 9
Xn3 = 8
Xn4 = 7
Xn5 = 6
Xn6 = 3
Xn7 = 4
Xpond = 10
If t_tel.Text = "NO APLICA" And t_cel.Text = "NO APLICA" Then
   Xelerror = 9
End If

Xelerror = 0
If labno.Caption = "NO" Then
   Xelerror = 60
Else
   If t_conv.Text <> "" Then
      If t_conv.Text <> "PART" Then
         Consulta_convenio
         If labno.Caption = "NO" Then
            Xelerror = 60
         End If
      End If
   End If
   If labno.Caption <> "NO" Then
      If cboactual.Text = "CIRCULO CATOLICO" Or cboactual.Text = "SMI" Or _
         cboactual.Text = "UNIVERSAL" Or _
         cboactual.Text = "EVANGELICO" Then
         MsgBox "No se permiten cambios entre mutualistas de convenio SAPP. No podrá grabar.", vbCritical
         labno.Caption = "NO"
         Xelerror = 60
      Else
         labno.Caption = ""
      End If
   End If
End If
If t_ced.Text <> "" And t_cced.Text <> "" Then
   If IsNumeric(t_ced.Text) = False Then
      MsgBox "La cédula debe contener solo números", vbInformation
      t_ced.Text = ""
   Else
      Xcedtex = Trim(str(t_ced.Text))
      Xlargo = Len(Xcedtex)
      If Xlargo = 6 Then
         Xcedtex = "0" & Trim(Xcedtex)
      End If
      Xced1 = Val(Mid(Trim(Xcedtex), 1, 1))
      Xced2 = Val(Mid(Xcedtex, 2, 1))
      Xced3 = Val(Mid(Xcedtex, 3, 1))
      Xced4 = Val(Mid(Xcedtex, 4, 1))
      Xced5 = Val(Mid(Xcedtex, 5, 1))
      Xced6 = Val(Mid(Xcedtex, 6, 1))
      Xced7 = Val(Mid(Xcedtex, 7, 1))
      Xced1 = Xced1 * Xn1
      Xced2 = Xced2 * Xn2
      Xced3 = Xced3 * Xn3
      Xced4 = Xced4 * Xn4
      Xced5 = Xced5 * Xn5
      Xced6 = Xced6 * Xn6
      Xced7 = Xced7 * Xn7
      Xtot = Xced1 + Xced2 + Xced3 + Xced4 + Xced5 + Xced6 + Xced7
      If Len(Trim(str(Xtot))) = 1 Then
         Xtottex = "0000" & Trim(str(Xtot))
      End If
      If Len(Trim(str(Xtot))) = 2 Then
         Xtottex = "000" & Trim(str(Xtot))
      End If
      If Len(Trim(str(Xtot))) = 3 Then
         Xtottex = "00" & Trim(str(Xtot))
      End If
      If Len(Trim(str(Xtot))) = 4 Then
         Xtottex = "0" & Trim(str(Xtot))
      End If
      Xtot = Val(Mid(Xtottex, 5, 1))
      If Xtot <> 0 Then
         Xtot = Xpond - Xtot
      Else
         Xtot = 0
      End If
      If Xtot <> t_cced.Text Then
         MsgBox "Hay un error en la cédula, verifique!", vbCritical
         Xelerror = 2
      End If
   End If
Else
   Xelerror = 2
End If
If t_mat.Text = "" Then
   t_mat.Text = 0
End If

If t_mat.Text <> "" Then
   If t_mat.Text > 0 Then
      data_buscasi.RecordSource = "select * from clientes where cl_codigo =" & t_mat.Text
      data_buscasi.Refresh
      If data_buscasi.Recordset.RecordCount > 0 Then
         data_buscasi.RecordSource = "select * from bloqueos where matricula =" & t_mat.Text
         data_buscasi.Refresh
         If data_buscasi.Recordset.RecordCount > 0 Then
            Xelerror = 43
         End If
      Else
         Xelerror = 44
      End If
   Else
      If t_ced.Text <> "" Then
         If t_ced.Text > 0 Then
            data_buscasi.RecordSource = "select * from bloqueos where cedula =" & t_ced.Text
            data_buscasi.Refresh
            If data_buscasi.Recordset.RecordCount > 0 Then
               Xelerror = 43
            End If
         End If
      End If
   End If
End If
If Trim(t_conv.Text) = "" Then
   t_conv.Text = "PART"
End If

If Trim(t_mat.Text) = "" Then
   Xelerror = 1
End If
If Trim(t_ced.Text) = "" Then
   Xelerror = 2
Else
   If t_ced.Text > 0 Then
   Else
      Xelerror = 2
   End If
End If
If Trim(t_cced.Text) = "" Then
   Xelerror = 3
End If
If Trim(t_conv.Text) = "" Then
   Xelerror = 4
End If
If Trim(t_nom.Text) = "" Then
   Xelerror = 5
End If
If Trim(t_tel.Text) = "" Then
   Xelerror = 6
End If
If Trim(t_cel.Text) = "" Then
   Xelerror = 7
End If
If Trim(cboactual.Text) = "" Then
   Xelerror = 8
End If
If Trim(cbozona.Text) = "" Then
   Xelerror = 8
End If
If Trim(cbofunc.Text) = "" Then
   Xelerror = 8
End If
If Trim(t_dir1.Text) = "" Then
   Xelerror = 9
End If
If Trim(cboactual.Text) = "" Then
   Xelerror = 10
End If
If Trim(cbomut.Text) = "" Then
   Xelerror = 11
End If
If Trim(cbogenera.Text) = "" Then
   Xelerror = 12
End If
If Trim(cbocontact.Text) = "" Then
   Xelerror = 13
End If
If Trim(labcodfunc.Caption) = "" Then
   Xelerror = 14
End If
If Xelerror = 0 Then
   Data1.RecordSource = "bloqueos"
   Data1.Refresh
   If XAlta = 1 Then
      Data1.Recordset.AddNew
      Data1.Recordset("fecha") = mfec.Text
      Data1.Recordset("hora") = mhora.Text
      Data1.Recordset("base") = t_base.Text
      Data1.Recordset("usuario") = labusu.Caption
      Data1.Recordset("matricula") = t_mat.Text
      Data1.Recordset("cedula") = t_ced.Text
      Data1.Recordset("codced") = t_cced.Text
      Data1.Recordset("convenio") = t_conv.Text
      Data1.Recordset("nombre") = t_nom.Text
      Data1.Recordset("telef") = t_tel.Text
      Data1.Recordset("celular") = t_cel.Text
      If Trim(t_otro.Text) <> "" Then
         Data1.Recordset("celular2") = t_otro.Text
      End If
      Data1.Recordset("dir1") = t_dir1.Text
      If Trim(t_dir2.Text) <> "" Then
         Data1.Recordset("dir2") = t_dir2.Text
      End If
      Data1.Recordset("mut_act") = cboactual.Text
      Data1.Recordset("zona") = cbozona.Text
      Data1.Recordset("mut_new") = cbomut.Text
      Data1.Recordset("genera") = cbogenera.Text
      Data1.Recordset("contacto") = cbocontact.Text
      Data1.Recordset("codfunc") = Val(labcodfunc.Caption)
      Data1.Recordset("nomfunc") = cbofunc.Text
      If mfing.Text <> "__/__/____" Then
         Data1.Recordset("fec_afil") = mfing.Text
      End If
      Data1.Recordset.Update
'      data_buscados.RecordSource = "select * from bloqueos order by fecha DESC"
      data_buscados.Refresh
      XAlta = 0
      Borra_campos
      labno.Caption = ""
      b_nuevo.Enabled = True
      b_edita.Enabled = True
      b_cancelar.Enabled = False
      b_graba.Enabled = False
      b_infos.Enabled = True
      Frame1.Enabled = False
   End If
Else
   If Xelerror = 43 Or Xelerror = 44 Then
      If Xelerror = 44 Then
         MsgBox "La matrícula no se encuentra, verifique!", vbCritical
      Else
         MsgBox "Ya figura esta cédula ingresada en captación. No se puede ingresar.", vbCritical
      End If
   Else
      If Xelerror = 60 Then
         MsgBox "No es posible el cambio entre mutualistas SAPP", vbCritical
      Else
         If Xelerror = 2 Then
            MsgBox "No se puede grabar por error en la cédula", vbCritical
         Else
            If Xelerror = 1 Then
               MsgBox "Hay un error en la cédula, VERIFIQUE!", vbCritical
            Else
               MsgBox "Hay algún campo que falta ingresar datos. Verifique!", vbInformation
            End If
         End If
      End If
   End If
End If

Xelerror = 0
Exit Sub

Querr:
       If Err.Number = 3197 Then
          MsgBox "Error al grabar, verifique si realizó modificaciones o presione cancelar", vbInformation
       Else
          MsgBox "Error al grabar, verifique datos." & Err.Description, vbInformation

       End If


End Sub

Private Sub b_infos_Click()
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
Promo = InputBox("INGRESE CÓDIGO DE PROMOTOR (CERO PARA LISTAR TODOS)", "CODIGO de PROMOTOR", 0)
If Trim(Promo) = "" Then
   Promo = "0"
End If
frm_bloqueos.MousePointer = 11
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0

If desde <> "" And hasta <> "" Then
   If Val(Promo) = 0 Then
      If ControlUsuario("frm_bloqueos") = 1 Then
         data_buscasi.RecordSource = "select * from bloqueos where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# order by codfunc,fecha"
      Else
         data_buscasi.RecordSource = "select * from bloqueos where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and usuario ='" & WElusuario & "' order by codfunc,fecha"
      End If
      data_buscasi.Refresh
   Else
      If ControlUsuario("frm_bloqueos") = 1 Then
         data_buscasi.RecordSource = "select * from bloqueos where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and codfunc =" & Val(Promo) & " order by fecha"
      Else
         data_buscasi.RecordSource = "select * from bloqueos where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and codfunc =" & Val(Promo) & " and usuario ='" & WElusuario & "' order by fecha"
      End If
      data_buscasi.Refresh
   End If
   If data_buscasi.Recordset.RecordCount > 0 Then
      data_buscasi.Recordset.MoveFirst
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      Xarchexel22.Name = Trim("Captados")
      Xlibexel22.SaveAs ("C:\planillas\InfoCaptados.xls")
      Xarchtex = "C:\planillas\InfoCaptados.xls"
      Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "M" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Cells(Xlin, XCol) = "INFORME DE CAPTADOS ORDENADOS POR PROMOTOR DESDE: " & desde & " HASTA: " & hasta
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
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 16
      Xarchexel22.Cells(Xlin, XCol) = "MUT. A INGRESAR"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "COD.PRO"
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "PROMOTOR"
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "CONVENIO"
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 16
      Xarchexel22.Cells(Xlin, XCol) = "MUT.ACTUAL"
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "BASE"
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 26
      Xarchexel22.Cells(Xlin, XCol) = "CELULAR/TELEF."
      XCol = XCol + 1
      Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "Nro.FACTURA"
      XCol = XCol + 1
      Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "FECHA AFIL."
      XCol = XCol + 1
      Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "CONTACTADO POR:"
      
      Xlin = Xlin + 1
      XCol = 1
        
      Do While Not data_buscasi.Recordset.EOF
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_buscasi.Recordset("fecha"), "dd/mm/yyyy")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Trim(str(data_buscasi.Recordset("cedula"))) & "-" & Trim(str(data_buscasi.Recordset("codced")))
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("nombre")
         XCol = XCol + 1
         If IsNull(data_buscasi.Recordset("mut_new")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("mut_new")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("codfunc")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("nomfunc")
         XCol = XCol + 1
         If IsNull(data_buscasi.Recordset("convenio")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("convenio")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("mut_act")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("base")
         XCol = XCol + 1
         If IsNull(data_buscasi.Recordset("celular2")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = "Cel:" & data_buscasi.Recordset("celular") & "--" & data_buscasi.Recordset("celular2") & "/Tel:" & data_buscasi.Recordset("telef")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "Cel:" & data_buscasi.Recordset("celular") & "/Tel:" & data_buscasi.Recordset("telef")
         End If
         XCol = XCol + 1
         data_busca.RecordSource = "select * from linmmdd where ced_socio =" & data_buscasi.Recordset("cedula") & " and cod_prod in (984,985,986,987,989) and fecha >=#" & Format("01/02/2021", "yyyy/mm/dd") & "#"
         data_busca.Refresh
         If data_busca.Recordset.RecordCount > 0 Then
            Xarchexel22.Cells(Xlin, XCol) = data_busca.Recordset("factura")
         End If
         XCol = XCol + 1
         If IsNull(data_buscasi.Recordset("fec_afil")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_buscasi.Recordset("fec_afil"), "dd/mm/yyyy")
         End If
         XCol = XCol + 1
         If IsNull(data_buscasi.Recordset("contacto")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_buscasi.Recordset("contacto")
         End If
         
         Xlin = Xlin + 1
         XCol = 1
         Xtotreg = Xtotreg + 1
         data_buscasi.Recordset.MoveNext
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
      frm_bloqueos.MousePointer = 0
   
   Else
      frm_bloqueos.MousePointer = 0
      MsgBox "No hay registros"
   End If
Else
   frm_bloqueos.MousePointer = 0
   MsgBox "Faltan fechas"
End If

End Sub

Private Sub b_nuevo_Click()
Frame1.Enabled = True
Borra_campos

mfec.Text = Date
mhora.Text = Format(Time, "HH:mm")
t_base.Text = frm_menu.data_parse.Recordset("base")

labusu.Caption = WElusuario

XAlta = 1
b_nuevo.Enabled = False
b_edita.Enabled = False
b_cancelar.Enabled = True
b_graba.Enabled = True
b_infos.Enabled = False
'b_edac.Enabled = False
t_mat.Text = 0
t_mat.SetFocus
labno.Caption = ""

End Sub

Private Sub cbodes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cbodes.ListIndex >= 0 Then
      cbohas.SetFocus
   Else
      MsgBox "Campo obligatorio"
      cbodes.SetFocus
   End If
End If

   
End Sub

Private Sub cbohas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cbohas.ListIndex >= 0 Then
      cboorig.SetFocus
   Else
      MsgBox "Campo obligatorio"
      cbohas.SetFocus
   End If
End If

End Sub

Private Sub cboorig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If cboorig.ListIndex >= 0 Then
      cbomot.SetFocus
   Else
      MsgBox "Campo obligatorio"
      cboorig.SetFocus
   End If
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   If Trim(t_accion.Text) = "" Then
      MsgBox "No hay acciones registradas para cerrar la solicitud", vbInformation
      Check1.Value = 0
   End If
End If

End Sub

Private Sub cboactual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_dir1.SetFocus
End If

End Sub

Private Sub cboactual_LostFocus()
If cboactual.Text <> "" Then
   Consulta_mutual
   If cboactual.Text = "CIRCULO CATOLICO" Or cboactual.Text = "SMI" Or _
      cboactual.Text = "UNIVERSAL" Or _
      cboactual.Text = "EVANGELICO" Then
      MsgBox "No se permiten cambios entre mutualistas de convenio SAPP. No podrá grabar.", vbCritical
      labno.Caption = "NO"
   Else
      labno.Caption = ""
   End If
Else
   labno.Caption = ""
End If


End Sub

Private Sub cbocontact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfing.SetFocus
End If

End Sub

Private Sub cbofunc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfing.SetFocus
End If

End Sub

Private Sub cbofunc_LostFocus()
If cbofunc.Text <> "" Then
   Consulta_vendedor
   
End If

End Sub

Private Sub cbogenera_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocontact.SetFocus
End If

End Sub

Private Sub cbomut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbogenera.SetFocus
End If

End Sub

Private Sub cbozona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomut.SetFocus
End If

End Sub

Private Sub cbozona_LostFocus()
If cbozona.Text <> "" Then
   Consulta_zonas
Else
   labcodzon.Caption = ""
End If

End Sub

Private Sub DBGrid1_DblClick()
Borra_campos

      mfec.Text = data_buscados.Recordset("fecha")
      mhora.Text = data_buscados.Recordset("hora")
      t_base.Text = data_buscados.Recordset("base")
      labusu.Caption = data_buscados.Recordset("usuario")
      t_mat.Text = data_buscados.Recordset("matricula")
      t_ced.Text = data_buscados.Recordset("cedula")
      t_cced.Text = data_buscados.Recordset("codced")
      t_conv.Text = data_buscados.Recordset("convenio")
      t_nom.Text = data_buscados.Recordset("nombre")
      t_tel.Text = data_buscados.Recordset("telef")
      t_cel.Text = data_buscados.Recordset("celular")
      If IsNull(data_buscados.Recordset("celular2")) = False Then
         t_otro.Text = data_buscados.Recordset("celular2")
      End If
      t_dir1.Text = data_buscados.Recordset("dir1")
      If IsNull(data_buscados.Recordset("dir2")) = False Then
         t_dir2.Text = data_buscados.Recordset("dir2")
      End If
      cboactual.Text = data_buscados.Recordset("mut_act")
      cbozona.Text = data_buscados.Recordset("zona")
      cbomut.Text = data_buscados.Recordset("mut_new")
      cbogenera.Text = data_buscados.Recordset("genera")
      cbocontact.Text = data_buscados.Recordset("contacto")
      labcodfunc.Caption = data_buscados.Recordset("codfunc")
      cbofunc.Text = data_buscados.Recordset("nomfunc")
      If IsNull(data_buscados.Recordset("fec_afil")) = False Then
         mfing.Text = data_buscados.Recordset("fec_afil")
      End If

End Sub

Private Sub Form_Load()
data_par.DatabaseName = App.path & "\paramb.mdb"
data_par.RecordSource = "paramb"
data_par.Refresh

data_buscasi.Connect = "odbc;dsn=" & Xconexrmt & ";"


data_orig.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"


Carga_zonas
Carga_mutua
Carga_vendedores

data_busca.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_buscados.Connect = "odbc;dsn=" & Xconexrmt & ";"
If ControlUsuario("frm_bloqueos") = 1 Then
   data_buscados.RecordSource = "select * from bloqueos order by fecha DESC"
Else
   data_buscados.RecordSource = "select * from bloqueos where usuario ='" & WElusuario & "' order by fecha DESC"
End If

data_buscados.Refresh

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub

Private Sub mffin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cborec.SetFocus
End If

End Sub

Private Sub mfing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbofunc.SetFocus
End If

End Sub

Private Sub t_cced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_conv.SetFocus
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
      MsgBox "Solo números"
      KeyAscii = 8
    End If
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_ced.Text = "" Then
      MsgBox "Campo obligatorio"
      t_mat.SetFocus
   Else
      t_cced.SetFocus
   
   End If
End If

End Sub

Private Sub t_ced_LostFocus()
If t_ced.Text <> "" Then
      If Trim(t_mat.Text) = "" Then
         t_mat.Text = 0
      End If
      If Trim(t_mat.Text) = "" Or t_mat.Text = 0 Then
         'If t_mat.Text > 0 Then
         data_buscasi.RecordSource = "select * from bloqueos where cedula =" & t_ced.Text
         data_buscasi.Refresh
         If data_buscasi.Recordset.RecordCount > 0 Then
            MsgBox "ESTA CÉDULA YA EXISTE REGISTRADA EN CAPTACIÓN, VERIFIQUE!", vbCritical
            Borra_camposSin
         Else
            data_busca.RecordSource = "Select * from clientes where cl_cedula =" & t_ced.Text
            data_busca.Refresh
            If data_busca.Recordset.RecordCount > 0 Then
                If data_busca.Recordset("estado") = 2 Then
                   MsgBox "Atención! El socio se encuentra de baja en el padrón. VERIFIQUE DATOS!", vbExclamation
                End If
                If IsNull(data_busca.Recordset("cl_codigo")) = False Then
                   t_mat.Text = data_busca.Recordset("cl_codigo")
                Else
                   t_mat.Text = ""
                End If
                If IsNull(data_busca.Recordset("cl_apellid")) = False Then
                   t_nom.Text = data_busca.Recordset("cl_apellid")
                Else
                   t_nom.Text = "NN"
                End If
                If IsNull(data_busca.Recordset("cl_cedula")) = False Then
                   t_ced.Text = data_busca.Recordset("cl_cedula")
                Else
                   t_ced.Text = 0
                End If
                If IsNull(data_busca.Recordset("cl_codced")) = False Then
                   t_cced.Text = data_busca.Recordset("cl_codced")
                Else
                   t_cced.Text = 0
                End If
                If IsNull(data_busca.Recordset("cl_codconv")) = False Then
                   t_conv.Text = data_busca.Recordset("cl_codconv")
                Else
                   t_conv.Text = 0
                End If
                If IsNull(data_busca.Recordset("cl_telefon")) = False Then
                   t_tel.Text = data_busca.Recordset("cl_telefon")
                Else
                   t_tel.Text = ""
                End If
                If IsNull(data_busca.Recordset("cl_dpto")) = False Then
                   t_cel.Text = data_busca.Recordset("cl_dpto")
                Else
                   t_cel.Text = 0
                End If
                If IsNull(data_busca.Recordset("cl_socmnom")) = False Then
                   cboactual.Text = data_busca.Recordset("cl_socmnom")
                Else
                   cboactual.Text = ""
                End If
                If IsNull(data_busca.Recordset("cl_direcci")) = False Then
                   t_dir1.Text = data_busca.Recordset("cl_direcci")
                Else
                   t_dir1.Text = ""
                End If
                If IsNull(data_busca.Recordset("cl_entre")) = False Then
                   t_dir2.Text = data_busca.Recordset("cl_entre")
                Else
                   t_dir2.Text = ""
                End If
                If IsNull(data_busca.Recordset("cl_zona")) = False Then
                   cbozona.Text = data_busca.Recordset("cl_zona")
                Else
                   cbozona.Text = ""
                End If
                If t_conv.Text <> "" Then
                   If t_conv.Text <> "PART" Then
                      Consulta_convenio
                      If labno.Caption = "NO" Then
                         MsgBox "No se permiten cambios entre mutualistas de convenio SAPP. No podrá grabar.", vbCritical
                      End If
                   End If
                End If
               t_cced.SetFocus
            Else
               MsgBox "Cédula no encontrada"
               t_cced.SetFocus
               If t_mat.Text = "" Or t_mat.Text = 0 Then
                  t_conv.Text = "PART"
               End If
            End If
         End If
      Else
         data_buscasi.RecordSource = "select * from bloqueos where cedula =" & t_ced.Text
         data_buscasi.Refresh
         If data_buscasi.Recordset.RecordCount > 0 Then
            MsgBox "ESTA CÉDULA YA EXISTE REGISTRADA EN CAPTACIÓN, VERIFIQUE!", vbCritical
            Borra_camposSin
         End If
      End If
End If

End Sub

Private Sub t_cel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   If t_cel.Text = "" Then
      MsgBox "Campo obligatorio. Ingrese NO APLICA"
   End If
   t_otro.SetFocus
End If

End Sub

Private Sub t_cel_LostFocus()
If t_cel.Text <> "" Then
   t_cel.Text = Trim(t_cel.Text)
End If

End Sub

Private Sub t_conv_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_dir1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_dir2.SetFocus
End If

End Sub

Private Sub t_dir2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   cbozona.SetFocus
End If

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If t_mat.Text = "" Then
'      MsgBox "Campo obligatorio"
'      t_mat.SetFocus
   Else
      If t_mat.Text > 0 Then
      
        data_buscasi.RecordSource = "select * from bloqueos where matricula =" & t_mat.Text
        data_buscasi.Refresh
        If data_buscasi.Recordset.RecordCount > 0 Then
           MsgBox "ESTA MATRICULA YA EXISTE REGISTRADA EN CAPTACIÓN, VERIFIQUE!", vbCritical
           Borra_camposSin
        Else
          data_busca.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
          data_busca.Refresh
          If data_busca.Recordset.RecordCount > 0 Then
             If data_busca.Recordset("estado") = 2 Then
                MsgBox "Atención! El socio se encuentra de baja en el padrón. VERIFIQUE DATOS!", vbExclamation
             End If
             If IsNull(data_busca.Recordset("cl_apellid")) = False Then
                t_nom.Text = data_busca.Recordset("cl_apellid")
             Else
                t_nom.Text = "NN"
             End If
             If IsNull(data_busca.Recordset("cl_cedula")) = False Then
                t_ced.Text = data_busca.Recordset("cl_cedula")
             Else
                t_ced.Text = 0
             End If
             If IsNull(data_busca.Recordset("cl_codced")) = False Then
                t_cced.Text = data_busca.Recordset("cl_codced")
             Else
                t_cced.Text = 0
             End If
             If IsNull(data_busca.Recordset("cl_codconv")) = False Then
                t_conv.Text = data_busca.Recordset("cl_codconv")
             Else
                t_conv.Text = ""
             End If
             If IsNull(data_busca.Recordset("cl_telefon")) = False Then
                t_tel.Text = data_busca.Recordset("cl_telefon")
             Else
                t_tel.Text = ""
             End If
             If IsNull(data_busca.Recordset("cl_dpto")) = False Then
                t_cel.Text = data_busca.Recordset("cl_dpto")
             Else
                t_cel.Text = 0
             End If
             If IsNull(data_busca.Recordset("cl_socmnom")) = False Then
                cboactual.Text = data_busca.Recordset("cl_socmnom")
             Else
                cboactual.Text = ""
             End If
             If IsNull(data_busca.Recordset("cl_direcci")) = False Then
                t_dir1.Text = data_busca.Recordset("cl_direcci")
             Else
                t_dir1.Text = ""
             End If
             If IsNull(data_busca.Recordset("cl_entre")) = False Then
                t_dir2.Text = data_busca.Recordset("cl_entre")
             Else
                t_dir2.Text = ""
             End If
             If IsNull(data_busca.Recordset("cl_zona")) = False Then
                cbozona.Text = data_busca.Recordset("cl_zona")
             Else
                cbozona.Text = ""
             End If
             If t_conv.Text <> "" Then
                If t_conv.Text <> "PART" Then
                   Consulta_convenio
                   If labno.Caption = "NO" Then
                      MsgBox "No se permiten cambios entre mutualistas de convenio SAPP. No podrá grabar.", vbCritical
                   End If
                End If
             End If
          Else
             MsgBox "Matrícula no encontrada"
             t_mat.Text = 0
             t_ced.SetFocus
          End If
        End If
     End If
   
   End If
   t_ced.SetFocus
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_tel.SetFocus
End If

End Sub

Private Sub t_nom_LostFocus()
If t_nom.Text <> "" Then
   t_nom.Text = Trim(t_nom.Text)
End If

End Sub

Private Sub t_otro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboactual.SetFocus
End If

End Sub

Private Sub t_tel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   t_cel.SetFocus
End If

End Sub

Public Sub Borra_campos()

mfec.Text = "__/__/____"
mhora.Text = "__:__"
labusu.Caption = ""
t_base.Text = ""
t_mat.Text = 0
t_ced.Text = ""
t_cced.Text = ""
t_conv.Text = ""
t_nom.Text = ""
t_tel.Text = ""
t_cel.Text = ""
t_otro.Text = ""
cboactual.Text = ""
t_dir1.Text = ""
t_dir2.Text = ""
cbozona.Text = ""
cbomut.ListIndex = -1
labcodzon.Caption = ""
cbogenera.ListIndex = -1
labcodfunc.Caption = ""
cbofunc.Text = ""
cbocontact.ListIndex = -1
mfing.Text = "__/__/____"

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

Public Sub Consulta_zonas()
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
Else
   MsgBox "Zona no encontrada"
   labcodzon.Caption = ""
   cbozona.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Carga_mutua()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm order by ca_nom"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("ca_nom")) = False Then
         cboactual.AddItem Xrecclii("ca_nom")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_mutual()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm where ca_nom ='" & cboactual.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
Else
   MsgBox "No se encuentra mutualista, VERIFIQUE!!", vbCritical
   cboactual.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Consulta_vendedor()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
If IsNumeric(cbofunc.Text) = True Then
   Xsqlpromo = "Select * from vende_func where idfunc =" & cbofunc.Text
Else
   Xsqlpromo = "Select * from vende_func where nombre ='" & cbofunc.Text & "'"
End If

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labcodfunc.Caption = Xrecclii("idfunc")
   cbofunc.Text = Xrecclii("nombre")
Else
   MsgBox "Promotor, No encontrado."
   labcodfunc.Caption = ""
   cbofunc.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Carga_vendedores()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from vende_func order by nombre"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("nombre")) = False Then
         cbofunc.AddItem Xrecclii("nombre")
      End If
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Private Sub t_tel_LostFocus()
If t_tel.Text <> "" Then
   t_tel.Text = Trim(t_tel.Text)
End If
   If t_tel.Text = "" Then
      MsgBox "Campo obligatorio. Ingrese NO APLICA"
   End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If Text1.Text <> "" Then
      If Option1.Value = True Then
         If ControlUsuario("frm_bloqueos") = 1 Then
            data_buscados.RecordSource = "select * from bloqueos where cedula >=" & Text1.Text & " order by cedula"
         Else
            data_buscados.RecordSource = "select * from bloqueos where cedula >=" & Text1.Text & " and usuario ='" & WElusuario & "' order by cedula"
         End If
         data_buscados.Refresh
      Else
         If Option2.Value = True Then
            If ControlUsuario("frm_bloqueos") = 1 Then
               data_buscados.RecordSource = "select * from bloqueos where codfunc =" & Text1.Text & " order by fecha DESC"
            Else
               data_buscados.RecordSource = "select * from bloqueos where codfunc =" & Text1.Text & " and usuario ='" & WElusuario & "' order by fecha DESC"
            End If
            data_buscados.Refresh
         Else
            If ControlUsuario("frm_bloqueos") = 1 Then
               data_buscados.RecordSource = "select * from bloqueos order by fecha DESC"
            Else
               data_buscados.RecordSource = "select * from bloqueos where usuario ='" & WElusuario & "' order by fecha DESC"
            End If
            data_buscados.Refresh
         End If
      End If
   Else
      If ControlUsuario("frm_bloqueos") = 1 Then
         data_buscados.RecordSource = "select * from bloqueos order by fecha DESC"
      Else
         data_buscados.RecordSource = "select * from bloqueos where usuario ='" & WElusuario & "' order by fecha DESC"
      End If
      data_buscados.Refresh
   End If
   DBGrid1.SetFocus
End If

End Sub
Public Sub Borra_camposSin()

t_mat.Text = 0
t_ced.Text = ""
t_cced.Text = ""
t_conv.Text = ""
t_nom.Text = ""
t_tel.Text = ""
t_cel.Text = ""
t_otro.Text = ""
cboactual.Text = ""
t_dir1.Text = ""
t_dir2.Text = ""
cbozona.Text = ""
cbomut.ListIndex = -1
labcodzon.Caption = ""
cbogenera.ListIndex = -1
labcodfunc.Caption = ""
cbofunc.Text = ""
cbocontact.ListIndex = -1
mfing.Text = "__/__/____"

End Sub


Public Sub Consulta_convenio()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open

Xsqlpromo = "Select * from convenio where cnv_codigo ='" & t_conv.Text & "' and cnv_grupo in ('CCOU','UNIVERSAL','SMI','H.EVANGELICO')"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labno.Caption = "NO"
Else
   labno.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
