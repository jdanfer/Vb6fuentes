VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_mejora 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mejora continua"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9735
   Icon            =   "frm_mejora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_parmam 
      Caption         =   "data_parmam"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H0080FFFF&
      Caption         =   "TITULO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6600
      TabIndex        =   44
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0080FFFF&
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   43
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080FFFF&
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6600
      TabIndex        =   42
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton b_nover 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   615
      Left            =   5160
      Picture         =   "frm_mejora.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Cancelar la visualización del cuadro descripción."
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton b_ver 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3960
      Picture         =   "frm_mejora.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Editar el cuadro DESCRIPCION para leer los datos ingresados."
      Top             =   7560
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   7080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_cargo 
      Caption         =   "data_cargo"
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
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver solo acciones en proceso."
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
      TabIndex        =   31
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton b_histo 
      BackColor       =   &H0000FF00&
      Caption         =   "Registrar ACCIONES"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Data data_graba 
      Caption         =   "data_graba"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton b_buscafec 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      Picture         =   "frm_mejora.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4920
      Width           =   495
   End
   Begin VB.Data data_accion 
      Caption         =   "data_accion"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_mejora.frx":1250
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "frm_mejora.frx":126A
      TabIndex        =   26
      Top             =   5880
      Width           =   9495
   End
   Begin VB.CommandButton b_infor 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_mejora.frx":1F95
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Informes"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_mejora.frx":251F
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancelar movimiento realizado"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_mejora.frx":2AA9
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Grabar datos"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      Picture         =   "frm_mejora.frx":3033
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Modificar datos de registro seleccionado"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_mejora.frx":35BD
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Ingresar nuevo registro"
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Datos de la acción solicitada"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin MSMask.MaskEdBox mfvence 
         Height          =   375
         Left            =   6960
         TabIndex        =   46
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frm_mejora.frx":3B47
         Left            =   2040
         List            =   "frm_mejora.frx":3B60
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   3960
         Width           =   3135
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Oportunidad de mejora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   36
         Top             =   3600
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "NO conformidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   35
         Top             =   3600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   3600
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mfecfin 
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_mejora.frx":3BC0
         Left            =   2040
         List            =   "frm_mejora.frx":3BCD
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4440
         Width           =   3135
      End
      Begin VB.TextBox txt_detal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2280
         Width           =   7095
      End
      Begin VB.TextBox txt_encab 
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
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   15
         Top             =   1920
         Width           =   7095
      End
      Begin VB.CommandButton b_elimin 
         Height          =   495
         Left            =   5160
         Picture         =   "frm_mejora.frx":3BF4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Elimina destinatario seleccionado"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   6120
         TabIndex        =   12
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton b_agreg 
         Height          =   495
         Left            =   5160
         Picture         =   "frm_mejora.frx":4036
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Agrega..."
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frm_mejora.frx":4478
         Left            =   2040
         List            =   "frm_mejora.frx":447A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   3015
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8160
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin MSMask.MaskEdBox mfecha 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   255
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
      Begin VB.TextBox txt_nro 
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "VENCIMIENTO:"
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
         Left            =   5400
         TabIndex        =   45
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label labid 
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Origen:"
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
         TabIndex        =   37
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Registro de:"
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
         TabIndex        =   33
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Destinatario/s"
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
         Left            =   6120
         TabIndex        =   32
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Conformidad:"
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
         TabIndex        =   18
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Título:"
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
         TabIndex        =   14
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Dirigido a:"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
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
         Height          =   375
         Left            =   7080
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FECHA:"
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
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NUMERO:"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FFFF&
      Caption         =   "Doble click para editar "
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
      TabIndex        =   28
      Top             =   7560
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar por:"
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
      Left            =   4800
      TabIndex        =   27
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label labusuario 
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
      Left            =   1800
      TabIndex        =   8
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Usuario actual:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   6000
      Picture         =   "frm_mejora.frx":447C
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   1575
   End
End
Attribute VB_Name = "frm_mejora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_agreg_Click()
Dim XX, Xban As Long
XX = 0
Xban = 0
If List1.ListCount >= 1 Then
   For XX = 1 To List1.ListCount
       List1.ListIndex = XX - 1
       If List1.List(List1.ListIndex) = Combo1.Text Then
          Xban = 1
       End If
   Next
Else
   Xban = 0
End If

If Combo1.ListIndex >= 0 And Xban <> 1 Then
   List1.AddItem Combo1.Text

End If

End Sub

Private Sub b_buscafec_Click()
Dim Xm1 As String
If Check2.Value = 1 Then
   Xm1 = InputBox("INGRESE FECHA A BUSCAR (formato: dd/mm/aaaa)")
   If Xm1 <> "" Then
      If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "AGUILLEN" Or WElusuario = "COMPUTOS" Or WElusuario = "MEUGENIA" Then
         data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_fnac >=#" & Format(Xm1, "yyyy/mm/dd") & "# and cl_nomcobr =" & 1 & " order by cl_fnac"
         data_accion.Refresh
      Else
         data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_fnac >=#" & Format(Xm1, "yyyy/mm/dd") & "# and cl_nomcobr =" & 1 & " and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac"
         data_accion.Refresh
      End If
   End If
Else
   If Check3.Value = 1 Then
      Xm1 = InputBox("INGRESE CODIGO A BUSCAR")
      If Xm1 <> "" Then
        If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "AGUILLEN" Or WElusuario = "COMPUTOS" Or WElusuario = "MEUGENIA" Then
           data_accion.RecordSource = "Select * from infor_sol where estado =" & Val(Xm1) & " and cl_nomcobr =" & 1 & " order by cl_fnac"
           data_accion.Refresh
        Else
           data_accion.RecordSource = "Select * from infor_sol where estado =" & Val(Xm1) & " and cl_nomcobr =" & 1 & " and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac"
           data_accion.Refresh
        End If
      End If
   Else
      If Check4.Value = 1 Then
'and cl_desc1 like '*" & Xm1 & "*' and
         Xm1 = InputBox("INGRESE TEXTO A BUSCAR EN TITULO")
         If Xm1 <> "" Then
            If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "AGUILLEN" Or WElusuario = "COMPUTOS" Or WElusuario = "MEUGENIA" Then
               data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_nomcobr =" & 1 & " and cl_desc1 like '*" & Xm1 & "*' order by cl_fnac"
               data_accion.Refresh
            Else
               data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_nomcobr =" & 1 & " and cl_desc1 like '*" & Xm1 & "*' and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac"
               data_accion.Refresh
            End If
         End If
      End If
   End If
End If
DBGrid1.SetFocus

End Sub

Private Sub b_cancela_Click()
'If XAlta = 1 Then
'   data_graba.Recordset.CancelUpdate
'End If
b_nuevo.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cancela.Enabled = False
b_buscafec.Enabled = True
b_infor.Enabled = True
DBGrid1.Enabled = True
borracamp
Frame1.Enabled = False

End Sub

Private Sub b_elimin_Click()
If List1.ListIndex >= 0 Then
   List1.RemoveItem List1.ListIndex
End If

End Sub

Private Sub b_graba_Click()
Dim XXdest As Long
Dim Xelnro As Double
Xelnro = txt_nro.Text
XXdest = 0
If XAlta = 1 Then
   If List1.ListCount >= 1 Then
      If Len(txt_encab.Text) > 5 Then
         If Len(txt_detal.Text) > 5 Then
            List1.ListIndex = 0
            For XXdest = 1 To List1.ListCount
                data_graba.Recordset.AddNew
                data_graba.Recordset("cl_etiquet") = 0
                data_graba.Recordset("cl_val2") = 7
                data_graba.Recordset("cl_codigo") = labid.Caption
                data_graba.Recordset("estado") = Xelnro
                data_graba.Recordset("cl_fnac") = mfecha.Text
                data_graba.Recordset("cl_ruc") = txt_hora.Text
                data_graba.Recordset("cl_nomcobr") = 1
'                data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
                data_cargo.RecordSource = "Select * from movil where chofer ='" & List1.List(List1.ListIndex) & "'"
                data_cargo.Refresh
                If data_cargo.Recordset.RecordCount > 0 Then
'                If Not data_cargo.Recordset.NoMatch Then
                   data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
                Else
                   data_graba.Recordset("cl_nom_sup") = WElusuario
                End If
                data_graba.Recordset("cl_descpag") = labusuario.Caption
                data_graba.Recordset("cl_desc2") = List1.List(List1.ListIndex)
                data_graba.Recordset("cl_desc1") = txt_encab.Text
                data_graba.Recordset("info_debit") = txt_detal.Text
                If Combo2.ListIndex >= 0 Then
                   data_graba.Recordset("cl_val1") = Combo2.ListIndex
                Else
                   data_graba.Recordset("cl_val1") = -1
                End If
                If mfecfin.Text <> "__/__/____" Then
                   data_graba.Recordset("cl_fultmov") = mfecfin.Text
                Else
                   
                End If
                If Option1.Value = True Then
                   data_graba.Recordset("cl_atrasop") = 1
                Else
                   If Option2.Value = True Then
                      data_graba.Recordset("cl_atrasop") = 2
                   Else
                      If Option3.Value = True Then
                         data_graba.Recordset("cl_atrasop") = 3
                      Else
                         data_graba.Recordset("cl_Atrasop") = 0
                      End If
                   End If
                End If
                If Combo3.ListIndex >= 0 Then
                   data_graba.Recordset("cl_grupo") = Combo3.ListIndex
                Else
                   data_graba.Recordset("cl_grupo") = -1
                End If
                data_graba.Recordset("cl_codconv") = "A"
                If mfvence.Text <> "__/__/____" Then
                   data_graba.Recordset("cl_fec1") = Format(mfvence.Text, "dd/mm/yyyy")
                End If
                data_graba.Recordset.Update
                Xelnro = Xelnro + 1
                If labid.Caption <> "" Then
                   labid.Caption = labid.Caption + 1
                End If
                If List1.ListCount - 1 = List1.ListIndex Then
                Else
                   List1.ListIndex = List1.ListIndex + 1
                End If
            Next
            b_nuevo.Enabled = True
            b_modif.Enabled = True
            b_graba.Enabled = False
            b_cancela.Enabled = False
            b_buscafec.Enabled = True
            b_infor.Enabled = True
            DBGrid1.Enabled = True
            Frame1.Enabled = False
            data_graba.Refresh
            data_accion.Refresh
            borracamp
            XAlta = 0
         Else
            MsgBox "Ingrese detalles"
         End If
      Else
         MsgBox "Ingrese título"
      End If
   Else
      MsgBox "Ingrese al menos un destinatario"
   End If
Else
   data_graba.Recordset.Edit
   List1.ListIndex = 0
   
   data_cargo.RecordSource = "Select * from movil where chofer ='" & List1.List(List1.ListIndex) & "'"
   data_cargo.Refresh
   If data_cargo.Recordset.RecordCount > 0 Then
'                If Not data_cargo.Recordset.NoMatch Then
      data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
   Else
      data_graba.Recordset("cl_nom_sup") = WElusuario
   End If
'   data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
'   If Not data_cargo.Recordset.NoMatch Then
'      data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
'   Else
'      data_graba.Recordset("cl_nom_sup") = WElusuario
'   End If
   data_graba.Recordset("cl_descpag") = labusuario.Caption
   data_graba.Recordset("cl_desc2") = List1.List(List1.ListIndex)
   data_graba.Recordset("cl_desc1") = txt_encab.Text
   data_graba.Recordset("info_debit") = txt_detal.Text
   If Combo2.ListIndex >= 0 Then
      data_graba.Recordset("cl_val1") = Combo2.ListIndex
   Else
      data_graba.Recordset("cl_val1") = -1
   End If
   If mfecfin.Text <> "__/__/____" Then
      data_graba.Recordset("cl_fultmov") = mfecfin.Text
   Else
'      data_graba.Recordset("cl_fecing") = Date
   End If
   If Option1.Value = True Then
      data_graba.Recordset("cl_atrasop") = 1
   Else
      If Option2.Value = True Then
         data_graba.Recordset("cl_atrasop") = 2
      Else
         If Option3.Value = True Then
            data_graba.Recordset("cl_atrasop") = 3
         Else
            data_graba.Recordset("cl_Atrasop") = 0
         End If
      End If
   End If
   If Combo3.ListIndex >= 0 Then
      data_graba.Recordset("cl_grupo") = Combo3.ListIndex
   Else
      data_graba.Recordset("cl_grupo") = -1
   End If
   If mfvence.Text <> "__/__/____" Then
      data_graba.Recordset("cl_fec1") = Format(mfvence.Text, "dd/mm/yyyy")
   Else
      If IsNull(data_graba.Recordset("cl_fec1")) = False Then
         data_graba.Recordset("cl_fec1") = Null
      End If
   End If
   
   data_graba.Recordset.Update
   b_nuevo.Enabled = True
   b_modif.Enabled = True
   b_graba.Enabled = False
   b_cancela.Enabled = False
   b_buscafec.Enabled = True
   b_infor.Enabled = True
   DBGrid1.Enabled = True
   Frame1.Enabled = False
   data_graba.Refresh
   data_accion.Refresh
   borracamp
End If


End Sub

Private Sub b_histo_Click()
frm_mejoracons.Show vbModal

End Sub

Private Sub b_infor_Click()
frm_infmejoras.Show vbModal

End Sub

Private Sub b_modif_Click()
If WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Or WElusuario = "AGUILLEN" Or WElusuario = "MEUGENIA" Then
   'If labusuario.Caption = data_accion.Recordset("cl_descpag") Then
    XAlta = 0
    Frame1.Enabled = True
    If Combo2.ListIndex >= 0 And WElusuario <> "COMPUTOS" Then
       MsgBox "ATENCION! EL REGISTRO YA FUE CERRADO", vbInformation, "Mejora continua"
       Frame1.Enabled = False
    Else
        b_nuevo.Enabled = False
        b_modif.Enabled = False
        b_graba.Enabled = True
        b_cancela.Enabled = True
        b_buscafec.Enabled = False
        b_infor.Enabled = False
        DBGrid1.Enabled = False
         borracamp
         data_graba.RecordSource = "Select * from infor_sol where estado =" & data_accion.Recordset("estado") & " and cl_nomcobr =" & 1
         data_graba.Refresh
         If data_graba.Recordset.RecordCount > 0 Then
            If IsNull(data_graba.Recordset("cl_val3")) = True Then
               Combo2.Enabled = False
               mfecfin.Enabled = False
            Else
               If data_graba.Recordset("cl_val3") = 1 Then
                  Combo2.Enabled = True
                  mfecfin.Enabled = True
               Else
                  Combo2.Enabled = False
                  mfecfin.Enabled = False
               End If
            End If
            igualaacc
         Else
            Frame1.Enabled = False
            b_nuevo.Enabled = True
            b_modif.Enabled = True
            b_graba.Enabled = False
            b_cancela.Enabled = False
            b_buscafec.Enabled = True
            b_infor.Enabled = True
            DBGrid1.Enabled = True
            Combo2.Enabled = False
            mfecfin.Enabled = False
         End If
         If WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Or WElusuario = "AGUILLEN" Or WElusuario = "MEUGENIA" Then
            Combo2.Enabled = True
            mfecfin.Enabled = True
         Else
            Combo2.Enabled = False
            mfecfin.Enabled = False
         End If
        'Else
        '    MsgBox "NO ES PROPIETARIO DE LA ACCION", vbCritical
        '    DBGrid1.SetFocus
        'End If
    End If
Else
   MsgBox "Usuario no autorizado para modificación, solo se habilita el cuadro Descripción para ver"
   Frame1.Enabled = True

End If

End Sub

Private Sub b_nover_Click()
b_nuevo.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cancela.Enabled = False
b_buscafec.Enabled = True
b_infor.Enabled = True
DBGrid1.Enabled = True
Combo1.Enabled = True
List1.Enabled = True
txt_encab.Enabled = True
txt_detal.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Combo3.Enabled = True
Combo2.Enabled = True
mfecfin.Enabled = True
Check1.Enabled = True
b_ver.Enabled = True
b_nover.Enabled = False
mfvence.Enabled = True
Frame1.Enabled = False

End Sub

Private Sub b_nuevo_Click()
If WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Or WElusuario = "AGUILLEN" Or WElusuario = "MEUGENIA" Then
    XAlta = 1
    b_nuevo.Enabled = False
    b_modif.Enabled = False
    b_graba.Enabled = True
    b_cancela.Enabled = True
    b_buscafec.Enabled = False
    b_infor.Enabled = False
    DBGrid1.Enabled = False
    Frame1.Enabled = True
    borracamp
    Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
    Data1.RecordSource = "Select * from infor_sol order by cl_codigo"
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
       Data1.Recordset.MoveLast
'       labid.Caption = Data1.Recordset("cl_codigo") + 1
       labid.Caption = data_parmam.Recordset("nro_mam2") + 1
       data_parmam.Recordset.Edit
       data_parmam.Recordset("nro_mam2") = data_parmam.Recordset("nro_mam2") + 1
       data_parmam.Recordset.Update
       data_parmam.Refresh
    Else
'       labid.Caption = 1000
       data_parmam.Recordset.Edit
       data_parmam.Recordset("nro_mam2") = data_parmam.Recordset("nro_mam2") + 1
       data_parmam.Recordset.Update
       data_parmam.Refresh
    End If
    If data_graba.Recordset.RecordCount > 0 Then
       data_graba.Recordset.MoveLast
       If data_graba.Recordset("estado") >= 70000 Then
'          txt_nro.Text = data_graba.Recordset("estado") + 1
          txt_nro.Text = data_parmam.Recordset("nro_mam2") + 1
          data_parmam.Recordset.Edit
          data_parmam.Recordset("nro_mam2") = data_parmam.Recordset("nro_mam2") + 1
          data_parmam.Recordset.Update
          data_parmam.Refresh
       Else
          txt_nro.Text = data_parmam.Recordset("nro_mam2") + 1
          data_parmam.Recordset.Edit
          data_parmam.Recordset("nro_mam2") = data_parmam.Recordset("nro_mam2") + 1
          data_parmam.Recordset.Update
          data_parmam.Refresh
          
       End If
    Else
       txt_nro.Text = data_parmam.Recordset("nro_mam2") + 1
       data_parmam.Recordset.Edit
       data_parmam.Recordset("nro_mam2") = data_parmam.Recordset("nro_mam2") + 1
       data_parmam.Recordset.Update
       data_parmam.Refresh
    
    End If
    mfecha.Text = Format(Date, "dd/mm/yyyy")
    txt_hora.Text = Format(Time, "HH:mm")
    Combo1.SetFocus
    Combo2.Enabled = False
    mfecfin.Enabled = False
Else
    MsgBox "Usuario no autorizado para crear registros"
End If

End Sub

Private Sub b_ver_Click()
b_nuevo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = False
b_cancela.Enabled = False
b_buscafec.Enabled = False
b_infor.Enabled = False
DBGrid1.Enabled = False
Frame1.Enabled = True
Combo1.Enabled = False
List1.Enabled = False
txt_encab.Enabled = False
txt_detal.Enabled = True
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Combo3.Enabled = False
Combo2.Enabled = False
mfecfin.Enabled = False
Check1.Enabled = False
b_ver.Enabled = False
b_nover.Enabled = True
mfvence.Enabled = False


End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "AGUILLEN" Or WElusuario = "COMPUTOS" Or WElusuario = "MEUGENIA" Then
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_codconv ='" & "A" & "' and cl_nomcobr =" & 1 & " order by cl_fnac DESC"
      data_accion.Refresh
   Else
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_codconv ='" & "A" & "' and cl_nomcobr =" & 1 & " and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac DESC"
      data_accion.Refresh
   End If
Else
   If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "AGUILLEN" Or WElusuario = "COMPUTOS" Or WElusuario = "MEUGENIA" Then
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_nomcobr =" & 1 & " order by cl_fnac DESC"
      data_accion.Refresh
   Else
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_nomcobr =" & 1 & " and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac DESC"
      data_accion.Refresh
   End If
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   If Check3.Value = 1 Then
      Check3.Value = 0
   End If
   If Check4.Value = 1 Then
      Check4.Value = 0
   End If
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   If Check2.Value = 1 Then
      Check2.Value = 0
   End If
   If Check4.Value = 1 Then
      Check4.Value = 0
   End If
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
   If Check2.Value = 1 Then
      Check2.Value = 0
   End If
   If Check3.Value = 1 Then
      Check3.Value = 0
   End If
End If

End Sub

Private Sub Combo1_Click()
b_agreg_Click

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo2.Enabled = True Then
      Combo2.SetFocus
   Else
      b_graba.SetFocus
   End If
End If

End Sub

Private Sub DBGrid1_DblClick()
borracamp
igualaacc

End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "DIRECTOR GENERAL"
Combo1.AddItem "GERENTE GENERAL"
Combo1.AddItem "DIRECCION TECNICA"
Combo1.AddItem "SUB-DIREC.TECNICA"
'Combo1.AddItem "GERENTE COMERCIAL"
'Combo1.AddItem "JEFE DE MEDICOS DE MOVIL"
Combo1.AddItem "JEFE CHOFERES Y MANT."
Combo1.AddItem "JEFE ADMINISTRACION"
Combo1.AddItem "JEFE DPTO.TI"
Combo1.AddItem "JEFE REGIONAL COSTA"
Combo1.AddItem "JEFE FARMACIA/ECONOMATO"
Combo1.AddItem "JEFE DESPACHO"
'Combo1.AddItem "ENCARGADO METAS"
'Combo1.AddItem "JEFE ATENCION AL CLIENTE"
Combo1.AddItem "SUB-JEFE FACTURACION"
Combo1.AddItem "JEFE CONTADURIA"
Combo1.AddItem "JEFE REGIONAL NORTE"
Combo1.AddItem "JEFE COMERCIAL"
Combo1.AddItem "RESPONSABLE CALIDAD"
Combo1.AddItem "JEFE ASISTENCIAL"
Combo1.AddItem "SUB-JEFE TESORERIA" 'Paola
'''Combo1.AddItem "SUB-JEFE CONTADURIA" 'Cris

Combo1.ListIndex = -1
List1.Clear
data_accion.Connect = "odbc;dsn=" & Xconexrmt & ";"
If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "AGUILLEN" Or WElusuario = "COMPUTOS" Or WElusuario = "MEUGENIA" Then
   data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_nomcobr =" & 1 & " order by cl_fnac DESC"
   data_accion.Refresh
Else
   data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_nomcobr =" & 1 & " and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac DESC"
   data_accion.Refresh
End If

data_graba.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_graba.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " order by estado"
data_graba.Refresh
data_cargo.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cargo.RecordSource = "movil"
data_cargo.Refresh

data_parmam.DatabaseName = App.path & "\paramam.mdb"
data_parmam.RecordSource = "parsec0"
data_parmam.Refresh

labusuario.Caption = WElusuario

End Sub


Public Function borracamp()
txt_nro.Text = ""
mfecha.Text = "__/__/____"
txt_hora.Text = ""
Combo1.ListIndex = -1
List1.Clear
txt_encab.Text = ""
txt_detal.Text = ""
mfecfin.Enabled = True
mfecfin.Text = "__/__/____"
mfecfin.Enabled = False
Combo2.Enabled = True
Combo2.ListIndex = -1
Combo2.Enabled = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Combo3.ListIndex = -1
mfvence.Text = "__/__/____"

End Function

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_detal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo3.SetFocus
End If

End Sub

Private Sub txt_encab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_detal.SetFocus
End If

End Sub

Public Function igualaacc()
If data_accion.Recordset.RecordCount > 0 Then
    If IsNull(data_accion.Recordset("estado")) = False Then
       txt_nro.Text = data_accion.Recordset("estado")
    Else
       txt_nro.Text = 0
    End If
    If IsNull(data_accion.Recordset("cl_fnac")) = False Then
       mfecha.Text = data_accion.Recordset("cl_fnac")
    Else
       mfecha.Text = "__/__/____"
    End If
    If IsNull(data_accion.Recordset("cl_ruc")) = False Then
       txt_hora.Text = data_accion.Recordset("cl_ruc")
    Else
       txt_hora.Text = ""
    End If
    If IsNull(data_accion.Recordset("cl_desc2")) = False Then
       List1.AddItem data_accion.Recordset("cl_desc2")
    Else
       List1.Clear
    End If
    If IsNull(data_accion.Recordset("cl_desc1")) = False Then
       txt_encab.Text = data_accion.Recordset("cl_desc1")
    Else
       txt_encab.Text = ""
    End If
    If IsNull(data_accion.Recordset("info_debit")) = False Then
       txt_detal.Text = data_accion.Recordset("info_debit")
    Else
       txt_detal.Text = ""
    End If
    If IsNull(data_accion.Recordset("cl_val1")) = False Then
       Combo2.Enabled = True
       Combo2.ListIndex = data_accion.Recordset("cl_val1")
       Combo2.Enabled = False
    Else
       Combo2.Enabled = True
       Combo2.ListIndex = -1
       Combo2.Enabled = False
    End If
    If IsNull(data_accion.Recordset("cl_fultmov")) = False Then
       mfecfin.Enabled = True
       mfecfin.Text = Format(data_accion.Recordset("cl_fultmov"), "dd/mm/yyyy")
       mfecfin.Enabled = False
    Else
       mfecfin.Enabled = True
       mfecfin.Text = "__/__/____"
       mfecfin.Enabled = False
    End If
    If IsNull(data_accion.Recordset("cl_fec1")) = False Then
       mfvence.Text = Format(data_accion.Recordset("cl_fec1"), "dd/mm/yyyy")
    Else
       mfvence.Text = "__/__/____"
    End If
    If IsNull(data_accion.Recordset("cl_atrasop")) = False Then
       If data_accion.Recordset("cl_atrasop") = 1 Then
          Option1.Value = True
       Else
          If data_accion.Recordset("cl_atrasop") = 2 Then
             Option2.Value = True
          Else
             If data_accion.Recordset("cl_atrasop") = 3 Then
                Option3.Value = True
             Else
                Option1.Value = False
                Option2.Value = False
                Option3.Value = False
             End If
          End If
       End If
    Else
       Option1.Value = False
       Option2.Value = False
       Option3.Value = False
    End If
    If IsNull(data_accion.Recordset("cl_grupo")) = False Then
       If data_accion.Recordset("cl_grupo") >= 0 Then
          Combo3.ListIndex = data_accion.Recordset("cl_grupo")
       Else
          Combo3.ListIndex = -1
       End If
    Else
       Combo3.ListIndex = -1
    End If
End If

End Function
