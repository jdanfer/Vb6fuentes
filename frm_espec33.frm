VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_espec 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horarios especialistas"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7140
   Enabled         =   0   'False
   Icon            =   "frm_espec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_lis 
      Caption         =   "data_lis"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton bborra 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Borrar Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      MaskColor       =   &H008080FF&
      Picture         =   "frm_espec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton b_rec 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Borrar Datos..."
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
      Height          =   615
      Left            =   5040
      MouseIcon       =   "frm_espec.frx":0884
      MousePointer    =   99  'Custom
      Picture         =   "frm_espec.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Borrar fechas de consultas ya realizadas"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Data data_aenviar 
      Caption         =   "data_aenviar"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data env_espe 
      Caption         =   "env_espe"
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
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_buscod 
      Caption         =   "data_buscod"
      Connect         =   "Access"
      DatabaseName    =   "C:\sapp\sapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "espec"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Windows\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton b_bustodo 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BUSCAR..."
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
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Opciones de envío"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   6855
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
      Begin VB.CommandButton b_acep 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Aceptar"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Por código"
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
         Left            =   2640
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Por Fecha"
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
         TabIndex        =   26
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Envía el CODIGO seleccionado"
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
         Left            =   2640
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de Horarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   6855
      Begin VB.CommandButton b_imp 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Imprimir"
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
         Left            =   4920
         Picture         =   "frm_espec.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton b_ing 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ingresar Fechas"
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
         Left            =   2520
         Picture         =   "frm_espec.frx":1412
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton b_cons 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Consulta Fechas"
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
         Picture         =   "frm_espec.frx":1854
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de Especialista"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton b_alta 
         Height          =   615
         Left            =   240
         Picture         =   "frm_espec.frx":1F96
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   735
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PARSEC0"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data data_espec 
         Caption         =   "data_espec"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "espec"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
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
         RecordSource    =   "medicos"
         Top             =   3360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton b_busca 
         Height          =   615
         Left            =   4560
         Picture         =   "frm_espec.frx":23D8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Eliminar registro"
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton b_canc 
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         Picture         =   "frm_espec.frx":281A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton b_graba 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2400
         Picture         =   "frm_espec.frx":2C5C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton b_modif 
         Height          =   615
         Left            =   1320
         Picture         =   "frm_espec.frx":309E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txt_espera 
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
         Left            =   4800
         TabIndex        =   15
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txt_cantp 
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
         Left            =   2520
         TabIndex        =   13
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txt_mmpp 
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
         Left            =   4800
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txt_mm 
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
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txt_hh 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txt_desc 
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
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txt_cod 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txt_base 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "MINUTOS"
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
         Left            =   5400
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6840
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ESPERA:"
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
         Left            =   3480
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CANT. de PACIENTES:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "C/PACIENTE"
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
         Left            =   3480
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "HORA COMIENZO"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "HORARIOS Y CANTIDAD DE PACIENTES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CODIGO:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "BASE:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_espec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
env_espe.DatabaseName = App.Path & "\env_espe.mdb"
env_espe.RecordSource = "env_esp"
env_espe.Refresh
data_aenviar.DatabaseName = App.Path & "\sapp.mdb"
data_aenviar.RecordSource = "espec"
data_aenviar.Refresh
If env_espe.Recordset.RecordCount > 0 Then
   env_espe.Recordset.MoveFirst
   Do While Not env_espe.Recordset.EOF
      env_espe.Recordset.Delete
      env_espe.Recordset.MoveNext
   Loop
End If
env_espe.RecordSource = "env_fechas"
env_espe.Refresh
If env_espe.Recordset.RecordCount > 0 Then
   env_espe.Recordset.MoveFirst
   Do While Not env_espe.Recordset.EOF
      env_espe.Recordset.Delete
      env_espe.Recordset.MoveNext
   Loop
End If
env_espe.RecordSource = "env_lista"
env_espe.Refresh
If env_espe.Recordset.RecordCount > 0 Then
   env_espe.Recordset.MoveFirst
   Do While Not env_espe.Recordset.EOF
      env_espe.Recordset.Delete
      env_espe.Recordset.MoveNext
   Loop
End If
env_espe.RecordSource = "env_esp"
env_espe.Refresh
If Option1.Value = True Then
   data_aenviar.DatabaseName = App.Path & "\espec.mdb"
   data_aenviar.RecordSource = "Select * from espec"
   data_aenviar.Refresh
   If data_aenviar.Recordset.RecordCount > 0 Then
      data_aenviar.Recordset.MoveFirst
      Do While Not data_aenviar.Recordset.EOF
         env_espe.Recordset.AddNew
         env_espe.Recordset("base") = data_aenviar.Recordset("base")
         env_espe.Recordset("codigo") = data_aenviar.Recordset("codigo")
         env_espe.Recordset("desc") = data_aenviar.Recordset("desc")
         env_espe.Recordset("horcom") = data_aenviar.Recordset("horcom")
         env_espe.Recordset("min") = data_aenviar.Recordset("min")
         env_espe.Recordset("cantp") = data_aenviar.Recordset("cantp")
         env_espe.Recordset("espera") = data_aenviar.Recordset("espera")
         env_espe.Recordset("cada") = data_aenviar.Recordset("cada")
         env_espe.Recordset.Update
         data_aenviar.Recordset.MoveNext
      Loop
   End If
   env_espe.RecordSource = "env_fechas"
   env_espe.Refresh
   data_aenviar.DatabaseName = App.Path & "\sapp.mdb"
   data_aenviar.RecordSource = "Select * from fechasesp where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
   data_aenviar.Refresh
   If data_aenviar.Recordset.RecordCount > 0 Then
      data_aenviar.Recordset.MoveFirst
      Do While Not data_aenviar.Recordset.EOF
         env_espe.Recordset.AddNew
         env_espe.Recordset("base") = data_aenviar.Recordset("base")
         env_espe.Recordset("cod") = data_aenviar.Recordset("cod")
         env_espe.Recordset("desc") = data_aenviar.Recordset("desc")
         env_espe.Recordset("fecha") = data_aenviar.Recordset("fecha")
         env_espe.Recordset("descfec") = data_aenviar.Recordset("descfec")
         env_espe.Recordset.Update
         data_aenviar.Recordset.MoveNext
      Loop
   End If
   env_espe.RecordSource = "env_lista"
   env_espe.Refresh
   data_aenviar.RecordSource = "Select * from lista where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
   data_aenviar.Refresh
   If data_aenviar.Recordset.RecordCount > 0 Then
      data_aenviar.Recordset.MoveFirst
      Do While Not data_aenviar.Recordset.EOF
         env_espe.Recordset.AddNew
         env_espe.Recordset("nro") = data_aenviar.Recordset("nro")
         env_espe.Recordset("fecha") = data_aenviar.Recordset("fecha")
         env_espe.Recordset("base") = data_aenviar.Recordset("base")
         env_espe.Recordset("cod") = data_aenviar.Recordset("cod")
         env_espe.Recordset("matric") = data_aenviar.Recordset("matric")
         env_espe.Recordset("nompac") = data_aenviar.Recordset("nompac")
'         env_espe.Recordset("obs") = data_aenviar.Recordset("obs")
         env_espe.Recordset("tel") = data_aenviar.Recordset("tel")
         env_espe.Recordset("horacom") = data_aenviar.Recordset("horacom")
         env_espe.Recordset.Update
         data_aenviar.Recordset.MoveNext
      Loop
   End If
   env_espe.DatabaseName = ""
   env_espe.RecordSource = ""
   env_espe.Refresh
   FileCopy App.Path & "\env_espe.mdb", "C:\datos\env_espe.mdb"
   MsgBox "Proceso finalizado", vbInformation, "Mensaje"
End If
If Option2.Value = True Then
   data_aenviar.RecordSource = "Select * from espec"
   data_aenviar.Refresh
   If data_aenviar.Recordset.RecordCount > 0 Then
      data_aenviar.Recordset.MoveFirst
      Do While Not data_aenviar.Recordset.EOF
         env_espe.Recordset.AddNew
         env_espe.Recordset("base") = data_aenviar.Recordset("base")
         env_espe.Recordset("codigo") = data_aenviar.Recordset("codigo")
         env_espe.Recordset("desc") = data_aenviar.Recordset("desc")
         env_espe.Recordset("horcom") = data_aenviar.Recordset("horcom")
         env_espe.Recordset("min") = data_aenviar.Recordset("min")
         env_espe.Recordset("cantp") = data_aenviar.Recordset("cantp")
         env_espe.Recordset("espera") = data_aenviar.Recordset("espera")
         env_espe.Recordset.Update
         data_aenviar.Recordset.MoveNext
      Loop
   End If
   env_espe.RecordSource = "env_fechas"
   env_espe.Refresh
   data_aenviar.RecordSource = "Select * from fechasesp where cod ='" & Trim(txt_cod.Text) & "'"
   data_aenviar.Refresh
   If data_aenviar.Recordset.RecordCount > 0 Then
      data_aenviar.Recordset.MoveFirst
      Do While Not data_aenviar.Recordset.EOF
         env_espe.Recordset.AddNew
         env_espe.Recordset("base") = data_aenviar.Recordset("base")
         env_espe.Recordset("cod") = data_aenviar.Recordset("cod")
         env_espe.Recordset("desc") = data_aenviar.Recordset("desc")
         env_espe.Recordset("fecha") = data_aenviar.Recordset("fecha")
         env_espe.Recordset("descfec") = data_aenviar.Recordset("descfec")
         env_espe.Recordset.Update
         data_aenviar.Recordset.MoveNext
      Loop
   End If
   env_espe.RecordSource = "env_lista"
   env_espe.Refresh
   data_aenviar.RecordSource = "Select * from lista where cod ='" & txt_cod.Text & "'"
   data_aenviar.Refresh
   If data_aenviar.Recordset.RecordCount > 0 Then
      data_aenviar.Recordset.MoveFirst
      Do While Not data_aenviar.Recordset.EOF
         env_espe.Recordset.AddNew
         env_espe.Recordset("nro") = data_aenviar.Recordset("nro")
         env_espe.Recordset("fecha") = data_aenviar.Recordset("fecha")
         env_espe.Recordset("base") = data_aenviar.Recordset("base")
         env_espe.Recordset("cod") = data_aenviar.Recordset("cod")
         env_espe.Recordset("matric") = data_aenviar.Recordset("matric")
         env_espe.Recordset("nompac") = data_aenviar.Recordset("nompac")
         env_espe.Recordset("obs") = data_aenviar.Recordset("obs")
         env_espe.Recordset("tel") = data_aenviar.Recordset("tel")
         env_espe.Recordset("horacom") = data_aenviar.Recordset("horacom")
         env_espe.Recordset.Update
         data_aenviar.Recordset.MoveNext
      Loop
   End If
   env_espe.DatabaseName = ""
   env_espe.RecordSource = ""
   env_espe.Refresh
   FileCopy App.Path & "\env_espe.mdb", "C:\datos\env_espe.mdb"
   MsgBox "Proceso finalizado", vbInformation, "Mensaje"
End If

End Sub

Private Sub b_alta_Click()
XAlta = 1
txt_base.Enabled = True
txt_cod.Enabled = True
txt_desc.Enabled = True
txt_hh.Enabled = True
txt_mm.Enabled = True
txt_mmpp.Enabled = True
txt_cantp.Enabled = True
txt_espera.Enabled = True

borraesp
b_alta.Enabled = False
b_bustodo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_canc.Enabled = True
b_busca.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False
txt_base.SetFocus
data_espec.Recordset.AddNew

End Sub

Private Sub b_busca_Click()
Dim XRES As String
XRES = MsgBox("Desea Borrar??", vbCritical + vbYesNo, "Mensaje")
If XRES = vbYes Then
   data_espec.Recordset.Delete
   data_espec.Refresh
   igualar
End If

End Sub

Private Sub b_bustodo_Click()
frm_busespe.Show vbModal

End Sub

Private Sub b_canc_Click()
If XAlta = 1 Then
   data_espec.Recordset.CancelUpdate
   XAlta = 0
   borraesp
   If data_espec.Recordset.RecordCount > 0 Then
      data_espec.Recordset.MoveLast
   End If
   igualar
   b_alta.Enabled = True
   b_bustodo.Enabled = True
   b_modif.Enabled = True
   b_graba.Enabled = False
   b_canc.Enabled = False
   b_busca.Enabled = True
   Frame2.Enabled = True
   Frame3.Enabled = True
   b_bustodo.SetFocus
   txt_base.Enabled = False
   txt_cod.Enabled = False
   txt_desc.Enabled = False
   txt_hh.Enabled = False
   txt_mm.Enabled = False
   txt_mmpp.Enabled = False
   txt_cantp.Enabled = False
   txt_espera.Enabled = False

Else
   XAlta = 0
   borraesp
   If data_espec.Recordset.RecordCount > 0 Then
      data_espec.Recordset.MoveLast
   End If
   igualar
   b_alta.Enabled = True
   b_bustodo.Enabled = True
   b_modif.Enabled = True
   b_graba.Enabled = False
   b_canc.Enabled = False
   b_busca.Enabled = True
   Frame2.Enabled = True
   Frame3.Enabled = True
   b_bustodo.SetFocus
   txt_base.Enabled = False
   txt_cod.Enabled = False
   txt_desc.Enabled = False
   txt_hh.Enabled = False
   txt_mm.Enabled = False
   txt_mmpp.Enabled = False
   txt_cantp.Enabled = False
   txt_espera.Enabled = False

End If
End Sub

Private Sub b_cons_Click()
frm_consfechas.Show vbModal

End Sub

Private Sub b_graba_Click()
If XAlta = 1 Then
   If txt_base.Text <> "" Then
      If txt_cod.Text <> "" Then
         data_buscod.Recordset.FindFirst "codigo ='" & txt_cod.Text & "'"
         If Not data_buscod.Recordset.NoMatch Then
            MsgBox "Ya existe CODIGO", vbCritical, "Mensaje"
            txt_cod.SetFocus
         Else
             data_espec.Recordset("base") = txt_base.Text
             data_espec.Recordset("codigo") = txt_cod.Text
             data_espec.Recordset("desc") = txt_desc.Text
             data_espec.Recordset("horcom") = Trim(txt_hh.Text) + ":" + Trim(txt_mm.Text)
             data_espec.Recordset("min") = txt_mmpp.Text
             data_espec.Recordset("cantp") = txt_cantp.Text
             data_espec.Recordset("cada") = txt_mmpp.Text
             data_espec.Recordset("espera") = txt_espera.Text
             data_espec.Recordset.Update
             XAlta = 0
             borraesp
             If data_espec.Recordset.RecordCount > 0 Then
                data_espec.Recordset.MoveLast
             End If
             igualar
             b_alta.Enabled = True
             b_bustodo.Enabled = True
             b_modif.Enabled = True
             b_graba.Enabled = False
             b_canc.Enabled = False
             b_busca.Enabled = True
             Frame2.Enabled = True
             Frame3.Enabled = True
             b_bustodo.SetFocus
            txt_base.Enabled = False
            txt_cod.Enabled = False
            txt_desc.Enabled = False
            txt_hh.Enabled = False
            txt_mm.Enabled = False
            txt_mmpp.Enabled = False
            txt_cantp.Enabled = False
            txt_espera.Enabled = False
         End If
      End If
   End If
Else
   If txt_base.Text <> "" Then
      If txt_cod.Text <> "" Then
         data_espec.Recordset.Edit
         data_espec.Recordset("base") = txt_base.Text
         data_espec.Recordset("codigo") = txt_cod.Text
         data_espec.Recordset("desc") = txt_desc.Text
         data_espec.Recordset("horcom") = Trim(txt_hh.Text) + ":" + Trim(txt_mm.Text)
         data_espec.Recordset("min") = txt_mmpp.Text
         data_espec.Recordset("cantp") = txt_cantp.Text
         data_espec.Recordset("cada") = txt_mmpp.Text
         data_espec.Recordset("espera") = txt_espera.Text
         data_espec.Recordset.Update
         XAlta = 0
         borraesp
         If data_espec.Recordset.RecordCount > 0 Then
            data_espec.Recordset.MoveLast
         End If
         igualar
         b_alta.Enabled = True
         b_bustodo.Enabled = True
         b_modif.Enabled = True
         b_graba.Enabled = False
         b_canc.Enabled = False
         b_busca.Enabled = True
         Frame2.Enabled = True
         Frame3.Enabled = True
         b_bustodo.SetFocus
        txt_base.Enabled = False
        txt_cod.Enabled = False
        txt_desc.Enabled = False
        txt_hh.Enabled = False
        txt_mm.Enabled = False
        txt_mmpp.Enabled = False
        txt_cantp.Enabled = False
        txt_espera.Enabled = False
      
      End If
   End If
End If
End Sub

Private Sub b_imp_Click()
frm_fechasesp.Show vbModal

End Sub

Private Sub b_ing_Click()
WCodesp = txt_cod.Text
WNomesp = txt_desc.Text

frm_creaesp.Show vbModal

End Sub

Private Sub b_modif_Click()
XAlta = 0
txt_base.Enabled = True
txt_cod.Enabled = False
txt_desc.Enabled = True
txt_hh.Enabled = True
txt_mm.Enabled = True
txt_mmpp.Enabled = True
txt_cantp.Enabled = True
txt_espera.Enabled = True

'borraesp
b_alta.Enabled = False
b_bustodo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_canc.Enabled = True
b_busca.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False
'txt_cod.SetFocus

End Sub

Private Sub b_rec_Click()
Dim Xquefec As String
Dim Xsionobo As String
Dim Baseborra As Database
Dim Sesionborra As Workspace
Set Sesionborra = Workspaces(0)

Xquefec = InputBox("Ingrese hasta que FECHA desea BORRAR?:", "Borrar listas de especialistas")
b_rec.Enabled = False
If Xquefec <> "" Then
   Xsionobo = MsgBox("Está seguro que desea borrar?", vbInformation + vbYesNo, "Borrar listas viejas")
   If Xsionobo = vbYes Then
      frm_espec.MousePointer = 11
      Set Baseborra = Sesionborra.OpenDatabase(App.Path & "\sapp.mdb")
      Baseborra.Execute = "Delete * from lista where fecha <=#" & Format(Xquefec, "yyyy/mm/dd") & "#"
      Baseborra.Execute = "Delete * from fechasesp where fecha <=#" & Format(Xquefec, "yyyy/mm/dd") & "#"
      frm_espec.MousePointer = 0
      MsgBox "Proceso finalizado", vbInformation, "Mensaje"
   End If
End If
b_rec.Enabled = True

End Sub

Private Sub bborra_Click()
Dim Xlafecborra As String
xlafechaborra = InputBox("Ingrese fecha a BORRAR: ", "Fechas de especialistas")
If xlafechaborra <> "" Then
   data_lis.RecordSource = "lista"
   data_lis.Refresh
   data_lis.Recordset.FindFirst "fecha =#" & Format(xlafechaborra, "yyyy/mm/dd") & "# and base =" & txt_base.Text & " and cod ='" & txt_cod.Text & "'"
   If Not data_lis.Recordset.NoMatch Then
      data_lis.RecordSource = "Select * from lista where fecha =#" & Format(xlafechaborra, "yyyy/mm/dd") & "# and base =" & txt_base.Text & " and cod ='" & txt_cod.Text & "'"
      data_lis.Refresh
      If data_lis.Recordset.RecordCount > 0 Then
         data_lis.Recordset.MoveFirst
         Do While Not data_lis.Recordset.EOF
            data_lis.Recordset.Delete
            data_lis.Recordset.MoveNext
         Loop
      End If
      data_lis.RecordSource = ""
      data_lis.Refresh
      data_lis.RecordSource = "fechasesp"
      data_lis.Refresh
      data_lis.Recordset.FindFirst "fecha =#" & Format(xlafechaborra, "yyyy/mm/dd") & "# and base =" & txt_base.Text & " and cod ='" & txt_cod.Text & "'"
      If Not data_lis.Recordset.NoMatch Then
         data_lis.Recordset.Delete
      End If
      data_lis.RecordSource = ""
      data_lis.Refresh
      MsgBox "Fecha eliminada", vbInformation, "Mensaje"
   Else
      MsgBox "Fecha no encontrada, verifique", vbCritical, "Mensaje"
   End If
End If


End Sub

Private Sub Form_Activate()
b_bustodo.SetFocus

End Sub

Public Function borraesp()
txt_base.Text = ""
txt_cod.Text = ""
txt_desc.Text = ""
txt_hh.Text = ""
txt_mm.Text = ""
txt_mmpp.Text = ""
txt_cantp.Text = ""
txt_espera.Text = ""

End Function

Public Function igualar()
If data_espec.Recordset.RecordCount > 0 Then
   txt_base.Text = data_espec.Recordset("base")
   txt_cod.Text = data_espec.Recordset("codigo")
   txt_desc.Text = data_espec.Recordset("desc")
   txt_hh.Text = Mid(data_espec.Recordset("horcom"), 1, 2)
   txt_mm.Text = Mid(data_espec.Recordset("horcom"), 4, 2)
   txt_mmpp.Text = data_espec.Recordset("min")
   txt_cantp.Text = data_espec.Recordset("cantp")
   txt_espera.Text = data_espec.Recordset("espera")
   
End If

End Function

Private Sub Form_Load()
data_med.DatabaseName = App.Path & "\sapp.mdb"
data_espec.DatabaseName = App.Path & "\sapp.mdb"
Data1.DatabaseName = App.Path & "\parsec0.mdb"
Data1.RecordSource = "parsec0"
Data1.Refresh
data_lis.DatabaseName = App.Path & "\sapp.mdb"

data_buscod.DatabaseName = App.Path & "\sapp.mdb"
Data2.DatabaseName = "c:\windows\usapp.mdb"
Data2.RecordSource = "usuarioact"
Data2.Refresh
If Data2.Recordset("nombre") = "JFERNAN" Or Data2.Recordset("nombre") = "CLAUDIA" Then
   Frame1.Enabled = True
   Frame3.Enabled = True
   bborra.Visible = True
   b_rec.Enabled = True
   data_espec.RecordSource = "Select * from espec order by base"
   data_espec.Refresh
Else
   Frame1.Enabled = False
   Frame3.Enabled = False
   bborra.Visible = False
   b_rec.Enabled = False
   data_espec.RecordSource = "Select * from espec where base =" & Data1.Recordset("base")
   data_espec.Refresh
End If
If data_espec.Recordset.RecordCount > 0 Then
   If IsNull(data_espec.Recordset("Base")) = True Then
      txt_base.Text = ""
   Else
      txt_base.Text = data_espec.Recordset("base")
   End If
   txt_cod.Text = data_espec.Recordset("codigo")
   txt_desc.Text = data_espec.Recordset("desc")
   txt_hh.Text = Mid(data_espec.Recordset("horcom"), 1, 2)
   txt_mm.Text = Mid(data_espec.Recordset("min"), 1, 2)
   If IsNull(data_espec.Recordset("cada")) = False Then
      txt_mmpp.Text = data_espec.Recordset("cada")
   Else
      txt_mmpp.Text = 15
   End If
   txt_cantp.Text = data_espec.Recordset("cantp")
   txt_espera.Text = data_espec.Recordset("espera")
   
End If

End Sub

Private Sub Option1_Click()
md.Visible = True
mh.Visible = True
Label9.Visible = False

End Sub

Private Sub Option2_Click()
md.Visible = False
mh.Visible = False
Label9.Visible = True

End Sub

Private Sub txt_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   txt_cod.SetFocus
End If

End Sub

Private Sub txt_cantp_KeyPress(KeyAscii As Integer)
If txt_cantp.Text = "" Then
Else
   If KeyAscii = 13 Then
      txt_espera.SetFocus
   End If
End If

End Sub

Private Sub txt_cantp_LostFocus()
If txt_cantp.Text = "" Then
   MsgBox "INGRESE CANTIDAD DE PACIENTES"
   txt_cantp.SetFocus
End If

End Sub

Private Sub txt_cod_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txt_desc.SetFocus
End If

End Sub

Private Sub txt_desc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_hh.SetFocus
End If

End Sub

Private Sub txt_espera_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub txt_espera_LostFocus()
If txt_espera.Text = "" Then
   txt_espera.Text = 0
End If

End Sub

Private Sub txt_hh_KeyPress(KeyAscii As Integer)
If txt_hh.Text = "" Then
Else
   If KeyAscii = 13 Then
      txt_mm.SetFocus
   End If
End If

End Sub

Private Sub txt_hh_LostFocus()
If txt_hh.Text = "" Then
   MsgBox "Ingrese Hora"
   txt_hh.SetFocus
End If

End Sub

Private Sub txt_mm_KeyPress(KeyAscii As Integer)
If txt_mm.Text = "" Then
Else
   If KeyAscii = 13 Then
      txt_mmpp.SetFocus
   End If
End If
End Sub

Private Sub txt_mm_LostFocus()
If txt_mm.Text = "" Then
   MsgBox "Ingrese MINUTOS"
   txt_mm.SetFocus
End If

End Sub

Private Sub txt_mmpp_KeyPress(KeyAscii As Integer)
If txt_mmpp.Text = "" Then
Else
   If KeyAscii = 13 Then
      txt_cantp.SetFocus
   End If
End If
End Sub

Private Sub txt_mmpp_LostFocus()
If txt_mmpp.Text = "" Then
   MsgBox "Ingrese DATO"
   txt_mmpp.SetFocus
End If

End Sub
