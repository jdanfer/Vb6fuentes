VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_solicitudbaja 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitudes de Baja"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10890
   Icon            =   "frm_solicitudbaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10890
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_busca 
      Caption         =   "data_busca"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
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
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CommandButton b_acciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acciones"
      Height          =   495
      Left            =   5400
      Picture         =   "frm_solicitudbaja.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Ingresar acciones al registro seleccionado"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton b_infos 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      Picture         =   "frm_solicitudbaja.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Informes"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_solicitudbaja.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Buscar"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_cancelar 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_solicitudbaja.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Cancelar acción"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_solicitudbaja.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Guardar datos"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_solicitudbaja.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Editar registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_solicitudbaja.frx":26C6
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Crear nuevo registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos para la Baja"
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
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10335
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Tiene contrato firmado con otro servicio de emergencia"
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
         Height          =   255
         Left            =   4200
         TabIndex        =   45
         Top             =   3480
         Width           =   5775
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
         Left            =   3840
         MaxLength       =   45
         TabIndex        =   44
         ToolTipText     =   "Teléfono alternativo (campo no obligatorio)"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Data data_par 
         Caption         =   "data_par"
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
         Top             =   480
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
         Left            =   7800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4440
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
         TabIndex        =   43
         ToolTipText     =   "Digite número de celular, campo obligatorio"
         Top             =   2400
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
         TabIndex        =   42
         Top             =   1920
         Width           =   1695
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
         TabIndex        =   41
         Top             =   1440
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
         TabIndex        =   40
         ToolTipText     =   "Digite cédula, campo obligatorio"
         Top             =   1440
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
         TabIndex        =   39
         ToolTipText     =   "Digite matrícula, campo obligatorio"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cborec 
         Enabled         =   0   'False
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   5040
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mffin 
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   5040
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Terminado"
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
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   5040
         Width           =   2295
      End
      Begin VB.TextBox t_accion 
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
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3720
         Width           =   9735
      End
      Begin VB.ComboBox cbohas 
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
         ItemData        =   "frm_solicitudbaja.frx":2C50
         Left            =   9000
         List            =   "frm_solicitudbaja.frx":2C9C
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "horario de contacto, campo obligatorio"
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cbodes 
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
         ItemData        =   "frm_solicitudbaja.frx":2D48
         Left            =   7800
         List            =   "frm_solicitudbaja.frx":2D94
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Horario de contacto, campo obligatorio"
         Top             =   2400
         Width           =   975
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
         TabIndex        =   20
         Top             =   1920
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
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cbomot 
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
         ItemData        =   "frm_solicitudbaja.frx":2E40
         Left            =   5640
         List            =   "frm_solicitudbaja.frx":2E42
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Motivo de la baja, campo obligatorio"
         Top             =   2880
         Width           =   4335
      End
      Begin VB.ComboBox cboorig 
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
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Origen de la solicitud, campo obligatorio"
         Top             =   2880
         Width           =   2415
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
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin MSMask.MaskEdBox mhora 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   960
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
         TabIndex        =   4
         Top             =   960
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
      Begin VB.TextBox t_cod 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C00000&
         Caption         =   "Resultado:"
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
         Left            =   4200
         TabIndex        =   30
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C00000&
         Caption         =   "Acciones registradas"
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
         TabIndex        =   26
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C00000&
         Caption         =   "Horario de contacto:"
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
         Left            =   5880
         TabIndex        =   23
         Top             =   2400
         Width           =   1935
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
         TabIndex        =   22
         Top             =   2400
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
         TabIndex        =   21
         Top             =   1920
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
         TabIndex        =   19
         Top             =   1920
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
         TabIndex        =   17
         Top             =   1440
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
         TabIndex        =   16
         Top             =   1440
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Motivo de baja:"
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
         Left            =   4200
         TabIndex        =   13
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Origen:"
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
         TabIndex        =   11
         Top             =   2880
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
         TabIndex        =   10
         Top             =   960
         Width           =   2175
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
         TabIndex        =   9
         Top             =   960
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
         TabIndex        =   7
         Top             =   960
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
         TabIndex        =   5
         Top             =   960
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
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Nro. Solicitud:"
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
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   5655
         Left            =   0
         Picture         =   "frm_solicitudbaja.frx":2E44
         Stretch         =   -1  'True
         Top             =   240
         Width           =   10335
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   8160
      Picture         =   "frm_solicitudbaja.frx":33CB
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1575
   End
End
Attribute VB_Name = "frm_solicitudbaja"
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

Private Sub b_busca_Click()
frm_solibajabusca.Show vbModal

End Sub

Private Sub b_cancelar_Click()
t_cod.Text = ""
mfec.Text = "__/__/____"
mhora.Text = "__:__"
labusu.Caption = ""
t_base.Text = ""
t_mat.Text = ""
t_ced.Text = ""
t_cced.Text = ""
t_conv.Text = ""
t_nom.Text = ""
t_tel.Text = ""
t_cel.Text = ""
Check2.Value = 0
cbodes.ListIndex = -1
cbohas.ListIndex = -1
cboorig.ListIndex = -1
cbomot.ListIndex = -1
Check1.Value = 0
mffin.Text = "__/__/____"
cborec.ListIndex = -1
t_otro.Text = ""
t_accion.Text = ""

      XAlta = 0
      b_edita.Enabled = True
      b_cancelar.Enabled = False
      b_graba.Enabled = False
      b_busca.Enabled = True
      b_acciones.Enabled = True
      b_infos.Enabled = True
      t_accion.Enabled = False
      Frame1.Enabled = False

End Sub


Private Sub b_edita_Click()
Frame1.Enabled = True
b_edita.Enabled = False
b_cancelar.Enabled = True
b_graba.Enabled = True
b_busca.Enabled = False
b_acciones.Enabled = False
b_infos.Enabled = False
XAlta = 2
t_nom.SetFocus
If WElusuario = "PINMEDIATO" Or WElusuario = "JFERNAN" Or WElusuario = "MCOSTA" Or WElusuario = "MPEREZ" Then
   Check1.Enabled = True
   mffin.Enabled = True
   cborec.Enabled = True
Else
   Check1.Enabled = False
   mffin.Enabled = False
   cborec.Enabled = False
End If
If t_accion.Text <> "" Then
   t_accion.Enabled = True
Else
   t_accion.Enabled = False
End If

End Sub

Private Sub b_graba_Click()
Dim Xelerror As Integer
Dim Fechaesta As Date
Fechaesta = Date - 31

On Error GoTo Querr

Xelerror = 0
If Trim(t_mat.Text) = "" Then
   Xelerror = 1
End If
If Trim(t_ced.Text) = "" Then
   Xelerror = 2
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
If cbodes.ListIndex < 0 Then
   Xelerror = 8
End If
If cbohas.ListIndex < 0 Then
   Xelerror = 9
End If
If cboorig.ListIndex < 0 Then
   Xelerror = 10
End If
If cbomot.ListIndex < 0 Then
   Xelerror = 11
End If
If Xelerror = 0 Then
   If XAlta = 1 Then
      If t_mat.Text <> "" Then
         data_buscasi.RecordSource = "select * from solic_bajas where matricula =" & t_mat.Text & " and fecha >=#" & Format(Fechaesta, "yyyy/mm/dd") & "#"
         data_buscasi.Refresh
         If data_buscasi.Recordset.RecordCount > 0 Then
            MsgBox "El socio ya tiene registro de BAJA. VERIFIQUE!", vbInformation
         Else
            data_par.Recordset.Edit
            data_par.Recordset("nro_accadm") = data_par.Recordset("nro_accadm") + 1
            data_par.Recordset.Update
            Data1.Recordset.AddNew
            Data1.Recordset("id") = data_par.Recordset("nro_accadm")
            Data1.Recordset("fecha") = mfec.Text
            Data1.Recordset("hora") = mhora.Text
            Data1.Recordset("base") = t_base.Text
            Data1.Recordset("usuario") = labusu.Caption
            Data1.Recordset("matricula") = t_mat.Text
            Data1.Recordset("cedula") = t_ced.Text
            Data1.Recordset("codced") = t_cced.Text
            Data1.Recordset("convenio") = t_conv.Text
            Data1.Recordset("nombre") = t_nom.Text
            Data1.Recordset("telefono") = t_tel.Text
            Data1.Recordset("celular") = t_cel.Text
            Data1.Recordset("hora1") = cbodes.Text
            Data1.Recordset("hora2") = cbohas.Text
            Data1.Recordset("origen") = cboorig.Text
            Data1.Recordset("motivo") = cbomot.Text
            Data1.Recordset("hora1id") = cbodes.ListIndex
            Data1.Recordset("hora2id") = cbohas.ListIndex
            Data1.Recordset("origid") = cboorig.ListIndex
            Data1.Recordset("motid") = cbomot.ListIndex
            Data1.Recordset("terminado") = Check1.Value
            If t_otro.Text <> "" Then
               Data1.Recordset("otrotel") = t_otro.Text
            End If
            If mffin.Text <> "__/__/____" Then
               Data1.Recordset("fechafin") = mffin.Text
            End If
            If cborec.ListIndex >= 0 Then
               Data1.Recordset("resultado") = cborec.Text
               Data1.Recordset("resulid") = cborec.ListIndex
            End If
            Data1.Recordset("contrato") = Check2.Value
            Data1.Recordset.Update
            t_cod.Text = data_par.Recordset("nro_accadm")
            XAlta = 0
            b_edita.Enabled = True
            b_cancelar.Enabled = False
            b_graba.Enabled = False
            b_busca.Enabled = True
            b_acciones.Enabled = True
            b_infos.Enabled = True
            t_accion.Enabled = False
            Frame1.Enabled = False
         End If
      Else
         MsgBox "No ingresó matrícula", vbExclamation
      End If
   Else
      Data1.RecordSource = "select * from solic_bajas where id =" & t_cod.Text
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.Edit
         Data1.Recordset("fecha") = mfec.Text
         Data1.Recordset("hora") = mhora.Text
         Data1.Recordset("base") = t_base.Text
         Data1.Recordset("usuario") = labusu.Caption
         Data1.Recordset("matricula") = t_mat.Text
         Data1.Recordset("cedula") = t_ced.Text
         Data1.Recordset("codced") = t_cced.Text
         Data1.Recordset("convenio") = t_conv.Text
         Data1.Recordset("nombre") = t_nom.Text
         Data1.Recordset("telefono") = t_tel.Text
         Data1.Recordset("celular") = t_cel.Text
         Data1.Recordset("hora1") = cbodes.Text
         Data1.Recordset("hora2") = cbohas.Text
         Data1.Recordset("origen") = cboorig.Text
         Data1.Recordset("motivo") = cbomot.Text
         Data1.Recordset("hora1id") = cbodes.ListIndex
         Data1.Recordset("hora2id") = cbohas.ListIndex
         Data1.Recordset("origid") = cboorig.ListIndex
         Data1.Recordset("motid") = cbomot.ListIndex
         Data1.Recordset("terminado") = Check1.Value
         If mffin.Text <> "__/__/____" Then
            Data1.Recordset("fechafin") = mffin.Text
         Else
            If IsNull(Data1.Recordset("fechafin")) = False Then
               Data1.Recordset("fechafin") = Null
            End If
         End If
         If cborec.ListIndex >= 0 Then
            Data1.Recordset("resultado") = cborec.Text
            Data1.Recordset("resulid") = cborec.ListIndex
         Else
            If IsNull(Data1.Recordset("resultado")) = False Then
               Data1.Recordset("resultado") = Null
               Data1.Recordset("resulid") = cborec.ListIndex
            End If
         End If
         If t_otro.Text <> "" Then
            Data1.Recordset("otrotel") = t_otro.Text
         End If
         Data1.Recordset("contrato") = Check2.Value
         Data1.Recordset.Update
         XAlta = 0
         b_edita.Enabled = True
         b_cancelar.Enabled = False
         b_graba.Enabled = False
         b_busca.Enabled = True
         b_acciones.Enabled = True
         b_infos.Enabled = True
         t_accion.Enabled = False
         Frame1.Enabled = False
      Else
         MsgBox "Error al grabar el registro, no se encuentra el código", vbCritical
         Unload Me
      End If
   End If
Else
   MsgBox "Hay algún campo que falta ingresar datos. Verifique!", vbInformation
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
frm_solibajainf.Show vbModal

End Sub

Private Sub b_nuevo_Click()
Frame1.Enabled = True
t_mat.Text = ""
t_ced.Text = ""
t_cced.Text = ""
t_conv.Text = ""
t_nom.Text = ""
t_tel.Text = ""
t_cel.Text = ""
Check2.Value = 0
cbodes.ListIndex = -1
cbohas.ListIndex = -1
cboorig.ListIndex = -1
cbomot.ListIndex = -1
Check1.Value = 0
mffin.Text = "__/__/____"
cborec.ListIndex = -1
t_otro.Text = ""
t_accion.Text = ""


t_mat.SetFocus

t_cod.Text = "......"
mfec.Text = Date
mhora.Text = Format(Time, "HH:mm")
t_base.Text = frm_menu.data_parse.Recordset("base")

labusu.Caption = WElusuario

XAlta = 1
b_edita.Enabled = False
b_cancelar.Enabled = True
b_graba.Enabled = True
b_busca.Enabled = False
b_acciones.Enabled = False
b_infos.Enabled = False
'b_edac.Enabled = False

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

Private Sub Form_Load()
data_par.DatabaseName = App.path & "\paramb.mdb"
data_par.RecordSource = "paramb"
data_par.Refresh

data_buscasi.Connect = "odbc;dsn=" & Xconexrmt & ";"


data_orig.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_orig.RecordSource = "select * from solbaj_orig"
data_orig.Refresh
data_orig.Recordset.MoveFirst
Do While Not data_orig.Recordset.EOF
   cboorig.AddItem data_orig.Recordset("descrip")
   data_orig.Recordset.MoveNext
Loop
data_orig.RecordSource = "select * from solbaj_rec"
data_orig.Refresh
data_orig.Recordset.MoveFirst
Do While Not data_orig.Recordset.EOF
   cborec.AddItem data_orig.Recordset("descrip")
   data_orig.Recordset.MoveNext
Loop

data_orig.RecordSource = "select * from motivos where mc_numero >='" & "B01" & "' and mc_numero <='" & "B50" & "'"
data_orig.Refresh
data_orig.Recordset.MoveFirst
Do While Not data_orig.Recordset.EOF
   cbomot.AddItem data_orig.Recordset("mc_desc")
   data_orig.Recordset.MoveNext
Loop
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from solic_bajas"
Data1.Refresh

data_busca.Connect = "odbc;dsn=" & Xconexrmt & ";"


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
      data_busca.RecordSource = "Select * from clientes where cl_cedula =" & t_ced.Text
      data_busca.Refresh
      If data_busca.Recordset.RecordCount > 0 Then
         If data_busca.Recordset("estado") = 2 Then
            MsgBox "Atención! El socio ya se encuentra de baja en el padrón.", vbExclamation
         End If
         data_buscasi.RecordSource = "select * from solic_bajas where cedula =" & t_ced.Text
         data_buscasi.Refresh
         If data_buscasi.Recordset.RecordCount > 0 Then
            data_buscasi.Recordset.MoveLast
            MsgBox "El socio ya tiene " & data_buscasi.Recordset.RecordCount & " registro de BAJA.", vbInformation
         End If
         
         If IsNull(data_busca.Recordset("cl_apellid")) = False Then
            t_nom.Text = data_busca.Recordset("cl_apellid")
         Else
            t_nom.Text = "NN"
         End If
         If IsNull(data_busca.Recordset("cl_codigo")) = False Then
            t_mat.Text = data_busca.Recordset("cl_codigo")
         Else
            t_mat.Text = 0
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
            t_tel.Text = Val(data_busca.Recordset("cl_telefon"))
         Else
            t_tel.Text = 0
         End If
         If IsNull(data_busca.Recordset("cl_dpto")) = False Then
            t_cel.Text = Val(data_busca.Recordset("cl_dpto"))
         Else
            t_cel.Text = 0
         End If
         t_cced.SetFocus
      Else
         MsgBox "Cédula no encontrada"
         t_ced.Text = Empty
         t_cced.SetFocus
      End If
   End If
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
      MsgBox "Solo números"
      KeyAscii = 8
    End If
End If

End Sub

Private Sub t_cel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_cel.Text = "" Then
      MsgBox "Campo obligatorio"
      t_cel.SetFocus
   Else
      cbodes.SetFocus
   End If
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
      MsgBox "Solo números"
      KeyAscii = 8
    End If
End If

End Sub

Private Sub t_conv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If t_mat.Text = "" Then
      MsgBox "Campo obligatorio"
      t_mat.SetFocus
   Else
      data_busca.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
      data_busca.Refresh
      If data_busca.Recordset.RecordCount > 0 Then
         If data_busca.Recordset("estado") = 2 Then
            MsgBox "Atención! El socio ya se encuentra de baja en el padrón.", vbExclamation
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
            t_tel.Text = Val(data_busca.Recordset("cl_telefon"))
         Else
            t_tel.Text = 0
         End If
         If IsNull(data_busca.Recordset("cl_dpto")) = False Then
            t_cel.Text = Val(data_busca.Recordset("cl_dpto"))
         Else
            t_cel.Text = 0
         End If
         t_ced.SetFocus
      Else
         MsgBox "Matrícula no encontrada"
         t_mat.Text = Empty
         t_ced.SetFocus
      End If
   End If
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
      MsgBox "Solo números"
      KeyAscii = 8
    End If
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_tel.SetFocus
End If

End Sub

Private Sub t_otro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbodes.SetFocus
End If

End Sub

Private Sub t_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cel.SetFocus
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
      MsgBox "Solo números"
      KeyAscii = 8
    End If
End If

End Sub
