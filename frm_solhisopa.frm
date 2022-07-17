VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_solhisopa 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ingreso de solicitudes de hisopados"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12465
   Icon            =   "frm_solhisopa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   12465
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_buscafact 
      Caption         =   "data_buscafact"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
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
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frm_solhisopa.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Descargar planilla excel"
      Top             =   7680
      Width           =   495
   End
   Begin VB.TextBox t_busca 
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
      MaxLength       =   7
      TabIndex        =   37
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sol_hisopos"
      Top             =   7080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3360
      Top             =   2640
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_solhisopa.frx":0B14
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "frm_solhisopa.frx":0B28
      TabIndex        =   27
      ToolTipText     =   "Haga doble click para editar un registro"
      Top             =   5280
      Width           =   12015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos de solicitud"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   12015
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Requiere de Certificado"
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
         Left            =   7920
         TabIndex        =   58
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CheckBox chderiva 
         BackColor       =   &H0000FFFF&
         Caption         =   "Derivar a CMT"
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
         TabIndex        =   57
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   9840
         TabIndex        =   56
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Data data_arch 
         Caption         =   "data_arch"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton b_subir 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11640
         Picture         =   "frm_solhisopa.frx":37E7
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Subir archivo de resultado en PDF"
         Top             =   3360
         Width           =   375
      End
      Begin MSMask.MaskEdBox mfcoord 
         Height          =   375
         Left            =   9600
         TabIndex        =   52
         Top             =   3360
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
      Begin VB.CheckBox chcoord 
         BackColor       =   &H00C00000&
         Caption         =   "Coordinado"
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
         Left            =   7920
         TabIndex        =   51
         Top             =   3360
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frm_solhisopa.frx":3D71
         Left            =   7920
         List            =   "frm_solhisopa.frx":3D73
         TabIndex        =   49
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Cancelado"
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
         Left            =   6120
         TabIndex        =   47
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Realizado en mutualista"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         TabIndex        =   46
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox t_obscierre 
         Height          =   615
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   3840
         Width           =   5775
      End
      Begin MSMask.MaskEdBox mfrea 
         Height          =   375
         Left            =   1920
         TabIndex        =   43
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16711680
         ForeColor       =   16777215
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
      Begin VB.ComboBox cbolocal 
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
         Left            =   8400
         TabIndex        =   39
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   3720
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Borrar"
         Height          =   495
         Left            =   10800
         Picture         =   "frm_solhisopa.frx":3D75
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Limpiar campos del formulario"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grabar"
         Height          =   495
         Left            =   10800
         Picture         =   "frm_solhisopa.frx":42FF
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Grabar formulario"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sintomático"
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
         Height          =   315
         Left            =   8160
         TabIndex        =   32
         Top             =   2880
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asintomático"
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
         Height          =   315
         Left            =   5880
         TabIndex        =   31
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ComboBox cbozon 
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
         ItemData        =   "frm_solhisopa.frx":4889
         Left            =   8400
         List            =   "frm_solhisopa.frx":4893
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mfcon 
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   2880
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
      Begin VB.TextBox t_tel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5400
         MaxLength       =   45
         TabIndex        =   22
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox t_cel 
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
         Left            =   1920
         MaxLength       =   45
         TabIndex        =   20
         Top             =   2400
         Width           =   1935
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
         ItemData        =   "frm_solhisopa.frx":48A3
         Left            =   9240
         List            =   "frm_solhisopa.frx":48B6
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   16
         Top             =   1920
         Width           =   5175
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   15
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox t_nom 
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
         Left            =   1920
         MaxLength       =   90
         TabIndex        =   12
         Top             =   1080
         Width           =   5175
      End
      Begin MSMask.MaskEdBox mfpos 
         Height          =   375
         Left            =   10200
         TabIndex        =   10
         Top             =   480
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
      Begin VB.TextBox t_mat 
         Alignment       =   1  'Right Justify
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
         Left            =   5880
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox t_codced 
         Alignment       =   1  'Right Justify
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
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   240
         TabIndex        =   55
         ToolTipText     =   "Doble click aquí para ver el resultado"
         Top             =   4440
         Width           =   5055
      End
      Begin VB.Label labid 
         Height          =   255
         Left            =   3120
         TabIndex        =   50
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Motivo:"
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
         Left            =   7920
         TabIndex        =   48
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Observación del cierre:"
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
         Height          =   615
         Left            =   240
         TabIndex        =   44
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha Realizado:"
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
         TabIndex        =   42
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label labyaesta 
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label labcodconv 
         Height          =   375
         Left            =   4320
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Estado Clínico:"
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
         Left            =   4200
         TabIndex        =   30
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zona:"
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
         Left            =   7320
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de contacto:"
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
         TabIndex        =   23
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Localidad:"
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
         Left            =   7320
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mutualista:"
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
         Left            =   7320
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre paciente:"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha posible de HNF"
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
         Left            =   7800
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Matrícula"
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
         Left            =   4320
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ingrese cédula:"
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
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Doble click aquí para abrir la agenda de HNF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      MouseIcon       =   "frm_solhisopa.frx":48EF
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   8160
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solicitud médica de HNF SAPP"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00404040&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   36
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label labusua 
      BackColor       =   &H00808080&
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
      Left            =   8520
      TabIndex        =   29
      Top             =   7800
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "Usuario actual:"
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
      Left            =   7080
      TabIndex        =   28
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label labfec 
      BackColor       =   &H00808080&
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
      Left            =   10560
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Fecha actual:"
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
      Left            =   9000
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Picture         =   "frm_solhisopa.frx":4E79
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2775
   End
End
Attribute VB_Name = "frm_solhisopa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdfpath, pdfpath1 As String
Public pdffile As ADODB.Stream
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub b_subir_Click()

If ControlUsuario("SubirArchivo") = 1 Then
   Shell App.path & "\subircovid.exe", vbNormalFocus
Else
   MsgBox "Usuario no autorizado."
End If

End Sub

Private Sub cbolocal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbozon.SetFocus
End If

End Sub

Private Sub cbolocal_LostFocus()
If cbolocal.Text <> "" Then
   Consulta_zonas
End If

End Sub

Private Sub cbomut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_dir1.SetFocus
End If

End Sub

Private Sub cbozon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cel.SetFocus
End If

End Sub

Private Sub chcoord_Click()
If chcoord.Value = 1 Then
   mfcoord.Text = Format(Date, "dd/mm/yyyy")
Else
   mfcoord.Text = "__/__/____"
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub Combo1_LostFocus()
If Combo1.Text <> "" Then
   Consulta_motivos
End If

End Sub

Private Sub Command1_Click()
Dim Sihayerror As Integer
On Error GoTo Algrabar

Sihayerror = 0
If t_ced.Text <> "" Then
    If Trim(labyaesta.Caption) <> "" Then
       If Trim(labyaesta.Caption) = "1" Then
          If Trim(labid.Caption) = "" Then
             Sihayerror = 2
          End If
       End If
    End If
    Consulta_siexiste
End If
If t_ced.Text = "" Then
   Sihayerror = 1
End If
If t_codced.Text = "" Then
   Sihayerror = 1
End If
If t_nom.Text = "" Then
   Sihayerror = 1
End If
If mfpos.Text = "__/__/____" Then
   Sihayerror = 1
End If
If cbomut.Text = "CCOU" Or cbomut.Text = "SMI" Or cbomut.Text = "CASA DE GALICIA" Or cbomut.Text = "H.EVANGELICO" Or cbomut.Text = "UNIVERSAL" Then
Else
   Sihayerror = 1
End If
If Trim(cbolocal.Text) = "" Then
   Sihayerror = 1
End If
If t_cel.Text = "" Then
   If t_tel.Text = "" Then
      Sihayerror = 1
   End If
End If
If t_dir1.Text = "" Then
   Sihayerror = 1
End If
If cbozon.Text = "" Then
   Sihayerror = 1
End If
'If mfcon.Text = "__/__/____" Then
'   Sihayerror = 1
'End If
If Option1.Value = False And Option2.Value = False Then
   Sihayerror = 1
End If
If cbomut.Text = "CASA DE GALICIA" Then
   If mfrea.Text = "__/__/____" Then
      mfrea.Text = Format(Date, "dd/mm/yyyy")
   End If
End If

If Sihayerror = 1 Or Sihayerror = 2 Then
   If Sihayerror = 2 Then
      MsgBox "Ya existe ingresada una solicitud con esta cédula.", vbCritical
      Borra_campos
   Else
      MsgBox "Falta datos, verifique.", vbCritical
   End If
Else
   If Trim(labid.Caption) <> "" Then
      If ControlUsuario(Command1.Name) = 1 Then
         data_graba.RecordSource = "select * from sol_hisopos where id =" & Val(labid.Caption)
         data_graba.Refresh
         If data_graba.Recordset.RecordCount > 0 Then
            data_graba.Recordset.MoveFirst
            data_graba.Recordset.Edit
            If labcodconv.Caption <> "" Then
               data_graba.Recordset("convenio") = labcodconv.Caption
            End If
            data_graba.Recordset("dir1") = t_dir1.Text
            If Trim(t_dir2.Text) <> "" Then
               data_graba.Recordset("dir2") = t_dir2.Text
            End If
            data_graba.Recordset("localid") = cbolocal.Text
            data_graba.Recordset("mutual") = cbomut.Text
            data_graba.Recordset("celular") = t_cel.Text
            data_graba.Recordset("telefono") = t_tel.Text
            data_graba.Recordset("zona") = cbozon.Text
            data_graba.Recordset("certif") = Check3.Value
            If mfcon.Text <> "__/__/____" Then
               data_graba.Recordset("fec_contact") = Format(mfcon.Text, "dd/mm/yyyy")
            Else
               If IsNull(data_graba.Recordset("fec_contact")) = False Then
                  data_graba.Recordset("fec_contact") = Null
               End If
            End If
            data_graba.Recordset("fec_posible") = Format(mfpos.Text, "dd/mm/yyyy")
            If Option1.Value = True Then
               data_graba.Recordset("estado_cli") = "Asintomático"
            End If
            If Option2.Value = True Then
               data_graba.Recordset("estado_cli") = "Sintomático"
            End If
            data_graba.Recordset("realiza_mut") = Check1.Value
            If mfrea.Text <> "__/__/____" Then
               data_graba.Recordset("fecha_fact") = mfrea.Text
            End If
            If Combo1.Text <> "" Then
               data_graba.Recordset("mot_cierre") = Combo1.Text
            End If
            If t_obscierre.Text <> "" Then
               data_graba.Recordset("obs_cierre") = t_obscierre.Text
            End If
            data_graba.Recordset("coord") = chcoord.Value
            If mfcoord.Text <> "__/__/____" Then
               data_graba.Recordset("fec_coord") = mfcoord.Text
            Else
               If IsNull(data_graba.Recordset("fec_coord")) = False Then
                  data_graba.Recordset("fec_coord") = Null
               End If
            End If
            If IsNull(data_graba.Recordset("deriva")) = False Then
               If data_graba.Recordset("deriva") <> chderiva.Value Then
                  data_graba.Recordset("deriva") = chderiva.Value
               End If
            Else
               data_graba.Recordset("deriva") = chderiva.Value
            End If
            data_graba.Recordset.Update
            Data1.Refresh
            Borra_campos
            labcodconv.Caption = ""
            labyaesta.Caption = ""
            If t_ced.Enabled = True Then
               t_ced.SetFocus
            End If
            MsgBox "Registro grabado correctamente."
         
         End If
      End If
   Else
      data_graba.Recordset.AddNew
      data_graba.Recordset("fecha") = Format(labfec.Caption, "dd/mm/yyyy")
      data_graba.Recordset("hora") = Format(Time, "HH:mm")
      data_graba.Recordset("cedula") = t_ced.Text
      data_graba.Recordset("codced") = t_codced.Text
      If t_mat.Text <> "" Then
         data_graba.Recordset("matricula") = t_mat.Text
      Else
         data_graba.Recordset("matricula") = 0
      End If
      data_graba.Recordset("nombre") = t_nom.Text
      If labcodconv.Caption <> "" Then
         data_graba.Recordset("convenio") = labcodconv.Caption
      End If
      data_graba.Recordset("dir1") = t_dir1.Text
      If Trim(t_dir2.Text) <> "" Then
         data_graba.Recordset("dir2") = t_dir2.Text
      End If
      data_graba.Recordset("localid") = cbolocal.Text
      data_graba.Recordset("mutual") = cbomut.Text
      data_graba.Recordset("celular") = t_cel.Text
      data_graba.Recordset("telefono") = t_tel.Text
      data_graba.Recordset("zona") = cbozon.Text
      data_graba.Recordset("certif") = Check3.Value
      If mfcon.Text <> "__/__/____" Then
         data_graba.Recordset("fec_contact") = Format(mfcon.Text, "dd/mm/yyyy")
      End If
      data_graba.Recordset("fec_posible") = Format(mfpos.Text, "dd/mm/yyyy")
      If Option1.Value = True Then
         data_graba.Recordset("estado_cli") = "Asintomático"
      End If
      If Option2.Value = True Then
         data_graba.Recordset("estado_cli") = "Sintomático"
      End If
      data_graba.Recordset("usua_sist") = WElusuario
      data_graba.Recordset("usua_nombre") = labusua.Caption
      data_graba.Recordset("base") = frm_menu.data_parse.Recordset("base")
      data_graba.Recordset("realiza_mut") = Check1.Value
      If mfrea.Text <> "__/__/____" Then
         data_graba.Recordset("fecha_fact") = mfrea.Text
      End If
      If Combo1.Text <> "" Then
         data_graba.Recordset("mot_cierre") = Combo1.Text
      End If
      If t_obscierre.Text <> "" Then
         data_graba.Recordset("obs_cierre") = t_obscierre.Text
      End If
      data_graba.Recordset("coord") = chcoord.Value
      If mfcoord.Text <> "__/__/____" Then
         data_graba.Recordset("fec_coord") = mfcoord.Text
      End If
      data_graba.Recordset("deriva") = chderiva.Value
      data_graba.Recordset.Update
      Data1.Refresh
      Borra_campos
      labcodconv.Caption = ""
      labyaesta.Caption = ""
      If t_ced.Enabled = True Then
         t_ced.SetFocus
      End If
      MsgBox "Registro grabado correctamente."
   End If
   
End If

Exit Sub

Algrabar:
        If Err.Number = 3081 Then
           MsgBox "Error :" & Err.Description
        Else
           MsgBox "Error:" & Err.Description
        End If
        
End Sub

Private Sub Command2_Click()

t_ced.Enabled = True
t_codced.Enabled = True
t_mat.Enabled = True
t_nom.Enabled = True

Borra_campos
Data2.RecordSource = "select * from usuarios where usuario ='" & Trim(WElusuario) & "'"
Data2.Refresh
If Data2.Recordset.RecordCount > 0 Then
   labusua.Caption = Data2.Recordset("nombre")
End If

t_ced.SetFocus

End Sub

Private Sub Command4_Click()
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
Promo = InputBox("Ingrese 0 (CERO) para listar todo o 1 (UNO) para listar sólo pendientes.", "Solicitud de hisopados", 0)

frm_solhisopa.MousePointer = 11
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0

If desde <> "" And hasta <> "" Then
   If Val(Promo) = 0 Then
      Data3.RecordSource = "select * from sol_hisopos where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# order by fecha,hora"
   Else
      Data3.RecordSource = "select * from sol_hisopos where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and fecha_fact is null order by fecha,hora"
   End If
   Data3.Refresh
   frm_solhisopa.MousePointer = 11
   If Data3.Recordset.RecordCount > 0 Then
      Data3.Recordset.MoveFirst
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      Xarchexel22.Name = Trim("SolicitudHNF")
      Xlibexel22.SaveAs ("C:\planillas\SolicitudHNF.xls")
      Xarchtex = "C:\planillas\SolicitudHNF.xls"
      Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Cells(Xlin, XCol) = "INFORME DE SOLICITUD HNF DESDE: " & desde & " HASTA: " & hasta
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
      Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "FECHA"
      XCol = XCol + 1
      Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "HORA"
      XCol = XCol + 1
      Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 30
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRE PACIENTE"
      XCol = XCol + 1
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 13
      Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "CELULAR"
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "TELEFONO"
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "FEC.CONTAC"
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 13
      Xarchexel22.Cells(Xlin, XCol) = "FEC.POSIBLE"
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "ESTADO"
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "ZONA"
      XCol = XCol + 1
      Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 18
      Xarchexel22.Cells(Xlin, XCol) = "LOCALIDAD"
      XCol = XCol + 1
      Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 40
      Xarchexel22.Cells(Xlin, XCol) = "DIRECCION"
      XCol = XCol + 1
      Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "MUTUALISTA"
      XCol = XCol + 1
      Xarchexel22.Range("N" & Trim(str(Xlin))).ColumnWidth = 25
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRE MEDICO SOL."
      XCol = XCol + 1
      Xarchexel22.Range("O" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "ENFERMERO"
      XCol = XCol + 1
      Xarchexel22.Range("P" & Trim(str(Xlin))).ColumnWidth = 13
      Xarchexel22.Cells(Xlin, XCol) = "FECHA FACT."
      XCol = XCol + 1
      Xarchexel22.Range("Q" & Trim(str(Xlin))).ColumnWidth = 9
      Xarchexel22.Cells(Xlin, XCol) = "BASE"
      XCol = XCol + 1
      Xarchexel22.Range("R" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "OTRO"
      XCol = XCol + 1
      Xarchexel22.Range("S" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "COORD?"
      XCol = XCol + 1
      Xarchexel22.Range("T" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "FECHA COORD"
      XCol = XCol + 1
      Xarchexel22.Range("U" & Trim(str(Xlin))).ColumnWidth = 25
      Xarchexel22.Cells(Xlin, XCol) = "ANÁLISIS FACTURADO"
      XCol = XCol + 1
      Xarchexel22.Range("V" & Trim(str(Xlin))).ColumnWidth = 15
      Xarchexel22.Cells(Xlin, XCol) = "REQUIERE CERT"
      
      Xlin = Xlin + 1
      XCol = 1
        
      Do While Not Data3.Recordset.EOF
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data3.Recordset("fecha"), "dd/mm/yyyy")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("hora")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("nombre")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Trim(str(Data3.Recordset("cedula"))) & "-" & Trim(str(Data3.Recordset("codced")))
         XCol = XCol + 1
         If IsNull(Data3.Recordset("celular")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = "'" & Data3.Recordset("celular")
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("telefono")) = False Then
            If Trim(Data3.Recordset("telefono")) <> "" Then
               Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("telefono")
            End If
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("fec_contact")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data3.Recordset("fec_contact"), "dd/mm/yyyy")
         End If
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data3.Recordset("fec_posible"), "dd/mm/yyyy")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("estado_cli")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("zona")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("localid")
         XCol = XCol + 1
         If IsNull(Data3.Recordset("dir2")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("dir1") & " " & Data3.Recordset("dir2")
         Else
            Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("dir1")
         End If
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("mutual")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("usua_nombre")
         XCol = XCol + 1
         If IsNull(Data3.Recordset("enf_realiza")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("enf_realiza")
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("fecha_fact")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data3.Recordset("fecha_fact"), "dd/mm/yyyy")
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("base_fact")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("base_fact")
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("realiza_mut")) = False Then
            If Data3.Recordset("realiza_mut") = 1 Then
               Xarchexel22.Cells(Xlin, XCol) = "EN MUTUALISTA"
            End If
         Else
            If IsNull(Data3.Recordset("cancelado")) = False Then
               If Data3.Recordset("cancelado") = 1 Then
                  Xarchexel22.Cells(Xlin, XCol) = "CANCELADO"
               End If
            End If
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("coord")) = False Then
            If Data3.Recordset("coord") = 1 Then
               Xarchexel22.Cells(Xlin, XCol) = "SI"
            Else
               Xarchexel22.Cells(Xlin, XCol) = "NO"
            End If
         Else
            Xarchexel22.Cells(Xlin, XCol) = "NO"
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("fec_coord")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data3.Recordset("fec_coord"), "dd/mm/yyyy")
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("nom_prod")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("nom_prod")
         Else
            data_buscafact.RecordSource = "select * from linmmdd where cod_prod in (30081,30084,30085) and fecha >=#" & Format(Data3.Recordset("fecha")) & "# and ced_socio =" & Data3.Recordset("cedula")
            data_buscafact.Refresh
            If data_buscafact.Recordset.RecordCount > 0 Then
               Xarchexel22.Cells(Xlin, XCol) = data_buscafact.Recordset("nom_prod")
            End If
         End If
         XCol = XCol + 1
         If IsNull(Data3.Recordset("certif")) = False Then
            If Data3.Recordset("certif") = 1 Then
               Xarchexel22.Cells(Xlin, XCol) = "SI"
            Else
               Xarchexel22.Cells(Xlin, XCol) = "NO"
            End If
         End If
         
         Xlin = Xlin + 1
         XCol = 1
         Xtotreg = Xtotreg + 1
         Data3.Recordset.MoveNext
      Loop
      Xlin = Xlin + 1
      XCol = 1
      Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
      Xlin = Xlin + 1
      XCol = 1
      Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
      frm_solhisopa.MousePointer = 0
      
      Xlibexel22.Save
      Xlibexel22.Close
      Xobjexel22.Quit
      Xlabrir3.Workbooks.Open Xarchtex, , False
      Xlabrir3.Visible = True
      Xlabrir3.WindowState = xlMaximized
   
   Else
      frm_solhisopa.MousePointer = 0
      MsgBox "No hay registros"
   End If
Else
   frm_solhisopa.MousePointer = 0
   MsgBox "Faltan fechas"
End If

End Sub

Private Sub Command5_Click()
'''t_id.Text = Data1.Recordset("nroid")

'''''Data3.RecordSource = "Select * from arch_orden where idsrv =" & Data2.Recordset("numero")
If data_arch.Recordset.RecordCount > 0 Then
   Set pdffile = New ADODB.Stream
   pdffile.Type = adTypeBinary
   pdffile.Open
   If IsNull(data_arch.Recordset("arch")) = False Then
      pdffile.Write data_arch.Recordset("arch").Value
      Dim pdfname As String
      pdfname = "temporal"
      pdffile.SaveToFile "" & App.path & "\laboratorio\" & pdfname & ".pdf", adSaveCreateOverWrite
      pdffile.Close
      Set pdffile = Nothing
   Else
      MsgBox "no hay archivo"
   End If
Else
   MsgBox "No existe registro"
End If

End Sub

Private Sub DBGrid1_DblClick()

t_ced.Enabled = False
t_codced.Enabled = False
t_mat.Enabled = False
t_nom.Enabled = False
labid.Caption = Data1.Recordset("id")
labfec.Caption = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
t_ced.Text = Data1.Recordset("cedula")
t_codced.Text = Data1.Recordset("codced")
If IsNull(Data1.Recordset("matricula")) = False Then
   t_mat.Text = Data1.Recordset("matricula")
Else
   t_mat.Text = ""
End If
If IsNull(Data1.Recordset("deriva")) = False Then
   chderiva.Value = Data1.Recordset("deriva")
Else
   chderiva.Value = 0
End If
If IsNull(Data1.Recordset("certif")) = False Then
   Check3.Value = Data1.Recordset("certif")
Else
   Check3.Value = 0
End If
t_nom.Text = Data1.Recordset("nombre")
If IsNull(Data1.Recordset("convenio")) = False Then
   labcodconv.Caption = Data1.Recordset("convenio")
Else
   labcodconv.Caption = ""
End If
If IsNull(Data1.Recordset("dir1")) = False Then
   t_dir1.Text = Data1.Recordset("dir1")
Else
   t_dir1.Text = ""
End If
If IsNull(Data1.Recordset("si_result")) = True Then
   Label21.Caption = "NO TIENE RESULTADO CARGADO."
Else
   Label21.Caption = "TIENE RESULTADO CARGADO."
End If

If IsNull(Data1.Recordset("dir2")) = False Then
   t_dir2.Text = Data1.Recordset("dir2")
End If
If IsNull(Data1.Recordset("localid")) = False Then
   cbolocal.Text = Data1.Recordset("localid")
Else
   cbolocal.Text = ""
End If
If IsNull(Data1.Recordset("mutual")) = False Then
   cbomut.Text = Data1.Recordset("mutual")
Else
   cbomut.Text = ""
End If
If IsNull(Data1.Recordset("celular")) = False Then
   t_cel.Text = Data1.Recordset("celular")
Else
   t_cel.Text = ""
End If
If IsNull(Data1.Recordset("telefono")) = False Then
   t_tel.Text = Data1.Recordset("telefono")
Else
   t_tel.Text = ""
End If
If IsNull(Data1.Recordset("zona")) = False Then
   cbozon.Text = Data1.Recordset("zona")
Else
   cbozon.Text = ""
End If
If IsNull(Data1.Recordset("fec_contact")) = False Then
   mfcon.Text = Format(Data1.Recordset("fec_contact"), "dd/mm/yyyy")
Else
   mfcon.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("fec_posible")) = False Then
   mfpos.Text = Format(Data1.Recordset("fec_posible"), "dd/mm/yyyy")
Else
   mfpos.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("estado_cli")) = False Then
   If Data1.Recordset("estado_cli") = "Asintomático" Then
      Option1.Value = True
   Else
      If Data1.Recordset("estado_cli") = "Sintomático" Then
         Option2.Value = True
      End If
   End If
End If
If IsNull(Data1.Recordset("usua_nombre")) = False Then
   labusua.Caption = Data1.Recordset("usua_nombre")
Else
   labusua.Caption = ""
End If
If IsNull(Data1.Recordset("realiza_mut")) = False Then
   Check1.Value = Data1.Recordset("realiza_mut")
Else
   Check1.Value = 0
End If
If IsNull(Data1.Recordset("cancelado")) = False Then
   Check2.Value = Data1.Recordset("cancelado")
Else
   Check2.Value = 0
End If
If IsNull(Data1.Recordset("fecha_fact")) = False Then
   mfrea.Text = Format(Data1.Recordset("fecha_fact"), "dd/mm/yyyy")
Else
   mfrea.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("mot_cierre")) = False Then
   Combo1.Text = Data1.Recordset("mot_cierre")
Else
   Combo1.Text = ""
End If
If IsNull(Data1.Recordset("obs_cierre")) = False Then
   t_obscierre.Text = Data1.Recordset("obs_cierre")
Else
   t_obscierre.Text = ""
End If
If IsNull(Data1.Recordset("coord")) = False Then
   chcoord.Value = Data1.Recordset("coord")
Else
   chcoord.Value = 0
End If
If IsNull(Data1.Recordset("fec_coord")) = False Then
   mfcoord.Text = Format(Data1.Recordset("fec_coord"), "dd/mm/yyyy")
Else
   mfcoord.Text = "__/__/____"
End If


End Sub

Private Sub Form_Load()
labfec.Caption = Format(Date, "dd/mm/yyyy")
data_buscafact.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from sol_hisopos order by fecha DESC"
Data1.Refresh
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_graba.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_graba.RecordSource = "select * from sol_hisopos"
data_graba.Refresh
data_arch.Connect = "odbc;dsn=sapparch;"

Data2.RecordSource = "select * from usuarios where usuario ='" & Trim(WElusuario) & "'"
Data2.Refresh
If Data2.Recordset.RecordCount > 0 Then
   labusua.Caption = Data2.Recordset("nombre")
Else
   labusua.Caption = "Sin Registrar"
End If

Carga_zonas
Carga_motivos

If ControlUsuario(Command1.Name) = 1 Then
   mfrea.Enabled = True
   Check1.Enabled = True
   Check2.Enabled = True
   t_obscierre.Enabled = True
   Combo1.Enabled = True
   chcoord.Enabled = True
   mfcoord.Enabled = True
   chderiva.Enabled = True
Else
   mfrea.Enabled = False
   Check1.Enabled = False
   Check2.Enabled = False
   t_obscierre.Enabled = True
   Combo1.Enabled = False
   chcoord.Enabled = False
   mfcoord.Enabled = False
   chderiva.Enabled = False
End If
Label21.Caption = ""

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub


Public Sub Borra_campos()
labid.Caption = ""
t_ced.Text = ""
t_codced.Text = ""
t_mat.Text = ""
mfpos.Text = "__/__/____"
t_nom.Text = ""
cbomut.Text = ""
t_dir1.Text = ""
t_dir2.Text = ""
cbolocal.Text = ""
t_cel.Text = ""
t_tel.Text = ""
cbozon.Text = ""
mfcon.Text = "__/__/____"
Option1.Value = False
Option2.Value = False
Check1.Value = 0
mfrea.Text = "__/__/____"
t_obscierre.Text = ""
Check2.Value = 0
Combo1.Text = ""
labcodconv.Caption = ""
chcoord.Value = 0
mfcoord.Text = "__/__/____"

End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub Label20_DblClick()
If frm_especcovid.Visible = True Then
   MsgBox "Ya está abierta la agenda."
Else
   frm_especcovid.Show
End If

End Sub

Private Sub Label21_DblClick()
Dim x, Xbandlab As Integer
Dim Xlac As String

Xlac = ""
Xbandlab = 0
On Error GoTo Noestaarch

frm_solhisopa.MousePointer = 11

If Dir(App.path & "\laboratorio\temporal.pdf") <> "" Then
   Kill App.path & "\laboratorio\temporal.pdf"
End If

'data_abre.Recordset.Edit
'data_abre.Recordset("numero") = data_buscasrv.Recordset("id")
'data_abre.Recordset.Update
'data_abre.Refresh
If IsNull(Data1.Recordset("id_result")) = False Then
    data_arch.RecordSource = "Select * from archs where id =" & Data1.Recordset("id_result")
    data_arch.Refresh
    If data_arch.Recordset.RecordCount > 0 Then
       Command5_Click
'       Shell App.path & "\archenf.exe", vbMinimizedFocus
       ShellExecute Me.hwnd, "open", App.path & "\laboratorio\temporal.pdf", "", "", 4
    Else
       frm_solhisopa.MousePointer = 0
       MsgBox "No tiene archivo escaneado"
    End If
    frm_solhisopa.MousePointer = 0
Else
   MsgBox "No figura archivo."
End If
frm_solhisopa.MousePointer = 0


Exit Sub

Noestaarch:
           If Err.Number = 53 Then
              MsgBox "No se ecuentra el archivo, verifique"
           Else
              MsgBox "Error al cargar el archivo"
           End If
           

End Sub

Private Sub mfcon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub

Private Sub mfpos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_busca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_busca.Text <> "" Then
      Data1.RecordSource = "select * from sol_hisopos where cedula =" & t_busca.Text
      Data1.Refresh
   Else
      Data1.RecordSource = "select * from sol_hisopos order by fecha DESC"
      Data1.Refresh
   End If
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
   t_codced.SetFocus
End If

End Sub

Private Sub t_ced_LostFocus()
   If t_ced.Text <> "" Then
      Data2.RecordSource = "select * from clientes where cl_cedula =" & t_ced.Text
      Data2.Refresh
      If Data2.Recordset.RecordCount > 0 Then
         If IsNull(Data2.Recordset("estado")) = False Then
            If Data2.Recordset("estado") = 2 Then
               If IsNull(Data2.Recordset("fecha_baja")) = False Then
                  MsgBox "VERIFIQUE DATOS! Socio de baja en padrón.", vbCritical
               End If
            End If
         End If
         t_codced.Text = Data2.Recordset("cl_codced")
         t_nom.Text = Data2.Recordset("cl_apellid")
         t_mat.Text = Data2.Recordset("cl_codigo")
         If IsNull(Data2.Recordset("cl_direcci")) = False Then
            t_dir1.Text = Data2.Recordset("cl_direcci")
         Else
            t_dir1.Text = "Ingresar dirección"
         End If
         If IsNull(Data2.Recordset("cl_entre")) = False Then
            t_dir2.Text = Data2.Recordset("cl_entre")
         Else
            t_dir2.Text = ""
         End If
         labcodconv.Caption = Data2.Recordset("cl_codconv")
         If IsNull(Data2.Recordset("cl_dpto")) = False Then
            t_cel.Text = Data2.Recordset("cl_dpto")
         Else
            t_cel.Text = ""
         End If
         If IsNull(Data2.Recordset("cl_telefon")) = False Then
            t_tel.Text = Data2.Recordset("cl_telefon")
         Else
            t_tel.Text = ""
         End If
         mfpos.Text = Format(Date + 1, "dd/mm/yyyy")
         If IsNull(Data2.Recordset("cl_zona")) = False Then
            cbolocal.Text = Data2.Recordset("cl_zona")
            If cbolocal.Text <> "*TODOS" Then
               Consulta_zonas2
               If cbolocal.Text = "CENTRO DE PANDO ZONA URBA" Then
                  cbozon.Text = "NORTE"
               End If
            Else
               cbozon.Text = ""
            End If
         Else
            cbolocal.Text = ""
            cbozon.Text = ""
         End If
         
         Data2.RecordSource = "select * from convenio where cnv_codigo ='" & Trim(labcodconv.Caption) & "'"
         Data2.Refresh
         If IsNull(Data2.Recordset("cnv_grupo")) = False Then
            If Trim(Data2.Recordset("cnv_grupo")) <> "" Then
               cbomut.Text = Data2.Recordset("cnv_grupo")
            Else
               cbomut.Text = ""
            End If
         Else
            cbomut.Text = ""
         End If
      '   Consulta_facturacion
      Else
'         MsgBox "Cédula no encontrada. Se necesita contar con alta de ficha. Consulte con recepción.", vbCritical
         MsgBox "Cédula no encontrada. Ingrese datos.", vbCritical
         MsgBox "ATENCIÓN!!! VERIFIQUE PREVIAMENTE MUTUALISTA DEL SOCIO Y SI ESTA ACTIVO!", vbCritical
         Borra_camposSinCed
      End If
   End If

End Sub

Private Sub t_cel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_tel.SetFocus
End If

End Sub

Private Sub t_codced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mat.SetFocus
End If

End Sub

Private Sub t_dir1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_dir2.SetFocus
End If

End Sub

Private Sub t_dir2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbolocal.SetFocus
End If

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfpos.SetFocus
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomut.SetFocus
End If

End Sub

Private Sub t_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfcon.SetFocus
End If

End Sub

Private Sub Timer1_Timer()
labfec.Caption = Format(Date, "dd/mm/yyyy")

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
         cbolocal.AddItem Xrecclii("zo_nombre")
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
             
Xsqlpromo = "Select * from zonas where zo_nombre ='" & Trim(cbolocal.Text) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
Else
   MsgBox "Zona no encontrada"
   cbolocal.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close
End Sub
Public Sub Consulta_zonas2()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Pregunta As String

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from zonas where zo_nombre ='" & Trim(cbolocal.Text) & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If Xrecclii("zo_grupo") >= 100 And Xrecclii("zo_grupo") <= 502 Then
      cbozon.Text = "SUR"
   Else
      If Xrecclii("zo_grupo") >= 600 And Xrecclii("zo_grupo") <= 699 Then
         cbozon.Text = "NORTE"
      Else
         If Xrecclii("zo_grupo") >= 700 And Xrecclii("zo_grupo") <= 799 Then
            cbozon.Text = "SUR"
         Else
            If Xrecclii("zo_grupo") = 999 Then
               cbozon.Text = ""
            Else
               cbozon.Text = "NORTE"
            End If
         End If
      End If
   End If
End If

Xrecclii.Close
ConbdSapp.Close
End Sub
Public Sub Consulta_siexiste()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim lafecha As Date
lafecha = Date - 1

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from sol_hisopos where cedula =" & t_ced.Text & " and fecha >='" & Format(lafecha, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labyaesta.Caption = "1"
Else
   labyaesta = ""
End If

Xrecclii.Close
ConbdSapp.Close
End Sub


Public Sub Carga_motivos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from solhnf_mot"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      Combo1.AddItem Xrecclii("descrip")
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close
End Sub
Public Sub Consulta_motivos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from solhnf_mot where descrip ='" & Combo1.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
Else
   MsgBox "No existe el motivo, verifique"
   Combo1.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close
End Sub

Public Sub Borra_camposSinCed()
labid.Caption = ""
t_mat.Text = ""
mfpos.Text = "__/__/____"
t_nom.Text = ""
cbomut.Text = ""
t_dir1.Text = ""
t_dir2.Text = ""
cbolocal.Text = ""
t_cel.Text = ""
t_tel.Text = ""
cbozon.Text = ""
mfcon.Text = "__/__/____"
Option1.Value = False
Option2.Value = False
Check1.Value = 0
mfrea.Text = "__/__/____"
t_obscierre.Text = ""
Check2.Value = 0
Combo1.Text = ""
labcodconv.Caption = ""
chcoord.Value = 0
mfcoord.Text = "__/__/____"

End Sub

Public Sub Consulta_facturacion()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from linmmdd where ced_socio ='" & t_ced.Text & "' and fecha ='" & Format(Date, "yyyy-mm-dd") & "' and nro_flia in (1,10,14)"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
Else
   MsgBox "No existe registro de asistencia facturada. No se puede realizar solicitud.", vbCritical
   Borra_campos
End If

Xrecclii.Close
ConbdSapp.Close
End Sub

