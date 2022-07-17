VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_calidadiso 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de calidad"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "frm_calidadiso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr2 
      Left            =   2880
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_llamod 
      Caption         =   "data_llamod"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "T.Largador"
      Height          =   495
      Left            =   5760
      TabIndex        =   28
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6840
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_res 
      Caption         =   "data_res"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_info 
      Caption         =   "data_info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "Salir"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Procesar..."
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Selección de datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Data data_consmed 
         Caption         =   "data_consmed"
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
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "CMT x Medic"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Llamados por Oper"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Data data_infu 
         Caption         =   "data_infu"
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
         Top             =   4320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command7 
         Caption         =   "CMT por telef"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc data_llam 
         Height          =   375
         Left            =   3840
         Top             =   5400
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         DataSourceName  =   "sappnew"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "data_llam"
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
      Begin VB.Data data_chof 
         Caption         =   "data_chof"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frm_calidadiso.frx":0442
         Left            =   2880
         List            =   "frm_calidadiso.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H0080FFFF&
         Caption         =   "Omitir 911 y 911B"
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
         Left            =   3720
         TabIndex        =   29
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H0080FFFF&
         Caption         =   "Tiempos Largador"
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
         Left            =   360
         TabIndex        =   27
         Top             =   3120
         Width           =   2655
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sin SAMC y Sin 711"
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
         Left            =   360
         TabIndex        =   26
         Top             =   6120
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Data data_tras 
         Caption         =   "data_tras"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox t_codmed 
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
         Left            =   2880
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox t_mov 
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
         Left            =   5280
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H0080FF80&
         Caption         =   "Informe desde respaldos"
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
         Left            =   360
         TabIndex        =   21
         Top             =   5760
         Width           =   3255
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Omitir promedios"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Generar planilla de resumen"
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
         Left            =   360
         TabIndex        =   18
         Top             =   5400
         Width           =   3255
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenar por largador"
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
         TabIndex        =   17
         Top             =   4440
         Width           =   2775
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenar por receptor"
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
         Left            =   360
         TabIndex        =   16
         Top             =   4440
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Tiempos Recepción/ Clasificación/Asignación"
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
         Left            =   360
         TabIndex        =   14
         Top             =   2400
         Width           =   2655
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
         ItemData        =   "frm_calidadiso.frx":0446
         Left            =   2760
         List            =   "frm_calidadiso.frx":0450
         TabIndex        =   13
         Top             =   4920
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "INFORME SOLO ZONA TALA O FUERA DE ZONA"
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
         Left            =   360
         TabIndex        =   10
         Top             =   4080
         Width           =   5895
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   3600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Resumen"
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
         Left            =   360
         TabIndex        =   6
         Top             =   3600
         Value           =   -1  'True
         Width           =   2535
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
         ItemData        =   "frm_calidadiso.frx":0466
         Left            =   2880
         List            =   "frm_calidadiso.frx":047C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   360
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
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   360
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
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "INDICADORES DE CHOFER:"
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
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "CODIGO DE MEDICO:"
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
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "MOVIL:"
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
         Height          =   375
         Left            =   4320
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   4
         X1              =   0
         X2              =   6480
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Llamados reales de:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Módulo a seleccionar:"
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
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Rango de fechas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   120
      Picture         =   "frm_calidadiso.frx":04DF
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "frm_calidadiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
   Check2.Value = 0
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Check1.Value = 0
End If

End Sub



Private Sub Check5_Click()
If Check6.Value = 1 Then
Else
   Check5.Value = 0
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xhh1, Xmm1, Xhh2, Xmm2, Xtmm, Xths As Long
Dim XtotminR, XtotminA, XtotminV, XtotminC As Long
Dim XtotlleR, XtotlleA, XtotlleV, XtotlleC, Xtotcan, Xtotcanmax, Xtotcanmax2, XtotrealR, XtotrealA, XtotrealV, XtotrealC As Long
Dim Xtorojos, Xtoama, Xtoverde, Xtocele, Xtocerti, Xtogra, Xtomeno2 As Long
Dim Xporreav, Xporrear, Xporreaa, Xporreac As Double
Dim Xtotrsalin, Xtotasalin, Xtotvsalin, Xrojoama, Xrojover As Long
Dim Xhh3, Xmm3, Xtot3, Xtot3tot As Long
Dim Xhh4a, Xmm4a, Xtot4a, Xtot4atot As Long
Dim Xhh5a, Xmm5a, Xtot5a, Xtot5atot, Xtot6atot As Long
Dim Xqueus As String
Dim Xveonum1 As Currency
Dim Xtexto As String
Xtomeno2 = 0
Xrojoama = 0
Xrojover = 0

Xtexto = "50"
frm_calidadiso.MousePointer = 11
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\inftras.mdb")
MiBaseact.Execute "Delete * from inflla"
data_tras.DatabaseName = App.path & "\inftras.mdb"
data_tras.RecordSource = "inflla"
data_tras.Refresh

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infarqc"
data_res.RecordSource = "infarqc"
data_res.Refresh

XtotrealV = 0
XtotrealR = 0
XtotrealA = 0
XtotrealC = 0
Xtot3tot = 0
Xqueus = ""
If Check7.Value = 1 Then
   data_llam.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
Else
'   data_llam.DatabaseName = App.Path & "\sapp.mdb"
   data_llam.ConnectionString = "dsn=" & Xconexrmt
End If

If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      If Combo1.Text = "DESPACHO" Then
         MiBaseact.Execute "Delete * from inflla"
         data_info.RecordSource = "inflla"
         data_info.Refresh
         
         If Check2.Value = 1 Then
            If Check3.Value = 1 Or Check4.Value = 1 Then
               Xqueus = InputBox("Ingrese NOMBRE DE USUARIO A SELECCIONAR (Ej.JFERNAN):", "Datos para seleccionar usuario")
            Else
               Xqueus = ""
            End If
         End If
         If Combo2.Text = "ROJOS" Then
            If Xqueus = "" Then
               If t_mov.Text = "" Then
                  If Check10.Value = 1 Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "R" & "' and codmed <>" & 959 & " and categ not in ('911','911B''MSP') and cancela is null and movilpas not in (2015) order by nrolla"
                     data_llam.Refresh
                  Else
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "R" & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
                     data_llam.Refresh
                  End If
               Else
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "R" & "' and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) order by nrolla"
                  data_llam.Refresh
               End If
            Else
               If Check3.Value = 1 Then
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "R" & "' and usuario ='" & Trim(UCase(Xqueus)) & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
                  data_llam.Refresh
               Else
                  If Check4.Value = 1 Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "R" & "' and timdes ='" & Trim(UCase(Xqueus)) & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
                     data_llam.Refresh
                  End If
               End If
            End If
            If t_codmed.Text <> "" Then
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "R" & "' and codmed =" & t_codmed.Text & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
               data_llam.Refresh
            End If
         Else
            If Combo2.Text = "AMARILLOS" Then
               If Xqueus = "" Then
                  If t_mov.Text = "" Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "A" & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
                     data_llam.Refresh
                  Else
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "A" & "' and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) order by nrolla"
                     data_llam.Refresh
                  End If
               Else
                  If Check3.Value = 1 Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "A" & "' and usuario ='" & Trim(UCase(Xqueus)) & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
                     data_llam.Refresh
                  Else
                     If Check4.Value = 1 Then
                        data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "A" & "' and timdes ='" & Trim(UCase(Xqueus)) & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
                        data_llam.Refresh
                     End If
                  End If
               End If
               If t_codmed.Text <> "" Then
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot ='" & "A" & "' and codmed =" & t_codmed.Text & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015) order by nrolla"
                  data_llam.Refresh
               End If
            Else
               If Xqueus = "" Then
                  If t_mov.Text = "" Then
                     If Check10.Value = 1 Then
                        data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmed <>" & 959 & " and categ not in ('911','911B','MSP') and cancela is null and codzon not in (4,6) and movilpas not in (2015,0) order by nrolla"
                        data_llam.Refresh
                     Else
                        If Check8.Value = 1 Then
                           data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmed <>" & 959 & " and cancela is null and categ in ('CAAMEP','CAAM') and codzon not in (4,6) and movilpas not in (2015,0) order by nrolla"
                           data_llam.Refresh
                        Else
                        
                           data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed <>" & 959 & " and cancela is null and categ not in ('MSP','50') and codzon not in (4,6) and movilpas not in (2015,0) order by nrolla"
                           data_llam.Refresh
                        End If
                     End If
                  Else
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and categ not in ('MSP') and codzon not in (4,6) and cancela is null and movilpas not in (2015,0) order by nrolla"
                     data_llam.Refresh
                  End If
               Else
                  If Check3.Value = 1 Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy/mm/dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and usuario ='" & Trim(UCase(Xqueus)) & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015,0) order by nrolla"
                     data_llam.Refresh
                  Else
                     If Check4.Value = 1 Then
                        data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and timdes ='" & Trim(UCase(Xqueus)) & "' and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015,0) order by nrolla"
                        data_llam.Refresh
                     End If
                  End If
               End If
               If t_codmed.Text <> "" Then
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmed =" & t_codmed.Text & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) and movilpas not in (2015,0) order by nrolla"
                  data_llam.Refresh
               End If
            End If
         End If
      End If
      If Combo1.Text = "CMT por Operador" Or Combo1.Text = "CMT por médico" Then
'         data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed <>" & 959 & " and cancela is null and categ not in ('MSP','50') and codzon not in (4,6) and movilpas not in (2015,0) order by nrolla"
'         data_llam.Refresh
         If Combo1.Text = "CMT por Operador" Then
            data_llam.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.hora,llamado.usuario," & _
            "llamado.unied,llamado.edad,llamado.matric,llamado.usuario,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codmed,llamado.codzon,llamado.obsmot," & _
            "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
            "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.enfer,llamado.ci," & _
            "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.hzona,resplla.timdes,resplla.mes from llamado " & _
            "inner join resplla on llamado.nrolla=resplla.nro where resplla.hzona is not null and " & _
            "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.cancela is null and llamado.categ not in ('MSP','50','55') and llamado.enfer not in (1) and llamado.codmed not in (959) order by llamado.usuario,llamado.fecha"
         Else
            If t_codmed.Text = "" Then
                data_llam.RecordSource = "select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and llamado.cancela is null and movilpas in (2015) and codmed not in (959) order by codmedcmt,fecha"
            Else
                data_llam.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.hora,llamado.usuario," & _
                "llamado.codmedcmt,llamado.unied,llamado.edad,llamado.matric,llamado.usuario,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codmed,llamado.codzon,llamado.obsmot," & _
                "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
                "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.enfer,llamado.ci," & _
                "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.hzona,resplla.timdes,resplla.mes from llamado " & _
                "inner join resplla on llamado.nrolla=resplla.nro where resplla.hzona is not null and " & _
                "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.cancela is null and llamado.categ not in ('MSP','50','55') and llamado.enfer not in (1) and llamado.codmed not in (959) and llamado.codmedcmt =" & t_codmed.Text & " order by llamado.codmedcmt,llamado.fecha"
            End If
         End If
         data_llam.Refresh
      End If
      If Combo1.Text = "Llamados por Operador" Then
         data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed <>" & 959 & " and cancela is null order by nrolla"
         data_llam.Refresh
      End If
      If Check1.Value = 1 Or Combo1.Text = "CMT por Operador" Or Combo1.Text = "Llamados por Operador" Or Combo1.Text = "CMT por médico" Then
         If Combo1.Text = "CMT por Operador" Then
            Command7_Click
         Else
            If Combo1.Text = "Llamados por Operador" Or Combo1.Text = "CMT por médico" Then
               If Combo1.Text = "Llamados por Operador" Then
                  Command8_Click
               Else
                  Command9_Click
               End If
            Else
               Command3_Click
            End If
         End If
      Else
          If Check2.Value = 1 Then
             Command4_Click
          Else
              If data_llam.Recordset.RecordCount > 0 Then
                 data_llam.Recordset.MoveFirst
                 Do While Not data_llam.Recordset.EOF
                    If data_llam.Recordset("codzon") = 4 Then
                    Else
                        If data_llam.Recordset("codmot") = "R" Then
                           If IsNull(data_llam.Recordset("colormot")) = False Then
                              If data_llam.Recordset("colormot") = "R" Then
                                 XtotrealR = XtotrealR + 1
                              End If
                           End If
                           If IsNull(data_llam.Recordset("horsali")) = False Then
                              Xhh1 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                              Xmm1 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                           End If
                           If IsNull(data_llam.Recordset("hora")) = False Then
                              Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                              Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                           End If
                           If IsNull(data_llam.Recordset("horpas")) = False Then
                              Xhh3 = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                              Xmm3 = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                           End If
                           If Xmm3 = Xmm1 Then
                              If Xhh3 = Xhh1 Then
                                 Xtot3 = 0
                              Else
                                 Xths = Xhh1 - Xhh3
                                 Xtot3 = Xmm1 - Xmm3 + 60
                                 If Xths = 2 Then
                                    Xtot3 = Xtot3 + 60
                                 End If
                                 If Xths = 3 Then
                                    Xtot3 = Xtot3 + 120
                                 End If
                              End If
                           Else
                              If Xhh3 = Xhh1 Then
                                 Xtot3 = Xmm1 - Xmm3
                              Else
                                 Xths = Xhh1 - Xhh3
                                 Xtot3 = Xmm1 - Xmm3 + 60
                                 If Xths = 2 Then
                                    Xtot3 = Xtot3 + 60
                                 End If
                                 If Xths = 3 Then
                                    Xtot3 = Xtot3 + 120
                                 End If
                              End If
                           End If
                           If Xtot3 > 3 Then
                              Xtot3tot = Xtot3tot + 1
                           End If
    ''''-------------------------------
                           If Xhh1 = Xhh2 Then
                              Xtmm = Xmm1 - Xmm2
                           Else
                              Xths = Xhh1 - Xhh2
                              If IsNull(data_llam.Recordset("fecsali")) = False Then
                                 If data_llam.Recordset("fecsali") > data_llam.Recordset("fecha") Then
                                    Xths = Xths + 24
                                 End If
                              End If
                              Xtmm = Xmm1 - Xmm2 + 60
                              If Xths = 2 Then
                                 Xtmm = Xtmm + 60
                              End If
                              If Xths = 3 Then
                                 Xtmm = Xtmm + 120
                              End If
                              If Xths = 4 Then
                                 Xtmm = Xtmm + 180
                              End If
                           End If
                           If Xtmm >= 0 Then
                              XtotminR = Xtot3
                           Else
                              XtotminR = 0
                           End If
            '' Llegada!!!
                           If IsNull(data_llam.Recordset("hor_llega")) = False Then
                              Xhh1 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                              Xmm1 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                           End If
                           If Combo3.ListIndex >= 0 Then
                              If IsNull(data_llam.Recordset("horsali")) = False Then
                                 Xhh2 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                                 Xmm2 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                              End If
                           Else
                              If IsNull(data_llam.Recordset("hora")) = False Then
                                 Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                 Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                              End If
                           End If
                           If Xhh1 = Xhh2 Then
                              Xtmm = Xmm1 - Xmm2
                           Else
                              Xths = Xhh1 - Xhh2
                              If IsNull(data_llam.Recordset("fec_llega")) = False Then
                                 If data_llam.Recordset("fec_llega") > data_llam.Recordset("fecha") Then
                                    Xths = Xths + 24
                                 End If
                              End If
                              Xtmm = Xmm1 - Xmm2 + 60
                              If Xths = 2 Then
                                 Xtmm = Xtmm + 60
                              End If
                              If Xths = 3 Then
                                 Xtmm = Xtmm + 120
                              End If
                              If Xths = 4 Then
                                 Xtmm = Xtmm + 180
                              End If
                           End If
                           If Xtmm >= 0 Then
                              XtotlleR = Xtmm
                           Else
                              XtotlleR = 0
                           End If
                           data_info.Recordset.AddNew
                           data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                           data_info.Recordset("nro") = data_llam.Recordset("nrolla")
                           data_info.Recordset("hora") = data_llam.Recordset("hora")
                           data_info.Recordset("matric") = data_llam.Recordset("matric")
                           data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                           data_info.Recordset("categ") = data_llam.Recordset("categ")
                           data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                           data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                           data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                           data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                           data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                           data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                           data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                           data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                           data_info.Recordset("hor_rea") = data_llam.Recordset("hor_rea")
                           data_info.Recordset("trasla") = data_llam.Recordset("trasla")
                           data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                           data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                           data_info.Recordset("activo") = data_llam.Recordset("activo")
                           data_info.Recordset("codmed") = data_llam.Recordset("codmed")
                           data_info.Recordset("nommed") = data_llam.Recordset("nommed")
                            If IsNull(data_llam.Recordset("cancela")) = False Then
                               If data_llam.Recordset("cancela") = 1 Then
                                  data_info.Recordset("cancela") = 1
                               Else
                                  data_info.Recordset("cancela") = 9
                               End If
                            Else
                               data_info.Recordset("cancela") = 9
                            End If
                           
                           If Check3.Value = 1 Then
                              data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                           Else
                              If Check4.Value = 1 Then
                                 data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                              End If
                           End If
                           data_info.Recordset("tmm") = XtotminR
                           data_info.Recordset("thh") = XtotlleR
                           data_info.Recordset.Update
                           Xtorojos = Xtorojos + 1
                        End If
                        If data_llam.Recordset("codmot") = "A" Then
                           If IsNull(data_llam.Recordset("colormot")) = False Then
                              If data_llam.Recordset("colormot") = "R" Then
                                 Xrojoama = Xrojoama + 1
                              End If
                           End If
                           If IsNull(data_llam.Recordset("colormot")) = False Then
                              If data_llam.Recordset("colormot") = "A" Then
                                 XtotrealA = XtotrealA + 1
                              End If
                           End If
                           If IsNull(data_llam.Recordset("horsali")) = False Then
                              Xhh1 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                              Xmm1 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                           End If
                           If IsNull(data_llam.Recordset("hora")) = False Then
                              Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                              Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                           End If
                           If IsNull(data_llam.Recordset("horpas")) = False Then
                              Xhh4a = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                              Xmm4a = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                           End If
                           If Xmm4a = Xmm1 Then
                              If Xhh4a = Xhh1 Then
                                 Xtot4a = 0
                              Else
                                Xths = Xhh1 - Xhh4a
                                Xtot4a = Xmm1 - Xmm4a + 60
                                If Xths = 2 Then
                                   Xtot4a = Xtot4a + 60
                                End If
                                If Xths = 3 Then
                                   Xtot4a = Xtot4a + 120
                                End If
                              End If
                           Else
                              If Xhh4a = Xhh1 Then
                                 Xtot4a = Xmm1 - Xmm4a
                              Else
                                 Xths = Xhh1 - Xhh4a
                                 Xtot4a = Xmm1 - Xmm4a + 60
                                 If Xths = 2 Then
                                    Xtot4a = Xtot4a + 60
                                 End If
                                 If Xths = 3 Then
                                    Xtot4a = Xtot4a + 120
                                 End If
                              End If
                           End If
                           If Xtot4a > 5 Then
                              Xtot4atot = Xtot4atot + 1
                           End If
                           
                           If Xhh1 = Xhh2 Then
                              Xtmm = Xmm1 - Xmm2
                           Else
                              Xths = Xhh1 - Xhh2
                              If IsNull(data_llam.Recordset("fecsali")) = False Then
                                 If data_llam.Recordset("fecsali") > data_llam.Recordset("fecha") Then
                                    Xths = Xths + 24
                                 End If
                              End If
                              Xtmm = Xmm1 - Xmm2 + 60
                              If Xths = 2 Then
                                 Xtmm = Xtmm + 60
                              End If
                              If Xths = 3 Then
                                 Xtmm = Xtmm + 120
                              End If
                              If Xths = 4 Then
                                 Xtmm = Xtmm + 180
                              End If
                           End If
                           If Xtmm >= 0 Then
                              XtotminA = Xtot4a
                           Else
                              XtotminA = 0
                           End If
            
            '' Llegada!!!
                           If IsNull(data_llam.Recordset("hor_llega")) = False Then
                              Xhh1 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                              Xmm1 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                           End If
                           If Combo3.ListIndex >= 0 Then
                              If IsNull(data_llam.Recordset("horsali")) = False Then
                                 Xhh2 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                                 Xmm2 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                              End If
                           Else
                              If IsNull(data_llam.Recordset("hora")) = False Then
                                 Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                 Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                              End If
                           End If
                           If Xhh1 = Xhh2 Then
                              Xtmm = Xmm1 - Xmm2
                           Else
                              Xths = Xhh1 - Xhh2
                              If IsNull(data_llam.Recordset("fec_llega")) = False Then
                                 If data_llam.Recordset("fec_llega") > data_llam.Recordset("fecha") Then
                                    Xths = Xths + 24
                                 End If
                              End If
                              Xtmm = Xmm1 - Xmm2 + 60
                              If Xths = 2 Then
                                 Xtmm = Xtmm + 60
                              End If
                              If Xths = 3 Then
                                 Xtmm = Xtmm + 120
                              End If
                              If Xths = 4 Then
                                 Xtmm = Xtmm + 180
                              End If
                           End If
                           If Xtmm >= 0 Then
                              XtotlleA = Xtmm
                           Else
                              XtotlleA = 0
                           End If
                           data_info.Recordset.AddNew
                           data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                           data_info.Recordset("nro") = data_llam.Recordset("nrolla")
                           data_info.Recordset("hora") = data_llam.Recordset("hora")
                           data_info.Recordset("matric") = data_llam.Recordset("matric")
                           data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                           data_info.Recordset("categ") = data_llam.Recordset("categ")
                           data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                           data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                           data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                           data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                           data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                           data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                           data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                           data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                           data_info.Recordset("hor_rea") = data_llam.Recordset("hor_rea")
                           data_info.Recordset("trasla") = data_llam.Recordset("trasla")
                           data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                           data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                           data_info.Recordset("activo") = data_llam.Recordset("activo")
                           data_info.Recordset("codmed") = data_llam.Recordset("codmed")
                           data_info.Recordset("nommed") = data_llam.Recordset("nommed")
                            If IsNull(data_llam.Recordset("cancela")) = False Then
                               If data_llam.Recordset("cancela") = 1 Then
                                  data_info.Recordset("cancela") = 1
                               Else
                                  data_info.Recordset("cancela") = 9
                               End If
                            Else
                               data_info.Recordset("cancela") = 9
                            End If
                           
                           If Check3.Value = 1 Then
                              data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                           Else
                              If Check4.Value = 1 Then
                                 data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                              End If
                           End If
                           data_info.Recordset("tmm") = XtotminA
                           data_info.Recordset("thh") = XtotlleA
                           data_info.Recordset.Update
                           Xtoama = Xtoama + 1
                        End If
                        If data_llam.Recordset("codmot") = "V" Or data_llam.Recordset("codmot") = "C" Then
                           If data_llam.Recordset("categ") = "UDEMM" Or _
                              data_llam.Recordset("categ") = "CERSEM" Or _
                              data_llam.Recordset("categ") = "CERCAS" Or _
                              data_llam.Recordset("categ") = "CERANT" Or _
                              data_llam.Recordset("categ") = "CERADU" Or _
                              data_llam.Recordset("categ") = "CERDGI" Or _
                              data_llam.Recordset("categ") = "CERESS" Or _
                              data_llam.Recordset("categ") = "CERSAP" Or _
                              data_llam.Recordset("categ") = "CERHEV" Or _
                              data_llam.Recordset("categ") = "CERIMP" Or _
                              data_llam.Recordset("categ") = "CERKEV" Or _
                              data_llam.Recordset("categ") = "CERMAT" Or _
                              data_llam.Recordset("categ") = "CERSEV" Or _
                              data_llam.Recordset("categ") = "CERVIS" Then
                              Xtocerti = Xtocerti + 1
                           Else
                              If IsNull(data_llam.Recordset("colormot")) = False Then
                                 If data_llam.Recordset("colormot") = "V" Then
                                    XtotrealV = XtotrealV + 1
                                 Else
                                    XtotrealC = XtotrealC + 1
                                 End If
                                 If data_llam.Recordset("colormot") = "R" Then
                                    Xrojover = Xrojover + 1
                                 End If
                              End If
                                 If IsNull(data_llam.Recordset("horsali")) = False Then
                                    Xhh1 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                                    Xmm1 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                                 End If
                                 If IsNull(data_llam.Recordset("hora")) = False Then
                                    Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                    Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                                 End If
                                 If IsNull(data_llam.Recordset("horpas")) = False Then
                                    Xhh5a = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                                    Xmm5a = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                                 End If
                                 If Xmm5a = Xmm1 Then
                                    If Xhh5a = Xhh1 Then
                                       Xtot5a = 0
                                    Else
                                       Xtot5a = Xhh1 - Xhh5a
                                    End If
                                 Else
                                    If Xhh5a = Xhh1 Then
                                       Xtot5a = Xmm1 - Xmm5a
                                    Else
                                       Xtot5a = Xmm1 - Xmm5a + 60
                                    End If
                                 End If
                                 If data_llam.Recordset("codmot") = "V" Then
                                    If Xtot5a > 5 Then
                                       Xtot5atot = Xtot5atot + 1
                                    End If
                                 Else
                                    If Xtot5a > 5 Then
                                       Xtot6atot = Xtot6atot + 1
                                    End If
                                 End If
                                 If Xhh1 = Xhh5a Then
                                    Xtmm = Xmm1 - Xmm5a
                                 Else
                                    Xths = Xhh1 - Xhh5a
                                    If IsNull(data_llam.Recordset("fecsali")) = False Then
                                       If data_llam.Recordset("fecsali") > data_llam.Recordset("fecha") Then
                                          Xths = Xths + 24
                                       End If
                                    End If
                                    Xtmm = Xmm1 - Xmm5a + 60
                                    If Xths = 2 Then
                                       Xtmm = Xtmm + 60
                                    End If
                                    If Xths = 3 Then
                                       Xtmm = Xtmm + 120
                                    End If
                                    If Xths = 4 Then
                                       Xtmm = Xtmm + 180
                                    End If
                                    If Xths = 5 Then
                                       Xtmm = Xtmm + 240
                                    End If
                                    If Xths = 6 Then
                                       Xtmm = Xtmm + 300
                                    End If
                                    If Xths = 7 Then
                                       Xtmm = Xtmm + 420
                                    End If
                                    If Xths = 8 Then
                                       Xtmm = Xtmm + 480
                                    End If
                                    If Xths = 9 Then
                                       Xtmm = Xtmm + 540
                                    End If
                                    If Xths = 10 Then
                                       Xtmm = Xtmm + 600
                                    End If
                                    If Xths > 10 Then
                                       Xtmm = Xtmm + 700
                                    End If
                                 
                                 End If
                                 If Xtmm >= 0 Then
                                    XtotminV = Xtmm
                                    XtotminC = Xtmm
                                 Else
                                    XtotminV = 0
                                    XtotminC = 0
                                 End If
                  '' Llegada!!!
                                 If IsNull(data_llam.Recordset("hor_llega")) = False Then
                                    Xhh1 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                                    Xmm1 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                                 End If
                                 If Combo3.ListIndex >= 0 Then
                                    If IsNull(data_llam.Recordset("horsali")) = False Then
                                       Xhh2 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                                       Xmm2 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                                    End If
                                 Else
                                    If IsNull(data_llam.Recordset("hora")) = False Then
                                       Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                       Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                                    End If
                                 End If
                                 If Xhh1 = Xhh2 Then
                                    Xtmm = Xmm1 - Xmm2
                                 Else
                                    Xths = Xhh1 - Xhh2
                                    If IsNull(data_llam.Recordset("fec_llega")) = False Then
                                       If data_llam.Recordset("fec_llega") > data_llam.Recordset("fecha") Then
                                          If Xhh2 <= Xhh1 Then
                                          Else
                                             Xths = Xths + 24
                                          End If
                                       End If
                                    End If
                                    Xtmm = Xmm1 - Xmm2 + 60
                                    If Xths = 2 Then
                                       Xtmm = Xtmm + 60
                                    End If
                                    If Xths = 3 Then
                                       Xtmm = Xtmm + 120
                                    End If
                                    If Xths = 4 Then
                                       Xtmm = Xtmm + 180
                                    End If
                                    If Xths = 5 Then
                                       Xtmm = Xtmm + 240
                                    End If
                                    If Xths = 6 Then
                                       Xtmm = Xtmm + 300
                                    End If
                                    If Xths = 7 Then
                                       Xtmm = Xtmm + 420
                                    End If
                                    If Xths = 8 Then
                                       Xtmm = Xtmm + 480
                                    End If
                                    If Xths = 9 Then
                                       Xtmm = Xtmm + 540
                                    End If
                                    If Xths = 10 Then
                                       Xtmm = Xtmm + 600
                                    End If
                                    If Xths > 10 Then
                                       Xtmm = Xtmm + 700
                                    End If
                                 
                                 End If
                                 If Xtmm >= 0 Then
                                    XtotlleV = Xtmm
                                    XtotlleC = Xtmm
                                 Else
                                    XtotlleV = 0
                                    XtotlleC = 0
                                 End If
                                 data_info.Recordset.AddNew
                                 data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                                 data_info.Recordset("nro") = data_llam.Recordset("nrolla")
                                 data_info.Recordset("hora") = data_llam.Recordset("hora")
                                 data_info.Recordset("matric") = data_llam.Recordset("matric")
                                 data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                                 data_info.Recordset("categ") = data_llam.Recordset("categ")
                                 data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                                 data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                                 data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                                 data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                                 data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                                 data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                                 data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                                 data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                                 data_info.Recordset("hor_rea") = data_llam.Recordset("hor_rea")
                                 data_info.Recordset("trasla") = data_llam.Recordset("trasla")
                                 data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                                 data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                                 data_info.Recordset("activo") = data_llam.Recordset("activo")
                                 data_info.Recordset("codmed") = data_llam.Recordset("codmed")
                                 data_info.Recordset("nommed") = data_llam.Recordset("nommed")
                                 If IsNull(data_llam.Recordset("cancela")) = False Then
                                   If data_llam.Recordset("cancela") = 1 Then
                                      data_info.Recordset("cancela") = 1
                                   Else
                                      data_info.Recordset("cancela") = 9
                                   End If
                                 Else
                                    data_info.Recordset("cancela") = 9
                                 End If
                                 If Check3.Value = 1 Then
                                    data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                                 Else
                                    If Check4.Value = 1 Then
                                       data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                                    End If
                                 End If
                                 If data_llam.Recordset("codmot") = "V" Then
                                    data_info.Recordset("tmm") = XtotminV
                                    data_info.Recordset("thh") = XtotlleV
                                    Xtoverde = Xtoverde + 1
                                 Else
                                    data_info.Recordset("tmm") = XtotminC
                                    data_info.Recordset("thh") = XtotlleC
                                    Xtocele = Xtocele + 1
                                 End If
                                 data_info.Recordset.Update
                              'End If
                           End If
                        End If
                        Xtogra = Xtogra + 1
                    End If
                    data_llam.Recordset.MoveNext
                    XtotminR = 0
                    XtotlleR = 0
                    XtotlleA = 0
                    XtotminA = 0
                    XtotminV = 0
                    XtotlleV = 0
                    XtotlleC = 0
                    XtotminC = 0
                 Loop
                 MiBaseact.Execute "Delete * from inflla where cancela =" & 1
                 data_info.Refresh
                 
                 Xtexto = "R"
                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
                 data_info.Refresh
                 
                 If data_info.Recordset.RecordCount > 0 Then
                    data_info.Recordset.MoveFirst
                    Do While Not data_info.Recordset.EOF
                       XtotminR = XtotminR + data_info.Recordset("tmm")
                       XtotlleR = XtotlleR + data_info.Recordset("thh")
                       Xtotcan = Xtotcan + 1
                       If IsNull(data_info.Recordset("tmm")) = True Then
                          data_info.Recordset.Edit
                          data_info.Recordset("tmm") = 0
                          data_info.Recordset.Update
                       End If
                       If data_info.Recordset("tmm") > 3 Then
                          Xtotcanmax = Xtotcanmax + 1
                       End If
                       If IsNull(data_info.Recordset("thh")) = True Then
                          data_info.Recordset.Edit
                          data_info.Recordset("thh") = 0
                          data_info.Recordset.Update
                       End If
                       If data_info.Recordset("thh") > 15 Then
                          Xtotcanmax2 = Xtotcanmax2 + 1
                       End If
                       
                       data_info.Recordset.MoveNext
                    Loop
                    data_res.Recordset.AddNew
        ' Totales de llamados por CLAVE
                    data_res.Recordset("mes") = Xtotcanmax2
                    data_res.Recordset("cob") = Xtogra
                    data_res.Recordset("quesob") = Xtorojos
                    data_res.Recordset("totuniv") = Xtoama
                    data_res.Recordset("totccou") = Xtoverde
                    If Xrojoama <= 0 Then
                       data_res.Recordset("promcel") = 0
                    Else
                       Xhh1 = Xrojoama / Xtoama * 100
                       data_res.Recordset("promcel") = Xhh1
                    End If
                    If Xrojover <= 0 Then
                       data_res.Recordset("cantc5m") = 0
                    Else
                       Xhh1 = Xrojover / Xtoverde * 100
                       data_res.Recordset("cantc5m") = Xhh1
                    End If
                    data_res.Recordset("comodo") = Xtocerti
                    If Xtorojos > 0 Then
                       Xhh1 = Xtotcanmax2 / Xtorojos * 100
                       Xhh1 = 100 - Xhh1
                    Else
                       data_res.Recordset("anoarq") = 0
                    End If
                    data_res.Recordset("anoarq") = Xhh1
                    
                    Xveonum1 = Xtorojos / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    data_res.Recordset("totevang") = XtotrealR
                    data_res.Recordset("totcudam") = XtotrealA
                    data_res.Recordset("totsmi") = XtotrealV
                    data_res.Recordset("totacom") = Xtot3tot
                    data_res.Recordset("totodon") = Xtot4atot
                    data_res.Recordset("totdeudas") = Xtot5atot
                    data_res.Recordset("celemas5") = Xtot6atot
                    data_res.Recordset("totceles") = Xtocele
                    If XtotrealR > 0 Then
                       Xporrear = XtotrealR / Xtorojos
                    Else
                       Xporrear = 0
                    End If
                    If XtotrealA > 0 Then
                       Xporreaa = XtotrealA / Xtoama
                    Else
                       Xporreaa = 0
                    End If
                    If XtotrealV > 0 Then
                       Xtoverde = Xtoverde + Xtocele
                       Xporreav = XtotrealV / Xtoverde
                       Xtoverde = Xtoverde - Xtocele
                    Else
                       Xporreav = 0
                    End If
                    If XtotrealC > 0 Then
                       If Xtocele > 0 Then
                          Xporreac = XtotrealC / Xtocele
                       Else
                          Xporreac = 0
                       End If
                    Else
                       Xporreac = 0
                    End If
                    
                    data_res.Recordset("totimpasa") = Xporrear * 100
                    data_res.Recordset("totcgali") = Xporreaa * 100
                    data_res.Recordset("totcima") = Xporreav * 100
                    data_res.Recordset("totcel") = Xporreac * 100
                                           
                    If Xveonum1 < 1 Then
                       data_res.Recordset("desc1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("desc1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    Xveonum1 = Xtoama / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("desc2") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("desc2") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    Xveonum1 = Xtoverde / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("desr") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("desr") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    Xveonum1 = Xtocerti / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("des21") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("des21") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
' Celestes
                    Xveonum1 = Xtocele / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("descele1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("descele1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    
        ' Fin de TOTALES
'                    data_res.Recordset("mes") = XtotminR / Xtotcan
                    data_res.Recordset("ano") = 0
'                    data_res.Recordset("anoarq") = Xtotcanmax
                    If Xtotcanmax > 0 Then
                       data_res.Recordset("mesarq") = Xtotcan / Xtotcanmax
                    Else
                       data_res.Recordset("mesarq") = 0
                    End If
'                    data_res.Recordset("comr") = XtotlleR / Xtotcan
                    data_res.Recordset("coma") = 0
'                    data_res.Recordset("comm") = Xtotcanmax2
                    If Xtotcanmax2 <= 0 Then
                       data_res.Recordset("comc") = 0
                    Else
                       data_res.Recordset("comc") = Xtotcan / Xtotcanmax2
                    End If
                    Xtotcanmax = 0
                    Xtotcanmax2 = 0
                    Xtotcan = 0
                    XtotminR = 0
                    XtotlleR = 0
                    data_res.Recordset.Update
                    data_res.Refresh
            
                 Else
                    data_res.Recordset.AddNew
        ' Totales de llamados por CLAVE
                    data_res.Recordset("mes") = Xtotcanmax2
                    data_res.Recordset("cob") = Xtogra
                    data_res.Recordset("quesob") = Xtorojos
                    data_res.Recordset("totuniv") = Xtoama
                    data_res.Recordset("totccou") = Xtoverde
                    data_res.Recordset("comodo") = Xtocerti
                    If Xtorojos > 0 Then
                       Xhh1 = Xtotcanmax2 / Xtorojos * 100
                       Xhh1 = 100 - Xhh1
                    Else
                       data_res.Recordset("mes") = 0
                       data_res.Recordset("totimpasa") = 0
                       data_res.Recordset("totacom") = 0
                    End If
                    data_res.Recordset("anoarq") = Xhh1
                    If Xtorojos > 0 Then
                    Else
                       data_res.Recordset("anoarq") = 0
                    End If
                    
                    Xveonum1 = Xtorojos / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    data_res.Recordset("totevang") = XtotrealR
                    data_res.Recordset("totcudam") = XtotrealA
                    data_res.Recordset("totsmi") = XtotrealV
                    data_res.Recordset("totacom") = Xtot3tot
                    data_res.Recordset("totodon") = Xtot4atot
                    data_res.Recordset("totdeudas") = Xtot5atot
                    data_res.Recordset("celemas5") = Xtot6atot
                    data_res.Recordset("totceles") = Xtocele
                    If XtotrealR > 0 Then
                       Xporrear = XtotrealR / Xtorojos
                    Else
                       Xporrear = 0
                    End If
                    If XtotrealA > 0 Then
                       Xporreaa = XtotrealA / Xtoama
                    Else
                       Xporreaa = 0
                    End If
                    If XtotrealV > 0 Then
                       Xtoverde = Xtoverde + Xtocele
                       Xporreav = XtotrealV / Xtoverde
                       Xtoverde = Xtoverde - Xtocele
                    Else
                       Xporreav = 0
                    End If
                    If XtotrealC > 0 Then
                       If Xtocele > 0 Then
                          Xporreac = XtotrealC / Xtocele
                       Else
                          Xporreac = 0
                       End If
                    Else
                       Xporreac = 0
                    End If
                    
                    data_res.Recordset("totimpasa") = Xporrear * 100
                    data_res.Recordset("totcgali") = Xporreaa * 100
                    data_res.Recordset("totcima") = Xporreav * 100
                    data_res.Recordset("totcel") = Xporreac * 100
                                           
                    If Xveonum1 < 1 Then
                       data_res.Recordset("desc1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("desc1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    Xveonum1 = Xtoama / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("desc2") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("desc2") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    Xveonum1 = Xtoverde / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("desr") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("desr") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    Xveonum1 = Xtocerti / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("des21") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("des21") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
' Celestes
                    Xveonum1 = Xtocele / Xtogra
                    Xveonum1 = Xveonum1 * 100
                    If Xveonum1 < 1 Then
                       data_res.Recordset("descele1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                    Else
                       data_res.Recordset("descele1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                    End If
                    
                    
        ' Fin de TOTALES
'                    data_res.Recordset("mes") = XtotminR / Xtotcan
                    data_res.Recordset("ano") = 0
'                    data_res.Recordset("anoarq") = Xtotcanmax
                    If Xtotcanmax > 0 Then
                       data_res.Recordset("mesarq") = Xtotcan / Xtotcanmax
                    Else
                       data_res.Recordset("mesarq") = 0
                    End If
'                    data_res.Recordset("comr") = XtotlleR / Xtotcan
                    data_res.Recordset("coma") = 0
'                    data_res.Recordset("comm") = Xtotcanmax2
                    If Xtotcanmax2 <= 0 Then
                       data_res.Recordset("comc") = 0
                    Else
                       data_res.Recordset("comc") = Xtotcan / Xtotcanmax2
                    End If
                    Xtotcanmax = 0
                    Xtotcanmax2 = 0
                    Xtotcan = 0
                    XtotminR = 0
                    XtotlleR = 0
                    data_res.Recordset.Update
                    data_res.Refresh
                 
                 
                 End If
                 
                 Xtexto = "A"
                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
                 data_info.Refresh
                 If data_info.Recordset.RecordCount > 0 Then
                    data_info.Recordset.MoveFirst
                    Do While Not data_info.Recordset.EOF
                       XtotminR = XtotminR + data_info.Recordset("tmm")
                       XtotlleR = XtotlleR + data_info.Recordset("thh")
                       Xtotcan = Xtotcan + 1
                       If data_info.Recordset("tmm") > 5 Then
                          Xtotcanmax = Xtotcanmax + 1
                       End If
                       If data_info.Recordset("thh") > 30 Then
                          Xtotcanmax2 = Xtotcanmax2 + 1
                       End If
                       data_info.Recordset.MoveNext
                    Loop
                    If data_res.Recordset.RecordCount > 0 Then
                       data_res.Recordset.Edit
                    Else
                       data_res.Recordset.AddNew
                    End If
                    data_res.Recordset("comr") = Xtotcanmax2
'                    data_res.Recordset("totrec") = XtotminR / Xtotcan
                    data_res.Recordset("totimp") = 0
'                    data_res.Recordset("comv") = Xtotcanmax
                    If Xtotcanmax > 0 Then
                       data_res.Recordset("totimpu") = Xtotcan / Xtotcanmax
                    Else
                       data_res.Recordset("totimpu") = 0
                    End If
'                    data_res.Recordset("totrecu") = XtotlleR / Xtotcan
                    data_res.Recordset("iva1") = 0
                    If Xtoama > 0 Then
                       Xhh1 = Xtotcanmax2 / Xtoama * 100
                       Xhh1 = 100 - Xhh1
                    Else
                       Xhh1 = 0
                    End If
                    data_res.Recordset("comm") = Xhh1

'                    data_res.Recordset("iva2") = Xtotcanmax2
                    If Xtotcanmax2 <= 0 Then
                       data_res.Recordset("ivatot") = 0
                    Else
                       data_res.Recordset("ivatot") = Xtotcan / Xtotcanmax2
                    End If
                    Xtotcanmax = 0
                    Xtotcanmax2 = 0
                    Xtotcan = 0
                    XtotminR = 0
                    XtotlleR = 0
                    data_res.Recordset.Update
                    data_res.Refresh
                 End If
                 
                 Xtexto = "V"
                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
'                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "' or codmot ='" & Trim("C") & "'"
                 data_info.Refresh
                 Dim Xtotmas3, Xtotmas30en, xhh, xmm, Xhhh, Xmmh, xdemh, xdemm As Long
                 Dim Xporcenver As Double
                 If data_info.Recordset.RecordCount > 0 Then
                    data_info.Recordset.MoveFirst
                    If Combo3.ListIndex >= 0 Then
                       MiBaseact.Execute "Delete * from inflla where codmot in ('Z')"
                       data_info.Refresh
                    Else
                       Do While Not data_info.Recordset.EOF
                          XtotminR = XtotminR + data_info.Recordset("tmm")
                          XtotlleR = XtotlleR + data_info.Recordset("thh")
                          Xtotcan = Xtotcan + 1
                          If data_info.Recordset("thh") > 120 Then
                             Xtotcanmax = Xtotcanmax + 1
                          Else
                             Xtomeno2 = Xtomeno2 + 1
                          End If
                          If data_info.Recordset("tmm") > 180 Then
                             Xtotmas3 = Xtotmas3 + 1
                          End If
                          If data_info.Recordset("thh") > 180 Then
                             Xtotcanmax2 = Xtotcanmax2 + 1
                          End If
                          If IsNull(data_info.Recordset("hor_llega")) = True Then
                          Else
                             xhh = Val(Mid(data_info.Recordset("hor_llega"), 1, 2))
                             xmm = Val(Mid(data_info.Recordset("hor_llega"), 4, 2))
                             If IsNull(data_info.Recordset("hor_rea")) = False Then
                                Xhhh = Val(Mid(data_info.Recordset("hor_rea"), 1, 2))
                                Xmmh = Val(Mid(data_info.Recordset("hor_rea"), 4, 2))
                             End If
                             xdemh = Xhhh - xhh
                             xdemm = Xmmh - xmm
                             If data_info.Recordset("fecha") < data_info.Recordset("fec_llega") Then
                                If xdemh < 0 Then
                                   xdemh = Xhhh - xhh
                                   xdemh = xdemh + 24
                                End If
                             Else
                                If IsNull(data_info.Recordset("fec_llega")) = True Then
                                   xdemh = Xhhh - xhh
                                   xdemh = xdemh + 24
                                Else
                                   If xdemh < 0 Then
                                      xdemh = xdemh + 24
                                   End If
                                End If
                             End If
                             If xdemh > 0 Then
                                If xdemm < 0 Then
                                   xdemm = xdemm + 60
                                   xdemh = xdemh - 1
                                End If
                             Else
                                If xdemm < 0 Then
                                   xdemm = xdemm + 60
                                End If
                             End If
                             data_info.Recordset.Edit
                             If xdemh > 9 Then
                                If xdemm > 9 Then
                                   data_info.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                                Else
                                   data_info.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                                End If
                             Else
                                If xdemm > 9 Then
                                   If xdemh < 0 Then
                                      xdemh = 0
                                   End If
                                   data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                                Else
                                   If xdemh < 0 Then
                                      xdemh = 0
                                   End If
                                   If xdemm < 0 Then
                                      xdemm = 0
                                   End If
                                   data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                                End If
                             End If
                             data_info.Recordset.Update
                          End If
                          data_info.Recordset.MoveNext
                       Loop
                       data_res.Recordset.Edit
                       data_res.Recordset("totrec") = Xtotcanmax2
                       data_res.Recordset("comuni") = XtotminR / Xtotcan
                       data_res.Recordset("comccou") = 0
                       data_res.Recordset("comimp") = Xtotcanmax
                       data_res.Recordset("comacom") = Xtotmas3
    '                   Xhh1 = Xtotmas3 / Xtoverde * 100
                       Xporcenver = Xtotcanmax2 / Xtoverde * 100
    '                   Xhh1 = 100 - Xhh1
                       Xporcenver = 100 - Xporcenver
                       
    '                    data_res.Recordset("comv") = Xhh1
                       data_res.Recordset("comv") = Xporcenver
                        
                       If Xtotcanmax <= 0 Then
                          data_res.Recordset("comgal") = 0
                       Else
                          data_res.Recordset("comgal") = Xtotcan / Xtotcanmax
                       End If
                       data_res.Recordset("comcima") = XtotlleR / Xtotcan
                       data_res.Recordset("comeva") = 0
                       data_res.Recordset("comcud") = Xtotcanmax2
                       If Xtotcanmax2 <= 0 Then
                          data_res.Recordset("comsmi") = 0
                       Else
                          data_res.Recordset("comsmi") = Xtotcan / Xtotcanmax2
                       End If
                       data_res.Recordset("comdeu") = Xtomeno2 / Xtoverde * 100
                       data_res.Recordset.Update
                       data_res.Refresh
                    End If
                 End If
                 
                 Xtexto = "C"
                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
'                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "' or codmot ='" & Trim("C") & "'"
                 data_info.Refresh
                 Xtotcanmax = 0
                 Xtomeno2 = 0
                 Xtotcanmax2 = 0
'                 Dim Xtotmas3, Xtotmas30en, xhh, xmm, Xhhh, Xmmh, xdemh, xdemm As Long
                 If data_info.Recordset.RecordCount > 0 Then
                    data_info.Recordset.MoveFirst
                    Do While Not data_info.Recordset.EOF
                       XtotminR = XtotminR + data_info.Recordset("tmm")
                       XtotlleR = XtotlleR + data_info.Recordset("thh")
                       Xtotcan = Xtotcan + 1
                       If data_info.Recordset("tmm") > 5 Then
                          Xtotcanmax = Xtotcanmax + 1
                       Else
                          Xtomeno2 = Xtomeno2 + 1
                       End If
                       If data_info.Recordset("thh") > 30 Then
                          Xtotcanmax2 = Xtotcanmax2 + 1
                       End If
                       If IsNull(data_info.Recordset("hor_llega")) = True Then
                       Else
                          xhh = Val(Mid(data_info.Recordset("hor_llega"), 1, 2))
                          xmm = Val(Mid(data_info.Recordset("hor_llega"), 4, 2))
                          If IsNull(data_info.Recordset("hor_rea")) = False Then
                             Xhhh = Val(Mid(data_info.Recordset("hor_rea"), 1, 2))
                             Xmmh = Val(Mid(data_info.Recordset("hor_rea"), 4, 2))
                          End If
                          xdemh = Xhhh - xhh
                          xdemm = Xmmh - xmm
                          If data_info.Recordset("fecha") < data_info.Recordset("fec_llega") Then
                             If xdemh < 0 Then
                                xdemh = Xhhh - xhh
                                xdemh = xdemh + 24
                             End If
                          Else
                             If IsNull(data_info.Recordset("fec_llega")) = True Then
                                xdemh = Xhhh - xhh
                                xdemh = xdemh + 24
                             Else
                                If xdemh < 0 Then
                                   xdemh = xdemh + 24
                                End If
                             End If
                          End If
                          If xdemh > 0 Then
                             If xdemm < 0 Then
                                xdemm = xdemm + 60
                                xdemh = xdemh - 1
                             End If
                          Else
                             If xdemm < 0 Then
                                xdemm = xdemm + 60
                             End If
                          End If
                          data_info.Recordset.Edit
                          If xdemh > 9 Then
                             If xdemm > 9 Then
                                data_info.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                             Else
                                data_info.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                             End If
                          Else
                             If xdemm > 9 Then
                                If xdemh < 0 Then
                                   xdemh = 0
                                End If
                                data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                             Else
                                If xdemh < 0 Then
                                   xdemh = 0
                                End If
                                data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                             End If
                          End If
                          data_info.Recordset.Update
                       End If
                       data_info.Recordset.MoveNext
                    Loop

                    data_res.Recordset.Edit
                    data_res.Recordset("promcel") = XtotminR / Xtotcan ' promedio
                    Xhh1 = Xtotcanmax2 / Xtocele * 100
                    Xhh1 = 100 - Xhh1
                    data_res.Recordset("iva2") = Xhh1
'                    data_res.Recordset("comccou") = 0
                    data_res.Recordset("cantc5m") = Xtotcanmax ' cantidad que pasaron los 5min de salida
                    If Xtotcanmax <= 0 Then
                       data_res.Recordset("promcmas5") = 0
                    Else
                       data_res.Recordset("promcmas5") = Xtotcan / Xtotcanmax
                    End If
                    data_res.Recordset("promcmas30") = XtotlleR / Xtotcan ' Promedio de la llegada
                    data_res.Recordset("cantc30m") = Xtotcanmax2 ' Cantidad que pasaron los 30m en llegada
                    If Xtotcanmax2 <= 0 Then
                       data_res.Recordset("promcmas302") = 0
                    Else
                       data_res.Recordset("promcmas302") = Xtotcan / Xtotcanmax2 'Porcentaje de desvío en llegada mas 30m
                    End If
                    data_res.Recordset("totrecu") = Xtotcanmax2
                    data_res.Recordset.Update
                    data_res.Refresh
                 End If
                    If Combo3.ListIndex >= 0 Then
                       data_info.RecordSource = "Select * from inflla"
                       data_info.Refresh
                       If data_info.Recordset.RecordCount > 0 Then
                          data_info.Recordset.MoveFirst
                          Do While Not data_info.Recordset.EOF
                             data_llamod.RecordSource = "Select * from resplla where nro =" & data_info.Recordset("nro")
                             data_llamod.Refresh
                             If data_llamod.Recordset.RecordCount > 0 Then
                                If IsNull(data_llamod.Recordset("movil_rea")) = False Then
                                   If data_llamod.Recordset("movil_rea") > 0 Then
                                      data_chof.RecordSource = "Select * from movil where nromov =" & data_llamod.Recordset("movil_rea")
                                      data_chof.Refresh
                                      If data_chof.Recordset.RecordCount > 0 Then
                                         data_info.Recordset.Edit
                                         data_info.Recordset("direcc") = data_chof.Recordset("chofer")
                                         data_info.Recordset("edad") = data_llamod.Recordset("movil_rea")
                                         data_info.Recordset.Update
                                      End If
                                   End If
                                End If
                             End If
                             data_info.Recordset.MoveNext
                          Loop
                       End If
                    End If
                 
                 If data_res.Recordset("mes") > 0 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("desa") = "NO"
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("desa") = "SI"
                    data_res.Recordset.Update
                 End If
                 If data_res.Recordset("comr") > 0 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("desm") = "NO"
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("desm") = "SI"
                    data_res.Recordset.Update
                 End If
                 If data_res.Recordset("totrec") > 0 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("desv") = "NO"
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("desv") = "SI"
                    data_res.Recordset.Update
                 End If
                 If data_res.Recordset("totrecu") > 0 Then 'ultimas modificaciones
                    data_res.Recordset.Edit
                    data_res.Recordset("des25") = "NO"
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des25") = "SI"
                    data_res.Recordset.Update
                 End If
'''''' hasta aquiiii
        ' amarillos
                 If data_res.Recordset("totrec") <= 3 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("des23") = "SI"
                    data_res.Recordset("des24") = ""
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des24") = "NO"
                    data_res.Recordset("des23") = ""
                    data_res.Recordset.Update
                 End If
        ' llegada
                 If data_res.Recordset("totrecu") < 20 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("des27") = "SI"
                    data_res.Recordset("des28") = ""
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des28") = "NO"
                    data_res.Recordset("des27") = ""
                    data_res.Recordset.Update
                 End If
                 If data_res.Recordset("iva2") > 0 Then 'ultimas modificaciones
                    data_res.Recordset.Edit
                    data_res.Recordset("des29") = "NO"
                    data_res.Recordset("des30") = ""
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des30") = "SI"
                    data_res.Recordset("des29") = ""
                    data_res.Recordset.Update
                 End If
        ' VERDES solo llegada
                 If data_res.Recordset("comcima") < 120 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("des31") = "SI"
                    data_res.Recordset("des32") = ""
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des32") = "NO"
                    data_res.Recordset("des31") = ""
                    data_res.Recordset.Update
                 End If
                 If data_res.Recordset("comcud") > 0 Then ' ultimos modificaciones
                    data_res.Recordset.Edit
                    data_res.Recordset("des33") = "NO"
                    data_res.Recordset("des34") = ""
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des34") = "SI"
                    data_res.Recordset("des33") = ""
                    data_res.Recordset.Update
                 End If
        'salida
                 If data_res.Recordset("comuni") < 7 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("des35") = "SI"
                    data_res.Recordset("des36") = ""
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des36") = "NO"
                    data_res.Recordset("des35") = ""
                    data_res.Recordset.Update
                 End If
                 If data_res.Recordset("comgal") < 5 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("des37") = "SI"
                    data_res.Recordset("des38") = ""
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des38") = "NO"
                    data_res.Recordset("des37") = ""
                    data_res.Recordset.Update
                 End If
'' CELESTES
                 If data_res.Recordset("promcel") <= 3 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("des26") = "SI"
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des26") = "NO"
                    data_res.Recordset.Update
                 End If

                 If data_res.Recordset("promcmas30") <= 20 Then
                    data_res.Recordset.Edit
                    data_res.Recordset("des37") = "SI"
                    data_res.Recordset.Update
                 Else
                    data_res.Recordset.Edit
                    data_res.Recordset("des37") = "NO"
                    data_res.Recordset.Update
                 End If

                 If Check5.Value = 1 Then
                    Command5_Click
                 End If
                 frm_calidadiso.MousePointer = 0
                 MsgBox "Proceso terminado"
                 data_res.RecordSource = "Select * from infarqc"
                 data_res.Refresh
                 data_info.RecordSource = "Select * from inflla"
                 data_info.Refresh
                 If Combo3.ListIndex >= 0 Then
                    cr2.ReportFileName = App.path & "\infcaldesch.rpt"
                    cr2.ReportTitle = "INFORME PROMEDIOS DEMORAS POR CHOFERES FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                    CrystalReport1.ReportTitle = "INFORME DEMORAS DE MOVILES POR CHOFERES FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                    If Option1.Value = True Then
                       CrystalReport1.ReportFileName = App.path & "\infcaldeschn.rpt"
                    Else
                       CrystalReport1.ReportFileName = App.path & "\infcaldeschdt.rpt"
                    End If
                    cr2.Action = 1
                    CrystalReport1.DiscardSavedData = True
                 Else
                    If Option1.Value = True Then
                       If Check6.Value = 1 Then
                          CrystalReport1.ReportFileName = App.path & "\infcaldes22n.rpt"
                          CrystalReport1.DiscardSavedData = True
                       Else
                          CrystalReport1.ReportFileName = App.path & "\infcaldes2n.rpt"
                          CrystalReport1.DiscardSavedData = True
                       End If
                       If t_mov.Text = "" Then
                          CrystalReport1.ReportTitle = "INFORME DEMORAS DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                       Else
                          CrystalReport1.ReportTitle = "INFORME DEMORAS DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text & " MOVIL: " & t_mov.Text
                       End If
                    Else
                       If Check6.Value = 1 Then
                          CrystalReport1.ReportFileName = App.path & "\infcaldes22.rpt"
                          CrystalReport1.DiscardSavedData = True
                       Else
                          CrystalReport1.ReportFileName = App.path & "\infcaldes2.rpt"
                          CrystalReport1.DiscardSavedData = True
                       End If
                       If t_mov.Text = "" And t_codmed.Text = "" Then
                          CrystalReport1.ReportTitle = "INFORME DEMORAS DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                       Else
                          If t_codmed.Text <> "" Then
                             CrystalReport1.ReportTitle = "INFORME DEMORAS DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text & " MEDICO: " & t_codmed.Text
                          Else
                             CrystalReport1.ReportTitle = "INFORME DEMORAS DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text & " MOVIL: " & t_mov.Text
                          End If
                       End If
                    End If
                 End If
                 CrystalReport1.Action = 1
              End If
          End If
      End If
   Else
      MsgBox "Verifique fechas"
   End If
Else
   MsgBox "Verifique fechas"
End If
   
   
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
          If data_llam.Recordset.RecordCount > 0 Then
             data_llam.Recordset.MoveFirst
             Do While Not data_llam.Recordset.EOF
                If data_llam.Recordset("codzon") = 3 Then
                    If data_llam.Recordset("codmot") = "R" Then
                       If IsNull(data_llam.Recordset("colormot")) = False Then
                          If data_llam.Recordset("colormot") = "R" Then
                             XtotrealR = XtotrealR + 1
                          End If
                       End If
                       If IsNull(data_llam.Recordset("horsali")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
''''' nuevoooo
                       If IsNull(data_llam.Recordset("horpas")) = False Then
                          Xhh3 = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                          Xmm3 = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                       End If
                       If Xmm3 = Xmm1 Then
                          If Xhh3 = Xhh1 Then
                             Xtot3 = 0
                          Else
                             Xths = Xhh1 - Xhh3
                             Xtot3 = Xmm1 - Xmm3 + 60
                             If Xths = 2 Then
                                Xtot3 = Xtot3 + 60
                             End If
                             If Xths = 3 Then
                                Xtot3 = Xtot3 + 120
                             End If
                          End If
                       Else
                          If Xhh3 = Xhh1 Then
                             Xtot3 = Xmm1 - Xmm3
                          Else
                             Xths = Xhh1 - Xhh3
                             Xtot3 = Xmm1 - Xmm3 + 60
                             If Xths = 2 Then
                                Xtot3 = Xtot3 + 60
                             End If
                             If Xths = 3 Then
                                Xtot3 = Xtot3 + 120
                             End If
                          End If
                       End If
                       If Xtot3 > 3 Then
                          Xtot3tot = Xtot3tot + 1
                       End If
''''-------------------------------
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          If IsNull(data_llam.Recordset("fecsali")) = False Then
                             If data_llam.Recordset("fecsali") > data_llam.Recordset("fecha") Then
                                Xths = Xths + 24
                             End If
                          End If
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotminR = Xtot3
                       Else
                          XtotminR = 0
                       End If
        '' Llegada!!!
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          If IsNull(data_llam.Recordset("fec_llega")) = False Then
                             If data_llam.Recordset("fec_llega") > data_llam.Recordset("fecha") Then
                                Xths = Xths + 24
                             End If
                          End If
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotlleR = Xtmm
                       Else
                          XtotlleR = 0
                       End If
                       data_info.Recordset.AddNew
                       data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                       data_info.Recordset("hora") = data_llam.Recordset("hora")
                       data_info.Recordset("matric") = data_llam.Recordset("matric")
                       data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                       data_info.Recordset("categ") = data_llam.Recordset("categ")
                       data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                       data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                       data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                       data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                       data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                       data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                       data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                       data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                       data_info.Recordset("hor_rea") = data_llam.Recordset("hor_rea")
                       data_info.Recordset("trasla") = data_llam.Recordset("trasla")
                       data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                       data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                       data_info.Recordset("activo") = data_llam.Recordset("activo")
                       data_info.Recordset("codmed") = data_llam.Recordset("codmed")
                       data_info.Recordset("nommed") = data_llam.Recordset("nommed")
                        If IsNull(data_llam.Recordset("cancela")) = False Then
                           If data_llam.Recordset("cancela") = 1 Then
                              data_info.Recordset("cancela") = 1
                           Else
                              data_info.Recordset("cancela") = 9
                           End If
                        Else
                           data_info.Recordset("cancela") = 9
                        End If
                       
                       If Check3.Value = 1 Then
                          data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                       Else
                          If Check4.Value = 1 Then
                             data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                          End If
                       End If
                       data_info.Recordset("tmm") = XtotminR
                       data_info.Recordset("thh") = XtotlleR
                       data_info.Recordset.Update
                       Xtorojos = Xtorojos + 1
                    End If
                    If data_llam.Recordset("codmot") = "A" Then
                       If IsNull(data_llam.Recordset("colormot")) = False Then
                          If data_llam.Recordset("colormot") = "R" Then
                             Xrojoama = Xrojoama + 1
                          End If
                       End If
                       If IsNull(data_llam.Recordset("colormot")) = False Then
                          If data_llam.Recordset("colormot") = "A" Then
                             XtotrealA = XtotrealA + 1
                          End If
                       End If
                       If IsNull(data_llam.Recordset("horsali")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("horpas")) = False Then
                          Xhh4a = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                          Xmm4a = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                       End If
                       If Xmm4a = Xmm1 Then
                          If Xhh4a = Xhh1 Then
                             Xtot4a = 0
                          Else
                            Xths = Xhh1 - Xhh4a
                            Xtot4a = Xmm1 - Xmm4a + 60
                            If Xths = 2 Then
                               Xtot4a = Xtot4a + 60
                            End If
                            If Xths = 3 Then
                               Xtot4a = Xtot4a + 120
                            End If
                          End If
                       Else
                          If Xhh4a = Xhh1 Then
                             Xtot4a = Xmm1 - Xmm4a
                          Else
                             Xths = Xhh1 - Xhh4a
                             Xtot4a = Xmm1 - Xmm4a + 60
                             If Xths = 2 Then
                                Xtot4a = Xtot4a + 60
                             End If
                             If Xths = 3 Then
                                Xtot4a = Xtot4a + 120
                             End If
                          End If
                       End If
                       If Xtot4a > 5 Then
                          Xtot4atot = Xtot4atot + 1
                       End If
                       
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          If IsNull(data_llam.Recordset("fecsali")) = False Then
                             If data_llam.Recordset("fecsali") > data_llam.Recordset("fecha") Then
                                Xths = Xths + 24
                             End If
                          End If
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotminA = Xtot4a
                       Else
                          XtotminA = 0
                       End If
        
        '' Llegada!!!
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          If IsNull(data_llam.Recordset("fec_llega")) = False Then
                             If data_llam.Recordset("fec_llega") > data_llam.Recordset("fecha") Then
                                Xths = Xths + 24
                             End If
                          End If
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotlleA = Xtmm
                       Else
                          XtotlleA = 0
                       End If
                       data_info.Recordset.AddNew
                       data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                       data_info.Recordset("hora") = data_llam.Recordset("hora")
                       data_info.Recordset("matric") = data_llam.Recordset("matric")
                       data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                       data_info.Recordset("categ") = data_llam.Recordset("categ")
                       data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                       data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                       data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                       data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                       data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                       data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                       data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                       data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                       data_info.Recordset("hor_rea") = data_llam.Recordset("hor_rea")
                       data_info.Recordset("trasla") = data_llam.Recordset("trasla")
                       data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                       data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                       data_info.Recordset("activo") = data_llam.Recordset("activo")
                       data_info.Recordset("codmed") = data_llam.Recordset("codmed")
                       data_info.Recordset("nommed") = data_llam.Recordset("nommed")
                        If IsNull(data_llam.Recordset("cancela")) = False Then
                           If data_llam.Recordset("cancela") = 1 Then
                              data_info.Recordset("cancela") = 1
                           Else
                              data_info.Recordset("cancela") = 9
                           End If
                        Else
                           data_info.Recordset("cancela") = 9
                        End If
                       
                       If Check3.Value = 1 Then
                          data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                       Else
                          If Check4.Value = 1 Then
                             data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                          End If
                       End If
                       data_info.Recordset("tmm") = XtotminA
                       data_info.Recordset("thh") = XtotlleA
                       data_info.Recordset.Update
                       Xtoama = Xtoama + 1
                    End If
                    If data_llam.Recordset("codmot") = "V" Or data_llam.Recordset("codmot") = "C" Then
                       If data_llam.Recordset("categ") = "UDEMM" Or _
                          data_llam.Recordset("categ") = "CERSEM" Or _
                          data_llam.Recordset("categ") = "CERCAS" Or _
                          data_llam.Recordset("categ") = "CERANT" Or _
                          data_llam.Recordset("categ") = "CERADU" Or _
                          data_llam.Recordset("categ") = "CERDGI" Or _
                          data_llam.Recordset("categ") = "CERESS" Or _
                          data_llam.Recordset("categ") = "CERSAP" Or _
                          data_llam.Recordset("categ") = "CERHEV" Or _
                          data_llam.Recordset("categ") = "CERIMP" Or _
                          data_llam.Recordset("categ") = "CERKEV" Or _
                          data_llam.Recordset("categ") = "CERMAT" Or _
                          data_llam.Recordset("categ") = "CERSEV" Or _
                          data_llam.Recordset("categ") = "CERVIS" Then
                          Xtocerti = Xtocerti + 1
                       Else
                          If IsNull(data_llam.Recordset("colormot")) = False Then
                             If data_llam.Recordset("colormot") = "V" Then
                                XtotrealV = XtotrealV + 1
                             Else
                                XtotrealC = XtotrealC + 1
                             End If
                             If data_llam.Recordset("colormot") = "R" Then
                                Xrojover = Xrojover + 1
                             End If
                          End If
                             If IsNull(data_llam.Recordset("horsali")) = False Then
                                Xhh1 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                                Xmm1 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                             End If
                             If IsNull(data_llam.Recordset("hora")) = False Then
                                Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                             End If
                             If IsNull(data_llam.Recordset("horpas")) = False Then
                                Xhh5a = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                                Xmm5a = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                             End If
                             If Xmm5a = Xmm1 Then
                                If Xhh5a = Xhh1 Then
                                   Xtot5a = 0
                                Else
                                   Xtot5a = Xhh1 - Xhh5a
                                End If
                             Else
                                If Xhh5a = Xhh1 Then
                                   Xtot5a = Xmm1 - Xmm5a
                                Else
                                   Xtot5a = Xmm1 - Xmm5a + 60
                                End If
                             End If
                             If data_llam.Recordset("codmot") = "V" Then
                                If Xtot5a > 10 Then
                                   Xtot5atot = Xtot5atot + 1
                                End If
                             Else
                                If Xtot5a > 5 Then
                                   Xtot6atot = Xtot6atot + 1
                                End If
                             End If
                             If Xhh1 = Xhh5a Then
                                Xtmm = Xmm1 - Xmm5a
                             Else
                                Xths = Xhh1 - Xhh5a
                                If IsNull(data_llam.Recordset("fecsali")) = False Then
                                   If data_llam.Recordset("fecsali") > data_llam.Recordset("fecha") Then
                                      Xths = Xths + 24
                                   End If
                                End If
                                
                                Xtmm = Xmm1 - Xmm5a + 60
                                If Xths = 2 Then
                                   Xtmm = Xtmm + 60
                                End If
                                If Xths = 3 Then
                                   Xtmm = Xtmm + 120
                                End If
                                If Xths = 4 Then
                                   Xtmm = Xtmm + 180
                                End If
                                If Xths = 5 Then
                                   Xtmm = Xtmm + 240
                                End If
                                If Xths = 6 Then
                                   Xtmm = Xtmm + 300
                                End If
                                If Xths = 7 Then
                                   Xtmm = Xtmm + 420
                                End If
                                If Xths = 8 Then
                                   Xtmm = Xtmm + 480
                                End If
                                If Xths = 9 Then
                                   Xtmm = Xtmm + 540
                                End If
                                If Xths = 10 Then
                                   Xtmm = Xtmm + 600
                                End If
                                If Xths > 10 Then
                                   Xtmm = Xtmm + 700
                                End If
                             
                             End If
                             If Xtmm >= 0 Then
                                XtotminV = Xtmm
                                XtotminC = Xtmm
                             Else
                                XtotminV = 0
                                XtotminC = 0
                             End If
              '' Llegada!!!
                             If IsNull(data_llam.Recordset("hor_llega")) = False Then
                                Xhh1 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                                Xmm1 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                             End If
                             If IsNull(data_llam.Recordset("hora")) = False Then
                                Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                             End If
                             If Xhh1 = Xhh2 Then
                                Xtmm = Xmm1 - Xmm2
                             Else
                                Xths = Xhh1 - Xhh2
                                If IsNull(data_llam.Recordset("fec_llega")) = False Then
                                   If data_llam.Recordset("fec_llega") > data_llam.Recordset("fecha") Then
                                      Xths = Xths + 24
                                   End If
                                End If
                                Xtmm = Xmm1 - Xmm2 + 60
                                If Xths = 2 Then
                                   Xtmm = Xtmm + 60
                                End If
                                If Xths = 3 Then
                                   Xtmm = Xtmm + 120
                                End If
                                If Xths = 4 Then
                                   Xtmm = Xtmm + 180
                                End If
                                If Xths = 5 Then
                                   Xtmm = Xtmm + 240
                                End If
                                If Xths = 6 Then
                                   Xtmm = Xtmm + 300
                                End If
                                If Xths = 7 Then
                                   Xtmm = Xtmm + 420
                                End If
                                If Xths = 8 Then
                                   Xtmm = Xtmm + 480
                                End If
                                If Xths = 9 Then
                                   Xtmm = Xtmm + 540
                                End If
                                If Xths = 10 Then
                                   Xtmm = Xtmm + 600
                                End If
                                If Xths > 10 Then
                                   Xtmm = Xtmm + 700
                                End If
                             
                             End If
                             If Xtmm >= 0 Then
                                XtotlleV = Xtmm
                                XtotlleC = Xtmm
                             Else
                                XtotlleV = 0
                                XtotlleC = 0
                             End If
                             data_info.Recordset.AddNew
                             data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                             data_info.Recordset("hora") = data_llam.Recordset("hora")
                             data_info.Recordset("matric") = data_llam.Recordset("matric")
                             data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                             data_info.Recordset("categ") = data_llam.Recordset("categ")
                             data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                             data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                             data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                             data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                             data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                             data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                             data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                             data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                             data_info.Recordset("hor_rea") = data_llam.Recordset("hor_rea")
                             data_info.Recordset("trasla") = data_llam.Recordset("trasla")
                             data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                             data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                             data_info.Recordset("activo") = data_llam.Recordset("activo")
                            data_info.Recordset("codmed") = data_llam.Recordset("codmed")
                            data_info.Recordset("nommed") = data_llam.Recordset("nommed")
                            If IsNull(data_llam.Recordset("cancela")) = False Then
                               If data_llam.Recordset("cancela") = 1 Then
                                  data_info.Recordset("cancela") = 1
                               Else
                                  data_info.Recordset("cancela") = 9
                               End If
                            Else
                               data_info.Recordset("cancela") = 9
                            End If
                             If Check3.Value = 1 Then
                                data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                             Else
                                If Check4.Value = 1 Then
                                   data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                                End If
                             End If
                             If data_llam.Recordset("codmot") = "V" Then
                                data_info.Recordset("tmm") = XtotminV
                                data_info.Recordset("thh") = XtotlleV
                                Xtoverde = Xtoverde + 1
                             Else
                                data_info.Recordset("tmm") = XtotminC
                                data_info.Recordset("thh") = XtotlleC
                                Xtocele = Xtocele + 1
                             End If
                             data_info.Recordset.Update
                          'End If
                       End If
                    End If
                    Xtogra = Xtogra + 1
                End If
                data_llam.Recordset.MoveNext
                XtotminR = 0
                XtotlleR = 0
                XtotlleA = 0
                XtotminA = 0
                XtotminV = 0
                XtotlleV = 0
                XtotlleC = 0
                XtotminC = 0
             Loop
             Dim MiBaseact As Database
             Dim Unasesact As Workspace
             Set Unasesact = Workspaces(0)
             Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
            
             MiBaseact.Execute "Delete * from inflla where cancela =" & 1
             data_info.Refresh
             
             Xtexto = "R"
             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
             data_info.Refresh
             
             If data_info.Recordset.RecordCount > 0 Then
                data_info.Recordset.MoveFirst
                Do While Not data_info.Recordset.EOF
                   XtotminR = XtotminR + data_info.Recordset("tmm")
                   XtotlleR = XtotlleR + data_info.Recordset("thh")
                   Xtotcan = Xtotcan + 1
                   If IsNull(data_info.Recordset("tmm")) = True Then
                      data_info.Recordset.Edit
                      data_info.Recordset("tmm") = 0
                      data_info.Recordset.Update
                   End If
                   If data_info.Recordset("tmm") > 3 Then
                      Xtotcanmax = Xtotcanmax + 1
                   End If
                   If IsNull(data_info.Recordset("thh")) = True Then
                      data_info.Recordset.Edit
                      data_info.Recordset("thh") = 0
                      data_info.Recordset.Update
                   End If
                   If data_info.Recordset("thh") > 15 Then
                      Xtotcanmax2 = Xtotcanmax2 + 1
                   End If
                   data_info.Recordset.MoveNext
                Loop
                data_res.Recordset.AddNew
    ' Totales de llamados por CLAVE
                data_res.Recordset("mes") = Xtotcanmax2
                data_res.Recordset("cob") = Xtogra
                data_res.Recordset("quesob") = Xtorojos
                data_res.Recordset("totuniv") = Xtoama
                data_res.Recordset("totccou") = Xtoverde
                If Xrojoama <= 0 Then
                   data_res.Recordset("promcel") = 0
                Else
                   Xhh1 = Xrojoama / Xtoama * 100
                   data_res.Recordset("promcel") = Xhh1
                End If
                If Xrojover <= 0 Then
                   data_res.Recordset("cantc5m") = 0
                Else
                   Xhh1 = Xrojover / Xtoverde * 100
                   data_res.Recordset("cantc5m") = Xhh1
                End If
                data_res.Recordset("comodo") = Xtocerti
                If Xtorojos > 0 Then
                   Xhh1 = Xtotcanmax2 / Xtorojos * 100
                   Xhh1 = 100 - Xhh1
                Else
                   data_res.Recordset("anoarq") = 0
                End If
                data_res.Recordset("anoarq") = Xhh1
                
                Xveonum1 = Xtorojos / Xtogra
                Xveonum1 = Xveonum1 * 100
                data_res.Recordset("totevang") = XtotrealR
                data_res.Recordset("totcudam") = XtotrealA
                data_res.Recordset("totsmi") = XtotrealV
                data_res.Recordset("totacom") = Xtot3tot
                data_res.Recordset("totodon") = Xtot4atot
                data_res.Recordset("totdeudas") = Xtot5atot
                data_res.Recordset("celemas5") = Xtot6atot
                data_res.Recordset("totceles") = Xtocele
                If XtotrealR > 0 Then
                   Xporrear = XtotrealR / Xtorojos
                Else
                   Xporrear = 0
                End If
                If XtotrealA > 0 Then
                   Xporreaa = XtotrealA / Xtoama
                Else
                   Xporreaa = 0
                End If
                If XtotrealV > 0 Then
                   Xtoverde = Xtoverde + Xtocele
                   Xporreav = XtotrealV / Xtoverde
                   Xtoverde = Xtoverde - Xtocele
                Else
                   Xporreav = 0
                End If
                If XtotrealC > 0 Then
                   If Xtocele > 0 Then
                      Xporreac = XtotrealC / Xtocele
                   Else
                      Xporreac = 0
                   End If
                Else
                   Xporreac = 0
                End If
                
                data_res.Recordset("totimpasa") = Xporrear * 100
                data_res.Recordset("totcgali") = Xporreaa * 100
                data_res.Recordset("totcima") = Xporreav * 100
                data_res.Recordset("totcel") = Xporreac * 100
                                       
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtoama / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc2") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc2") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtoverde / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("desr") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desr") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtocerti / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("des21") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("des21") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
' Celestes
                Xveonum1 = Xtocele / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("descele1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("descele1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                
                
    ' Fin de TOTALES
'                    data_res.Recordset("mes") = XtotminR / Xtotcan
                data_res.Recordset("ano") = 0
'                    data_res.Recordset("anoarq") = Xtotcanmax
                If Xtotcanmax > 0 Then
                   data_res.Recordset("mesarq") = Xtotcan / Xtotcanmax
                Else
                   data_res.Recordset("mesarq") = 0
                End If
'                    data_res.Recordset("comr") = XtotlleR / Xtotcan
                data_res.Recordset("coma") = 0
'                    data_res.Recordset("comm") = Xtotcanmax2
                If Xtotcanmax2 <= 0 Then
                   data_res.Recordset("comc") = 0
                Else
                   data_res.Recordset("comc") = Xtotcan / Xtotcanmax2
                End If
                Xtotcanmax = 0
                Xtotcanmax2 = 0
                Xtotcan = 0
                XtotminR = 0
                XtotlleR = 0
                data_res.Recordset.Update
                data_res.Refresh
        
             Else
                data_res.Recordset.AddNew
    ' Totales de llamados por CLAVE
                data_res.Recordset("mes") = Xtotcanmax2
                data_res.Recordset("cob") = Xtogra
                data_res.Recordset("quesob") = Xtorojos
                data_res.Recordset("totuniv") = Xtoama
                data_res.Recordset("totccou") = Xtoverde
                data_res.Recordset("comodo") = Xtocerti
                If Xtorojos > 0 Then
                   Xhh1 = Xtotcanmax2 / Xtorojos * 100
                   Xhh1 = 100 - Xhh1
                Else
                   data_res.Recordset("mes") = 0
                   data_res.Recordset("totimpasa") = 0
                   data_res.Recordset("totacom") = 0
                End If
                data_res.Recordset("anoarq") = Xhh1
                If Xtorojos > 0 Then
                Else
                   data_res.Recordset("anoarq") = 0
                End If
                
                Xveonum1 = Xtorojos / Xtogra
                Xveonum1 = Xveonum1 * 100
                data_res.Recordset("totevang") = XtotrealR
                data_res.Recordset("totcudam") = XtotrealA
                data_res.Recordset("totsmi") = XtotrealV
                data_res.Recordset("totacom") = Xtot3tot
                data_res.Recordset("totodon") = Xtot4atot
                data_res.Recordset("totdeudas") = Xtot5atot
                data_res.Recordset("celemas5") = Xtot6atot
                data_res.Recordset("totceles") = Xtocele
                If XtotrealR > 0 Then
                   Xporrear = XtotrealR / Xtorojos
                Else
                   Xporrear = 0
                End If
                If XtotrealA > 0 Then
                   Xporreaa = XtotrealA / Xtoama
                Else
                   Xporreaa = 0
                End If
                If XtotrealV > 0 Then
                   Xtoverde = Xtoverde + Xtocele
                   Xporreav = XtotrealV / Xtoverde
                   Xtoverde = Xtoverde - Xtocele
                Else
                   Xporreav = 0
                End If
                If XtotrealC > 0 Then
                   If Xtocele > 0 Then
                      Xporreac = XtotrealC / Xtocele
                   Else
                      Xporreac = 0
                   End If
                Else
                   Xporreac = 0
                End If
                
                data_res.Recordset("totimpasa") = Xporrear * 100
                data_res.Recordset("totcgali") = Xporreaa * 100
                data_res.Recordset("totcima") = Xporreav * 100
                data_res.Recordset("totcel") = Xporreac * 100
                                       
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtoama / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc2") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc2") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtoverde / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("desr") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desr") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtocerti / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("des21") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("des21") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
' Celestes
                Xveonum1 = Xtocele / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("descele1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("descele1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                
                
    ' Fin de TOTALES
'                    data_res.Recordset("mes") = XtotminR / Xtotcan
                data_res.Recordset("ano") = 0
'                    data_res.Recordset("anoarq") = Xtotcanmax
                If Xtotcanmax > 0 Then
                   data_res.Recordset("mesarq") = Xtotcan / Xtotcanmax
                Else
                   data_res.Recordset("mesarq") = 0
                End If
'                    data_res.Recordset("comr") = XtotlleR / Xtotcan
                data_res.Recordset("coma") = 0
'                    data_res.Recordset("comm") = Xtotcanmax2
                If Xtotcanmax2 <= 0 Then
                   data_res.Recordset("comc") = 0
                Else
                   data_res.Recordset("comc") = Xtotcan / Xtotcanmax2
                End If
                Xtotcanmax = 0
                Xtotcanmax2 = 0
                Xtotcan = 0
                XtotminR = 0
                XtotlleR = 0
                data_res.Recordset.Update
                data_res.Refresh
             
             
             End If
             
             Xtexto = "A"
             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
             data_info.Refresh
             If data_info.Recordset.RecordCount > 0 Then
                data_info.Recordset.MoveFirst
                Do While Not data_info.Recordset.EOF
                   XtotminR = XtotminR + data_info.Recordset("tmm")
                   XtotlleR = XtotlleR + data_info.Recordset("thh")
                   Xtotcan = Xtotcan + 1
                   If data_info.Recordset("tmm") > 5 Then
                      Xtotcanmax = Xtotcanmax + 1
                   End If
                   If data_info.Recordset("thh") > 30 Then
                      Xtotcanmax2 = Xtotcanmax2 + 1
                   End If
                   data_info.Recordset.MoveNext
                Loop
                If data_res.Recordset.RecordCount > 0 Then
                   data_res.Recordset.Edit
                Else
                   data_res.Recordset.AddNew
                End If
                data_res.Recordset("comr") = Xtotcanmax2
'                    data_res.Recordset("totrec") = XtotminR / Xtotcan
                data_res.Recordset("totimp") = 0
'                    data_res.Recordset("comv") = Xtotcanmax
                If Xtotcanmax > 0 Then
                   data_res.Recordset("totimpu") = Xtotcan / Xtotcanmax
                Else
                   data_res.Recordset("totimpu") = 0
                End If
'                    data_res.Recordset("totrecu") = XtotlleR / Xtotcan
                data_res.Recordset("iva1") = 0
                If Xtoama > 0 Then
                   Xhh1 = Xtotcanmax2 / Xtoama * 100
                   Xhh1 = 100 - Xhh1
                Else
                   Xhh1 = 0
                End If
                data_res.Recordset("comm") = Xhh1

'                    data_res.Recordset("iva2") = Xtotcanmax2
                If Xtotcanmax2 <= 0 Then
                   data_res.Recordset("ivatot") = 0
                Else
                   data_res.Recordset("ivatot") = Xtotcan / Xtotcanmax2
                End If
                Xtotcanmax = 0
                Xtotcanmax2 = 0
                Xtotcan = 0
                XtotminR = 0
                XtotlleR = 0
                data_res.Recordset.Update
                data_res.Refresh
             End If
             
             Xtexto = "V"
             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
'                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "' or codmot ='" & Trim("C") & "'"
             data_info.Refresh
             Dim Xtotmas3, Xtotmas30en, xhh, xmm, Xhhh, Xmmh, xdemh, xdemm As Long
             Dim Xporcenver As Double
             If data_info.Recordset.RecordCount > 0 Then
                data_info.Recordset.MoveFirst
                Do While Not data_info.Recordset.EOF
                   XtotminR = XtotminR + data_info.Recordset("tmm")
                   XtotlleR = XtotlleR + data_info.Recordset("thh")
                   Xtotcan = Xtotcan + 1
                   If data_info.Recordset("thh") > 120 Then
                      Xtotcanmax = Xtotcanmax + 1
                   Else
                      Xtomeno2 = Xtomeno2 + 1
                   End If
                   If data_info.Recordset("tmm") > 180 Then
                      Xtotmas3 = Xtotmas3 + 1
                   End If
                   If data_info.Recordset("thh") > 180 Then
                      Xtotcanmax2 = Xtotcanmax2 + 1
                   End If
                   If IsNull(data_info.Recordset("hor_llega")) = True Then
                   Else
                      xhh = Val(Mid(data_info.Recordset("hor_llega"), 1, 2))
                      xmm = Val(Mid(data_info.Recordset("hor_llega"), 4, 2))
                      If IsNull(data_info.Recordset("hor_rea")) = False Then
                         Xhhh = Val(Mid(data_info.Recordset("hor_rea"), 1, 2))
                         Xmmh = Val(Mid(data_info.Recordset("hor_rea"), 4, 2))
                      End If
                      xdemh = Xhhh - xhh
                      xdemm = Xmmh - xmm
                      If data_info.Recordset("fecha") < data_info.Recordset("fec_llega") Then
                         If xdemh < 0 Then
                            xdemh = Xhhh - xhh
                            xdemh = xdemh + 24
                         End If
                      Else
                         If IsNull(data_info.Recordset("fec_llega")) = True Then
                            xdemh = Xhhh - xhh
                            xdemh = xdemh + 24
                         Else
                            If xdemh < 0 Then
                               xdemh = xdemh + 24
                            End If
                         End If
                      End If
                      If xdemh > 0 Then
                         If xdemm < 0 Then
                            xdemm = xdemm + 60
                            xdemh = xdemh - 1
                         End If
                      Else
                         If xdemm < 0 Then
                            xdemm = xdemm + 60
                         End If
                      End If
                      data_info.Recordset.Edit
                      If xdemh > 9 Then
                         If xdemm > 9 Then
                            data_info.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                         Else
                            data_info.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                         End If
                      Else
                         If xdemm > 9 Then
                            If xdemh < 0 Then
                               xdemh = 0
                            End If
                            data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                         Else
                            If xdemh < 0 Then
                               xdemh = 0
                            End If
                            data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                         End If
                      End If
                      data_info.Recordset.Update
                   End If
                   data_info.Recordset.MoveNext
                Loop
                data_res.Recordset.Edit
                data_res.Recordset("totrec") = Xtotcanmax2
                data_res.Recordset("comuni") = XtotminR / Xtotcan
                data_res.Recordset("comccou") = 0
                data_res.Recordset("comimp") = Xtotcanmax
                data_res.Recordset("comacom") = Xtotmas3
'                    Xhh1 = Xtotmas3 / Xtoverde * 100
                Xporcenver = Xtotcanmax2 / Xtoverde * 100
'                    Xhh1 = 100 - Xhh1
                Xporcenver = 100 - Xporcenver
                
'                    data_res.Recordset("comv") = Xhh1
                data_res.Recordset("comv") = Xporcenver
                
                If Xtotcanmax <= 0 Then
                   data_res.Recordset("comgal") = 0
                Else
                   data_res.Recordset("comgal") = Xtotcan / Xtotcanmax
                End If
                data_res.Recordset("comcima") = XtotlleR / Xtotcan
                data_res.Recordset("comeva") = 0
                data_res.Recordset("comcud") = Xtotcanmax2
                If Xtotcanmax2 <= 0 Then
                   data_res.Recordset("comsmi") = 0
                Else
                   data_res.Recordset("comsmi") = Xtotcan / Xtotcanmax2
                End If
                data_res.Recordset("comdeu") = Xtomeno2 / Xtoverde * 100
                data_res.Recordset.Update
                data_res.Refresh
             End If
             
             Xtexto = "C"
             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
'                 data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "' or codmot ='" & Trim("C") & "'"
             data_info.Refresh
             Xtotcanmax = 0
             Xtomeno2 = 0
             Xtotcanmax2 = 0
'                 Dim Xtotmas3, Xtotmas30en, xhh, xmm, Xhhh, Xmmh, xdemh, xdemm As Long
             If data_info.Recordset.RecordCount > 0 Then
                data_info.Recordset.MoveFirst
                Do While Not data_info.Recordset.EOF
                   XtotminR = XtotminR + data_info.Recordset("tmm")
                   XtotlleR = XtotlleR + data_info.Recordset("thh")
                   Xtotcan = Xtotcan + 1
                   If data_info.Recordset("tmm") > 5 Then
                      Xtotcanmax = Xtotcanmax + 1
                   Else
                      Xtomeno2 = Xtomeno2 + 1
                   End If
                   If data_info.Recordset("thh") > 30 Then
                      Xtotcanmax2 = Xtotcanmax2 + 1
                   End If
                   If IsNull(data_info.Recordset("hor_llega")) = True Then
                   Else
                      xhh = Val(Mid(data_info.Recordset("hor_llega"), 1, 2))
                      xmm = Val(Mid(data_info.Recordset("hor_llega"), 4, 2))
                      If IsNull(data_info.Recordset("hor_rea")) = False Then
                         Xhhh = Val(Mid(data_info.Recordset("hor_rea"), 1, 2))
                         Xmmh = Val(Mid(data_info.Recordset("hor_rea"), 4, 2))
                      End If
                      xdemh = Xhhh - xhh
                      xdemm = Xmmh - xmm
                      If data_info.Recordset("fecha") < data_info.Recordset("fec_llega") Then
                         If xdemh < 0 Then
                            xdemh = Xhhh - xhh
                            xdemh = xdemh + 24
                         End If
                      Else
                         If IsNull(data_info.Recordset("fec_llega")) = True Then
                            xdemh = Xhhh - xhh
                            xdemh = xdemh + 24
                         Else
                            If xdemh < 0 Then
                               xdemh = xdemh + 24
                            End If
                         End If
                      End If
                      If xdemh > 0 Then
                         If xdemm < 0 Then
                            xdemm = xdemm + 60
                            xdemh = xdemh - 1
                         End If
                      Else
                         If xdemm < 0 Then
                            xdemm = xdemm + 60
                         End If
                      End If
                      data_info.Recordset.Edit
                      If xdemh > 9 Then
                         If xdemm > 9 Then
                            data_info.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                         Else
                            data_info.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                         End If
                      Else
                         If xdemm > 9 Then
                            If xdemh < 0 Then
                               xdemh = 0
                            End If
                            data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                         Else
                            If xdemh < 0 Then
                               xdemh = 0
                            End If
                            data_info.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                         End If
                      End If
                      data_info.Recordset.Update
                   End If
                   data_info.Recordset.MoveNext
                Loop
                data_res.Recordset.Edit
                data_res.Recordset("promcel") = XtotminR / Xtotcan ' promedio
                Xhh1 = Xtotcanmax2 / Xtocele * 100
                Xhh1 = 100 - Xhh1
                data_res.Recordset("iva2") = Xhh1
'                    data_res.Recordset("comccou") = 0
                data_res.Recordset("cantc5m") = Xtotcanmax ' cantidad que pasaron los 5min de salida
                If Xtotcanmax <= 0 Then
                   data_res.Recordset("promcmas5") = 0
                Else
                   data_res.Recordset("promcmas5") = Xtotcan / Xtotcanmax
                End If
                data_res.Recordset("promcmas30") = XtotlleR / Xtotcan ' Promedio de la llegada
                data_res.Recordset("cantc30m") = Xtotcanmax2 ' Cantidad que pasaron los 30m en llegada
                If Xtotcanmax2 <= 0 Then
                   data_res.Recordset("promcmas302") = 0
                Else
                   data_res.Recordset("promcmas302") = Xtotcan / Xtotcanmax2 'Porcentaje de desvío en llegada mas 30m
                End If
                data_res.Recordset("totrecu") = Xtotcanmax2
                data_res.Recordset.Update
                data_res.Refresh
             End If
             
             If Check5.Value = 1 Then
                Command5_Click
             End If
             frm_calidadiso.MousePointer = 0
             MsgBox "Proceso terminado"
             If Option1.Value = True Then
                CrystalReport1.ReportFileName = App.path & "\infcaldes2n.rpt"
                CrystalReport1.DiscardSavedData = True
                CrystalReport1.ReportTitle = "INFORME DEMORAS DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
             Else
                CrystalReport1.ReportFileName = App.path & "\infcaldes2.rpt"
                CrystalReport1.DiscardSavedData = True
                CrystalReport1.ReportTitle = "INFORME DEMORAS DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
             End If
             CrystalReport1.Action = 1
          End If
            
End Sub

Private Sub Command4_Click()
Dim Xhh88, Xmm88, Xhh89, Xmm89, Xtmm88, Xths88, Xtotlle88 As Integer
Dim Xhh98, Xmm98, Xhh99, Xmm99, Xtmm98, Xths98, Xtotlle98 As Integer
Dim XtotrecepR, XtotrecepAC, XtotdespaR, XtotdespaA, XtotdespaC As Double
Dim Xdifclaves As Double
Xdifclaves = 0

XtotrecepR = 0
XtotdespaR = 0
XtotrecepAC = 0
XtotdespaA = 0
XtotdespaC = 0
Dim Xtotceles As Long
Xtotceles = 0

          If data_llam.Recordset.RecordCount > 0 Then
             data_llam.Recordset.MoveFirst
             Do While Not data_llam.Recordset.EOF
                If data_llam.Recordset("codzon") = 4 Then
                Else
                    If data_llam.Recordset("codmot") = "R" Then
                       If IsNull(data_llam.Recordset("colormot")) = False Then
                          If data_llam.Recordset("colormot") = "R" Then
                             XtotrealR = XtotrealR + 1
                          End If
                       End If
' Tiempos entre recepción y Clasificación
                       
                       If IsNull(data_llam.Recordset("activo")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotminR = Xtmm
                       Else
                          XtotminR = 0
                       End If
' Tiempos entre Clasificación y Asignación
                       If IsNull(data_llam.Recordset("horpas")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("activo")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                          If Xths = 5 Then
                             Xtmm = Xtmm + 240
                          End If
                          If Xths = 6 Then
                             Xtmm = Xtmm + 300
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotlleR = Xtmm
                       Else
                          XtotlleR = 0
                       End If
''' DEMORA EN LLEGADA (ENTRE QUE LLAMO Y LLEGO
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh88 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm88 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh89 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm89 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh88 = Xhh89 Then
                          Xtmm88 = Xmm88 - Xmm89
                       Else
                          Xths88 = Xhh88 - Xhh89
                          Xtmm88 = Xmm88 - Xmm89 + 60
                          If Xths88 = 2 Then
                             Xtmm88 = Xtmm88 + 60
                          End If
                          If Xths88 = 3 Then
                             Xtmm88 = Xtmm88 + 120
                          End If
                          If Xths88 = 4 Then
                             Xtmm88 = Xtmm88 + 180
                          End If
                          If Xths88 = 5 Then
                             Xtmm88 = Xtmm88 + 240
                          End If
                          If Xths88 = 6 Then
                             Xtmm88 = Xtmm88 + 300
                          End If
                       
                       End If
                       If Xtmm88 >= 0 Then
                          Xtotlle88 = Xtmm88
                       Else
                          Xtotlle88 = 0
                       End If
''' demora en llegada (entre que salió y llegó)
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh98 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm98 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("horsali")) = False Then
                          Xhh99 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                          Xmm99 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                       End If
                       If Xhh98 = Xhh99 Then
                          Xtmm98 = Xmm98 - Xmm99
                       Else
                          Xths98 = Xhh98 - Xhh99
                          Xtmm98 = Xmm98 - Xmm99 + 60
                          If Xths98 = 2 Then
                             Xtmm98 = Xtmm98 + 60
                          End If
                          If Xths98 = 3 Then
                             Xtmm98 = Xtmm98 + 120
                          End If
                          If Xths98 = 4 Then
                             Xtmm98 = Xtmm98 + 180
                          End If
                          If Xths98 = 5 Then
                             Xtmm98 = Xtmm98 + 240
                          End If
                          If Xths98 = 6 Then
                             Xtmm98 = Xtmm98 + 300
                          End If
                       
                       End If
                       If Xtmm98 >= 0 Then
                          Xtotlle98 = Xtmm98
                       Else
                          Xtotlle98 = 0
                       End If
                       data_info.Recordset.AddNew
                       data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                       data_info.Recordset("hora") = data_llam.Recordset("hora")
                       data_info.Recordset("matric") = data_llam.Recordset("matric")
                       data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                       data_info.Recordset("categ") = data_llam.Recordset("categ")
                       data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                       data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                       data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                       data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                       data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                       data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                       data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                       data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                       data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                       data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                       data_info.Recordset("activo") = data_llam.Recordset("activo")
                       If IsNull(data_llam.Recordset("cancela")) = False Then
                          If data_llam.Recordset("cancela") = 1 Then
                             data_info.Recordset("cancela") = 1
                             XtotrecepR = XtotrecepR - XtotminR
                             XtotdespaR = XtotdespaR - XtotlleR
                             Xtorojos = Xtorojos - 1
                             If IsNull(data_llam.Recordset("colormot")) = False Then
                                If data_llam.Recordset("colormot") = "R" Then
                                   XtotrealR = XtotrealR - 1
                                End If
                             End If
                          Else
                             data_info.Recordset("cancela") = 9
                          End If
                       Else
                          data_info.Recordset("cancela") = 9
                       End If
                       If Check3.Value = 1 Then
                          data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                       Else
                          If Check4.Value = 1 Then
                             data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                          End If
                       End If
                       data_info.Recordset("tmm") = XtotminR
                       data_info.Recordset("thh") = XtotlleR
                       XtotrecepR = XtotrecepR + XtotminR
                       XtotdespaR = XtotdespaR + XtotlleR
                       
                       data_info.Recordset("mes") = Xtotlle88
                       data_info.Recordset("ano") = Xtotlle98
                       
                       data_info.Recordset.Update
                       Xtorojos = Xtorojos + 1
                    End If
                    If data_llam.Recordset("codmot") = "A" Then
                       If IsNull(data_llam.Recordset("colormot")) = False Then
                          If data_llam.Recordset("colormot") = "A" Then
                             XtotrealA = XtotrealA + 1
                          End If
                          If data_llam.Recordset("colormot") = "R" Then
                             Xrojoama = Xrojoama + 1
                          End If
                       End If
                       
                       If IsNull(data_llam.Recordset("activo")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotminA = Xtmm
                       Else
                          XtotminA = 0
                       End If
                       
                       If IsNull(data_llam.Recordset("horpas")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("activo")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotlleA = Xtmm
                       Else
                          XtotlleA = 0
                       End If
''' DEMORA EN LLEGADA (ENTRE QUE LLAMO Y LLEGO
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh88 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm88 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh89 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm89 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh88 = Xhh89 Then
                          Xtmm88 = Xmm88 - Xmm89
                       Else
                          Xths88 = Xhh88 - Xhh89
                          Xtmm88 = Xmm88 - Xmm89 + 60
                          If Xths88 = 2 Then
                             Xtmm88 = Xtmm88 + 60
                          End If
                          If Xths88 = 3 Then
                             Xtmm88 = Xtmm88 + 120
                          End If
                          If Xths88 = 4 Then
                             Xtmm88 = Xtmm88 + 180
                          End If
                       End If
                       If Xtmm88 >= 0 Then
                          Xtotlle88 = Xtmm88
                       Else
                          Xtotlle88 = 0
                       End If
''' demora en llegada (entre que salió y llegó)
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh98 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm98 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("horsali")) = False Then
                          Xhh99 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                          Xmm99 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                       End If
                       If Xhh98 = Xhh99 Then
                          Xtmm98 = Xmm98 - Xmm99
                       Else
                          Xths98 = Xhh98 - Xhh99
                          Xtmm98 = Xmm98 - Xmm99 + 60
                          If Xths98 = 2 Then
                             Xtmm98 = Xtmm98 + 60
                          End If
                          If Xths98 = 3 Then
                             Xtmm98 = Xtmm98 + 120
                          End If
                          If Xths98 = 4 Then
                             Xtmm98 = Xtmm98 + 180
                          End If
                       End If
                       If Xtmm98 >= 0 Then
                          Xtotlle98 = Xtmm98
                       Else
                          Xtotlle98 = 0
                       End If
                       data_info.Recordset.AddNew
                       data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                       data_info.Recordset("hora") = data_llam.Recordset("hora")
                       data_info.Recordset("matric") = data_llam.Recordset("matric")
                       data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                       data_info.Recordset("categ") = data_llam.Recordset("categ")
                       data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                       data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                       data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                       data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                       data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                       data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                       data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                       data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                       data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                       data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                       data_info.Recordset("activo") = data_llam.Recordset("activo")
                       If IsNull(data_llam.Recordset("cancela")) = False Then
                          If data_llam.Recordset("cancela") = 1 Then
                             data_info.Recordset("cancela") = 1
                             XtotrecepAC = XtotrecepAC - XtotminA
                             XtotdespaA = XtotdespaA - XtotlleA
                             
                             Xtoama = Xtoama - 1
                             If IsNull(data_llam.Recordset("colormot")) = False Then
                                If data_llam.Recordset("colormot") = "A" Then
                                   XtotrealA = XtotrealA - 1
                                End If
                                If data_llam.Recordset("colormot") = "R" Then
                                   Xrojoama = Xrojoama - 1
                                End If
                             End If
                          Else
                             data_info.Recordset("cancela") = 9
                          End If
                       Else
                          data_info.Recordset("cancela") = 9
                       End If
                       If Check3.Value = 1 Then
                          data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                       Else
                          If Check4.Value = 1 Then
                             data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                          End If
                       End If
                       data_info.Recordset("tmm") = XtotminA
                       data_info.Recordset("thh") = XtotlleA
                       XtotrecepAC = XtotrecepAC + XtotminA
                       XtotdespaA = XtotdespaA + XtotlleA
                       
                       data_info.Recordset("mes") = Xtotlle88
                       data_info.Recordset("ano") = Xtotlle98
                       data_info.Recordset.Update
                       Xtoama = Xtoama + 1
                    End If
                    If data_llam.Recordset("codmot") = "C" Then
                       If IsNull(data_llam.Recordset("colormot")) = False Then
                          If data_llam.Recordset("colormot") = "C" Then
                             XtotrealA = XtotrealA + 1
                          End If
                       End If
                       
                       If IsNull(data_llam.Recordset("activo")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotminA = Xtmm
                       Else
                          XtotminA = 0
                       End If
                       
                       If IsNull(data_llam.Recordset("horpas")) = False Then
                          Xhh1 = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                          Xmm1 = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("activo")) = False Then
                          Xhh2 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                          Xmm2 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                       End If
                       If Xhh1 = Xhh2 Then
                          Xtmm = Xmm1 - Xmm2
                       Else
                          Xths = Xhh1 - Xhh2
                          Xtmm = Xmm1 - Xmm2 + 60
                          If Xths = 2 Then
                             Xtmm = Xtmm + 60
                          End If
                          If Xths = 3 Then
                             Xtmm = Xtmm + 120
                          End If
                          If Xths = 4 Then
                             Xtmm = Xtmm + 180
                          End If
                       End If
                       If Xtmm >= 0 Then
                          XtotlleA = Xtmm
                       Else
                          XtotlleA = 0
                       End If
''' DEMORA EN LLEGADA (ENTRE QUE LLAMO Y LLEGO
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh88 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm88 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("hora")) = False Then
                          Xhh89 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                          Xmm89 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                       End If
                       If Xhh88 = Xhh89 Then
                          Xtmm88 = Xmm88 - Xmm89
                       Else
                          Xths88 = Xhh88 - Xhh89
                          Xtmm88 = Xmm88 - Xmm89 + 60
                          If Xths88 = 2 Then
                             Xtmm88 = Xtmm88 + 60
                          End If
                          If Xths88 = 3 Then
                             Xtmm88 = Xtmm88 + 120
                          End If
                          If Xths88 = 4 Then
                             Xtmm88 = Xtmm88 + 180
                          End If
                       End If
                       If Xtmm88 >= 0 Then
                          Xtotlle88 = Xtmm88
                       Else
                          Xtotlle88 = 0
                       End If
''' demora en llegada (entre que salió y llegó)
                       If IsNull(data_llam.Recordset("hor_llega")) = False Then
                          Xhh98 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                          Xmm98 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_llam.Recordset("horsali")) = False Then
                          Xhh99 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                          Xmm99 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                       End If
                       If Xhh98 = Xhh99 Then
                          Xtmm98 = Xmm98 - Xmm99
                       Else
                          Xths98 = Xhh98 - Xhh99
                          Xtmm98 = Xmm98 - Xmm99 + 60
                          If Xths98 = 2 Then
                             Xtmm98 = Xtmm98 + 60
                          End If
                          If Xths98 = 3 Then
                             Xtmm98 = Xtmm98 + 120
                          End If
                          If Xths98 = 4 Then
                             Xtmm98 = Xtmm98 + 180
                          End If
                       End If
                       If Xtmm98 >= 0 Then
                          Xtotlle98 = Xtmm98
                       Else
                          Xtotlle98 = 0
                       End If
                       data_info.Recordset.AddNew
                       data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                       data_info.Recordset("hora") = data_llam.Recordset("hora")
                       data_info.Recordset("matric") = data_llam.Recordset("matric")
                       data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                       data_info.Recordset("categ") = data_llam.Recordset("categ")
                       data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                       data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                       data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                       data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                       data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                       data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                       data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                       data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                       data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                       data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                       data_info.Recordset("activo") = data_llam.Recordset("activo")
                       If IsNull(data_llam.Recordset("cancela")) = False Then
                          If data_llam.Recordset("cancela") = 1 Then
                             data_info.Recordset("cancela") = 1
                             XtotrecepAC = XtotrecepAC - XtotminA
                             XtotdespaC = XtotdespaC - XtotlleA
                             Xtotceles = Xtotceles - 1
                             Xtoama = Xtoama - 1
                             If data_llam.Recordset("colormot") = "C" Then
                                XtotrealA = XtotrealA - 1
                             End If
                          Else
                             data_info.Recordset("cancela") = 9
                          End If
                       Else
                          data_info.Recordset("cancela") = 9
                       End If
                       
                       If Check3.Value = 1 Then
                          data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                       Else
                          If Check4.Value = 1 Then
                             data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                          End If
                       End If
                       
                       data_info.Recordset("tmm") = XtotminA
                       data_info.Recordset("thh") = XtotlleA
                       XtotrecepAC = XtotrecepAC + XtotminA
                       XtotdespaC = XtotdespaC + XtotlleA
                       
                       data_info.Recordset("mes") = Xtotlle88
                       data_info.Recordset("ano") = Xtotlle98
                       data_info.Recordset.Update
                       Xtoama = Xtoama + 1
                       Xtotceles = Xtotceles + 1
                    End If
                    
                    If data_llam.Recordset("codmot") = "V" Then
                       If data_llam.Recordset("categ") = "UDEMM" Or _
                          data_llam.Recordset("categ") = "CERADU" Or _
                          data_llam.Recordset("categ") = "CERANT" Or _
                          data_llam.Recordset("categ") = "CERCAS" Or _
                          data_llam.Recordset("categ") = "CERDGI" Or _
                          data_llam.Recordset("categ") = "CERESS" Or _
                          data_llam.Recordset("categ") = "CERSAP" Or _
                          data_llam.Recordset("categ") = "CERHEV" Or _
                          data_llam.Recordset("categ") = "CERIMP" Or _
                          data_llam.Recordset("categ") = "CERKEV" Or _
                          data_llam.Recordset("categ") = "CERMAT" Or _
                          data_llam.Recordset("categ") = "CERSEM" Or _
                          data_llam.Recordset("categ") = "CERSEV" Or _
                          data_llam.Recordset("categ") = "CERVIS" Then
                          Xtocerti = Xtocerti + 1
                       Else
                          If IsNull(data_llam.Recordset("colormot")) = False Then
                             If data_llam.Recordset("colormot") = "V" Then
                                XtotrealV = XtotrealV + 1
                             End If
                             If data_llam.Recordset("colormot") = "R" Then
                                Xrojover = Xrojover + 1
                             End If
                          End If
                          
'                          If Mid(data_llam.Recordset("categ"), 1, 4) <> "TALA" Then
'                             Xtoverde = Xtoverde + 1
'                          Else
                             If IsNull(data_llam.Recordset("activo")) = False Then
                                Xhh1 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                                Xmm1 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                             End If
                             If IsNull(data_llam.Recordset("hora")) = False Then
                                Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                             End If
                             If Xhh1 = Xhh2 Then
                                Xtmm = Xmm1 - Xmm2
                             Else
                                Xths = Xhh1 - Xhh2
                                Xtmm = Xmm1 - Xmm2 + 60
                                If Xths = 2 Then
                                   Xtmm = Xtmm + 60
                                End If
                                If Xths = 3 Then
                                   Xtmm = Xtmm + 120
                                End If
                                If Xths = 4 Then
                                   Xtmm = Xtmm + 180
                                End If
                                If Xths = 5 Then
                                   Xtmm = Xtmm + 240
                                End If
                                If Xths = 6 Then
                                   Xtmm = Xtmm + 300
                                End If
                             End If
                             If Xtmm >= 0 Then
                                XtotminV = Xtmm
                             Else
                                XtotminV = 0
                             End If
              '' Llegada!!!
                             If IsNull(data_llam.Recordset("horpas")) = False Then
                                Xhh1 = Val(Mid(data_llam.Recordset("horpas"), 1, 2))
                                Xmm1 = Val(Mid(data_llam.Recordset("horpas"), 4, 2))
                             End If
                             If IsNull(data_llam.Recordset("activo")) = False Then
                                Xhh2 = Val(Mid(data_llam.Recordset("activo"), 1, 2))
                                Xmm2 = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                             End If
                             If Xhh1 = Xhh2 Then
                                Xtmm = Xmm1 - Xmm2
                             Else
                                Xths = Xhh1 - Xhh2
                                Xtmm = Xmm1 - Xmm2 + 60
                                If Xths = 2 Then
                                   Xtmm = Xtmm + 60
                                End If
                                If Xths = 3 Then
                                   Xtmm = Xtmm + 120
                                End If
                                If Xths = 4 Then
                                   Xtmm = Xtmm + 180
                                End If
                                If Xths = 5 Then
                                   Xtmm = Xtmm + 240
                                End If
                                If Xths = 6 Then
                                   Xtmm = Xtmm + 300
                                End If
                             End If
                             If Xtmm >= 0 Then
                                XtotlleV = Xtmm
                             Else
                                XtotlleV = 0
                             End If
                             
''' DEMORA EN LLEGADA (ENTRE QUE LLAMO Y LLEGO
                               If IsNull(data_llam.Recordset("hor_llega")) = False Then
                                  Xhh88 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                                  Xmm88 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                               End If
                               If IsNull(data_llam.Recordset("hora")) = False Then
                                  Xhh89 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
                                  Xmm89 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                               End If
                               If Xhh88 = Xhh89 Then
                                  Xtmm88 = Xmm88 - Xmm89
                               Else
                                  Xths88 = Xhh88 - Xhh89
                                  Xtmm88 = Xmm88 - Xmm89 + 60
                                  If Xths88 = 2 Then
                                     Xtmm88 = Xtmm88 + 60
                                  End If
                                  If Xths88 = 3 Then
                                     Xtmm88 = Xtmm88 + 120
                                  End If
                                  If Xths88 = 4 Then
                                     Xtmm88 = Xtmm88 + 180
                                  End If
                                  If Xths88 = 5 Then
                                     Xtmm88 = Xtmm88 + 240
                                  End If
                                  If Xths88 = 6 Then
                                     Xtmm88 = Xtmm88 + 300
                                  End If
                               
                               End If
                               If Xtmm88 >= 0 Then
                                  Xtotlle88 = Xtmm88
                               Else
                                  Xtotlle88 = 0
                               End If
        ''' demora en llegada (entre que salió y llegó)
                               If IsNull(data_llam.Recordset("hor_llega")) = False Then
                                  Xhh98 = Val(Mid(data_llam.Recordset("hor_llega"), 1, 2))
                                  Xmm98 = Val(Mid(data_llam.Recordset("hor_llega"), 4, 2))
                               End If
                               If IsNull(data_llam.Recordset("horsali")) = False Then
                                  Xhh99 = Val(Mid(data_llam.Recordset("horsali"), 1, 2))
                                  Xmm99 = Val(Mid(data_llam.Recordset("horsali"), 4, 2))
                               End If
                               If Xhh98 = Xhh99 Then
                                  Xtmm98 = Xmm98 - Xmm99
                               Else
                                  Xths98 = Xhh98 - Xhh99
                                  Xtmm98 = Xmm98 - Xmm99 + 60
                                  If Xths98 = 2 Then
                                     Xtmm98 = Xtmm98 + 60
                                  End If
                                  If Xths98 = 3 Then
                                     Xtmm98 = Xtmm98 + 120
                                  End If
                                  If Xths98 = 4 Then
                                     Xtmm98 = Xtmm98 + 180
                                  End If
                                  If Xths98 = 5 Then
                                     Xtmm98 = Xtmm98 + 240
                                  End If
                                  If Xths98 = 6 Then
                                     Xtmm98 = Xtmm98 + 300
                                  End If
                               
                               End If
                               If Xtmm98 >= 0 Then
                                  Xtotlle98 = Xtmm98
                               Else
                                  Xtotlle98 = 0
                               End If
                             
                             data_info.Recordset.AddNew
                             data_info.Recordset("fecha") = data_llam.Recordset("fecha")
                             data_info.Recordset("hora") = data_llam.Recordset("hora")
                             data_info.Recordset("matric") = data_llam.Recordset("matric")
                             data_info.Recordset("nombre") = data_llam.Recordset("nombre")
                             data_info.Recordset("categ") = data_llam.Recordset("categ")
                             data_info.Recordset("codzon") = data_llam.Recordset("codzon")
                             data_info.Recordset("codmot") = data_llam.Recordset("codmot")
                             data_info.Recordset("fecsali") = data_llam.Recordset("fecsali")
                             data_info.Recordset("horsali") = data_llam.Recordset("horsali")
                             data_info.Recordset("fecpas") = data_llam.Recordset("fecpas")
                             data_info.Recordset("horpas") = data_llam.Recordset("horpas")
                             data_info.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
                             data_info.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
                             data_info.Recordset("movilpas") = data_llam.Recordset("movilpas")
                             data_info.Recordset("colormot") = data_llam.Recordset("colormot")
                             data_info.Recordset("activo") = data_llam.Recordset("activo")
                             If Check3.Value = 1 Then
                                data_info.Recordset("timdes") = data_llam.Recordset("usuario")
                             Else
                                If Check4.Value = 1 Then
                                   data_info.Recordset("timdes") = data_llam.Recordset("timdes")
                                End If
                             End If
                            If IsNull(data_llam.Recordset("cancela")) = False Then
                               If data_llam.Recordset("cancela") = 1 Then
                                  data_info.Recordset("cancela") = 1
                                  XtotrecepAC = XtotrecepAC - XtotminV
                                  Xtoverde = Xtoverde - 1
                                  If IsNull(data_llam.Recordset("colormot")) = False Then
                                     If data_llam.Recordset("colormot") = "V" Then
                                        XtotrealV = XtotrealV - 1
                                     End If
                                     If data_llam.Recordset("colormot") = "R" Then
                                        Xrojover = Xrojover - 1
                                     End If
                                  End If
                               
                               Else
                                  data_info.Recordset("cancela") = 9
                               End If
                            Else
                               data_info.Recordset("cancela") = 9
                            End If
                             data_info.Recordset("tmm") = XtotminV
                             data_info.Recordset("thh") = XtotlleV
                             XtotrecepAC = XtotrecepAC + XtotminV
                             data_info.Recordset("mes") = Xtotlle88
                             data_info.Recordset("ano") = Xtotlle98
                             data_info.Recordset.Update
                             Xtoverde = Xtoverde + 1
                             
'                          End If
                       End If
                    End If
                End If
                data_llam.Recordset.MoveNext
                XtotminR = 0
                XtotlleR = 0
                XtotlleA = 0
                XtotminA = 0
                XtotminV = 0
                XtotlleV = 0
                Xtogra = Xtogra + 1
             Loop
             Dim MiBaseact As Database
             Dim Unasesact As Workspace
             Set Unasesact = Workspaces(0)
             Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
            
             Xtexto = "R"
             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
             data_info.Refresh
             
             If data_info.Recordset.RecordCount > 0 Then
                data_info.Recordset.MoveFirst
                Do While Not data_info.Recordset.EOF
                   XtotminR = XtotminR + data_info.Recordset("tmm")
                   XtotlleR = XtotlleR + data_info.Recordset("thh")
                   Xtotcan = Xtotcan + 1
                   If data_info.Recordset("tmm") <= 1 Then
                      Xtotcanmax = Xtotcanmax + 1
                   End If
                   If data_info.Recordset("thh") < 1 Then
                      Xtotcanmax2 = Xtotcanmax2 + 1
                   End If
                   If IsNull(data_info.Recordset("codmot")) = False Then
                      If IsNull(data_info.Recordset("colormot")) = False Then
                         If data_info.Recordset("codmot") = data_info.Recordset("colormot") Then
                            XtotlleA = XtotlleA + 1
                         Else
                         End If
                      End If
                   End If
                   data_info.Recordset.MoveNext
                Loop
                data_res.Recordset.AddNew
    ' Totales de llamados por CLAVE
                data_res.Recordset("cob") = Xtogra
                data_res.Recordset("quesob") = Xtorojos
                data_res.Recordset("totuniv") = Xtoama
                data_res.Recordset("totccou") = Xtoverde
                data_res.Recordset("comodo") = Xtocerti
                data_res.Recordset("totevang") = Xtotcanmax
                data_res.Recordset("totcudam") = Xtotcanmax2
                data_res.Recordset("mes") = XtotlleA
                Dim Xtotvac As Long
                Dim Xtoamacele As Long
                Xtoamacele = Xtoama - Xtotceles
                
                Xtotvac = Xtoverde + Xtoama
                Xveonum1 = Xtorojos - XtotlleA
                If Xtorojos > 0 Then
                   data_res.Recordset("comgal") = XtotrecepR / Xtorojos
                Else
                   data_res.Recordset("comgal") = 0
                End If
                If Xtotvac > 0 Then
                   data_res.Recordset("comcima") = XtotrecepAC / Xtotvac
                Else
                   data_res.Recordset("comcima") = 0
                End If
                If Xtorojos > 0 Then
                   data_res.Recordset("comeva") = XtotdespaR / Xtorojos
                Else
                   data_res.Recordset("comeva") = 0
                End If
                If Xtoamacele > 0 Then
                   data_res.Recordset("comcud") = XtotdespaA / Xtoamacele
                Else
                   data_res.Recordset("comcud") = 0
                End If
                If Xtotceles > 0 Then
                   data_res.Recordset("comsmi") = XtotdespaC / Xtotceles
                Else
                   data_res.Recordset("comsmi") = 0
                End If
                data_res.Recordset("anoarq") = Xveonum1 / Xtorojos * 100
                data_res.Recordset("comm") = Xtotcanmax / Xtorojos * 100
                data_res.Recordset("totimpu") = Xtotcanmax2 / Xtorojos * 100
                
                If Xrojoama <= 0 Then
                   data_res.Recordset("promcel") = 0
                   data_res.Recordset("des38") = "0"
                Else
                   Xdifclaves = Xrojoama / Xtoamacele * 100
                   data_res.Recordset("promcel") = Xdifclaves
                   data_res.Recordset("des38") = Trim(str(Format(Xdifclaves, "Standard")))
                End If
                If Xrojover <= 0 Then
                   data_res.Recordset("cantc5m") = 0
                   data_res.Recordset("des37") = "0"
                Else
                   Xdifclaves = Xrojover / Xtoverde * 100
                   data_res.Recordset("cantc5m") = Xdifclaves
                   data_res.Recordset("des37") = Trim(str(Format(Xdifclaves, "Standard")))
                End If
                
                If XtotrealR > 0 Then
                   Xporrear = XtotrealR / Xtorojos
                Else
                   Xporrear = 0
                End If
                If XtotrealA > 0 Then
                   Xporreaa = XtotrealA / Xtoama
                Else
                   Xporreaa = 0
                End If
                If XtotrealV > 0 Then
                   Xporreav = XtotrealV / Xtoverde
                Else
                   Xporreav = 0
                End If
                
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtoama / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc2") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc2") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtoverde / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("desr") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desr") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = Xtocerti / Xtogra
                Xveonum1 = Xveonum1 * 100
                If Xveonum1 < 1 Then
                   data_res.Recordset("des21") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("des21") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                
    ' Fin de TOTALES
                Xtotcanmax = 0
                Xtotcanmax2 = 0
                Xtotcan = 0
                XtotminR = 0
                XtotlleR = 0
                XtotlleA = 0
                data_res.Recordset.Update
                data_res.Refresh
             Else
                data_res.Recordset.AddNew
    ' Totales de llamados por CLAVE
                data_res.Recordset("cob") = 0
                data_res.Recordset("quesob") = 0
                data_res.Recordset("totuniv") = 0
                data_res.Recordset("totccou") = 0
                data_res.Recordset("comodo") = 0
                data_res.Recordset("totevang") = 0
                data_res.Recordset("totcudam") = 0
                data_res.Recordset("mes") = 0
                data_res.Recordset("anoarq") = 0
                data_res.Recordset("comm") = 0
                data_res.Recordset("totimpu") = 0
                Xveonum1 = 0
                Xveonum1 = 0
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc1") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc1") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = 0
                Xveonum1 = 0
                If Xveonum1 < 1 Then
                   data_res.Recordset("desc2") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desc2") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = 0
                Xveonum1 = 0
                If Xveonum1 < 1 Then
                   data_res.Recordset("desr") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("desr") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                Xveonum1 = 0
                Xveonum1 = 0
                If Xveonum1 < 1 Then
                   data_res.Recordset("des21") = "0" & Trim(str(Format(Xveonum1, "Standard")) & " %")
                Else
                   data_res.Recordset("des21") = Trim(str(Format(Xveonum1, "Standard")) & " %")
                End If
                
    ' Fin de TOTALES
                data_res.Recordset("mes") = 0
                data_res.Recordset("ano") = 0
                data_res.Recordset("mesarq") = 0
                data_res.Recordset("comr") = 0
                data_res.Recordset("coma") = 0
                data_res.Recordset("comm") = 0
                If Xtotcanmax2 <= 0 Then
                   data_res.Recordset("comc") = 0
                Else
                   data_res.Recordset("comc") = 0
                End If
                Xtotcanmax = 0
                Xtotcanmax2 = 0
                Xtotcan = 0
                XtotminR = 0
                XtotlleR = 0
                data_res.Recordset.Update
                data_res.Refresh
             End If
             
             Xtexto = "A"
             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
             data_info.Refresh
             XtotlleA = 0
             If data_info.Recordset.RecordCount > 0 Then
                data_info.Recordset.MoveFirst
                Do While Not data_info.Recordset.EOF
                   XtotminR = XtotminR + data_info.Recordset("tmm")
                   XtotlleR = XtotlleR + data_info.Recordset("thh")
                   Xtotcan = Xtotcan + 1
                   If data_info.Recordset("tmm") <= 1 Then
                      Xtotcanmax = Xtotcanmax + 1
                   End If
                   If data_info.Recordset("thh") <= 10 Then
                      Xtotcanmax2 = Xtotcanmax2 + 1
                   End If
                   If IsNull(data_info.Recordset("codmot")) = False Then
                      If IsNull(data_info.Recordset("colormot")) = False Then
                         If data_info.Recordset("codmot") = data_info.Recordset("colormot") Then
                            XtotlleA = XtotlleA + 1
                         Else
                         End If
                      End If
                   End If
                   
                   data_info.Recordset.MoveNext
                Loop
                If data_res.Recordset.RecordCount > 0 Then
                   data_res.Recordset.Edit
                Else
                   data_res.Recordset.AddNew
                End If
                data_res.Recordset("totrec") = Xtotcanmax
                data_res.Recordset("totimp") = Xtotcanmax2
                data_res.Recordset("ano") = XtotlleA
                Xveonum1 = Xtoama - XtotlleA
                data_res.Recordset("comr") = Xveonum1 / Xtoama * 100
                data_res.Recordset("comc") = Xtotcanmax / Xtoama * 100
                data_res.Recordset("totrecu") = Xtotcanmax2 / Xtoama * 100
                Xtotcanmax = 0
                Xtotcanmax2 = 0
                Xtotcan = 0
                XtotminR = 0
                XtotlleR = 0
                XtotlleA = 0
                data_res.Recordset.Update
                data_res.Refresh
             End If
             
'             Xtexto = "C"
'             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
'             data_info.Refresh
'             If data_info.Recordset.RecordCount > 0 Then
'                data_info.Recordset.MoveFirst
'                Do While Not data_info.Recordset.EOF
'                   XtotminR = XtotminR + data_info.Recordset("tmm")
'                   XtotlleR = XtotlleR + data_info.Recordset("thh")
'                   Xtotcan = Xtotcan + 1
'                   If data_info.Recordset("tmm") <= 1 Then
'                      Xtotcanmax = Xtotcanmax + 1
'                   End If
'                   If data_info.Recordset("thh") <= 10 Then
'                      Xtotcanmax2 = Xtotcanmax2 + 1
'                   End If
'                   If IsNull(data_info.Recordset("codmot")) = False Then
'                      If IsNull(data_info.Recordset("colormot")) = False Then
'                         If data_info.Recordset("codmot") = data_info.Recordset("colormot") Then
'                            XtotlleA = XtotlleA + 1
'                         Else
'                         End If
'                      End If
'                   End If
'
'                   data_info.Recordset.MoveNext
'                Loop
'                If data_res.Recordset.RecordCount > 0 Then
'                   data_res.Recordset.Edit
'                Else
'                   data_res.Recordset.AddNew
'                End If
'                data_res.Recordset("totrec") = Xtotcanmax
'                data_res.Recordset("totimp") = Xtotcanmax2
'                data_res.Recordset("ano") = XtotlleA
'                Xveonum1 = Xtoama - XtotlleA
'                data_res.Recordset("comr") = Xveonum1 / Xtoama * 100
'                data_res.Recordset("comc") = Xtotcanmax / Xtoama * 100
'                data_res.Recordset("totrecu") = Xtotcanmax2 / Xtoama * 100
'                Xtotcanmax = 0
'                Xtotcanmax2 = 0
'                Xtotcan = 0
'                XtotminR = 0
'                XtotlleR = 0
'                XtotlleA = 0
'                data_res.Recordset.Update
'                data_res.Refresh
'             End If
             
             Xtexto = "V"
             data_info.RecordSource = "Select * from inflla where codmot ='" & Trim(Xtexto) & "'"
             data_info.Refresh
             XtotlleA = 0
             If data_info.Recordset.RecordCount > 0 Then
                data_info.Recordset.MoveFirst
                Do While Not data_info.Recordset.EOF
                   XtotminR = XtotminR + data_info.Recordset("tmm")
                   XtotlleR = XtotlleR + data_info.Recordset("thh")
                   Xtotcan = Xtotcan + 1
                   If data_info.Recordset("tmm") <= 1 Then
                      Xtotcanmax = Xtotcanmax + 1
                   Else
                      Xtomeno2 = Xtomeno2 + 1
                   End If
                   If data_info.Recordset("thh") <= 60 Then
                      Xtotcanmax2 = Xtotcanmax2 + 1
                   End If
                   If IsNull(data_info.Recordset("codmot")) = False Then
                      If IsNull(data_info.Recordset("colormot")) = False Then
                         If data_info.Recordset("codmot") = data_info.Recordset("colormot") Then
                            XtotlleA = XtotlleA + 1
                         Else
                         End If
                      End If
                   End If
                   data_info.Recordset.MoveNext
                Loop
                If data_res.Recordset.RecordCount > 0 Then
                   data_res.Recordset.Edit
                Else
                   data_res.Recordset.AddNew
                End If
                Xveonum1 = Xtoverde - XtotlleA
                data_res.Recordset("coma") = Xveonum1 / Xtoverde * 100
                data_res.Recordset("comuni") = Xtotcanmax
                data_res.Recordset("comccou") = Xtotcanmax2
                data_res.Recordset("mesarq") = XtotlleA
                data_res.Recordset("comv") = Xtotcanmax / Xtoverde * 100
                data_res.Recordset("iva1") = Xtotcanmax2 / Xtoverde * 100
                data_res.Recordset.Update
                data_res.Refresh
             End If
             If data_res.Recordset.RecordCount > 0 Then
             Else
                data_res.Recordset.AddNew
                data_res.Recordset("comm") = 0
                data_res.Recordset("comc") = 0
                data_res.Recordset("comv") = 0
                data_res.Recordset("quesob") = 0
                data_res.Recordset("totuniv") = 0
                data_res.Recordset("totccou") = 0
                data_res.Recordset("totevang") = 0
                data_res.Recordset("totrec") = 0
                data_res.Recordset("comuni") = 0
                data_res.Recordset.Update
             End If
             data_res.Recordset.Edit
             'totales de llamados
             Xveonum1 = data_res.Recordset("quesob") + data_res.Recordset("totuniv") + data_res.Recordset("totccou")
             'totales de recibidos en <=90"
             Xtotcan = data_res.Recordset("totevang") + data_res.Recordset("totrec") + data_res.Recordset("comuni")
             data_res.Recordset("iva2") = Xtotcan / Xveonum1 * 100
             data_res.Recordset.Update
             If Check5.Value = 1 Then
                Command5_Click
             End If
             frm_calidadiso.MousePointer = 0
             MsgBox "Proceso terminado"
             If Option1.Value = True Then
                If Check3.Value = 1 Then
                   CrystalReport1.ReportFileName = App.path & "\infcaldes5n.rpt"
                   CrystalReport1.ReportTitle = "INFORME DE LLAMADOS POR USUARIO RECEPTOR Y POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                Else
                   If Check4.Value = 1 Then
                      CrystalReport1.ReportFileName = App.path & "\infcaldes5n.rpt"
                      CrystalReport1.DiscardSavedData = True
                      CrystalReport1.ReportTitle = "INFORME DE LLAMADOS POR USUARIO LARGADOR Y POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                   Else
                      CrystalReport1.ReportFileName = App.path & "\infcaldes5n.rpt"
                      CrystalReport1.DiscardSavedData = True
                      CrystalReport1.ReportTitle = "INFORME RECEPCION/CLASIFICACION/ASGFERNANDEZN DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                   End If
                End If
             Else
                If Check3.Value = 1 Then
                   CrystalReport1.ReportFileName = App.path & "\infcaldes5.rpt"
                   CrystalReport1.DiscardSavedData = True
                   CrystalReport1.ReportTitle = "INFORME DE LLAMADOS POR USUARIO RECEPTOR POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                Else
                   If Check4.Value = 1 Then
                      CrystalReport1.ReportFileName = App.path & "\infcaldes5.rpt"
                      CrystalReport1.DiscardSavedData = True
                      CrystalReport1.ReportTitle = "INFORME DE LLAMADOS POR USUARIO LARGADOR POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                   Else
                      CrystalReport1.ReportFileName = App.path & "\infcaldes3.rpt"
                      CrystalReport1.DiscardSavedData = True
                      CrystalReport1.ReportTitle = "INFORME RECEPCION/CLASIFICACION/ASGFERNANDEZN DE LLAMADOS POR CLAVE FECHA: " & mfd.Text & " HASTA : " & mfh.Text
                   End If
                End If
             End If
             CrystalReport1.Action = 1
          End If


End Sub

Private Sub Command5_Click()
Dim Xobjexel2 As Excel.Application
Dim Xlibexel2 As Excel.Workbook
Dim Xarchexel2 As New Excel.Worksheet
Dim Xdiferen As Double
Dim Xarchtex2 As String
Dim Xlin, XCol, Xtotglla, Xtotgllagt, Xpromed, Xtotgraldos As Double
Dim Xdiass As Long
Dim Xfecontrol As Date
Xdiass = 1

If Check5.Value = 1 Then
   Set Xobjexel2 = New Excel.Application
   Set Xlibexel2 = Xobjexel2.Workbooks.Add
   Set Xarchexel2 = Xlibexel2.Worksheets.Add
'   Xlibexel2.SaveAs ("C:\planillas\analisis" & Trim(Str(Month(mfd.Text))) & Trim(Str(Year(mfd.Text))) & ".xls")
   Xlibexel2.SaveAs ("C:\planillas\indicadores.xls")
   
   Xarchtex2 = "C:\planillas\indicadores.xls"
End If

Xlin = 1
XCol = 1
Xarchexel2.Name = "Indicadores"
Xarchexel2.Cells(Xlin, XCol) = "SAPP S.A."
Xlin = Xlin + 1
XCol = 2
Xarchexel2.Range("A1", "C3").Font.Size = 16
If Month(mfd.Text) = Month(mfh.Text) Then
   Xarchexel2.Cells(Xlin, XCol) = "INDICADORES ASISTENCIALES: " & Year(mfh.Text)
Else
   Xarchexel2.Cells(Xlin, XCol) = "INDICADORES ASISTENCIALES: " & mfd.Text & " AL " & mfh.Text
End If
'Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 120)
XCol = 1
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "ACTIVIDAD DE MOVILES: "
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "MES: " & Month(mfh.Text) & "/" & Year(mfh.Text)

'Xnrocan = Xnrocan + Xlin
Xarchexel2.Range("A5" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel2.Range("A5", "D" & Trim(str(35))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Xarchexel2.Range("A5", "D" & Trim(str(35))).Borders(xlInsideVertical).LineStyle = xlContinuous
Xarchexel2.Range("A5", "D" & Trim(str(35))).Borders(xlEdgeBottom).LineStyle = xlContinuous
Xarchexel2.Range("A5", "D" & Trim(str(35))).Borders(xlEdgeTop).LineStyle = xlContinuous
Xarchexel2.Range("A5", "D" & Trim(str(35))).Borders(xlEdgeLeft).LineStyle = xlContinuous
Xarchexel2.Range("A5", "D" & Trim(str(35))).Borders(xlEdgeRight).LineStyle = xlContinuous
Xarchexel2.Range("A5" & Trim(str(Xlin)), "D" & Trim(str(Xlin))).Interior.color = RGB(50, 110, 120)
'Xarchexel2.Range("A5" & Trim(Str(Xlin))).ColumnWidth = 43
Xarchexel2.Range("B5" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel2.Range("C5" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel2.Range("D5" & Trim(str(Xlin))).ColumnWidth = 15
'Xarchexel2.Range("E4" & Trim(Str(Xlin))).ColumnWidth = 4
Dim Xelgrantot As Long

If Check1.Value = 1 Then
   If t_mov.Text = "" Then
      If Check8.Value = 1 Then
         data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 3 & " and movilpas =" & t_mov.Text & " and codmed not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043,959) and cancela is null order by codmot"
         data_llam.Refresh
      Else
         If Check10.Value = 1 Then
            data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 3 & " and codmed <>" & 959 & " and categ not in ('911','911B') and cancela is null order by codmot"
            data_llam.Refresh
         Else
            If t_codmed.Text = "" Then
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 3 & " and codmed <>" & 959 & " and cancela is null order by codmot"
               data_llam.Refresh
            Else
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 3 & " and codmed =" & t_codmed.Text & " and cancela is null order by codmot"
               data_llam.Refresh
            End If
         End If
      End If
   Else
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "# And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 3 & " and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by codmot"
      data_llam.Refresh
   End If
          
   ' End If
Else
'''15/11 se modifica para agregar CMT al total de llamados
   ''Data1.DatabaseName = App.path & "\informes.mdb"
   
   data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Trim(Xtexto) & "' and codmot in ('R','C','A','V') and codmed <>" & 959 & " and cancela is null and categ not in ('MSP') and codzon not in (4,6) order by nrolla"
'   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# order by codmot"
   data_llam.Refresh
   If data_llam.Recordset.RecordCount > 0 Then
      data_llam.Recordset.MoveLast
      Xelgrantot = data_llam.Recordset.RecordCount
   End If

End If
data_res.Recordset.MoveFirst
Data1.DatabaseName = App.path & "\informes.mdb"
If Check1.Value = 1 Then
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "R" & "' and codzon in (3) order by codmot"
Else
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "R" & "' and codmed <>" & 959 & " and movilpas <>" & 99 & " and codzon not in (4,6) order by codmot"
' Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "R" & "' and codzon in (1,2,3) order by codmot"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
End If
Xtotglla = Data1.Recordset.RecordCount

XCol = 1
Xlin = Xlin + 2

Xarchexel2.Cells(Xlin, XCol) = "Asistencias Clave 1:"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de Salida <=3':"
XCol = XCol + 1
If IsNull(data_res.Recordset("totacom")) = False Then
   Xdiferen = Xtotglla - data_res.Recordset("totacom")
Else
   Xdiferen = Xtotglla - 0
End If

If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard"))) & "%"
'Xdiass = 1
Xlin = Xlin + 1
XCol = 2
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de respuesta <=15':"
XCol = XCol + 1
'If IsNull(data_res.Recordset("comm")) = False Then
'   Xdiferen = Xtotglla - data_res.Recordset("comm")
'Else
'   Xdiferen = Xtotglla - 0
'End If
If IsNull(data_res.Recordset("mes")) = False Then
   Xdiferen = Xtotglla - data_res.Recordset("mes")
Else
   Xdiferen = Xtotglla - 0
End If

If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If

'Xarchexel2.Cells(Xlin, XCol) = Trim(Str(Format(Xpromed, "Standard"))) & "%"
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(data_res.Recordset("anoarq"), "Standard"))) & "%"

Xlin = Xlin + 2
XCol = 1
If Check1.Value = 1 Then
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "A" & "' and codzon in(3) order by codmot"
Else
  Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "A" & "' and codmed <>" & 959 & " and movilpas <>" & 99 & " and categ not in ('MSP','50') and codzon not in (4,6) order by codmot"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
End If
Xtotglla = Data1.Recordset.RecordCount

Xarchexel2.Cells(Xlin, XCol) = "Asistencias Clave 2:"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de Salida <=5':"
XCol = XCol + 1
If IsNull(data_res.Recordset("totodon")) = False Then
   Xdiferen = Xtotglla - data_res.Recordset("totodon")
End If
If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If

Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard"))) & "%"
'Xdiass = 1
Xlin = Xlin + 1
XCol = 2
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de respuesta <=30':"
XCol = XCol + 1
If Xtotglla = 0 Then
   Xdiferen = 0
   Xpromed = 0
Else
   Xdiferen = Xtotglla - data_res.Recordset("comr")
   Xpromed = Xdiferen / Xtotglla * 100
End If

'Xarchexel2.Cells(Xlin, XCol) = Trim(Str(Format(Xpromed, "Standard"))) & "%"
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(data_res.Recordset("comm"), "Standard"))) & "%"

Xlin = Xlin + 2
XCol = 1
'' celestes
Xlin = Xlin + 2
XCol = 1
If Check1.Value = 1 Then
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "C" & "' and codzon =" & 3 & " order by codmot"
Else
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "C" & "' and codmed <>" & 959 & " and movilpas <>" & 99 & " and categ not in ('MSP','50') and codzon not in (4,6) order by codmot"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
End If
Xtotglla = Data1.Recordset.RecordCount

Xarchexel2.Cells(Xlin, XCol) = "Asistencias Clave C:"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de Salida <=5':"
XCol = XCol + 1
Xdiferen = Xtotglla - data_res.Recordset("cantc5m")
If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If

Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard"))) & "%"
'Xdiass = 1
Xlin = Xlin + 1
XCol = 2
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de respuesta <=30':"
XCol = XCol + 1
Xdiferen = Xtotglla - data_res.Recordset("cantc30m")
If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If

'Xarchexel2.Cells(Xlin, XCol) = Trim(Str(Format(Xpromed, "Standard"))) & "%"
If IsNull(data_res.Recordset("iva2")) = False Then
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(data_res.Recordset("iva2"), "Standard"))) & "%"
Else
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format("0", "Standard"))) & "%"
End If
Xlin = Xlin + 2
XCol = 1

''--- fin celestes
If Check1.Value = 1 Then
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "' and codzon in (3) order by codmot"
Else
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "' and codmed <>" & 959 & " and movilpas <>" & 99 & " and categ not in ('MSP','50') and codzon not in (4,6) order by codmot"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
End If
Xtotglla = Data1.Recordset.RecordCount

Xarchexel2.Cells(Xlin, XCol) = "Asistencias Clave 3:"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de Salida < 5':"
XCol = XCol + 1
If IsNull(data_res.Recordset("totdeudas")) = False Then
   Xdiferen = Xtotglla - data_res.Recordset("totdeudas")
Else
   Xdiferen = Xtotglla
End If
If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If

Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard"))) & "%"
'Xdiass = 1
Xlin = Xlin + 1
XCol = 2
Xarchexel2.Cells(Xlin, XCol) = "Tiempo de respuesta:"
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = ">=90% máximo 2 horas:"
XCol = XCol + 1
Xpromed = data_res.Recordset("comdeu")

If IsNull(Xpromed) = True Then
   Xpromed = 0
End If

Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard"))) & "%"

XCol = 2
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "Tiempo máximo <= 3horas:"
XCol = XCol + 1
'Xdiferen = Xtotglla - data_res.Recordset("comacom")
If IsNull(data_res.Recordset("totrec")) = True Then
   Xdiferen = 0
Else
   Xdiferen = Xtotglla - data_res.Recordset("totrec")
End If
If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(data_res.Recordset("comv"), "Standard"))) & "%"

'Xlin = Xlin + 2
'XCol = 1
'Xarchexel2.Cells(Xlin, XCol) = "Tiempo de atención en"
'Xlin = Xlin + 1
'Xarchexel2.Cells(Xlin, XCol) = "el lugar de asistencia"
'XCol = XCol + 1
'Xarchexel2.Cells(Xlin, XCol) = "<= 30 minutos: "
'XCol = XCol + 1
'Xdiferen = 0
'If Check1.Value = 1 Then
'   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "' and codzon in (3) order by codmot"
'Else
'   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "' and codmed <>" & 959 & " and movilpas <>" & 99 & " and categ not in ('MSP','50') and codzon not in (4,6) order by codmot"
'End If
'Data1.Refresh
'If Data1.Recordset.RecordCount > 0 Then
'   Data1.Recordset.MoveLast
'End If
'Xtotglla = Data1.Recordset.RecordCount

Dim Xtotgllados As Double
Xtotgllados = Data1.Recordset.RecordCount
If Check1.Value = 1 Then
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "' and totend <='" & "00:30" & "' and codzon in (3) order by codmot"
Else
   Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and codmot ='" & "V" & "' and totend <='" & "00:30" & "' and codzon in (1,2,3,6) order by codmot"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
End If
Xdiferen = Data1.Recordset.RecordCount
If Xtotglla = 0 Then
   Xpromed = 0
Else
   Xpromed = Xdiferen / Xtotglla * 100
End If

'Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard"))) & "%"
Xlin = Xlin + 2


XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TRASLADOS"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "Total realizados:"

If t_mov.Text = "" Then
'   data_llam.RecordSource = "Select * from llamado where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# And trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13) order by codmot"
'not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043)
   If Check8.Value = 1 Then
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,4,6) and codmed not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043,959) and cancela is null order by codmot"
      data_llam.Refresh
   Else
'      data_llam.RecordSource = "Select * from llamado where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# And trasla in (1,2,4,6) and codmed <>" & 959 & " and movilpas not in (620,602,601) order by codmot"
      If t_codmed.Text = "" Then
'         data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,4,6) and cancela is null order by codmot"
         
         data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,5,10) and cancela is null order by codmot"
         data_llam.Refresh
      Else
         data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16) and codmed =" & t_codmed.Text & " and cancela is null order by codmot"
         data_llam.Refresh
      End If
   End If
Else
   If t_codmed.Text = "" Then
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by codmot"
      data_llam.Refresh
   Else
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16) and movilpas =" & t_mov.Text & " and codmed =" & t_codmed.Text & " and cancela is null order by codmot"
      data_llam.Refresh
   End If
End If

 Dim MiBaseact As Database
 Dim Unasesact As Workspace
 Set Unasesact = Workspaces(0)
 Set MiBaseact = Unasesact.OpenDatabase(App.path & "\inftras.mdb")

 MiBaseact.Execute "Delete * from inflla"
 data_tras.Refresh

If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveFirst
   Do While Not data_llam.Recordset.EOF
      data_tras.Recordset.AddNew
      data_tras.Recordset("fecha") = data_llam.Recordset("fecha")
      data_tras.Recordset("hora") = data_llam.Recordset("hora")
      data_tras.Recordset("nombre") = data_llam.Recordset("nombre")
      data_tras.Recordset("categ") = data_llam.Recordset("categ")
      data_tras.Recordset.Update
      data_llam.Recordset.MoveNext
   Loop
'   data_llam.Recordset.MoveLast
End If
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = data_llam.Recordset.RecordCount
Xlin = Xlin + 1
XCol = 2
Xarchexel2.Cells(Xlin, XCol) = "Promedio diario:"
XCol = XCol + 1
Xtotglla = data_llam.Recordset.RecordCount
Xtotgllados = data_llam.Recordset.RecordCount
If Month(mfd.Text) = Month(mfh.Text) Then
   Xpromed = data_llam.Recordset.RecordCount / Day(mfh.Text)
Else
   Xpromed = data_llam.Recordset.RecordCount / 365
End If
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard")))
XCol = 2
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "% Coordinados a Montevideo"
If t_mov.Text = "" Then
   If Check8.Value = 1 Then
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (7,9,10,11) and codmed not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043) and cancela is null order by codmot"
      data_llam.Refresh
   Else
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (7,8) and cancela is null order by codmot"
      data_llam.Refresh
   End If
Else
   data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (7,8) and movilpas =" & t_mov.Text & " and cancela is null order by codmot"
   data_llam.Refresh
End If
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveLast
End If
Xtotglla = data_llam.Recordset.RecordCount
Xpromed = Xtotglla / Xtotgllados * 100
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard")))

Xlin = Xlin + 1
XCol = 2
Xarchexel2.Cells(Xlin, XCol) = "% Traslados en zona"
If t_mov.Text = "" Then
'not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043)
   If Check8.Value = 1 Then
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (3) and codmed not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043) and cancela is null order by codmot"
      data_llam.Refresh
   Else
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (3) and cancela is null order by codmot"
      data_llam.Refresh
   End If
Else
   data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And trasla in (3) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by codmot"
   data_llam.Refresh
End If
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveLast
End If
Xtotglla = data_llam.Recordset.RecordCount
Xpromed = Xtotglla / Xtotgllados * 100
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard")))
XCol = 2
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "% VP+AP:"
Xlin = Xlin + 2
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "Relación Traslado/Llamado"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "Total:"
XCol = XCol + 1
Data1.RecordSource = "Select * from inflla where fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# order by codmot"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
End If

Xpromed = Xtotgllados / Xelgrantot * 100
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard")))
Xlin = Xlin + 1
XCol = 2
Xtotgllados = Data1.Recordset.RecordCount
Xarchexel2.Cells(Xlin, XCol) = "Norte (con Tala y SJ)"
If t_mov.Text = "" Then
   If Check8.Value = 1 Then
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon in (2,3,5) And trasla in (1,2,3,4,5,6,7,8,9,10,11,16) and codmed not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043,959) and cancela is null order by codmot"
      data_llam.Refresh
   Else
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon in (2,3,5) And trasla in (2,5,9,10) and codmed <>" & 959 & " and cancela is null order by codmot"
      data_llam.Refresh
   End If
Else
   data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon in (2,3,5) And trasla in (1,2,3,4,5,6,7,8,9,10,11,16) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by codmot"
   data_llam.Refresh
End If
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveLast
End If
XCol = XCol + 1
Xpromed = data_llam.Recordset.RecordCount / Xelgrantot * 100
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard")))
XCol = 2
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "Tala, SJ:"
If t_mov.Text = "" Then
   data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon in (3,5) And trasla in (5,10) and cancela is null order by codmot"
   data_llam.Refresh
Else
   data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon in (3,5) And trasla in (1,2,3,4,5,6,7,8,9,10,11,16) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by codmot"
   data_llam.Refresh
End If
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveLast
End If
XCol = XCol + 1
Xpromed = data_llam.Recordset.RecordCount / Xelgrantot * 100
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard")))
Xlin = Xlin + 1
XCol = 2
Xarchexel2.Cells(Xlin, XCol) = "Costa:"
If t_mov.Text = "" Then
   If Check8.Value = 1 Then
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 1 & " And trasla in (1,2,3,4,5,6,7,8,9,10,11) and codmed not in (1024,1030,1021,1035,1033,1032,1018,1040,976,977,1020,994,1014,1029,1031,1046,1044,1045,1043,959) and cancela is null order by codmot"
      data_llam.Refresh
   Else
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 1 & " And trasla in (1) and codmed <>" & 959 & " and cancela is null order by codmot"
      data_llam.Refresh
   End If
Else
   data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codzon =" & 1 & " And trasla in (1,2,3,4,5,6,7,8,9,10,11) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by codmot"
   data_llam.Refresh
End If
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveLast
End If
XCol = XCol + 1
Xpromed = data_llam.Recordset.RecordCount / Xelgrantot * 100
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xpromed, "Standard")))

Xlin = Xlin + 1
XCol = 1
'Xarchexel2.Cells(Xlin, XCol) = "Coordinados:"
XCol = XCol + 1
'Xarchexel2.Cells(Xlin, XCol) = "Tiempo de respuesta <=2 horas:"
Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "Total llamados:"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = Trim(str(Format(Xelgrantot, "Standard")))

DoEvents
Xlibexel2.Save
Xlibexel2.Close
Xobjexel2.Quit
MsgBox "El archivo INDICADORES.xls se guardó en la carpeta PLANILLAS", vbInformation, "SAPP"


End Sub

Private Sub Command7_Click()
Dim Xobjexel2 As Excel.Application
Dim Xlibexel2 As Excel.Workbook
Dim Xarchexel2 As New Excel.Worksheet

Dim Xarchtex2, XnombreUs, Fechatexto As String
Dim Xlin, XCol, Xtotglla, Xtotgraldos As Double
Dim Xdiass, Xhastaqdia As Long
Dim Xfecontrol As Date
Dim Xelanio, Xelmes As Integer
Xelanio = Year(mfd.Text)
Xelmes = Month(mfd.Text)
Xdiass = 1


Xhastaqdia = 30
If Month(mfd.Text) = 1 Or Month(mfd.Text) = 3 Or Month(mfd.Text) = 5 Or Month(mfd.Text) = 7 Or Month(mfd.Text) = 8 Or Month(mfd.Text) = 10 Or Month(mfd.Text) = 12 Then
   Xhastaqdia = 31
Else
  If Month(mfd.Text) = 2 And Year(mfd.Text) = 2020 Then
     Xhastaqdia = 29
  Else
     If Month(mfd.Text) = 2 And Year(mfd.Text) = 2021 Then
        Xhastaqdia = 28
     Else
        If Month(mfd.Text) = 2 And Year(mfd.Text) = 2022 Then
           Xhastaqdia = 28
        Else
           If Month(mfd.Text) = 2 And Year(mfd.Text) = 2023 Then
              Xhastaqdia = 28
           Else
              Xhastaqdia = 29
           End If
        End If
     End If
  End If
End If

Set Xobjexel2 = New Excel.Application
Set Xlibexel2 = Xobjexel2.Workbooks.Add
Set Xarchexel2 = Xlibexel2.Worksheets.Add
'   Xlibexel2.SaveAs ("C:\planillas\analisis" & Trim(Str(Month(mfd.Text))) & Trim(Str(Year(mfd.Text))) & ".xls")
Xlibexel2.SaveAs ("C:\planillas\CMT_por_Telefonista.xls")
Xarchtex2 = "C:\planillas\CMT_por_Telefonista.xls"
data_infu.RecordSource = "infarq"
data_infu.Refresh
If data_infu.Recordset.RecordCount > 0 Then
   data_infu.Recordset.MoveFirst
   Do While Not data_infu.Recordset.EOF
      data_infu.Recordset.Delete
      data_infu.Recordset.MoveNext
   Loop
   data_infu.Refresh
End If
data_info.RecordSource = "inflla"
data_info.Refresh
If data_info.Recordset.RecordCount > 0 Then
   data_info.Recordset.MoveFirst
   Do While Not data_info.Recordset.EOF
      data_info.Recordset.Delete
      data_info.Recordset.MoveNext
   Loop
   data_info.Refresh
End If

Xlin = 1
XCol = 1
Xarchexel2.Name = "CMT_PASADOS"
Xarchexel2.Cells(Xlin, XCol) = "SAPP S.A."
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "FECHA ACTUAL: " & Format(Date, "dd/mm/yyyy")

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Range("A1", "C3").Font.Size = 16
Xarchexel2.Cells(Xlin, XCol) = "MES: " & Month(mfd.Text) & "/" & Year(mfh.Text) & " ---CMT PASADOS"
'Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 120)
Xarchexel2.Range("A3").Interior.color = RGB(115, 120, 0)

XCol = 5
Xarchexel2.Cells(Xlin, XCol) = "DIAS"

XCol = 1
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "OPERADOR TELEFÓNICO"

'Xnrocan = Xnrocan + Xlin
Xarchexel2.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlInsideVertical).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeBottom).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeTop).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeLeft).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeRight).LineStyle = xlContinuous
Xarchexel2.Range("B4" & Trim(str(Xlin)), "AN" & Trim(str(Xlin))).Interior.color = RGB(24, 101, 244)
Xarchexel2.Range("A4" & Trim(str(Xlin))).ColumnWidth = 45
Xarchexel2.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("C4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("D4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("E4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("F4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("G4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("H4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("I4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("J4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("K4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("L4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("M4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("N4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("O4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("P4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Q4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("R4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("S4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("T4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("U4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("V4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("W4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("X4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Y4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Z4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AA4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AB4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AC4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AD4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AE4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AF4" & Trim(str(Xlin))).ColumnWidth = 4

XCol = 2
Do While Xdiass <= Xhastaqdia
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xdiass))
   Xdiass = Xdiass + 1
   XCol = XCol + 1
Loop
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL"
XCol = 1
Xlin = Xlin + 1
Xdiass = 1
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveFirst
   Do While Not data_llam.Recordset.EOF
      data_info.Recordset.AddNew
      data_info.Recordset("fecha") = data_llam.Recordset("fecha")
      data_info.Recordset("usuario") = Trim(data_llam.Recordset("usuario"))
      data_info.Recordset("nombre") = data_llam.Recordset("nombre")
      data_info.Recordset("codmot") = data_llam.Recordset("codmot")
      data_info.Recordset.Update
      data_llam.Recordset.MoveNext
   Loop
   data_info.RecordSource = "select * from inflla order by usuario"
   data_info.Refresh
   If data_info.Recordset.RecordCount > 0 Then
      data_info.Recordset.MoveFirst
      XnombreUs = Trim(data_info.Recordset("usuario"))
      Do While Not data_info.Recordset.EOF
         If Trim(data_info.Recordset("usuario")) = Trim(XnombreUs) Then
        
         Else
            data_info.Recordset.MovePrevious
            data_infu.Recordset.AddNew
            data_infu.Recordset("nombre") = Trim(data_info.Recordset("usuario"))
            data_infu.Recordset.Update
            data_info.Recordset.MoveNext
         End If
         XnombreUs = Trim(data_info.Recordset("usuario"))
         data_info.Recordset.MoveNext
      Loop
      data_info.Recordset.MovePrevious
      data_infu.Recordset.AddNew
      data_infu.Recordset("nombre") = Trim(data_info.Recordset("usuario"))
      data_infu.Recordset.Update
   End If
   data_infu.Refresh
   If data_infu.Recordset.RecordCount > 0 Then
      data_infu.Recordset.MoveFirst
      Xtotglla = 0
      Do While Not data_infu.Recordset.EOF
         Xarchexel2.Cells(Xlin, XCol) = data_infu.Recordset("nombre")
         XCol = XCol + 1
         Do While Xdiass <= Xhastaqdia
            If Xdiass < 10 Then
               If Month(mfd.Text) < 10 Then
                  Fechatexto = "0" & Xdiass & "/0" & Month(mfd.Text) & "/" & Year(mfd.Text)
               Else
                  Fechatexto = "0" & Xdiass & "/" & Month(mfd.Text) & "/" & Year(mfd.Text)
               End If
            Else
               If Month(mfd.Text) < 10 Then
                  Fechatexto = Xdiass & "/0" & Month(mfd.Text) & "/" & Year(mfd.Text)
               Else
                  Fechatexto = Xdiass & "/" & Month(mfd.Text) & "/" & Year(mfd.Text)
               End If
            End If
            data_info.RecordSource = "select * from inflla where usuario ='" & data_infu.Recordset("nombre") & "' and fecha =#" & Format(Fechatexto, "yyyy/mm/dd") & "#"
            data_info.Refresh
            If data_info.Recordset.RecordCount > 0 Then
               data_info.Recordset.MoveLast
               Xarchexel2.Cells(Xlin, XCol) = data_info.Recordset.RecordCount
               Xtotglla = Xtotglla + data_info.Recordset.RecordCount
            Else
               Xarchexel2.Cells(Xlin, XCol) = 0
            End If
            XCol = XCol + 1
            Xdiass = Xdiass + 1
         Loop
         XCol = XCol + 1
         Xarchexel2.Cells(Xlin, XCol) = Xtotglla
         Xtotglla = 0
         Xdiass = 1
         XCol = 1
         data_infu.Recordset.MoveNext
         Xlin = Xlin + 1
      Loop
      
   End If
Else
   MsgBox "No hay registros", vbInformation
End If
       
DoEvents
Xlibexel2.Save
Xlibexel2.Close
Xobjexel2.Quit
frm_calidadiso.MousePointer = 0
MsgBox "Proceso terminado, el archivo quedó guardado en la carpeta planillas con el nombre CMT_por_Telefonista.xls", vbInformation

'Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus

End Sub

Private Sub Command8_Click()
Dim Xobjexel2 As Excel.Application
Dim Xlibexel2 As Excel.Workbook
Dim Xarchexel2 As New Excel.Worksheet

Dim Xarchtex2, XnombreUs, Fechatexto As String
Dim Xlin, XCol, Xtotglla, Xtotgraldos As Double
Dim Xdiass, Xhastaqdia As Long
Dim Xfecontrol As Date
Dim Xelanio, Xelmes As Integer
Xelanio = Year(mfd.Text)
Xelmes = Month(mfd.Text)
Xdiass = 1


Xhastaqdia = 30
If Month(mfd.Text) = 1 Or Month(mfd.Text) = 3 Or Month(mfd.Text) = 5 Or Month(mfd.Text) = 7 Or Month(mfd.Text) = 8 Or Month(mfd.Text) = 10 Or Month(mfd.Text) = 12 Then
   Xhastaqdia = 31
Else
  If Month(mfd.Text) = 2 And Year(mfd.Text) = 2020 Then
     Xhastaqdia = 29
  Else
     If Month(mfd.Text) = 2 And Year(mfd.Text) = 2021 Then
        Xhastaqdia = 28
     End If
  End If
End If

Set Xobjexel2 = New Excel.Application
Set Xlibexel2 = Xobjexel2.Workbooks.Add
Set Xarchexel2 = Xlibexel2.Worksheets.Add
'   Xlibexel2.SaveAs ("C:\planillas\analisis" & Trim(Str(Month(mfd.Text))) & Trim(Str(Year(mfd.Text))) & ".xls")
Xlibexel2.SaveAs ("C:\planillas\Llamados_por_Telefonista.xls")
Xarchtex2 = "C:\planillas\Llamados_por_Telefonista.xls"
data_infu.RecordSource = "infarq"
data_infu.Refresh
If data_infu.Recordset.RecordCount > 0 Then
   data_infu.Recordset.MoveFirst
   Do While Not data_infu.Recordset.EOF
      data_infu.Recordset.Delete
      data_infu.Recordset.MoveNext
   Loop
   data_infu.Refresh
End If
data_info.RecordSource = "inflla"
data_info.Refresh
If data_info.Recordset.RecordCount > 0 Then
   data_info.Recordset.MoveFirst
   Do While Not data_info.Recordset.EOF
      data_info.Recordset.Delete
      data_info.Recordset.MoveNext
   Loop
   data_info.Refresh
End If

Xlin = 1
XCol = 1
Xarchexel2.Name = "LLAMADOS"
Xarchexel2.Cells(Xlin, XCol) = "SAPP S.A."
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "FECHA ACTUAL: " & Format(Date, "dd/mm/yyyy")

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Range("A1", "C3").Font.Size = 16
Xarchexel2.Cells(Xlin, XCol) = "MES: " & Month(mfd.Text) & "/" & Year(mfh.Text) & " ---LLAMADOS POR OPERADOR"
'Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 120)
Xarchexel2.Range("A3").Interior.color = RGB(115, 120, 0)

XCol = 5
Xarchexel2.Cells(Xlin, XCol) = "DIAS"

XCol = 1
Xlin = Xlin + 1
Xarchexel2.Cells(Xlin, XCol) = "OPERADOR TELEFÓNICO"

'Xnrocan = Xnrocan + Xlin
Xarchexel2.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlInsideVertical).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeBottom).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeTop).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeLeft).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(30))).Borders(xlEdgeRight).LineStyle = xlContinuous
Xarchexel2.Range("B4" & Trim(str(Xlin)), "AN" & Trim(str(Xlin))).Interior.color = RGB(24, 101, 244)
Xarchexel2.Range("A4" & Trim(str(Xlin))).ColumnWidth = 45
Xarchexel2.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("C4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("D4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("E4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("F4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("G4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("H4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("I4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("J4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("K4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("L4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("M4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("N4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("O4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("P4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Q4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("R4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("S4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("T4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("U4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("V4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("W4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("X4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Y4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Z4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AA4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AB4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AC4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AD4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AE4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AF4" & Trim(str(Xlin))).ColumnWidth = 4

XCol = 2
Do While Xdiass <= Xhastaqdia
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xdiass))
   Xdiass = Xdiass + 1
   XCol = XCol + 1
Loop
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL"
XCol = 1
Xlin = Xlin + 1
Xdiass = 1
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveFirst
   Do While Not data_llam.Recordset.EOF
      data_info.Recordset.AddNew
      data_info.Recordset("fecha") = data_llam.Recordset("fecha")
      data_info.Recordset("usuario") = Trim(data_llam.Recordset("usuario"))
      data_info.Recordset("nombre") = data_llam.Recordset("nombre")
      data_info.Recordset("codmot") = data_llam.Recordset("codmot")
      data_info.Recordset.Update
      data_llam.Recordset.MoveNext
   Loop
   data_info.RecordSource = "select * from inflla order by usuario"
   data_info.Refresh
   If data_info.Recordset.RecordCount > 0 Then
      data_info.Recordset.MoveFirst
      XnombreUs = Trim(data_info.Recordset("usuario"))
      Do While Not data_info.Recordset.EOF
         If Trim(data_info.Recordset("usuario")) = Trim(XnombreUs) Then
        
         Else
            data_info.Recordset.MovePrevious
            data_infu.Recordset.AddNew
            data_infu.Recordset("nombre") = Trim(data_info.Recordset("usuario"))
            data_infu.Recordset.Update
            data_info.Recordset.MoveNext
         End If
         XnombreUs = Trim(data_info.Recordset("usuario"))
         data_info.Recordset.MoveNext
      Loop
      data_info.Recordset.MovePrevious
      data_infu.Recordset.AddNew
      data_infu.Recordset("nombre") = Trim(data_info.Recordset("usuario"))
      data_infu.Recordset.Update
   End If
   data_infu.Refresh
   If data_infu.Recordset.RecordCount > 0 Then
      data_infu.Recordset.MoveFirst
      Xtotglla = 0
      Do While Not data_infu.Recordset.EOF
         Xarchexel2.Cells(Xlin, XCol) = data_infu.Recordset("nombre")
         XCol = XCol + 1
         Do While Xdiass <= Xhastaqdia
            If Xdiass < 10 Then
               If Month(mfd.Text) < 10 Then
                  Fechatexto = "0" & Xdiass & "/0" & Month(mfd.Text) & "/" & Year(mfd.Text)
               Else
                  Fechatexto = "0" & Xdiass & "/" & Month(mfd.Text) & "/" & Year(mfd.Text)
               End If
            Else
               If Month(mfd.Text) < 10 Then
                  Fechatexto = Xdiass & "/0" & Month(mfd.Text) & "/" & Year(mfd.Text)
               Else
                  Fechatexto = Xdiass & "/" & Month(mfd.Text) & "/" & Year(mfd.Text)
               End If
            End If
            data_info.RecordSource = "select * from inflla where usuario ='" & data_infu.Recordset("nombre") & "' and fecha =#" & Format(Fechatexto, "yyyy/mm/dd") & "#"
            data_info.Refresh
            If data_info.Recordset.RecordCount > 0 Then
               data_info.Recordset.MoveLast
               Xarchexel2.Cells(Xlin, XCol) = data_info.Recordset.RecordCount
               Xtotglla = Xtotglla + data_info.Recordset.RecordCount
            Else
               Xarchexel2.Cells(Xlin, XCol) = 0
            End If
            XCol = XCol + 1
            Xdiass = Xdiass + 1
         Loop
         XCol = XCol + 1
         Xarchexel2.Cells(Xlin, XCol) = Xtotglla
         Xtotglla = 0
         Xdiass = 1
         XCol = 1
         data_infu.Recordset.MoveNext
         Xlin = Xlin + 1
      Loop
      
   End If
Else
   MsgBox "No hay registros", vbInformation
End If
       
DoEvents
Xlibexel2.Save
Xlibexel2.Close
Xobjexel2.Quit
frm_calidadiso.MousePointer = 0
MsgBox "Proceso terminado, el archivo quedó guardado en la carpeta planillas con el nombre Llamados_por_Telefonista.xls", vbInformation

End Sub

Private Sub Command9_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xqdia, Xqdia3, Xcannocumple, Xcannocumple3, Xlin, XCol As Integer
Dim Xtotreg As Long
Dim Xarchtex As String
Dim Textofecha As String

Dim Xlabrir3 As New Excel.Application
Dim Xlabancp As Integer
Dim Sincumplir As Integer
Dim Sincumplir3 As Integer

Dim Xhh1, Xmm1, Xhh2, Xmm2, Xtoths As Integer

Xlin = 1
XCol = 1
Xtotreg = 0
Sincumplir = 0
Sincumplir3 = 0

frm_calidadiso.MousePointer = 11

Set Xobjexel22 = New Excel.Application

Set Xlibexel22 = Xobjexel22.Workbooks.Add
Set Xarchexel22 = Xlibexel22.Worksheets.Add

Xarchexel22.Name = Trim(Combo1.Text)

Xlibexel22.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls")
Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls"

Xqdia = 0
Xcanxdia = 0
Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
Xlin = Xlin + 1
XCol = XCol + 1
Xarchexel22.Range("A1", "C3").Font.Size = 16
Xarchexel22.Cells(Xlin, XCol) = "PLANILLA DE INDICADORES CMT POR MÉDICO " & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text

Xarchexel22.Range("B" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)

XCol = 1
Xlin = Xlin + 2
Xnrocan = Xnrocan + Xlin
Xarchexel22.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "FECHA"
XCol = XCol + 1
Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "HORA"
XCol = XCol + 1
Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
XCol = XCol + 1
Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "CATEGORIA"
XCol = XCol + 1
Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "HORA PAS."
XCol = XCol + 1
Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "H.REALIZA"
XCol = XCol + 1
Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "TOT.DEMORA"
XCol = XCol + 1
Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "COD.MED."
XCol = XCol + 1
Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 22
Xarchexel22.Cells(Xlin, XCol) = "NOMBRE MEDICO"

Xlin = Xlin + 1
XCol = 1
        
If data_llam.Recordset.RecordCount > 0 Then
   data_llam.Recordset.MoveFirst
   Do While Not data_llam.Recordset.EOF
      Xtotreg = Xtotreg + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llam.Recordset("fecha"), "dd/mm/yyyy")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Format(data_llam.Recordset("hora"), "HH:mm")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("nombre")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("categ")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Format(data_llam.Recordset("activo"), "HH:mm")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Format(data_llam.Recordset("hor_rea"), "HH:mm")
        
      If IsNull(data_llam.Recordset("hor_rea")) = False Then
         Xhh1 = Val(Mid(data_llam.Recordset("hor_rea"), 1, 2))
         Xmm1 = Val(Mid(data_llam.Recordset("hor_rea"), 4, 2))
      End If
      If IsNull(data_llam.Recordset("hora")) = False Then
         Xhh2 = Val(Mid(data_llam.Recordset("hora"), 1, 2))
         Xmm2 = Val(Mid(data_llam.Recordset("hora"), 4, 2))
      End If
        
      If Xhh1 = Xhh2 Then
         Xtmm = Xmm1 - Xmm2
      Else
         Xths = Xhh1 - Xhh2
'         If IsNull(data_llam.Recordset("fecsali")) = False Then
'            If data_llam.Recordset("fecsali") > data_llam.Recordset("fecha") Then
'               Xths = Xths + 24
'            End If
'         End If
         Xtmm = Xmm1 - Xmm2 + 60
         If Xths = 2 Then
            Xtmm = Xtmm + 60
         End If
         If Xths = 3 Then
            Xtmm = Xtmm + 120
         End If
         If Xths = 4 Then
            Xtmm = Xtmm + 180
         End If
         If Xths = 5 Then
            Xtmm = Xtmm + 240
         End If
      End If
      If Xtmm > 120 Then
         Sincumplir = Sincumplir + 1
      End If
      If Xtmm > 180 Then
         Sincumplir3 = Sincumplir3 + 1
      End If
      
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Xtmm
      XCol = XCol + 1
      If IsNull(data_llam.Recordset("codmedcmt")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_llam.Recordset("codmedcmt")
         data_consmed.RecordSource = "select * from medicos_esp where id =" & data_llam.Recordset("codmedcmt")
         data_consmed.Refresh
         If data_consmed.Recordset.RecordCount > 0 Then
            XCol = XCol + 1
            Xarchexel22.Cells(Xlin, XCol) = data_consmed.Recordset("nom_med")
         Else
            XCol = XCol + 1
            Xarchexel22.Cells(Xlin, XCol) = "Sin Dato"
         End If
      End If
            
      data_llam.Recordset.MoveNext
        
      Xlin = Xlin + 1
      XCol = 1
   Loop
   
   If Sincumplir > 0 Then
      Xcannocumple = Sincumplir / Xtotreg * 100
   Else
      Xcannocumple = 0
   End If
   If Sincumplir3 > 0 Then
      Xcannocumple3 = Sincumplir3 / Xtotreg * 100
   Else
      Xcannocumple3 = 0
   End If
   
   Xlin = Xlin + 1
   XCol = 1
   
   Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS: " & Xtotreg
   Xlin = Xlin + 1
   XCol = 1
   Xqdia = 100 - Xcannocumple
   Xqdia3 = 100 - Xcannocumple3
   
   Xarchexel22.Cells(Xlin, XCol) = "PORCENTAJE DE CMT CUMPLIDO (<=2HS): " & Format(Xqdia, "Standard") & " %"
   Xlin = Xlin + 1
   Xarchexel22.Cells(Xlin, XCol) = "PORCENTAJE DE CMT CUMPLIDO (<=3HS): " & Format(Xqdia3, "Standard") & " %"
   
   
   frm_calidadiso.MousePointer = 0
   MsgBox "Proceso terminado"
   Xlibexel22.Save
   Xlibexel22.Close
   Xobjexel22.Quit
        
   Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls"
   Xlabrir3.Workbooks.Open Xarchtex, , False
   Xlabrir3.Visible = True
   Xlabrir3.WindowState = xlMaximized
Else
   frm_calidadiso.MousePointer = 0
   MsgBox "No hay registros"
End If

End Sub

Private Sub Form_Load()
data_res.DatabaseName = App.path & "\informes.mdb"
'data_res.RecordSource = "infarqc"
'data_res.Refresh

If Check7.Value = 1 Then
'   data_llam.DatabaseName = App.Path & "\llamado.mdb"
   data_llam.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
Else
   data_llam.ConnectionString = "DSN=" & Xconexrmt
End If
'data_llam.RecordSource = "llamado"
'data_llam.Refresh
data_info.DatabaseName = App.path & "\informes.mdb"
'data_info.RecordSource = "inflla"
'data_info.Refresh

data_consmed.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data1.DatabaseName = App.path & "\informes.mdb"
'Data1.RecordSource = "inflla"
'Data1.Refresh
data_infu.DatabaseName = App.path & "\informes.mdb"

data_chof.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_chof.RecordSource = "Select * from movil where nromov >=" & 14 & " and nromov <" & 999 & " and medico is null order by chofer"
data_chof.Refresh
If data_chof.Recordset.RecordCount > 0 Then
   data_chof.Recordset.MoveFirst
   Combo3.AddItem "TODOS"
   Do While Not data_chof.Recordset.EOF
      Combo3.AddItem data_chof.Recordset("chofer")
      data_chof.Recordset.MoveNext
   Loop
End If

data_llamod.Connect = "odbc;dsn=" & Xconexrmt & ";"


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
