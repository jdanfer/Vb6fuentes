VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_abmper 
   BackColor       =   &H00FF0000&
   Caption         =   "Mantenimiento de datos del personal"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_abmper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   13395
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
      Height          =   390
      ItemData        =   "frm_abmper.frx":058A
      Left            =   3480
      List            =   "frm_abmper.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   58
      Top             =   120
      Width           =   1815
   End
   Begin VB.Data data_nroid2 
      Caption         =   "data_nroid2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_us 
      Caption         =   "data_us"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11040
      Top             =   0
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
      Connect         =   "DSN=sappper"
      OLEDBString     =   "DSN=sappper"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmm1 
      Left            =   6120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ver todo el personal"
      Height          =   375
      Left            =   8280
      TabIndex        =   42
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5280
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_cargo 
      Caption         =   "data_cargo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_buscap 
      Caption         =   "data_buscap"
      Connect         =   "odbc;dsn=sappper"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * from personas order by cedtext"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox t_busca 
      Height          =   390
      Left            =   4680
      TabIndex        =   37
      Top             =   3720
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      ItemData        =   "frm_abmper.frx":05A4
      Left            =   2160
      List            =   "frm_abmper.frx":05AE
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   3720
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_abmper.frx":05C4
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "frm_abmper.frx":05DE
      TabIndex        =   34
      Top             =   4080
      Width           =   12375
   End
   Begin VB.Data data_per 
      Caption         =   "data_per"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   6735
      Left            =   12600
      TabIndex        =   28
      Top             =   240
      Width           =   735
      Begin VB.CommandButton b_archivos 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":14BD
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Archivos alojados en el servidor"
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton b_evalu2 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":1A47
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Evaluación del desempeño"
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton b_eli 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":1FD1
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton b_imp 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":255B
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Informe ordenado por apellido"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton b_can 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":2AE5
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton b_gua 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":306F
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton b_ed 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":35F9
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton b_nue 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   120
         Picture         =   "frm_abmper.frx":3B83
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos"
      Enabled         =   0   'False
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12495
      Begin VB.ComboBox Combo3 
         Height          =   390
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Data data_solobus 
         Caption         =   "data_solobus"
         Connect         =   "ODBC;DSN=sappper;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox t_profe 
         Height          =   390
         Left            =   6480
         TabIndex        =   54
         Top             =   3840
         Width           =   3975
      End
      Begin VB.TextBox t_nro 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   10680
         TabIndex        =   52
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox t_tel 
         Height          =   390
         Left            =   8520
         TabIndex        =   49
         Top             =   4320
         Width           =   3255
      End
      Begin VB.TextBox t_dir 
         Height          =   390
         Left            =   2280
         TabIndex        =   47
         Top             =   4320
         Width           =   4455
      End
      Begin VB.TextBox t_arch 
         Height          =   390
         Left            =   7680
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton b_bfto 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   11400
         Picture         =   "frm_abmper.frx":410D
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Guardar la foto seleccionada"
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton b_busfto 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   10680
         Picture         =   "frm_abmper.frx":4697
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Buscar una foto"
         Top             =   2520
         Width           =   375
      End
      Begin VB.ComboBox cbodep 
         Height          =   390
         ItemData        =   "frm_abmper.frx":4C21
         Left            =   2280
         List            =   "frm_abmper.frx":4C2E
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Data data_jefes 
         Caption         =   "data_jefes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mfbaja 
         Height          =   375
         Left            =   8880
         TabIndex        =   27
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboestudios 
         Height          =   390
         ItemData        =   "frm_abmper.frx":4C5C
         Left            =   2280
         List            =   "frm_abmper.frx":4C6F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ComboBox cbohijo 
         Height          =   390
         ItemData        =   "frm_abmper.frx":4CA7
         Left            =   10200
         List            =   "frm_abmper.frx":4CB1
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3960
         Width           =   1095
      End
      Begin VB.ComboBox cbosexo 
         Height          =   390
         ItemData        =   "frm_abmper.frx":4CBD
         Left            =   5880
         List            =   "frm_abmper.frx":4CC7
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ComboBox cboestado 
         Height          =   390
         ItemData        =   "frm_abmper.frx":4CE0
         Left            =   2520
         List            =   "frm_abmper.frx":4CF3
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3360
         Width           =   2535
      End
      Begin MSMask.MaskEdBox mfnac 
         Height          =   375
         Left            =   8760
         TabIndex        =   17
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbojefe 
         Height          =   390
         Left            =   2280
         TabIndex        =   15
         Top             =   3360
         Width           =   3735
      End
      Begin VB.ComboBox cbocargo 
         Height          =   390
         Left            =   2280
         TabIndex        =   13
         Top             =   2160
         Width           =   4695
      End
      Begin MSMask.MaskEdBox ming 
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_apel2 
         Height          =   390
         Left            =   6720
         TabIndex        =   9
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox t_apel1 
         Height          =   390
         Left            =   2280
         TabIndex        =   8
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox t_nom2 
         Height          =   390
         Left            =   6720
         TabIndex        =   6
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox t_nom1 
         Height          =   390
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox t_codc 
         Height          =   390
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Jefatura:"
         Height          =   375
         Left            =   240
         TabIndex        =   59
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label labid2 
         Height          =   495
         Left            =   4800
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Profesión:"
         Height          =   375
         Left            =   5160
         TabIndex        =   53
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NRO.FUNCIONARIO:"
         Height          =   375
         Left            =   9840
         TabIndex        =   51
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Teléfono/s:"
         Height          =   375
         Left            =   6960
         TabIndex        =   48
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dirección:"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   2295
         Left            =   10680
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tipo de contrato:"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha de BAJA:"
         Height          =   375
         Left            =   6840
         TabIndex        =   26
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nivel de estudios:"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hijos?"
         Height          =   375
         Left            =   8040
         TabIndex        =   22
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sexo:"
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Estado civil:"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nacimiento:"
         Height          =   375
         Left            =   8040
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Jefe Responsable:"
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cargo:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ingreso:"
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Apellidos:"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nombres:"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CEDULA:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione período de evaluación:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080FFFF&
      Caption         =   "Haga doble click para seleccionar un funcionario"
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   6840
      Width           =   7455
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Buscar por:"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   8160
      Picture         =   "frm_abmper.frx":4D28
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1815
   End
End
Attribute VB_Name = "frm_abmper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdfpath, pdfpath1 As String
Public pdffile As ADODB.Stream



Private Sub b_archivos_Click()
'Dim Xcc As String
'Xcc = t_ced.Text & t_codc.Text

'Wxelnrocedev = Val(Xcc)

'frm_archper.Show vbModal

End Sub

Private Sub b_bfto_Click()
b_bfto.Enabled = False
b_busfto.Enabled = False

If t_ced.Text <> "" And t_codc.Text <> "" Then
   Adodc1.RecordSource = "Select * from personas where id2 =" & Val(labid2.Caption)
   Adodc1.Refresh
   If Adodc1.Recordset.RecordCount > 0 Then
      If pdfpath <> "" Then
'         Adodc1.Recordset.AddNew
         Set pdffile = New ADODB.Stream
         pdffile.Type = adTypeBinary
         pdffile.Open
         pdffile.LoadFromFile pdfpath
         Adodc1.Recordset.Fields("foto") = pdffile.Read
         Adodc1.Recordset.Update
         pdffile.Close
         Set pdffile = Nothing
'        Kill "d:\laboratorios\" & t_nom.Text & ".pdf"
         MsgBox "Guardado"
         b_can_Click
      Else
         MsgBox "No hay archivo"
      End If
   Else
      MsgBox "No se encuentra persona"
   End If
Else
   MsgBox "Seleccione un documento"
End If
b_bfto.Enabled = True
b_busfto.Enabled = True

End Sub

Private Sub b_busfto_Click()
With cmm1
     .FileName = ""
     .Filter = "JPG (*.jpg;) | *.jpg;"
     .ShowOpen
     If Len(.FileName) <> 0 Then
        pdfpath = .FileName
        pdfpath1 = .FileTitle
        t_arch.Text = .FileTitle
     End If
'     t_id.Text = 10
End With

End Sub

Private Sub b_can_Click()
XAlta = 0
borracamp
b_nue.Enabled = True
b_gua.Enabled = False
b_can.Enabled = False
b_ed.Enabled = True
b_eli.Enabled = True
DBGrid1.Enabled = True
Frame1.Enabled = False


End Sub

Private Sub b_ed_Click()
XAlta = 0
Frame1.Enabled = True
t_nom1.SetFocus
b_nue.Enabled = False
b_gua.Enabled = True
b_can.Enabled = True
b_ed.Enabled = False
b_eli.Enabled = False
DBGrid1.Enabled = False
b_evalu2.Enabled = False
b_archivos.Enabled = False

End Sub

Private Sub b_eli_Click()
If WElusuario = "JFERNAN" Then
   If t_ced.Text <> "" And t_codc.Text <> "" Then
      Xlac = t_ced.Text & t_codc.Text
      data_per.RecordSource = "Select * from personas where id =" & Val(Xlac)
      data_per.Refresh
      If data_per.Recordset.RecordCount > 0 Then
         data_per.Recordset.Delete
         MsgBox "Registro eliminado"
         data_buscap.Refresh
      End If
   End If
End If
         

End Sub

Private Sub b_evalu1_Click()

End Sub

Private Sub b_evalu2_Click()
b_evalu2.Enabled = False

If frm_evalua1.Visible = True Then
   MsgBox "Ya está abierto"
Else

    If cbocargo.Text <> "" And t_ced.Text <> "" Then
       data_cargo.RecordSource = "Select * from cargos where descrip ='" & cbocargo.Text & "'"
       data_cargo.Refresh
       If data_cargo.Recordset.RecordCount > 0 Then
          XWquecargo = data_cargo.Recordset("tipo")
          Wxquepreg = data_cargo.Recordset("codpreg")
          data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
          data_us.Refresh
          If data_us.Recordset.RecordCount > 0 Then
             data_solobus.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
             data_solobus.Refresh
             If data_solobus.Recordset.RecordCount > 0 Then
                Wxeljefeid = data_solobus.Recordset("id")
             Else
                MsgBox "No se encontró el usuario, comunique al administrador"
                Unload Me
             End If
          Else
             MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
             Unload Me
          End If
        
          If IsNull(data_solobus.Recordset("cargod")) = False Then
             data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_solobus.Recordset("cargod") & "'"
             data_cargo.Refresh
             If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Then
                XWquecargo = data_cargo.Recordset("tipo")
'                Wxquepreg = data_cargo.Recordset("codpreg")
             Else
                XWquecargo = data_cargo.Recordset("tipo")
             End If
          Else
        
          End If

          Wxelnroid2 = data_buscap.Recordset("id2")
          Wxelnrocedev = data_buscap.Recordset("id")
'          Wxeljefeid = data_buscap.Recordset("jefe")
'          If IsNull(data_buscap.Recordset("jefe")) = False Then
'             Wxeljefeid = data_buscap.Recordset("jefe")
'          Else
'             Wxeljefeid = 0
'          End If
    '      If XWquecargo = 1 Then
    '         Wxeljefeid = data_cargo.Recordset("id")
    '      End If
          frm_evalua1.Show vbModal
       Else
          MsgBox "No se encuentra cargo, comunique a Administración", vbInformation
          XWquecargo = 0
          Wxquepreg = 0
          Wxelnroid2 = 0
       End If
    Else
       MsgBox "No tiene registrado el cargo, comunique a Administración. O verifique si seleccionó funcionario.", vbInformation
       XWquecargo = 0
       Wxquepreg = 0
       Wxelnroid2 = 0
    End If
End If

b_evalu2.Enabled = True


End Sub

Private Sub b_gua_Click()
On Error GoTo Errevalgraba

b_gua.Enabled = False

If XAlta = 1 Then
   If t_ced.Text <> "" And t_codc.Text <> "" Then
      Xlac = t_ced.Text & t_codc.Text
      data_per.RecordSource = "Select * from personas where id =" & Val(Xlac)
      data_per.Refresh
      If data_per.Recordset.RecordCount > 0 Then
         MsgBox "Ya existe funcionario con ésta cédula", vbCritical
      Else
         If t_nom1.Text <> "" And t_apel1.Text <> "" Then
            data_per.Recordset.AddNew
            data_per.Recordset("id") = Val(Xlac)
            data_per.Recordset("id2") = labid2.Caption
            data_per.Recordset("cedtext") = t_ced.Text & "-" & t_codc.Text
            data_per.Recordset("nom1") = t_nom1.Text
            data_per.Recordset("ape1") = t_apel1.Text
            If t_dir.Text <> "" Then
               data_per.Recordset("direc") = t_dir.Text
            End If
            If t_tel.Text <> "" Then
               data_per.Recordset("tel") = t_tel.Text
            End If
            If t_nom2.Text <> "" Then
               data_per.Recordset("nom2") = t_nom2.Text
            End If
            If t_profe.Text <> "" Then
               data_per.Recordset("profesion") = t_profe.Text
            End If
            If t_nro.Text <> "" Then
               data_per.Recordset("nro") = t_nro.Text
            End If
            If t_apel2.Text <> "" Then
               data_per.Recordset("ape2") = t_apel2.Text
            End If
            If ming.Text <> "__/__/____" Then
               data_per.Recordset("fecing") = Format(ming.Text, "dd/mm/yyyy")
            End If
            If cbocargo.ListIndex >= 0 Then
               data_cargo.RecordSource = "Select * from cargos where descrip ='" & cbocargo.Text & "'"
               data_cargo.Refresh
               If data_cargo.Recordset.RecordCount > 0 Then
                  data_per.Recordset("cargo") = data_cargo.Recordset("id")
                  data_per.Recordset("cargod") = data_cargo.Recordset("descrip")
               End If
            End If
            If cbojefe.ListIndex >= 0 Then
               data_jefes.RecordSource = "Select * from jefes where descrip ='" & cbojefe.Text & "'"
               data_jefes.Refresh
               If data_jefes.Recordset.RecordCount > 0 Then
                  data_per.Recordset("jefe") = data_jefes.Recordset("id")
                  data_per.Recordset("jefed") = data_jefes.Recordset("descrip")
               End If
            End If
            If mfnac.Text <> "__/__/____" Then
               data_per.Recordset("fecnac") = Format(mfnac.Text, "dd/mm/yyyy")
            End If
            data_per.Recordset("estcivil") = cboestado.ListIndex
            data_per.Recordset("sexo") = cbosexo.ListIndex
            data_per.Recordset("hijos") = cbohijo.ListIndex
            data_per.Recordset("tipo") = cbodep.ListIndex
            data_per.Recordset("nivelest") = cboestudios.ListIndex
            If mfbaja.Text <> "__/__/____" Then
               data_per.Recordset("fechabaja") = Format(mfbaja.Text, "dd/mm/yyyy")
            End If
            data_per.Recordset.Update
            b_nue.Enabled = True
            b_gua.Enabled = False
            b_can.Enabled = False
            b_ed.Enabled = True
            b_eli.Enabled = True
            DBGrid1.Enabled = True
            b_evalu2.Enabled = True
            b_archivos.Enabled = True
            
            data_buscap.Refresh
            Frame1.Enabled = False
            XAlta = 0
         Else
            MsgBox "Ingrese nombre y apellido", vbInformation
         End If
      End If
   Else
      MsgBox "Ingrese cédula", vbInformation
   End If
Else
   If t_ced.Text <> "" And t_codc.Text <> "" Then
'      Xlac = t_ced.Text & t_codc.Text
      data_per.RecordSource = "Select * from personas where id2 =" & Val(labid2.Caption)
      data_per.Refresh
      If data_per.Recordset.RecordCount > 0 Then
         If t_nom1.Text <> "" And t_apel1.Text <> "" Then
            data_per.Recordset.Edit
            data_per.Recordset("nom1") = t_nom1.Text
            data_per.Recordset("ape1") = t_apel1.Text
            If t_dir.Text <> "" Then
               data_per.Recordset("direc") = t_dir.Text
            End If
            If t_tel.Text <> "" Then
               data_per.Recordset("tel") = t_tel.Text
            End If
            If t_nom2.Text <> "" Then
               data_per.Recordset("nom2") = t_nom2.Text
            Else
               If IsNull(data_per.Recordset("nom2")) = False Then
                  data_per.Recordset("nom2") = Null
               End If
            End If
            If t_apel2.Text <> "" Then
               data_per.Recordset("ape2") = t_apel2.Text
            Else
               If IsNull(data_per.Recordset("ape2")) = False Then
                  data_per.Recordset("ape2") = Null
               End If
            End If
            If cbocargo.ListIndex >= 0 Then
               data_cargo.RecordSource = "Select * from cargos where descrip ='" & cbocargo.Text & "'"
               data_cargo.Refresh
               If data_cargo.Recordset.RecordCount > 0 Then
                  data_per.Recordset("cargo") = data_cargo.Recordset("id")
                  data_per.Recordset("cargod") = data_cargo.Recordset("descrip")
               End If
            End If
            data_per.Recordset.Update
            b_nue.Enabled = True
            b_gua.Enabled = False
            b_can.Enabled = False
            b_ed.Enabled = True
            b_eli.Enabled = True
            b_evalu2.Enabled = True
            b_archivos.Enabled = True
            DBGrid1.Enabled = True
            Frame1.Enabled = False
            data_buscap.Refresh
            XAlta = 0
         Else
            MsgBox "Ingrese nombre y apellido", vbInformation
         End If
      Else
         MsgBox "No se encuentra documento", vbCritical
      End If
   Else
      MsgBox "Ingrese cédula", vbInformation
   End If
End If

b_gua.Enabled = True

Exit Sub

Errevalgraba:
             If Err.Number = 3155 Then
                MsgBox "Error al grabar, verifique datos"
             Else
                MsgBox "Error al grabar"
             End If


End Sub

Private Sub b_imp_Click()
If WElusuario = "JFERNAN" Or WElusuario = "SPEREZ" Or WElusuario = "DARIOH" Then

    data_inf.RecordSource = "infcli"
    data_inf.Refresh
    If data_inf.Recordset.RecordCount > 0 Then
       data_inf.Recordset.MoveFirst
       Do While Not data_inf.Recordset.EOF
          data_inf.Recordset.Delete
          data_inf.Recordset.MoveNext
       Loop
       data_inf.Refresh
    End If
    data_buscap.Recordset.MoveFirst
    Do While Not data_buscap.Recordset.EOF
       data_inf.Recordset.AddNew
       data_inf.Recordset("cl_telefon") = data_buscap.Recordset("cedtext")
       data_inf.Recordset("cl_apellid") = data_buscap.Recordset("ape1")
       data_inf.Recordset("cl_direcci") = data_buscap.Recordset("nom1")
       data_inf.Recordset("cl_localid") = data_buscap.Recordset("nom2")
       data_inf.Recordset("cl_nombre") = data_buscap.Recordset("ape2")
       data_inf.Recordset("cl_fecing") = data_buscap.Recordset("fecing")
       data_inf.Recordset("cl_email") = Mid(data_buscap.Recordset("cargod"), 1, 30)
       If IsNull(data_buscap.Recordset("tipo")) = True Then
          data_inf.Recordset("cl_codigo") = 0
          data_inf.Recordset("cl_celular") = "Dependiente"
       Else
          data_inf.Recordset("cl_codigo") = data_buscap.Recordset("tipo")
          If data_buscap.Recordset("tipo") = 0 Then
             data_inf.Recordset("cl_celular") = "Dependiente"
          End If
          If data_buscap.Recordset("tipo") = 1 Then
             data_inf.Recordset("cl_celular") = "No Dependiente"
          End If
          If data_buscap.Recordset("tipo") = 2 Then
             data_inf.Recordset("cl_celular") = "Tercerizado"
          End If
          
       End If
          
       data_inf.Recordset.Update
       data_inf.Refresh
       data_buscap.Recordset.MoveNext
       
    Loop
    data_buscap.Recordset.MoveFirst
    'MsgBox "Proceso Terminado"
    data_inf.RecordSource = "Select * from infcli"
    data_inf.Refresh
    
    cr1.ReportFileName = App.path & "\infperso.rpt"
    cr1.Action = 1
Else
    MsgBox "Opción no autorizada, puede realizar informes desde la ficha de evaluación.", vbInformation
End If


End Sub

Private Sub b_nue_Click()
XAlta = 1
Frame1.Enabled = True
borracamp
t_ced.SetFocus
b_nue.Enabled = False
b_gua.Enabled = True
b_can.Enabled = True
b_ed.Enabled = False
b_eli.Enabled = False
DBGrid1.Enabled = False
b_evalu2.Enabled = False
b_archivos.Enabled = False

data_nroid2.RecordSource = "Select * from personas order by id2 DESC"
data_nroid2.Refresh
If data_nroid2.Recordset.RecordCount > 0 Then
   labid2.Caption = data_nroid2.Recordset("id2") + 1
Else
   labid2.Caption = 1
End If


End Sub

Private Sub cbocargo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   cbojefe.SetFocus
'End If

End Sub

Private Sub cbodep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_profe.SetFocus
End If

End Sub

Private Sub cboestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbosexo.SetFocus
End If

End Sub

Private Sub cboestudios_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfbaja.SetFocus
End If

End Sub

Private Sub cbohijo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboestudios.SetFocus
End If

End Sub

Private Sub cbojefe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfnac.SetFocus
End If

End Sub

Private Sub cbosexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbohijo.SetFocus
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   data_buscap.RecordSource = "Select * from personas order by id"
Else
   data_buscap.RecordSource = "Select * from personas where tipo =" & 1 & " order by id"
End If
data_buscap.Refresh

End Sub

Private Sub Combo2_Click()
'094477160
'Combo2.Clear

If Combo2.Text = "2016" Then
   data_solobus.Connect = "ODBC;DSN=eval2015;"
   data_buscap.Connect = "ODBC;DSN=eval2015;"
   Adodc1.ConnectionString = "ODBC;DSN=eval2015;"
   
   data_per.Connect = "ODBC;DSN=eval2015;"
   data_cargo.Connect = "ODBC;DSN=eval2015;"
   data_cargo.RecordSource = "Select * from cargos order by descrip"
   data_cargo.Refresh
   If data_cargo.Recordset.RecordCount > 0 Then
      data_cargo.Recordset.MoveFirst
      Do While Not data_cargo.Recordset.EOF
         cbocargo.AddItem data_cargo.Recordset("descrip")
         data_cargo.Recordset.MoveNext
      Loop
   End If
   data_us.Connect = "odbc;dsn=" & Xconexrmt & ";"
   
   data_nroid2.Connect = "ODBC;DSN=eval2015;"
    
   data_jefes.Connect = "ODBC;DSN=eval2015;"
   data_jefes.RecordSource = "Select * from jefes order by id"
   data_jefes.Refresh
   If data_jefes.Recordset.RecordCount > 0 Then
      data_jefes.Recordset.MoveFirst
      Do While Not data_jefes.Recordset.EOF
         cbojefe.AddItem data_jefes.Recordset("descrip")
         data_jefes.Recordset.MoveNext
      Loop
   End If
    
   data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
   data_us.Refresh
   If data_us.Recordset.RecordCount > 0 Then
      data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
      data_buscap.Refresh
      If data_buscap.Recordset.RecordCount > 0 Then
         Wxeljefeid = data_buscap.Recordset("id")
      Else
         MsgBox "No se encontró el usuario, comunique al administrador"
         Unload Me
      End If
   Else
      MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
      Unload Me
   End If
    
   data_inf.DatabaseName = App.path & "\informes.mdb"
   If IsNull(data_buscap.Recordset("cargod")) = False Then
      data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_buscap.Recordset("cargod") & "'"
      data_cargo.Refresh
      If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Or WElusuario = "JFERNAN" Then
         If data_cargo.Recordset("tipo") = 2 And WElusuario <> "JFERNAN" Then
            data_buscap.RecordSource = "Select * from personas where jefed ='" & data_cargo.Recordset("descrip") & "' or cargod ='" & data_cargo.Recordset("descrip") & "' order by id"
            data_buscap.Refresh
            XWquecargo = data_cargo.Recordset("tipo")
            Wxquepreg = data_cargo.Recordset("codpreg")
            If WElusuario = "JFERNAN" Or WElusuario = "DARIOH" Then
               Check1.Visible = True
            Else
               Check1.Visible = False
            End If
         Else
            XWquecargo = data_cargo.Recordset("tipo")
            Wxquepreg = data_cargo.Recordset("codpreg")
            data_buscap.RecordSource = "Select * from personas order by id"
            data_buscap.Refresh
            Check1.Visible = False
         End If
      Else
         XWquecargo = data_cargo.Recordset("tipo")
         Wxquepreg = data_cargo.Recordset("codpreg")
         Check1.Visible = False
         b_nue.Enabled = False
         b_ed.Enabled = False
         b_gua.Enabled = False
         b_can.Enabled = False
         b_imp.Enabled = False
         t_busca.Enabled = False
         b_eli.Enabled = False
         data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
         data_buscap.Refresh
      End If
   Else
   
    
   End If
Else
   data_solobus.Connect = "ODBC;DSN=sappper;"
   data_buscap.Connect = "ODBC;DSN=sappper;"
   Adodc1.ConnectionString = "ODBC;DSN=sappper;"
   
   data_per.Connect = "ODBC;DSN=sappper;"
   data_cargo.Connect = "ODBC;DSN=sappper;"
   data_cargo.RecordSource = "Select * from cargos order by descrip"
   data_cargo.Refresh
   If data_cargo.Recordset.RecordCount > 0 Then
      data_cargo.Recordset.MoveFirst
      Do While Not data_cargo.Recordset.EOF
         cbocargo.AddItem data_cargo.Recordset("descrip")
         data_cargo.Recordset.MoveNext
      Loop
   End If
   data_us.Connect = "odbc;dsn=" & Xconexrmt & ";"
   
   data_nroid2.Connect = "ODBC;DSN=sappper;"
    
   data_jefes.Connect = "ODBC;DSN=sappper;"
   data_jefes.RecordSource = "Select * from jefes order by id"
   data_jefes.Refresh
   If data_jefes.Recordset.RecordCount > 0 Then
      data_jefes.Recordset.MoveFirst
      Do While Not data_jefes.Recordset.EOF
         cbojefe.AddItem data_jefes.Recordset("descrip")
         data_jefes.Recordset.MoveNext
      Loop
   End If
    
   data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
   data_us.Refresh
   If data_us.Recordset.RecordCount > 0 Then
      data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
      data_buscap.Refresh
      If data_buscap.Recordset.RecordCount > 0 Then
         Wxeljefeid = data_buscap.Recordset("id")
      Else
         MsgBox "No se encontró el usuario, comunique al administrador"
         Unload Me
      End If
   Else
      MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
      Unload Me
   End If
    
   data_inf.DatabaseName = App.path & "\informes.mdb"
   If IsNull(data_buscap.Recordset("cargod")) = False Then
      data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_buscap.Recordset("cargod") & "'"
      data_cargo.Refresh
      If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Or WElusuario = "JFERNAN" Then
         If data_cargo.Recordset("tipo") = 2 And WElusuario <> "JFERNAN" Then
            data_buscap.RecordSource = "Select * from personas where jefed ='" & data_cargo.Recordset("descrip") & "' or cargod ='" & data_cargo.Recordset("descrip") & "' order by id"
            data_buscap.Refresh
            XWquecargo = data_cargo.Recordset("tipo")
            Wxquepreg = data_cargo.Recordset("codpreg")
            If WElusuario = "JFERNAN" Or WElusuario = "DARIOH" Then
               Check1.Visible = True
            Else
               Check1.Visible = False
            End If
         Else
            XWquecargo = data_cargo.Recordset("tipo")
            Wxquepreg = data_cargo.Recordset("codpreg")
            data_buscap.RecordSource = "Select * from personas order by id"
            data_buscap.Refresh
            Check1.Visible = False
         End If
      Else
         XWquecargo = data_cargo.Recordset("tipo")
         Wxquepreg = data_cargo.Recordset("codpreg")
         Check1.Visible = False
         b_nue.Enabled = False
         b_ed.Enabled = False
         b_gua.Enabled = False
         b_can.Enabled = False
         b_imp.Enabled = False
         t_busca.Enabled = False
         b_eli.Enabled = False
         data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
         data_buscap.Refresh
      End If
   Else
   
    
   End If

End If
frm_abmper.MousePointer = 0

End Sub

Private Sub DBGrid1_DblClick()
verdatper

End Sub

Private Sub Form_Load()

frm_abmper.MousePointer = 11

data_per.Connect = "ODBC;DSN=sappper;"
data_cargo.Connect = "ODBC;DSN=sappper;"
data_cargo.RecordSource = "Select * from cargos order by descrip"
data_cargo.Refresh
If data_cargo.Recordset.RecordCount > 0 Then
   data_cargo.Recordset.MoveFirst
   Do While Not data_cargo.Recordset.EOF
      cbocargo.AddItem data_cargo.Recordset("descrip")
      Combo3.AddItem data_cargo.Recordset("descrip")
      data_cargo.Recordset.MoveNext
   Loop
End If
data_us.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_nroid2.Connect = "ODBC;DSN=sappper;"

data_jefes.Connect = "ODBC;DSN=sappper;"
data_jefes.RecordSource = "Select * from jefes order by id"
data_jefes.Refresh
If data_jefes.Recordset.RecordCount > 0 Then
   data_jefes.Recordset.MoveFirst
   Do While Not data_jefes.Recordset.EOF
      cbojefe.AddItem data_jefes.Recordset("descrip")
      data_jefes.Recordset.MoveNext
   Loop
End If

data_us.RecordSource = "Select * from cap_ciap where des_cap ='" & WElusuario & "'"
data_us.Refresh
If data_us.Recordset.RecordCount > 0 Then
   data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
   data_buscap.Refresh
   If data_buscap.Recordset.RecordCount > 0 Then
      Wxeljefeid = data_buscap.Recordset("id")
   Else
      MsgBox "No se encontró el usuario, comunique al administrador"
      Unload Me
   End If
Else
   MsgBox "No se encuentra usuario registrado, comunique al administrador", vbInformation
   Unload Me
End If


data_inf.DatabaseName = App.path & "\informes.mdb"
If IsNull(data_buscap.Recordset("cargod")) = False Then
   data_cargo.RecordSource = "Select * from cargos where descrip ='" & data_buscap.Recordset("cargod") & "'"
   data_cargo.Refresh
   If data_cargo.Recordset("tipo") = 2 Or data_cargo.Recordset("tipo") = 3 Or WElusuario = "DARIOH" Or WElusuario = "JFERNAN" Then
      If data_cargo.Recordset("tipo") = 2 And WElusuario <> "JFERNAN" Then
         data_buscap.RecordSource = "Select * from personas where jefed ='" & data_cargo.Recordset("descrip") & "' or cargod ='" & data_cargo.Recordset("descrip") & "' order by id"
         data_buscap.Refresh
         XWquecargo = data_cargo.Recordset("tipo")
         Wxquepreg = data_cargo.Recordset("codpreg")
         If WElusuario = "JFERNAN" Or WElusuario = "DARIOH" Then
            Check1.Visible = True
         Else
            Check1.Visible = False
         End If
      Else
         XWquecargo = data_cargo.Recordset("tipo")
         Wxquepreg = data_cargo.Recordset("codpreg")
         data_buscap.RecordSource = "Select * from personas order by id"
         data_buscap.Refresh
         Check1.Visible = False
      End If
   Else
      XWquecargo = data_cargo.Recordset("tipo")
      Wxquepreg = data_cargo.Recordset("codpreg")
      Check1.Visible = False
      b_nue.Enabled = False
      b_ed.Enabled = False
      b_gua.Enabled = False
      b_can.Enabled = False
      b_imp.Enabled = False
      t_busca.Enabled = False
      b_eli.Enabled = False
      data_buscap.RecordSource = "Select * from personas where id =" & Val(data_us.Recordset("cod_cap"))
      data_buscap.Refresh
   End If
Else


End If

frm_abmper.MousePointer = 0

Combo2.ListIndex = 1

End Sub

Private Sub Form_Resize()
With Image2
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub mfbaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbodep.SetFocus
End If

End Sub

Private Sub mfnac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboestado.SetFocus
End If

End Sub

Private Sub ming_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocargo.SetFocus
End If

End Sub

Private Sub t_apel1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_apel2.SetFocus
End If

End Sub

Private Sub t_apel2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocargo.SetFocus
End If

End Sub

Private Sub t_busca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo1.ListIndex = 0 Then
      If Check1.Value = 1 Then
         data_buscap.RecordSource = "Select * from personas where cedtext >='" & t_busca.Text & "' order by cedtext"
         data_buscap.Refresh
      Else
         ' ve solo las personas que tiene a cargo + sí mismo
         data_buscap.RecordSource = "SELECT * from personas WHERE " _
                                    & " (jefed = '" & data_cargo.Recordset("descrip") & "' " _
                                    & " or id = " & data_us.Recordset("cod_cap") & ") " _
                                    & " and cedtext Like  '*" & t_busca.Text & "*'"
         data_buscap.Refresh
      End If
   Else
      If Combo1.ListIndex = 1 Then
         If Check1.Value = 1 Then
            data_buscap.RecordSource = "Select * from personas where ape1 >='" & t_busca.Text & "' order by ape1"
            data_buscap.Refresh
         Else
            ' ve solo las personas que tiene a cargo + sí mismo
            data_buscap.RecordSource = "SELECT * from personas WHERE " _
                                    & " (jefed = '" & data_cargo.Recordset("descrip") & "' " _
                                    & " or id = " & data_us.Recordset("cod_cap") & ") " _
                                    & " and ape1 Like  '*" & t_busca.Text & "*'"
            data_buscap.Refresh
         End If
      End If
   End If
   DBGrid1.SetFocus
End If

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codc.SetFocus
End If

End Sub

Private Sub t_codc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom1.SetFocus
End If

End Sub

Private Sub t_codc_LostFocus()
Dim Xlac As String
If XAlta = 1 Then
    If t_ced.Text <> "" And t_codc.Text <> "" Then
       Xlac = t_ced.Text & t_codc.Text
       data_per.RecordSource = "Select * from personas where id =" & Val(Xlac)
       data_per.Refresh
       If data_per.Recordset.RecordCount > 0 Then
          MsgBox "Ya existe funcionario con ésta cédula", vbCritical
       End If
    Else
       MsgBox "No ingresó documento", vbCritical
    End If
End If


End Sub

Private Sub t_dir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_tel.SetFocus
End If

End Sub

Private Sub t_nom1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom2.SetFocus
End If

End Sub

Private Sub t_nom2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_apel1.SetFocus
End If

End Sub

Public Sub borracamp()
t_ced.Text = ""
t_codc.Text = ""
t_nom1.Text = ""
t_nom2.Text = ""
t_apel1.Text = ""
t_apel2.Text = ""
t_nro.Text = ""
t_profe.Text = ""
ming.Text = "__/__/____"
cbocargo.Text = ""
cbojefe.Text = ""
mfnac.Text = "__/__/____"
cboestado.ListIndex = -1
cbosexo.ListIndex = -1
cbohijo.ListIndex = -1
cboestudios.ListIndex = -1
mfbaja.Text = "__/__/____"
cbodep.ListIndex = -1
t_dir.Text = ""
t_tel.Text = ""


End Sub

Public Sub verdatper()
If IsNull(data_buscap.Recordset("id")) = False Then
   If Len(data_buscap.Recordset("id")) = 6 Then
      t_ced.Text = Mid(data_buscap.Recordset("id"), 1, 5)
      t_codc.Text = Mid(data_buscap.Recordset("id"), 6, 1)
   End If
   If Len(data_buscap.Recordset("id")) = 7 Then
      t_ced.Text = Mid(data_buscap.Recordset("id"), 1, 6)
      t_codc.Text = Mid(data_buscap.Recordset("id"), 7, 1)
   End If
   If Len(data_buscap.Recordset("id")) = 8 Then
      t_ced.Text = Mid(data_buscap.Recordset("id"), 1, 7)
      t_codc.Text = Mid(data_buscap.Recordset("id"), 8, 1)
   End If
Else
   t_ced.Text = 0
   t_codc.Text = 0
End If
If IsNull(data_buscap.Recordset("nom1")) = False Then
   t_nom1.Text = data_buscap.Recordset("nom1")
Else
   t_nom1.Text = ""
End If
If IsNull(data_buscap.Recordset("id2")) = False Then
   labid2.Caption = data_buscap.Recordset("id2")
Else
   labid2.Caption = 1
End If

If IsNull(data_buscap.Recordset("nom2")) = False Then
   t_nom2.Text = data_buscap.Recordset("nom2")
Else
   t_nom2.Text = ""
End If
If IsNull(data_buscap.Recordset("ape1")) = False Then
   t_apel1.Text = data_buscap.Recordset("ape1")
Else
   t_apel1.Text = ""
End If
If IsNull(data_buscap.Recordset("nro")) = False Then
   t_nro.Text = data_buscap.Recordset("nro")
Else
   t_nro.Text = ""
End If
If IsNull(data_buscap.Recordset("profesion")) = False Then
   t_profe.Text = data_buscap.Recordset("profesion")
Else
   t_profe.Text = ""
End If
If IsNull(data_buscap.Recordset("ape2")) = False Then
   t_apel2.Text = data_buscap.Recordset("ape2")
Else
   t_apel2.Text = ""
End If
If IsNull(data_buscap.Recordset("fecnac")) = False Then
   mfnac.Text = Format(data_buscap.Recordset("fecnac"), "dd/mm/yyyy")
Else
   mfnac.Text = "__/__/____"
End If
If IsNull(data_buscap.Recordset("fecing")) = False Then
   ming.Text = Format(data_buscap.Recordset("fecing"), "dd/mm/yyyy")
Else
   ming.Text = "__/__/____"
End If
If IsNull(data_buscap.Recordset("fechabaja")) = False Then
   mfbaja.Text = Format(data_buscap.Recordset("fechabaja"), "dd/mm/yyyy")
Else
   mfbaja.Text = "__/__/____"
End If
If IsNull(data_buscap.Recordset("cargod")) = False Then
   cbocargo.Text = data_buscap.Recordset("cargod")
Else
   cbocargo.Text = ""
End If
If IsNull(data_buscap.Recordset("jefed")) = False Then
   cbojefe.Text = data_buscap.Recordset("jefed")
Else
   cbojefe.Text = ""
End If
If IsNull(data_buscap.Recordset("estcivil")) = False Then
   cboestado.ListIndex = data_buscap.Recordset("estcivil")
Else
   cboestado.ListIndex = -1
End If
If IsNull(data_buscap.Recordset("tipo")) = False Then
   cbodep.ListIndex = data_buscap.Recordset("tipo")
Else
   cbodep.ListIndex = 0
End If

If IsNull(data_buscap.Recordset("sexo")) = False Then
   cbosexo.ListIndex = data_buscap.Recordset("sexo")
Else
   cbosexo.ListIndex = -1
End If
If IsNull(data_buscap.Recordset("hijos")) = False Then
   cbohijo.ListIndex = data_buscap.Recordset("hijos")
Else
   cbohijo.ListIndex = -1
End If
If IsNull(data_buscap.Recordset("nivelest")) = False Then
   cboestudios.ListIndex = data_buscap.Recordset("nivelest")
Else
   cboestudios.ListIndex = -1
End If
If IsNull(data_buscap.Recordset("direc")) = False Then
   t_dir.Text = data_buscap.Recordset("direc")
Else
   t_dir.Text = ""
End If
If IsNull(data_buscap.Recordset("tel")) = False Then
   t_tel.Text = data_buscap.Recordset("tel")
Else
   t_tel.Text = ""
End If


Set pdffile = New ADODB.Stream
pdffile.Type = adTypeBinary
pdffile.Open
If IsNull(data_buscap.Recordset("foto")) = False Then
   pdffile.Write data_buscap.Recordset("foto").Value
   Dim pdfname As String
   pdfname = "temporal"
   pdffile.SaveToFile "" & App.path & "\fotos\" & pdfname & ".jpg", adSaveCreateOverWrite
   pdffile.Close
   Set pdffile = Nothing
   Image1.Picture = LoadPicture(App.path & "\fotos\temporal.jpg")

'AcroRd32
'      Shell Data1.Recordset("desc") & " " & App.Path & "\laboratorio\temporal" & ".pdf", vbMaximizedFocus
'      Shell "c:\Program Files (x86)\Adobe\Reader 9.0\Reader\AcroRd32.exe " & App.Path & "\laboratorio\" & pdfname & ".pdf", vbMaximizedFocus

Else
   Image1.Picture = LoadPicture(App.path & "\fotos\sinfoto.jpg")
   
'   MsgBox "no hay archivo"
End If

   

End Sub

Private Sub t_nro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom1.SetFocus
End If

End Sub

Private Sub t_profe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_dir.SetFocus
End If

End Sub

Private Sub t_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If b_gua.Enabled = True Then
      b_gua.SetFocus
   End If
End If

End Sub
