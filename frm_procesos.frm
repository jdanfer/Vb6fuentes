VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11235
   Icon            =   "frm_procesos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc data_mdbn 
      Height          =   375
      Left            =   1800
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_mdbn"
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
   Begin VB.Data data_mdb 
      Caption         =   "data_mdb"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Promo Ami"
      Height          =   615
      Left            =   10440
      TabIndex        =   50
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Resp LINMM"
      Height          =   495
      Left            =   600
      TabIndex        =   49
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Data data_deu 
      Caption         =   "data_deu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
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
      Top             =   8160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Deb Visa /BROU"
      Height          =   615
      Left            =   7680
      TabIndex        =   48
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Udemm"
      Height          =   735
      Left            =   8160
      TabIndex        =   44
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Dig Verific/ Corral"
      Height          =   1095
      Left            =   10320
      TabIndex        =   43
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command37 
      Caption         =   "SMI Espec 0320"
      Height          =   855
      Left            =   10320
      TabIndex        =   42
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox t_ivasql 
      Height          =   375
      Left            =   9600
      TabIndex        =   41
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox t_iva 
      Height          =   285
      Left            =   9600
      TabIndex        =   40
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Llam 4/2/20"
      Height          =   495
      Left            =   2400
      TabIndex        =   39
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command35 
      Caption         =   "resplla a resplla"
      Height          =   615
      Left            =   9840
      TabIndex        =   38
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command34 
      Caption         =   "EMI controles 5.22"
      Height          =   615
      Left            =   9840
      TabIndex        =   37
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Receta Elect informes"
      Height          =   735
      Left            =   5640
      TabIndex        =   36
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command32 
      Caption         =   "UDE"
      Height          =   615
      Left            =   10440
      TabIndex        =   35
      Top             =   2880
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cap_ciap"
      Top             =   6000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_procesos.frx":0442
      Height          =   2055
      Left            =   600
      OleObjectBlob   =   "frm_procesos.frx":0456
      TabIndex        =   34
      Top             =   6240
      Width           =   7455
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Precios conv 29/12"
      Height          =   615
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Control Deu1218"
      Height          =   495
      Left            =   3000
      TabIndex        =   32
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command29 
      Caption         =   "evalua 19/2"
      Height          =   615
      Left            =   2400
      TabIndex        =   31
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Pro Amig / Ced a string"
      Height          =   855
      Left            =   2400
      TabIndex        =   30
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Arqueos 11/1/22"
      Height          =   495
      Left            =   3240
      TabIndex        =   29
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Mut infos edad"
      Height          =   495
      Left            =   5040
      TabIndex        =   28
      Top             =   120
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   495
      Left            =   3240
      TabIndex        =   27
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Controles Ced HCE"
      Height          =   495
      Left            =   5760
      TabIndex        =   26
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command24 
      Caption         =   "13-3 Estud Borra 02"
      Height          =   975
      Left            =   9480
      TabIndex        =   25
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Aocios no se atend HC depuracion 23/8"
      Height          =   615
      Left            =   600
      TabIndex        =   23
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command22 
      Caption         =   "cli repet / Borra FactE 7/12"
      Height          =   855
      Left            =   9240
      TabIndex        =   22
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Comentarios version"
      Height          =   495
      Left            =   7560
      TabIndex        =   21
      Top             =   360
      Width           =   1575
   End
   Begin VB.Data data_rec 
      Caption         =   "data_rec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_temp 
      Caption         =   "data_temp"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   "D:\sappmys\sappmysql"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "mutual"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_emi3 
      Caption         =   "data_emi3"
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
      Width           =   2775
   End
   Begin VB.CommandButton Command20 
      Caption         =   "14Modif conv cli, mes pag"
      Height          =   855
      Left            =   9360
      TabIndex        =   20
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Varios 021221"
      Height          =   735
      Left            =   10080
      TabIndex        =   19
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "24Deud/ Comer linmedic"
      Height          =   615
      Left            =   9360
      TabIndex        =   18
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Inf.Consult 4/19"
      Height          =   615
      Left            =   7680
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      Caption         =   "BAJAS auto/ CCOU SJ"
      Height          =   615
      Left            =   7560
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton Command15 
      Caption         =   "CGALICIA"
      Height          =   855
      Left            =   7560
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6960
      TabIndex        =   14
      Text            =   "28/01/2010"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Reservas cht"
      Height          =   615
      Left            =   5760
      TabIndex        =   13
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "RedPagos"
      Height          =   615
      Left            =   7560
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "20Borrar llamados o edit"
      Height          =   615
      Left            =   5640
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Códigos de Ced. HCE"
      Height          =   855
      Left            =   5640
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Arreglar MAM"
      Height          =   615
      Left            =   5640
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.CommandButton Command9 
      Caption         =   "RECUPERA"
      Height          =   615
      Left            =   3480
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Tesorero P815"
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Estudios modif"
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir..."
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Data datai 
      Caption         =   "datai"
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
      Top             =   5640
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Whatsapp 220322"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Arregla Caja/linmm/pagos"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Llamados"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BAJAS x motivo"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Em Pruebas rapidas 22/11"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Data data_sql 
      Caption         =   "data_sql"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Label labd 
      Height          =   375
      Left            =   8640
      TabIndex        =   47
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label labm 
      Height          =   375
      Left            =   9480
      TabIndex        =   46
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label laba 
      Height          =   375
      Left            =   8280
      TabIndex        =   45
      Top             =   7200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400


Private Sub Command1_Click()
Dim xim As Double
data_mdb.Connect = "ODBC;DSN=sappnew;"
data_mdb.RecordSource = "Select * from emi1121 where documento =" & 1192183
data_mdb.Refresh
data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.RecordSource = "deudas"
data_sql.Refresh

data_sql.Recordset.AddNew
data_sql.Recordset("cod_cnv") = data_mdb.Recordset("cod_cnv")
data_sql.Recordset("nom_cnv") = data_mdb.Recordset("nom_cnv")
data_sql.Recordset("tipocta") = data_mdb.Recordset("tipocta")
data_sql.Recordset("cliente") = data_mdb.Recordset("cliente")
data_sql.Recordset("nombre") = data_mdb.Recordset("apellidos")
data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
data_sql.Recordset("tipodoc") = data_mdb.Recordset("tipodoc")
data_sql.Recordset("documento") = data_mdb.Recordset("documento")
data_sql.Recordset("importe") = data_mdb.Recordset("importe")
data_sql.Recordset("moneda") = data_mdb.Recordset("moneda")
data_sql.Recordset("origen") = data_mdb.Recordset("origen")
data_sql.Recordset("mes") = data_mdb.Recordset("mes")
data_sql.Recordset("ano") = data_mdb.Recordset("ano")
data_sql.Recordset("nro_cobr") = data_mdb.Recordset("nro_cobr")
data_sql.Recordset("nom_cobr") = data_mdb.Recordset("nom_cobr")
data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
data_sql.Recordset("estado_cta") = 1
data_sql.Recordset("tiquet") = data_mdb.Recordset("tiquet")
data_sql.Recordset("deudas") = data_mdb.Recordset("deudas")
data_sql.Recordset("total") = data_mdb.Recordset("total")
data_sql.Recordset("servi") = data_mdb.Recordset("servi")
data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
data_sql.Recordset("iva") = data_mdb.Recordset("iva")
data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
data_sql.Recordset("promo") = data_mdb.Recordset("promo")
data_sql.Recordset("descimp") = data_mdb.Recordset("descimp")
data_sql.Recordset("descpor") = data_mdb.Recordset("descpor")
data_sql.Recordset.Update


MsgBox "Terminado"

End Sub

Private Sub Command10_Click()
Dim Xelnroasis, Xelnroadar As String
Dim Xelnroadarn As Long
'Xelnroasis = InputBox("Ingrese nro actual")
'Xelnroadar = InputBox("Ingrese nuevo nro")
Xelnroadar = ""
Xelnroasis = ""
Xelnroasis = InputBox("Ingrese CODIGO:")

Xelnroadar = InputBox("Ingrese NUEVO CODIGO:")
'81152
data_mdb.DatabaseName = App.Path & "\sapp.mdb"
''data_mdb.RecordSource = "Select * from infor_sol where cl_nom_sup ='" & "VICTORIAM" & "'"
data_mdb.RecordSource = "Select * from env_soc where cl_codigo =" & Val(Xelnroasis)
'data_mdb.RecordSource = "Select * from movil where nromov =" & 11
'data_mdb.RecordSource = "Select * from infor_sol where estado =" & Xelnroasis
'data_mdb.RecordSource = "Select * from infor_sol where cl_nrovend =" & 80596 & " and cl_val3 =" & 0

'data_mdb.RecordSource = "Select * from env_soc where cl_codigo =" & Xelnroasis
'ULTIMO DADO:6798906

data_mdb.Refresh
'data_mdb.Recordset.MoveLast
'data_mdb.Recordset.MoveFirst
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Xelnroadarn = Val(Xelnroadar)
'   Do While Not data_mdb.Recordset.EOF
'      data_mdb.Recordset.Delete
'   data_mdb.Recordset.MoveFirst
'   MsgBox "Total:" & data_mdb.Recordset.RecordCount
'   Do While Not data_mdb.Recordset.EOF
'      If data_mdb.Recordset("cl_codigo") = 74901053 Then
'      Else
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cl_codigo") = Xelnroadarn
'      data_mdb.Recordset("chofer") = "ASESOR GESTION CAP.HUMANO"
'      data_mdb.Recordset("medico") = "VICTORIAM"
         data_mdb.Recordset.Update
'         Xelnroadarn = Xelnroadarn + 1
'      End If
'      data_mdb.Recordset.MoveNext
'   data_mdb.Recordset("cl_num") = Xelnroadar
'   data_mdb.Recordset("cl_fultmov") = Null
'   data_mdb.Recordset.Update
'    data_mdb.Recordset.Delete
'      data_mdb.Recordset.MoveNext
'   Loop
'    data_mdb.Recordset.Delete
'   data_mdb.Recordset.Delete
Else
    MsgBox "Sin registros"

End If
MsgBox "Terminado"


End Sub

Private Sub Command11_Click()
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, xtot, Xlacedu As Long
Dim Xcedtex, Xtottex, Xcodced As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo As Long
Dim Xelnrof As Long
Dim Xrruu As Double
Dim Buscarut, XcedGraba As String

Xn1 = 2
Xn2 = 9
Xn3 = 8
Xn4 = 7
Xn5 = 6
Xn6 = 3
Xn7 = 4
Xcodced = ""
XcedGraba = ""
Form1.MousePointer = 11
Xpond = 10

data_mdb.DatabaseName = App.Path & "\informes.mdb"
data_mdb.RecordSource = "infcli"
data_mdb.Refresh
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_mdb.Recordset.Delete
      data_mdb.Recordset.MoveNext
   Loop
End If
Data1.Connect = "odbc;dsn=sappnew;"

data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.RecordSource = "select * from cabezal_hc"
data_sql.Refresh
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      If Len(Trim(data_sql.Recordset("cb_ced"))) = 7 Or Len(Trim(data_sql.Recordset("cb_ced"))) = 8 Then
         Xcedtex = Trim(data_sql.Recordset("cb_ced"))
      'Xcedtex = "3717847"
         Xlargo = Len(Xcedtex)
         If Xlargo = 7 Then
            Xcodced = Mid(Trim(Xcedtex), 7, 1)
            Xcedtex = Mid(Xcedtex, 1, 6)
            XcedGraba = Xcedtex
            Xcedtex = "0" & Trim(Xcedtex)
         Else
            Xcodced = Mid(Trim(Xcedtex), 8, 1)
            Xcedtex = Mid(Xcedtex, 1, 7)
            XcedGraba = Xcedtex
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
         xtot = Xced1 + Xced2 + Xced3 + Xced4 + Xced5 + Xced6 + Xced7
         If Len(Trim(Str(xtot))) = 1 Then
            Xtottex = "0000" & Trim(Str(xtot))
         End If
         If Len(Trim(Str(xtot))) = 2 Then
            Xtottex = "000" & Trim(Str(xtot))
         End If
         If Len(Trim(Str(xtot))) = 3 Then
            Xtottex = "00" & Trim(Str(xtot))
         End If
         If Len(Trim(Str(xtot))) = 4 Then
            Xtottex = "0" & Trim(Str(xtot))
         End If
         xtot = Val(Mid(Xtottex, 5, 1))
         If xtot <> 0 Then
            xtot = Xpond - xtot
         Else
            xtot = 0
         End If
         If Val(Xcodced) = Val(xtot) Then
         Else
            
            XcedGraba = XcedGraba & Trim(Str(xtot))
            data_mdb.Recordset.AddNew
            data_mdb.Recordset("cl_codigo") = data_sql.Recordset("cb_mat")
            data_mdb.Recordset("cl_apellid") = data_sql.Recordset("cb_nom1")
            data_mdb.Recordset("cl_codconv") = data_sql.Recordset("cb_codconv")
            data_mdb.Recordset("cl_cedula") = Val(XcedGraba)
            data_mdb.Recordset("cl_dpto") = data_sql.Recordset("cb_ced")
            
            data_mdb.Recordset.Update
            
            data_sql.Recordset.Edit
            data_sql.Recordset("cb_ced") = XcedGraba
            data_sql.Recordset.Update
            Data1.RecordSource = "select * from cabezal_hcdig where mat =" & data_sql.Recordset("cb_mat")
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.MoveFirst
               Do While Not Data1.Recordset.EOF
                  If Trim(Data1.Recordset("cedtext")) <> Trim(XcedGraba) Then
                     Data1.Recordset.Edit
                     Data1.Recordset("cedtext") = Trim(XcedGraba)
                     Data1.Recordset("codced") = Val(xtot)
                     Data1.Recordset.Update
                  End If
                  Data1.Recordset.MoveNext
               Loop
            End If
         End If
      End If
      data_sql.Recordset.MoveNext
   Loop
End If

Form1.MousePointer = 0
MsgBox "Terminado"

      

End Sub

Private Sub Command12_Click()
Dim Xdesdefec, Xhastafec As String
'Xdesdefec = InputBox("Ingrese desde Fecha:", "Desde")
'Xhastafec = InputBox("Ingrese hasta Fecha:", "Hasta")
'data_mdb.DatabaseName = App.Path & "\sapp.mdb"
'data_mdb.RecordSource = "llamado"
'data_mdb.Refresh
Xdesdefec = InputBox("Ingrese nro:")

data_sql.DatabaseName = App.Path & "\sapp.mdb"
'data_sql.RecordSource = "Select * from llamado where fecha =#" & Format("06/06/2015", "yyyy/mm/dd") & "#"
data_sql.RecordSource = "Select * from llamado where nrolla =" & Val(Xdesdefec)
data_sql.Refresh
Form1.MousePointer = 11
'10674087

If data_sql.Recordset.RecordCount > 0 Then
'   data_sql.Recordset.MoveFirst
'   Do While Not data_sql.Recordset.EOF
      data_sql.Recordset.Edit
'      data_sql.Recordset("pend") = 2
'      data_sql.Recordset("codmot") = "V"
      data_sql.Recordset("totend") = Null
'      data_sql.Recordset("movilpas") = 2015
'      data_sql.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
'      data_sql.Recordset("horpas") = Format(Time, "HH:mm")
'      data_sql.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
'      data_sql.Recordset("horsali") = Format(Time, "HH:mm")
'      data_sql.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
'      data_sql.Recordset("hor_llega") = Format(Time, "HH:mm")
'      data_sql.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
'      data_sql.Recordset("hor_rea") = Format(Time, "HH:mm")
'      data_sql.Recordset("diag") = "CMT DERIVADO A BASE"
'      data_sql.Recordset("colormot") = "V"
                   'data_lla.Recordset("codmed") = txt_codmed.Text
                   'data_lla.Recordset("nommed") = dbcbomed.Text
      data_sql.Recordset.Update

'      If IsNull(data_sql.Recordset("fecpas")) = False Then
'         data_sql.Recordset("fecpas") = Format("08/09/2015", "dd/mm/yyyy")
'      End If
'      If IsNull(data_sql.Recordset("fecsali")) = False Then
'         data_sql.Recordset("fecsali") = Format("08/09/2015", "dd/mm/yyyy")
'      End If
'      If IsNull(data_sql.Recordset("fec_rea")) = False Then
'         data_sql.Recordset("fec_rea") = Format("08/09/2015", "dd/mm/yyyy")
'      End If
'      If IsNull(data_sql.Recordset("fec_llega")) = False Then
'         data_sql.Recordset("fec_llega") = Format("08/09/2015", "dd/mm/yyyy")
'      End If
'      data_sql.Recordset("nrolla") = 99888877
'      data_sql.Recordset("nro") = 99888877
'      data_sql.Recordset.Update
'      data_sql.Recordset.Delete
'      data_sql.Recordset.MoveNext
'   Loop
End If
'data_sql.Refresh

'data_sql.RecordSource = "Select * from resplla where fecha =#" & Format("06/06/2015", "yyyy/mm/dd") & "#"
'data_sql.RecordSource = "Select * from llamado where nrolla in (10674138, 10674139, 10674141)"

'data_sql.Refresh
'Form1.MousePointer = 11
'10674087

'If data_sql.Recordset.RecordCount > 0 Then
'   data_sql.Recordset.MoveFirst
'   Do While Not data_sql.Recordset.EOF
'      data_sql.Recordset.Edit
'      data_sql.Recordset("nrolla") = 99888877
'      data_sql.Recordset("nro") = 99888877
'      data_sql.Recordset.Update
'      data_sql.Recordset.Delete
'      data_sql.Recordset.MoveNext
'   Loop
'End If
'data_sql.Refresh

Form1.MousePointer = 0
MsgBox "Terminado"

End Sub

Private Sub Command13_Click()
Dim Xlineat As String
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook
Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir3 As New Excel.Application
Dim Xmat As Long
Xmat = 0
Form1.MousePointer = 11
Xlineat = ""
Xtotreg = 0
data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.RecordSource = "select deudas.nro_cobr,deudas.nombre,deudas.fecha_pago,deudas.ano,deudas.mes,deudas.total,deudas.fecha,deudas.documento,deudas.servi,deudas.cliente," & _
"deudas.fecha_pago,clientes.cl_codigo,clientes.cl_cedula,clientes.cl_codced,clientes.estado from deudas inner join clientes on deudas.cliente=clientes.cl_codigo" & _
" where clientes.estado in (1) and deudas.fecha_pago is null and deudas.nro_cobr in (221) order by deudas.cliente,deudas.fecha"
data_sql.Refresh
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Xlin = 1
   XCol = 1
   Set Xobjexel22 = New Excel.Application
   Set Xlibexel22 = Xobjexel22.Workbooks.Add
   Set Xarchexel22 = Xlibexel22.Worksheets.Add
   Xarchexel22.Name = Trim("Sapp")
   Xlibexel22.SaveAs ("C:\planillas\RedPagos-20210504.csv")
   Xarchtex = "C:\planillas\RedPagos-20210504.csv"
   Xmat = data_sql.Recordset("cliente")
   Do While Not data_sql.Recordset.EOF
      Xlineat = Trim(Str(data_sql.Recordset("ano"))) & "," & Trim(Str(data_sql.Recordset("mes"))) & "," & Trim(Str(Xtotreg)) & "," & Trim(Str(data_sql.Recordset("cl_cedula"))) & Trim(Str(data_sql.Recordset("cl_codced"))) & "," & _
      data_sql.Recordset("nombre") & ",0," & Trim(Str(Val(data_sql.Recordset("total")))) & "," & Format(data_sql.Recordset("fecha"), "dd/mm/yyyy") & "," & _
      Format(data_sql.Recordset("fecha"), "dd/mm/yyyy") & "," & Trim(Str(data_sql.Recordset("documento"))) & "," & Trim(Str(1)) & "," & Format(data_sql.Recordset("servi"), "###0.00")
      Xmat = data_sql.Recordset("cliente")
      data_sql.Recordset.MoveNext
      If data_sql.Recordset.EOF = True Then
      Else
         If Xmat = data_sql.Recordset("cliente") Then
            Xtotreg = Xtotreg + 1
         Else
            Xtotreg = 0
         End If
      End If
      Xarchexel22.Cells(Xlin, XCol) = Xlineat
      Xlin = Xlin + 1
      Xlineat = ""
   Loop
   Xlibexel22.Save
   Xlibexel22.Close
   Xobjexel22.Quit
   Xlabrir3.Workbooks.Open Xarchtex, , False
   Xlabrir3.Visible = True
   Xlabrir3.WindowState = xlMaximized
End If

Form1.MousePointer = 0
MsgBox "Proceso terminado"

End Sub

Private Sub Command14_Click()
''''frm_consrep.Show vbModal
data_sql.Connect = "ODBC;DSN=sappnew;"
Data1.Connect = "ODBC;DSN=sappnew;"

'data_sql.DatabaseName = App.Path & "\sapp.mdb"
'data_sql.RecordSource = "Select * from deudas where cliente =" & 94253 & " and documento =" & 3270752
'data_sql.Refresh
'Dim Xxmat, Xxmes, Xxano, Xcontalas As Long
'Xxmat = 0
'Xxmes = 0
'Xxano = 0
'Xcontalas = 0
'data_sql.Recordset.MoveFirst

'data_mdb.DatabaseName = App.Path & "\sociosp.mdb"
'data_mdb.RecordSource = "sociosp"
'data_mdb.Refresh

Form1.MousePointer = 11
'data_sql.RecordSource = "SELECT * FROM t_fechas where especial in ('PEDIATRIA') and fecha_cons >=#" & Format("01/04/2022", "yyyy/mm/dd") & "# and nom_pac is null and cancela in ('NO') and sepuede is null order by cod_med,base,fecha_cons,nro"
'data_sql.Refresh
'2 solo No presencial ???
'1 solo presencial ???
'0 no habilitado
'3 habilita todo
Dim BanderaPed As Integer
BanderaPed = 0
'If data_sql.Recordset.RecordCount > 0 Then
'   data_sql.Recordset.MoveFirst
'   Do While Not data_sql.Recordset.EOF
'      If data_sql.Recordset("nro") >= 5 Then
'         If Trim(Mid(data_sql.Recordset("hora"), 4, 2)) = "30" Then
'            If BanderaPed = 0 Then
'               data_sql.Recordset.Edit
'               data_sql.Recordset("sepuede") = 1
'               data_sql.Recordset.Update
'               BanderaPed = 1
'            Else
'               data_sql.Recordset.Edit
'               data_sql.Recordset("sepuede") = 2
'               data_sql.Recordset.Update
'               BanderaPed = 0
'            End If
'         Else
'            data_sql.Recordset.Edit
'            data_sql.Recordset("sepuede") = 0
'            data_sql.Recordset.Update
'         End If
'      End If
'      data_sql.Recordset.MoveNext
'   Loop
'Else
'   MsgBox "No hay datos Ped."
'End If

'MED:GRAL
'modificar Temesio el 04/04 ??
'data_sql.RecordSource = "SELECT * FROM t_fechas where especial in ('MED.GRAL.') and base not in (98,99) and fecha_cons >=#" & Format("01/04/2022", "yyyy/mm/dd") & "# and nom_pac is null and cancela in ('NO') and sepuede is null order by base,fecha_cons,nro"
'data_sql.Refresh

'ESPECIALISTAS
'sacar las otras especialidades
Form1.MousePointer = 11
data_sql.RecordSource = "SELECT * FROM t_fechas where especial not in ('PEDIATRIA','MED.GRAL.','HNF','LABORATORIO','NUTRICIONISTA','ODONTOLOGIA','ECOGRAFIAS','SICOLOGIA','RADIOLOGIA','VACUNACION') and fecha_cons >=#" & Format("01/04/2022", "yyyy/mm/dd") & "# and nom_pac is null and cancela in ('NO') and sepuede in (3) order by cod_med,base,fecha_cons,nro"
data_sql.Refresh
'2 solo No presencial ???
'1 solo presencial
'0 no habilitado
'3 habilita todo
BanderaPed = 0
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      If data_sql.Recordset("nro") = 5 Then
         data_sql.Recordset.Edit
         data_sql.Recordset("sepuede") = 2
         data_sql.Recordset.Update
      Else
         data_sql.Recordset.Edit
         data_sql.Recordset("sepuede") = 1
         data_sql.Recordset.Update
      End If
      data_sql.Recordset.MoveNext
   Loop
Else
   MsgBox "No hay datos Espec."
End If

Form1.MousePointer = 0

MsgBox "Proceso terminado "

End Sub

Private Sub Command15_Click()

data_sql.Connect = "ODBC;DSN=sappnew;"
Data1.Connect = "odbc;dsn=sappnew;"
Form1.MousePointer = 11
'data_mdb.DatabaseName = App.Path & "\cgalmod.mdb"
'data_mdb.RecordSource = "cgalicia"
'data_mdb.Refresh

'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      data_sql.RecordSource = "select * from clientes where cl_codigo =" & data_mdb.Recordset("matricula")
'      data_sql.Refresh
'      If data_sql.Recordset.RecordCount > 0 Then
'         If data_sql.Recordset("cl_codconv") <> data_mdb.Recordset("convenio") Then
'            Data1.RecordSource = "select * from convenio where cnv_codigo ='" & data_mdb.Recordset("convenio") & "'"
''            Data1.Refresh
'            If Data1.Recordset.RecordCount > 0 Then
'               data_sql.Recordset.Edit
'               data_sql.Recordset("cl_codconv") = data_mdb.Recordset("convenio")
'               data_sql.Recordset("cl_nomconv") = Mid(Data1.Recordset("cnv_desc"), 1, 30)
'               data_sql.Recordset.Update
'               Data1.RecordSource = "select * from abmsocio where cl_codigo =" & data_mdb.Recordset("matricula")
'               Data1.Refresh
'               Data1.Recordset.AddNew
'               Data1.Recordset("usuario") = "JFERNAN"
'               Data1.Recordset("fecha") = Date
''               Data1.Recordset("hora") = Format(Time, "HH:mm")
'               Data1.Recordset("cl_codigo") = data_mdb.Recordset("matricula")
'               Data1.Recordset("desc") = "MODIF"
'               Data1.Recordset("cl_motivo") = "MODIFICACION ESPECIAL"
'               Data1.Recordset("convenio") = data_mdb.Recordset("convenio")
'               Data1.Recordset("base") = 18
'               Data1.Recordset.Update
            
'            Else
'               data_mdb.Recordset.Edit
'               data_mdb.Recordset("mod") = "NO"
'               data_mdb.Recordset.Update
'            End If
'         End If
'      End If
'      data_mdb.Recordset.MoveNext
'   Loop
'End If

data_mdb.DatabaseName = App.Path & "\cgalbajas.mdb"
data_mdb.RecordSource = "clicgal"
data_mdb.Refresh

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_sql.RecordSource = "select * from clientes where cl_codigo =" & data_mdb.Recordset("cl_codigo") & " and fecha_baja is null"
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.Edit
         data_sql.Recordset("estado") = 2
         data_sql.Recordset("fecha_baja") = Date
         data_sql.Recordset.Update
         Data1.RecordSource = "select * from abmsocio where cl_codigo =" & data_mdb.Recordset("cl_codigo")
         Data1.Refresh
         Data1.Recordset.AddNew
         Data1.Recordset("usuario") = "JFERNAN"
         Data1.Recordset("fecha") = Date
         Data1.Recordset("hora") = Format(Time, "HH:mm")
         Data1.Recordset("cl_codigo") = data_mdb.Recordset("cl_codigo")
         Data1.Recordset("desc") = "BAJA"
         Data1.Recordset("cl_motivo") = "SIN DATOS"
         Data1.Recordset("convenio") = data_mdb.Recordset("cl_codconv")
         Data1.Recordset("base") = 18
         Data1.Recordset.Update
         data_mdb.Recordset.Edit
         data_mdb.Recordset("obs") = "SI"
         data_mdb.Recordset.Update
      End If
      data_mdb.Recordset.MoveNext
   Loop
End If

Form1.MousePointer = 0
MsgBox "Proceso terminado"


End Sub

Private Sub Command16_Click()
Dim Xmat As Long
Dim Xcedtex, Xtottex, Xtelefon As String
Dim Xlargo As Integer
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7 As Integer
Dim xtot As Long
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xlacedu As Long
Xtelefon = ""
Xn1 = 2
Xn2 = 9
Xn3 = 8
Xn4 = 7
Xn5 = 6
Xn6 = 3
Xn7 = 4
Xpond = 10

data_mdb.DatabaseName = App.Path & "\ccouatl.mdb"
data_mdb.RecordSource = "ccouatl"
data_mdb.Refresh
Xmat = 5059501

Data1.Connect = "odbc;dsn=sappnew;"
Data1.RecordSource = "select * from abmsocio"
Data1.Refresh

data_sql.Connect = "odbc;dsn=sappnew;"
'data_sql.RecordSource = "emi1220"
'data_sql.Refresh

'Data1.Connect = "odbc;dsn=sappnew;"
'CIRCULO CATOLICO FONASA SAN JACINTO
data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   data_sql.RecordSource = "select * from clientes where cl_cedula =" & data_mdb.Recordset("ced")
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
      If data_mdb.Recordset("accion") = "SI" Then
         data_sql.Recordset.Edit
         data_sql.Recordset("cl_cedula") = data_mdb.Recordset("ced")
         data_sql.Recordset("cl_codced") = data_mdb.Recordset("dv")
         data_sql.Recordset("estado") = 1
         data_sql.Recordset("fecha_baja") = Null
         data_sql.Recordset("cl_sexo") = 2
         data_sql.Recordset("cl_codconv") = "CCFAT"
         data_sql.Recordset("cl_nomconv") = Mid("CIRCULO CATOLICO FONASA ATLANTIDA", 1, 30)
         data_sql.Recordset("cl_socmnom") = "CIRCULO CATOLICO"
         If IsNull(data_mdb.Recordset("fnac")) = False Then
            data_sql.Recordset("cl_fnac") = data_mdb.Recordset("fnac")
         End If
         If IsNull(data_mdb.Recordset("direc")) = False Then
            data_sql.Recordset("cl_direcci") = data_mdb.Recordset("direc")
         End If
         If IsNull(data_mdb.Recordset("telef1")) = False Then
            If IsNull(data_mdb.Recordset("telef2")) = False Then
               Xtelefon = Trim(data_mdb.Recordset("telef1")) & "//" & Trim(Str(data_mdb.Recordset("telef2")))
               data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
            Else
               Xtelefon = Trim(data_mdb.Recordset("telef1"))
               data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
            End If
         Else
            Xtelefon = ""
         End If
         data_sql.Recordset("cl_apellid") = Trim(data_mdb.Recordset("ape")) & " " & Trim(data_mdb.Recordset("nom"))
         data_sql.Recordset("cl_fecing") = Date
         data_sql.Recordset("cl_nrovend") = 738
         data_sql.Recordset("cl_nomvend") = "MUTUALISTA"
         data_sql.Recordset("fecha_sys") = Date
         data_sql.Recordset("cl_referen") = "NO APLICA"
         data_sql.Recordset.Update
         Data1.Recordset.AddNew
         Data1.Recordset("usuario") = "COMPUTOS"
         Data1.Recordset("fecha") = Date
         Data1.Recordset("hora") = Format(Time, "HH:mm")
         Data1.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
         Data1.Recordset("desc") = "MODIF"
         Data1.Recordset("cl_motivo") = "MODIFICACION ESPECIAL"
         Data1.Recordset("convenio") = "CCFAT"
         Data1.Recordset("base") = 18
         Data1.Recordset.Update
      Else
        If data_mdb.Recordset("accion") = "DI" Or data_mdb.Recordset("accion") = "LL" Then
           data_sql.Recordset.Edit
           data_sql.Recordset("cl_cedula") = data_mdb.Recordset("ced")
           data_sql.Recordset("cl_codced") = data_mdb.Recordset("dv")
           data_sql.Recordset("estado") = 1
           data_sql.Recordset("fecha_baja") = Null
           data_sql.Recordset("cl_sexo") = 2
           data_sql.Recordset("cl_codconv") = "CCFAT"
           data_sql.Recordset("cl_nomconv") = Mid("CIRCULO CATOLICO FONASA ATLANTIDA", 1, 30)
           data_sql.Recordset("cl_socmnom") = "CIRCULO CATOLICO"
           If IsNull(data_mdb.Recordset("fnac")) = False Then
              data_sql.Recordset("cl_fnac") = data_mdb.Recordset("fnac")
           End If
           If IsNull(data_mdb.Recordset("telef1")) = False Then
              If IsNull(data_mdb.Recordset("telef2")) = False Then
                 Xtelefon = Trim(data_mdb.Recordset("telef1")) & "//" & Trim(Str(data_mdb.Recordset("telef2")))
                 data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
              Else
                 Xtelefon = Trim(data_mdb.Recordset("telef1"))
                 data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
              End If
           Else
              Xtelefon = ""
           End If
           data_sql.Recordset("cl_apellid") = Trim(data_mdb.Recordset("ape")) & " " & Trim(data_mdb.Recordset("nom"))
           data_sql.Recordset("cl_fecing") = Date
           data_sql.Recordset("cl_nrovend") = 738
           data_sql.Recordset("cl_nomvend") = "MUTUALISTA"
           data_sql.Recordset("fecha_sys") = Date
           data_sql.Recordset("cl_referen") = "NO APLICA"
           data_sql.Recordset.Update
           Data1.Recordset.AddNew
           Data1.Recordset("usuario") = "COMPUTOS"
           Data1.Recordset("fecha") = Date
           Data1.Recordset("hora") = Format(Time, "HH:mm")
           Data1.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
           Data1.Recordset("desc") = "MODIF"
           Data1.Recordset("cl_motivo") = "MODIFICACION ESPECIAL"
           Data1.Recordset("convenio") = "CCFAT"
           Data1.Recordset("base") = 18
           Data1.Recordset.Update
        Else
            If data_mdb.Recordset("accion") = "CO" Then
               data_sql.Recordset.Edit
               data_sql.Recordset("cl_cedula") = data_mdb.Recordset("ced")
               data_sql.Recordset("cl_codced") = data_mdb.Recordset("dv")
               data_sql.Recordset("estado") = 1
               data_sql.Recordset("fecha_baja") = Null
               data_sql.Recordset("cl_sexo") = 2
               data_sql.Recordset("cl_codconv") = "CCFATA"
               data_sql.Recordset("cl_nomconv") = Mid("CIRCULO CATOLICO FONASA ATLANTIDA AMBULATORIO", 1, 30)
               data_sql.Recordset("cl_socmnom") = "CIRCULO CATOLICO"
               If IsNull(data_mdb.Recordset("fnac")) = False Then
                  data_sql.Recordset("cl_fnac") = data_mdb.Recordset("fnac")
               End If
               If IsNull(data_mdb.Recordset("telef1")) = False Then
                  If IsNull(data_mdb.Recordset("telef2")) = False Then
                     Xtelefon = Trim(data_mdb.Recordset("telef1")) & "//" & Trim(Str(data_mdb.Recordset("telef2")))
                     data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
                  Else
                     Xtelefon = Trim(data_mdb.Recordset("telef1"))
                     data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
                  End If
               Else
                  Xtelefon = ""
               End If
               data_sql.Recordset("cl_apellid") = Trim(data_mdb.Recordset("ape")) & " " & Trim(data_mdb.Recordset("nom"))
               data_sql.Recordset("cl_fecing") = Date
               data_sql.Recordset("cl_nrovend") = 738
               data_sql.Recordset("cl_nomvend") = "MUTUALISTA"
               data_sql.Recordset("fecha_sys") = Date
               data_sql.Recordset("cl_referen") = "NO APLICA"
               data_sql.Recordset.Update
               Data1.Recordset.AddNew
               Data1.Recordset("usuario") = "COMPUTOS"
               Data1.Recordset("fecha") = Date
               Data1.Recordset("hora") = Format(Time, "HH:mm")
               Data1.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
               Data1.Recordset("desc") = "MODIF"
               Data1.Recordset("cl_motivo") = "MODIFICACION ESPECIAL"
               Data1.Recordset("convenio") = "CCFATA"
               Data1.Recordset("base") = 18
               Data1.Recordset.Update
            Else
               data_sql.Recordset.Edit
               data_sql.Recordset("cl_cedula") = data_mdb.Recordset("ced")
               data_sql.Recordset("cl_codced") = data_mdb.Recordset("dv")
               data_sql.Recordset("estado") = 1
               data_sql.Recordset("fecha_baja") = Null
               data_sql.Recordset("cl_sexo") = 2
               data_sql.Recordset("cl_socmnom") = "CIRCULO CATOLICO"
               If IsNull(data_mdb.Recordset("fnac")) = False Then
                  data_sql.Recordset("cl_fnac") = data_mdb.Recordset("fnac")
               End If
               If IsNull(data_mdb.Recordset("telef1")) = False Then
                  If IsNull(data_mdb.Recordset("telef2")) = False Then
                     Xtelefon = Trim(data_mdb.Recordset("telef1")) & "//" & Trim(Str(data_mdb.Recordset("telef2")))
                     data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
                  Else
                     Xtelefon = Trim(data_mdb.Recordset("telef1"))
                     data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
                  End If
               Else
                  Xtelefon = ""
               End If
               data_sql.Recordset("cl_apellid") = Trim(data_mdb.Recordset("ape")) & " " & Trim(data_mdb.Recordset("nom"))
               data_sql.Recordset("fecha_sys") = Date
               data_sql.Recordset("cl_referen") = "NO APLICA"
               data_sql.Recordset.Update
               Data1.Recordset.AddNew
               Data1.Recordset("usuario") = "COMPUTOS"
               Data1.Recordset("fecha") = Date
               Data1.Recordset("hora") = Format(Time, "HH:mm")
               Data1.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
               Data1.Recordset("desc") = "MODIF"
               Data1.Recordset("cl_motivo") = "MODIFICACION ESPECIAL"
               Data1.Recordset("convenio") = "CCFATA"
               Data1.Recordset("base") = 18
               Data1.Recordset.Update
            End If
        End If
      End If
   Else 'HAGO ALTA
      Xmat = Xmat + 1
      data_sql.Recordset.AddNew
      data_sql.Recordset("cl_codigo") = Xmat
      data_sql.Recordset("cl_cedula") = data_mdb.Recordset("ced")
      data_sql.Recordset("cl_codced") = data_mdb.Recordset("dv")
      data_sql.Recordset("estado") = 1
      data_sql.Recordset("fecha_baja") = Null
      data_sql.Recordset("cl_apellid") = Trim(data_mdb.Recordset("ape")) & " " & Trim(data_mdb.Recordset("nom"))
      data_sql.Recordset("cl_sexo") = 2
      data_sql.Recordset("cl_codconv") = "CCFAT"
      data_sql.Recordset("cl_nomconv") = Mid("CIRCULO CATOLICO FONASA ATLANTIDA", 1, 30)
      data_sql.Recordset("cl_socmnom") = "CIRCULO CATOLICO"
      If IsNull(data_mdb.Recordset("fnac")) = False Then
         data_sql.Recordset("cl_fnac") = data_mdb.Recordset("fnac")
      End If
      If IsNull(data_mdb.Recordset("direc")) = False Then
         data_sql.Recordset("cl_direcci") = data_mdb.Recordset("direc")
      End If
      If IsNull(data_mdb.Recordset("telef1")) = False Then
         If IsNull(data_mdb.Recordset("telef2")) = False Then
            Xtelefon = data_mdb.Recordset("telef1") & "//" & data_mdb.Recordset("telef2")
         Else
            Xtelefon = data_mdb.Recordset("telef1")
         End If
      Else
         Xtelefon = ""
      End If
      If Trim(Xtelefon) <> "" Then
         data_sql.Recordset("cl_telefon") = Mid(Xtelefon, 1, 20)
      End If
      data_sql.Recordset("cl_fecing") = Date
      data_sql.Recordset("cl_nrovend") = 738
      data_sql.Recordset("cl_nomvend") = "MUTUALISTA"
      data_sql.Recordset("fecha_sys") = Date
      data_sql.Recordset("cl_referen") = "NO APLICA"
      data_sql.Recordset.Update
      
      Data1.Recordset.AddNew
      Data1.Recordset("usuario") = "COMPUTOS"
      Data1.Recordset("fecha") = Date
      Data1.Recordset("hora") = Format(Time, "HH:mm")
      Data1.Recordset("cl_codigo") = Xmat
      Data1.Recordset("desc") = "ALTA"
      Data1.Recordset("cl_motivo") = "ALTA DE FICHA"
      Data1.Recordset("convenio") = "CCFAT"
      Data1.Recordset("base") = 18
      Data1.Recordset.Update
      
   End If
   data_mdb.Recordset.Edit
   data_mdb.Recordset("estado") = "SI"
   data_mdb.Recordset.Update
   data_mdb.Recordset.MoveNext
Loop
'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      data_mdb.Recordset.Delete
'      data_mdb.Recordset.MoveNext
'
'   Loop
'End If

'data_sql.Recordset.MoveFirst
'Do While Not data_sql.Recordset.EOF
'   Data1.RecordSource = "select * from clientes where cl_codigo =" & data_sql.Recordset("cliente")
'   Data1.Refresh
'   If Data1.Recordset.RecordCount > 0 Then
'      data_mdb.Recordset.AddNew
'      data_mdb.Recordset("cliente") = data_sql.Recordset("cliente")
'      data_mdb.Recordset("apellidos") = data_sql.Recordset("apellidos")
'      data_mdb.Recordset("cod_cnv") = data_sql.Recordset("cod_cnv")
'      data_mdb.Recordset("nom_cnv") = data_sql.Recordset("nom_cnv")
'      data_mdb.Recordset("grupo") = data_sql.Recordset("grupo")
'      data_mdb.Recordset("zona") = data_sql.Recordset("zona")
'      data_mdb.Recordset("origen") = Data1.Recordset("cl_socmnom")
'      data_mdb.Recordset.Update
'   End If
'   data_sql.Recordset.MoveNext
'Loop
'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      data_sql.RecordSource = "select * from clientes where cl_codigo =" & data_mdb.Recordset("cl_codigo")
'      data_sql.Refresh
'      If data_sql.Recordset.RecordCount > 0 Then
'         If data_sql.Recordset("cl_codconv") = "CASH" Then
'            data_sql.Recordset.Edit
'            data_sql.Recordset("fecha_baja") = Date
'            data_sql.Recordset("estado") = 2
'            data_sql.Recordset.Update
'            Data1.Recordset.AddNew
'            Data1.Recordset("usuario") = "AUTOMAT"
'            Data1.Recordset("fecha") = Date
'            Data1.Recordset("hora") = Format(Time, "HH:mm")
'            Data1.Recordset("cl_codigo") = data_mdb.Recordset("cl_codigo")
'            Data1.Recordset("desc") = "BAJA"
'            Data1.Recordset("cl_motivo") = "SIN DATOS"
'            Data1.Recordset("convenio") = data_sql.Recordset("cl_codconv")
'            Data1.Recordset("base") = 18
'            Data1.Recordset.Update
'         End If
'      End If
'      data_mdb.Recordset.MoveNext
'   Loop
'End If
MsgBox "Proceso terminado"


End Sub

Private Sub Command17_Click()
Dim Xmat As Double

data_mdb.DatabaseName = App.Path & "\asistot.mdb"
data_mdb.RecordSource = "Select * from asistot order by cod_cli"
data_mdb.Refresh
data_sql.DatabaseName = App.Path & "\asistot.mdb"
data_mdb.Recordset.MoveFirst
Xmat = data_mdb.Recordset("cod_cli")

Do While Not data_mdb.Recordset.EOF
   If Xmat = data_mdb.Recordset("cod_cli") Then
      data_mdb.Recordset.Edit
      data_mdb.Recordset("total") = 999888
      data_mdb.Recordset.Update
   Else
      data_mdb.Recordset.MovePrevious
      data_mdb.Recordset.Edit
      data_mdb.Recordset("total") = 0
      data_mdb.Recordset.Update
      data_sql.RecordSource = "Select * from asistot where cod_cli =" & data_mdb.Recordset("cod_cli") & " and nro_flia =" & 1
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.MoveLast
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cantmg") = data_sql.Recordset.RecordCount
         data_mdb.Recordset.Update
      End If
      
      data_sql.RecordSource = "Select * from asistot where cod_cli =" & data_mdb.Recordset("cod_cli") & " and nro_flia =" & 2
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.MoveLast
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cantenf") = data_sql.Recordset.RecordCount
         data_mdb.Recordset.Update
      End If
      
      data_sql.RecordSource = "Select * from asistot where cod_cli =" & data_mdb.Recordset("cod_cli") & " and nro_flia =" & 3
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.MoveLast
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cantlab") = data_sql.Recordset.RecordCount
         data_mdb.Recordset.Update
      End If
      
      data_sql.RecordSource = "Select * from asistot where cod_cli =" & data_mdb.Recordset("cod_cli") & " and nro_flia =" & 5
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.MoveLast
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cantecos") = data_sql.Recordset.RecordCount
         data_mdb.Recordset.Update
      End If
      
      data_sql.RecordSource = "Select * from asistot where cod_cli =" & data_mdb.Recordset("cod_cli") & " and nro_flia =" & 7
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.MoveLast
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cantrx") = data_sql.Recordset.RecordCount
         data_mdb.Recordset.Update
      End If
      
      data_sql.RecordSource = "Select * from asistot where cod_cli =" & data_mdb.Recordset("cod_cli") & " and nro_flia =" & 10
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.MoveLast
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cantesp") = data_sql.Recordset.RecordCount
         data_mdb.Recordset.Update
      End If
      
      data_sql.RecordSource = "Select * from asistot where cod_cli =" & data_mdb.Recordset("cod_cli") & " and nro_flia =" & 14
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.MoveLast
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cantped") = data_sql.Recordset.RecordCount
         data_mdb.Recordset.Update
      End If
      
      data_mdb.Recordset.MoveNext
   End If
   Xmat = data_mdb.Recordset("cod_cli")
   data_mdb.Recordset.MoveNext
Loop


MsgBox "Terminado"


End Sub

Private Sub Command18_Click()
'data_sql.Connect = ""
'data_sql.DatabaseName = App.Path & "\comercios.mdb"
'data_sql.RecordSource = "comer"
'data_sql.Refresh
Dim Xeln As Long
Xeln = 375
data_mdb.Connect = "ODBC;DSN=sapp;"
data_mdb.DatabaseName = ""
data_mdb.RecordSource = "Select * from deudas where tipodoc ='" & "CRE" & "' and documento =" & 0
data_mdb.Refresh

data_sql.Connect = "ODBC;DSN=sapp;"

''Data1.DatabaseName = App.Path & "\sapp.mdb"
'Dim Xf As Date

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_sql.RecordSource = "Select * from clientes where cl_codigo =" & data_mdb.Recordset("cliente")
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         If IsNull(data_sql.Recordset("saldo_cc")) = False Then
            data_sql.Recordset.Edit
            data_sql.Recordset("saldo_cc") = data_sql.Recordset("saldo_cc") - data_mdb.Recordset("tot_lin")
            data_sql.Recordset.Update
         End If
      End If
     
'      data_mdb.Recordset.Edit
'      data_mdb.Recordset("dias") = 3
'      data_mdb.Recordset("vto") = Format("28/09/2015", "dd/mm/yyyy")
'      data_mdb.Recordset("nom_medic") = data_mdb.Recordset("nom_medic") & "-"
'      data_mdb.Recordset("base") = 0
'      data_mdb.Recordset("nro") = data_sql.Recordset("cod")
'      data_mdb.Recordset("obsmot") = data_sql.Recordset("rsoc")
'      data_mdb.Recordset("obs") = data_sql.Recordset("direc")
'      data_mdb.Recordset("mat") = data_sql.Recordset("codd")
'      data_mdb.Recordset("usuario") = data_sql.Recordset("local")
'      data_mdb.Recordset("referen") = data_sql.Recordset("telef")
'      data_mdb.Recordset("accion") = data_sql.Recordset("cp")
'      data_mdb.Recordset.Update
'      data_mdb.Recordset.AddNew
'      data_mdb.Recordset("base") = 97
'      data_mdb.Recordset("nro") = Xeln
'      data_mdb.Recordset("obsmot") = data_sql.Recordset("rut")
''      data_mdb.Recordset("obs") = data_sql.Recordset("direc")
'      data_mdb.Recordset("mat") = data_sql.Recordset("cod")
'      data_mdb.Recordset("usuario") = data_sql.Recordset("nomdep")
'      data_mdb.Recordset.Update
'      data_mdb.Recordset.MoveNext
'      Xeln = Xeln + 1
     data_mdb.Recordset.MoveNext
   Loop
End If

'data_mdb.Recordset.MoveFirst
'Do While Not data_mdb.Recordset.EOF
'   data_mdb.Recordset.Delete
'   data_mdb.Recordset.MoveNext
'Loop


MsgBox "Terminado"

End Sub

Private Sub Command19_Click()


'data_mdbn.ConnectionString = "DSN=sappnew"
Dim Xced As Long

data_mdb.DatabaseName = App.Path & "\evangmet.mdb"
data_mdb.RecordSource = "evangmet"
data_mdb.Refresh


data_sql.Connect = "odbc;dsn=sappnew;"

data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   If Len(data_mdb.Recordset("numero")) >= 8 Then
      If Len(data_mdb.Recordset("numero")) = 8 Then
         Xced = Val(Mid(Trim(data_mdb.Recordset("numero")), 1, 6))
      Else
         Xced = Val(Mid(Trim(data_mdb.Recordset("numero")), 1, 7))
      End If
      data_sql.RecordSource = "select * from clientes where cl_cedula =" & Xced
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_mdb.Recordset.Edit
         data_mdb.Recordset("conv_sapp") = data_sql.Recordset("cl_codconv")
         If IsNull(data_sql.Recordset("fecha_baja")) = True Then
            data_mdb.Recordset("estado") = "ACTIVO"
         Else
            data_mdb.Recordset("estado") = "BAJA"
         End If
         data_mdb.Recordset.Update
      Else
         data_mdb.Recordset.Edit
         data_mdb.Recordset("estado") = "NO EXISTE"
         data_mdb.Recordset.Update
      End If
   Else
      data_sql.RecordSource = "select * from clientes where cl_nrosocm ='" & data_mdb.Recordset("numero") & "'"
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_mdb.Recordset.Edit
         data_mdb.Recordset("conv_sapp") = data_sql.Recordset("cl_codconv")
         If IsNull(data_sql.Recordset("fecha_baja")) = True Then
            data_mdb.Recordset("estado") = "ACTIVO"
         Else
            data_mdb.Recordset("estado") = "BAJA"
         End If
         data_mdb.Recordset.Update
      Else
         data_mdb.Recordset.Edit
         data_mdb.Recordset("estado") = "NO EXISTE"
         data_mdb.Recordset.Update
      End If
   
   End If
   data_mdb.Recordset.MoveNext
Loop

Form1.MousePointer = 0
MsgBox "Terminado"

End Sub

Private Sub Command2_Click()
Dim Xpesos As Double
Xpesos = 0

Form1.MousePointer = 11
data_mdb.DatabaseName = App.Path & "\bajas2.mdb"
data_mdb.RecordSource = "bajas2"
data_mdb.Refresh

data_sql.Connect = "odbc;dsn=sappnew;"

data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   Xpesos = 0
   data_sql.RecordSource = "select * from abmsocio where cl_codigo =" & data_mdb.Recordset("cl_codigo") & " and desc ='" & "BAJA" & "' and fecha >=#" & Format(data_mdb.Recordset("fecha_baja"), "yyyy/mm/dd") & "#"
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
      data_mdb.Recordset.Edit
      data_mdb.Recordset("motivo") = data_sql.Recordset("cl_motivo")
      data_mdb.Recordset.Update
   End If
   data_sql.RecordSource = "select * from deudas where fecha_pago is null and cliente =" & data_mdb.Recordset("cl_codigo") & " and mes >" & 0 & " order by ano,mes"
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
      data_sql.Recordset.MoveFirst
      Do While Not data_sql.Recordset.EOF
         Xpesos = Xpesos + data_sql.Recordset("total")
         data_sql.Recordset.MoveNext
      Loop
      data_sql.Recordset.MovePrevious
      data_mdb.Recordset.Edit
      data_mdb.Recordset("ult_mesp") = data_sql.Recordset("mes")
      data_mdb.Recordset("ult_aniop") = data_sql.Recordset("ano")
      data_mdb.Recordset("pesos") = Xpesos
      data_mdb.Recordset.Update
   End If
   Xpesos = 0
   data_sql.RecordSource = "select * from deudas where fecha_pago is null and cliente =" & data_mdb.Recordset("cl_codigo") & " and mes =" & 0
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
      data_sql.Recordset.MoveFirst
      Do While Not data_sql.Recordset.EOF
         Xpesos = Xpesos + data_sql.Recordset("total")
         data_sql.Recordset.MoveNext
      Loop
      data_mdb.Recordset.Edit
      data_mdb.Recordset("servi") = Xpesos
      data_mdb.Recordset.Update
   End If
   
   data_mdb.Recordset.MoveNext
Loop

Form1.MousePointer = 0
MsgBox "Proceso terminado"

End Sub

Private Sub Command20_Click()
Dim Xma As Long
Dim Xca As Integer
'data_mdb.Connect = "ODBC;DSN=sapp;"
'data_mdb.RecordSource = "Select * from clientes where cl_codconv ='" & "CAAN" & "'"
'data_mdb.RecordSource = "Select * from clientes where cl_codigo =" & 10114245

'data_mdb.Refresh
Command20.Enabled = False
Dim Xlafec As Date

'data_sql.Connect = "ODBC;DSN=sapp;"
data_sql.DatabaseName = App.Path & "\informes.mdb"
data_sql.RecordSource = "infvtas"
data_sql.Refresh
data_mdb.DatabaseName = App.Path & "\sapp.mdb"
data_conv.Connect = ""
data_conv.DatabaseName = App.Path & "\informes.mdb"
data_conv.RecordSource = "infcli"
data_conv.Refresh
If data_conv.Recordset.RecordCount > 0 Then
   data_conv.Recordset.MoveFirst
   Do While Not data_conv.Recordset.EOF
      data_conv.Recordset.Delete
      data_conv.Recordset.MoveNext
   Loop
End If

If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
'   Xma = data_mdb.Recordset("cod_cli")
   Do While Not data_sql.Recordset.EOF
      Xlafec = data_sql.Recordset("fecha") + 3
      data_mdb.RecordSource = "Select * from linmmdd where cod_cli =" & data_sql.Recordset("cod_cli") & " and fecha >#" & Format(data_sql.Recordset("fecha"), "yyyy/mm/dd") & "# and fecha <=#" & Format(Xlafec, "yyyy/mm/dd") & "# and nro_flia =" & 1 & " order by fecha"
      data_mdb.Refresh
      If data_mdb.Recordset.RecordCount > 0 Then
         data_conv.Recordset.AddNew
         data_conv.Recordset("cl_codigo") = data_mdb.Recordset("cod_cli")
         data_conv.Recordset("cl_apellid") = data_mdb.Recordset("nom_cli")
         data_conv.Recordset("cl_fecing") = data_mdb.Recordset("fecha")
         data_conv.Recordset("cl_nrocobr") = data_mdb.Recordset("nro_med_a")
         data_conv.Recordset("cl_nomcobr") = Mid(data_mdb.Recordset("nom_med_a"), 1, 25)
         data_conv.Recordset.Update
      End If
      data_sql.Recordset.MoveNext
   Loop
End If
      
Form1.MousePointer = 0
MsgBox "Terminado"


End Sub

Private Sub Command21_Click()
Dim Ximp, Xiv As Double

Data1.Connect = "odbc;dsn=sappnew;"

Data1.RecordSource = "version"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.Edit
   Data1.Recordset("obs") = "**Modificaciones en versión 01072022v1:" & vbCrLf & "**Se modifica envío de correos en Serv.AP**" & vbCrLf & "**Se modifica en sistema afiliaciones control de cédulas que tiene afiliación.**" & vbCrLf & "**Se modifica control de facturación medicación mutual.**"
   Data1.Recordset.Update
End If

Form1.MousePointer = 0
MsgBox "Terminado"

End Sub

Private Sub Command22_Click()
'data_mdb.Connect = "odbc;DSN=sapp;"
Dim Xlacee As Double
Dim Xcanc As Integer
Dim Xnroaborrar As String
Xnroaborrar = ""

Xcanc = 0
'data_mdb.DatabaseName = App.Path & "\informes.mdb"
data_mdb.Connect = "ODBC;DSN=sapp;"
'data_mdb.RecordSource = ""
'data_mdb.Refresh
Xnroaborrar = InputBox("Ingrese nro")
data_mdb.RecordSource = "Select * from clientes where cl_codigo =" & Xnroaborrar
data_mdb.Refresh

If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
      data_mdb.Recordset.Delete
'      data_mdb.Recordset.MoveNext
'   Loop
End If

'data_mdb.RecordSource = "Select * from caja where base =" & 38
'data_mdb.Refresh
'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      data_mdb.Recordset.Delete
'      data_mdb.Recordset.MoveNext
'   Loop
'End If


'data_sql.DatabaseName = App.Path & "\sapp.mdb"
'data_sql.RecordSource = "Select * from clientes order by cl_cedula"
'data_sql.Refresh
'If data_sql.Recordset.RecordCount > 0 Then
'   data_sql.Recordset.MoveNext
'   Xlacee = Int(data_sql.Recordset("cl_cedula"))
'   Do While Not data_sql.Recordset.EOF
'      If IsNull(data_sql.Recordset("cl_cedula")) = False Then
'         If data_sql.Recordset("cl_Cedula") <> 0 Then
'            If Xlacee = Int(data_sql.Recordset("cl_cedula")) Then
'               Xcanc = Xcanc + 1
'            Else
'               If Xcanc >= 1 Then
'                  data_sql.Recordset.MovePrevious
'                  data_mdb.Recordset.AddNew
'                  data_mdb.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
 '                 data_mdb.Recordset("cl_apellid") = data_sql.Recordset("cl_apellid")
'                  data_mdb.Recordset("cl_cedula") = data_sql.Recordset("cl_cedula")
'                  data_mdb.Recordset("cl_codced") = data_sql.Recordset("cl_codced")
'                  data_mdb.Recordset("cl_codconv") = data_sql.Recordset("cl_codconv")
'                  data_mdb.Recordset.Update
'                  data_sql.Recordset.MoveNext
'                  Xcanc = 0
'               End If
'            End If
'         End If
'      End If
'      Xlacee = Int(data_sql.Recordset("cl_cedula"))
'      data_sql.Recordset.MoveNext
'   Loop
'End If

MsgBox "Proceso terminado"

End Sub

Private Sub Command23_Click()
Dim Xca As Long
On Error GoTo queeshc
data_sql.Connect = "ODBC;DSN=sapp;"
data_sql.RecordSource = "Select * from clientes where estado =" & 1
data_sql.Refresh
'data_sql.RecordSource = "Select * from clientes where estado <>" & 2 & " and cl_fecing >=#" & Format("01/12/2012", "yyyy/mm/dd") & "#"
'data_sql.RecordSource = "Select * from clientes where cl_codconv in ('CCNOS','CCNRE','UNIVS','HEVAN','HEVANO','HEVANR','SMIN','SMINR','IMPNO','GANOS','CASANR','CASANO') and estado =" & 2
'data_sql.RecordSource = "Select * from clientes"
'data_sql.Refresh
'data_sql.Recordset.MoveLast
'xCA = data_sql.Recordset.RecordCount

data_conv.DatabaseName = App.Path & "\sapp.mdb"

data_mdb.Connect = "ODBC;DSN=sapp;"
'data_mdb.RecordSource = "Select * from linmmdd where fecha >=#" & CDate("01/04/2011") & "# and cod_prod <>" & 999
'data_mdb.Refresh
'data_mdb.RecordSource = "Select * from abmsocio where fecha >=#" & Format("01/05/2015", "yyyy/mm/dd") & "# and cl_motivo ='" & "FALLECIDO" & "'"
'data_mdb.Refresh

datai.DatabaseName = App.Path & "\mutuales.mdb"
datai.RecordSource = "infno"
datai.Refresh

If datai.Recordset.RecordCount > 0 Then
   datai.Recordset.MoveFirst
   Do While Not datai.Recordset.EOF
      datai.Recordset.Delete
      datai.Recordset.MoveNext
   Loop
End If

Command23.Enabled = False


If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
'      data_mdb.Recordset.FindFirst "cod_cli =" & data_sql.Recordset("cliente")
'      data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_sql.Recordset("cl_codconv") & "'"
'      data_conv.Refresh
      data_mdb.RecordSource = "Select * from linmmdd where fecha >=#" & Format("01/08/2015", "yyyy/mm/dd") & "# and fecha <=#" & Format("05/09/2016", "yyyy/mm/dd") & "# and cod_cli =" & data_sql.Recordset("cl_codigo") & " and cod_prod in (10001,2,10003,10005,10007,10008,10016,14001) order by fecha"
      data_mdb.Refresh
      If data_mdb.Recordset.RecordCount > 0 Then
      Else
         data_mdb.RecordSource = "Select * from linmmdd where fecha >=#" & Format("01/01/2014", "yyyy/mm/dd") & "# and fecha <=#" & Format("31/07/2015", "yyyy/mm/dd") & "# and cod_cli =" & data_sql.Recordset("cl_codigo") & " and cod_prod in (10001,2,10003,10005,10007,10008,10016,14001) order by fecha"
         data_mdb.Refresh
         If data_mdb.Recordset.RecordCount > 0 Then
            data_mdb.Recordset.MoveLast
            datai.Recordset.AddNew
            datai.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
            datai.Recordset("cl_apellid") = data_sql.Recordset("cl_apellid")
            datai.Recordset("cl_cedula") = Int(data_sql.Recordset("cl_cedula"))
            datai.Recordset("cl_codced") = data_sql.Recordset("cl_codced")
            datai.Recordset("cl_codconv") = data_sql.Recordset("cl_codconv")
            datai.Recordset("cl_nomconv") = data_sql.Recordset("cl_nomconv")
            datai.Recordset("cl_fecing") = data_sql.Recordset("cl_fecing")
            datai.Recordset("cl_zona") = data_sql.Recordset("cl_zona")
            datai.Recordset("fecha_baja") = data_sql.Recordset("fecha_baja")
            datai.Recordset("cl_fultmov") = data_mdb.Recordset("fecha")
            datai.Recordset("cl_nrovend") = data_mdb.Recordset("base")
            datai.Recordset.Update
         End If
      End If
      data_sql.Recordset.MoveNext
   Loop
End If
                  
'

MsgBox "Proceso terminado"

Exit Sub

queeshc:
        If Err.Number = 3155 Then
           MsgBox "Al grabar"
        Else
           MsgBox "ERROR:" & Err.Number
        End If
        
End Sub

Private Sub Command24_Click()

data_mdb.Connect = "ODBC;DSN=sapp;"
data_mdb.RecordSource = "Select * from estudios where codest in (20057)"
data_mdb.Refresh
'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      data_mdb.Recordset.Delete
'      data_mdb.Recordset.MoveNext
'   Loop
'End If
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      If data_mdb.Recordset("codest") = 20057 And data_mdb.Recordset("id") = 1150 Then
         data_mdb.Recordset.Edit
         data_mdb.Recordset("codest") = 20119
         data_mdb.Recordset.Update
      End If
'      If data_mdb.Recordset("codest") = 20044 And data_mdb.Recordset("id") = 1093 Then
'         data_mdb.Recordset.Edit
'         data_mdb.Recordset("codest") = 20116
'         data_mdb.Recordset.Update
'      End If
'      If data_mdb.Recordset("codest") = 20045 And data_mdb.Recordset("id") = 1094 Then
'         data_mdb.Recordset.Edit
'         data_mdb.Recordset("codest") = 20117
'         data_mdb.Recordset.Update
'      End If
'      If data_mdb.Recordset("codest") = 20046 And data_mdb.Recordset("id") = 1095 Then
'         data_mdb.Recordset.Edit
'         data_mdb.Recordset("codest") = 20118
'         data_mdb.Recordset.Update
'      End If
      data_mdb.Recordset.MoveNext
   Loop
   
End If
'data_mdb.Recordset.Edit
'data_mdb.Recordset("id") = 10
'data_mdb.Recordset("hora") = "32.05"
'data_mdb.Recordset("descrip") = "3.38590"
'data_mdb.Recordset.Update

'data_sql.DatabaseName = App.Path & "\acom.mdb"
'data_sql.RecordSource = "acomp"
'data_sql.Refresh

'Data1.Connect = ""
'Data1.DatabaseName = App.Path & "\sapp.mdb"

'If data_sql.Recordset.RecordCount > 0 Then
'   data_sql.Recordset.MoveFirst
'   Do While Not data_sql.Recordset.EOF
'      data_mdb.Recordset.FindFirst "matricula =" & data_sql.Recordset("mat") & " and nrorec =" & data_sql.Recordset("factura")
'      If Not data_mdb.Recordset.NoMatch Then
'         data_mdb.Recordset.Delete
'      End If
'      data_sql.Recordset.MoveNext
'   Loop
'End If

''''''Dim Xdesdefec, Xhastafec As String
''''''Xdesdefec = InputBox("Ingrese desde Fecha:", "Desde")
''''''Xhastafec = InputBox("Ingrese hasta Fecha:", "Hasta")
'data_mdb.DatabaseName = App.Path & "\llamado.mdb"
'data_mdb.RecordSource = "llamado"
'data_mdb.Refresh
''''''''data_sql.DatabaseName = App.Path & "\sapp.mdb"
'''''''data_sql.RecordSource = "Select * from llamado where fecha >=#" & Format(Xdesdefec, "yyyy/mm/dd") & "# And fecha <=#" & Format(Xhastafec, "yyyy/mm/dd") & "#"
'''''''data_sql.Refresh
Form1.MousePointer = 11
'data_mdb.Recordset.Edit
'data_mdb.Recordset("id") = 20
'data_mdb.Recordset.Update

'data_mdb.Recordset.AddNew
'data_mdb.Recordset("MC_NUMERO") = "C10"
'data_mdb.Recordset("MC_DESC") = "LIMITACION DE SERVICIO"
'data_mdb.Recordset("id") = 21
'data_mdb.Recordset.Update

'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      If Len(data_mdb.Recordset("cedula")) = 6 Then
'         Xce = Val(Mid(data_mdb.Recordset("cedula"), 1, 5))
'      End If
'      If Len(data_mdb.Recordset("cedula")) = 7 Then
'         Xce = Val(Mid(data_mdb.Recordset("cedula"), 1, 6))
'      End If
'      If Len(data_mdb.Recordset("cedula")) = 8 Then
'         Xce = Val(Mid(data_mdb.Recordset("cedula"), 1, 7))
'      End If
'      data_sql.RecordSource = "Select * from acomp where cl_cedula =" & Xce
'      data_sql.Refresh
'      If data_sql.Recordset.RecordCount > 0 Then
'         data_mdb.Recordset.Edit
'         data_mdb.Recordset("ingreso") = data_sql.Recordset("cl_fecing")
'         Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_sql.Recordset("cl_codigo")
'         Data1.Refresh
'         If Data1.Recordset.RecordCount > 0 Then
'            If IsNull(Data1.Recordset("fecha_baja")) = False Then
'               data_mdb.Recordset("baja") = Data1.Recordset("fecha_baja")
'            End If
'            data_mdb.Recordset("categ") = Data1.Recordset("cl_codconv")
'            data_mdb.Recordset("nomcat") = Data1.Recordset("cl_nomconv")
'            data_mdb.Recordset.Update
'         Else
'            data_mdb.Recordset("categ") = "NO"
'            data_mdb.Recordset("nomcat") = "NO ENCONTRADO EN SAPP"
'            data_mdb.Recordset.Update
'         End If
'      Else
'         Data1.RecordSource = "Select * from clientes where cl_cedula =" & Xce
'         Data1.Refresh
'         If Data1.Recordset.RecordCount > 0 Then
'            data_mdb.Recordset.Edit
'            data_mdb.Recordset("ingreso") = Data1.Recordset("cl_fecing")
'            If IsNull(Data1.Recordset("fecha_baja")) = False Then
'               data_mdb.Recordset("baja") = Data1.Recordset("fecha_baja")
'            End If
'            data_mdb.Recordset("categ") = Data1.Recordset("cl_codconv")
'            data_mdb.Recordset("nomcat") = Data1.Recordset("cl_nomconv")
'            data_mdb.Recordset.Update
'         Else
'            data_mdb.Recordset.Edit
'            data_mdb.Recordset("categ") = "NO"
'            data_mdb.Recordset("nomcat") = "NO ENCONTRADO EN SAPP"
'            data_mdb.Recordset.Update
'         End If
'      End If
'      data_mdb.Recordset.MoveNext
   
'   Loop
'End If

Form1.MousePointer = 0
MsgBox "Terminado"

End Sub

Private Sub Command25_Click()
Dim Cedula, Codigo As String

data_mdb.Connect = "odbc;dsn=sappnew;"
data_mdb.RecordSource = "cabezal_hc"
data_mdb.Refresh

Form1.MousePointer = 11
data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   

Loop

Form1.MousePointer = 0
MsgBox "Proceso terminado"


End Sub

Private Sub Command26_Click()
frm_infmspmutual.Show vbModal


End Sub

Private Sub Command27_Click()

data_mdb.Connect = "odbc;DSN=sappnew;"
data_mdb.RecordSource = "select * from arqueo where cob in (683) and arqueo in ('C')"
data_mdb.Refresh

data_sql.Connect = "odbc;DSN=sappnew;"
'data_sql.RecordSource = "arqueo"
'data_sql.Refresh

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
'   data_mdb.Recordset.Delete
   Do While Not data_mdb.Recordset.EOF
      data_sql.RecordSource = "select * from deudas where documento =" & data_mdb.Recordset("nrorec")
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.Edit
         data_sql.Recordset("fecha_pago") = Date
         data_sql.Recordset.Update
'      data_sql.RecordSource = "select * from arqueo where nrorec =" & data_mdb.Recordset("documento")
'      data_sql.Refresh
'      If data_sql.Recordset.RecordCount > 0 Then
'      Else
'         data_sql.Recordset.AddNew
'         data_sql.Recordset("matricula") = data_mdb.Recordset("cliente")
'         data_sql.Recordset("nombre") = data_mdb.Recordset("apellidos")
'         data_sql.Recordset("mes") = data_mdb.Recordset("mes")
'         data_sql.Recordset("ano") = data_mdb.Recordset("ano")
'         data_sql.Recordset("color") = data_mdb.Recordset("color_rec")
'         data_sql.Recordset("cat") = data_mdb.Recordset("cod_cnv")
'         data_sql.Recordset("nomcat") = data_mdb.Recordset("nom_cnv")
'         data_sql.Recordset("arqueo") = "E"
'         data_sql.Recordset("importe") = data_mdb.Recordset("importe")
'         data_sql.Recordset("fecha") = Date
'         data_sql.Recordset("nrorec") = data_mdb.Recordset("documento")
'         data_sql.Recordset("usuar") = "JFERNAN"
'         data_sql.Recordset("moneda") = data_mdb.Recordset("moneda")
'         data_sql.Recordset("cob") = data_mdb.Recordset("nro_cobr")
'         data_sql.Recordset("nomcob") = data_mdb.Recordset("nom_cobr")
'         If IsNull(data_mdb.Recordset("grupo")) = False Then
'            data_sql.Recordset("codzon") = data_mdb.Recordset("grupo")
'         Else
'            data_sql.Recordset("codzon") = 0
'         End If
'         data_sql.Recordset("codsup") = data_mdb.Recordset("nro_superv")
'         data_sql.Recordset("codpro") = data_mdb.Recordset("nro_vende")
'         data_sql.Recordset("tiquet") = data_mdb.Recordset("tiquet")
'         data_sql.Recordset("total") = data_mdb.Recordset("total")
'         data_sql.Recordset("varia") = data_mdb.Recordset("deudas")
'         data_sql.Recordset("iva") = data_mdb.Recordset("iva")
'         data_sql.Recordset("deudas") = data_mdb.Recordset("deudas")
'         data_sql.Recordset("servi") = 0
'         data_sql.Recordset.Update
      End If
      data_mdb.Recordset.MoveNext
   Loop
   MsgBox "Proceso terminado"
Else
    MsgBox "No hay datos"
End If


End Sub

Private Sub Command28_Click()

Form1.MousePointer = 11
data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.RecordSource = "select * from clientes where cl_cedula >" & 0 & " and cl_dpto is not null and cl_cedula_t is null and cl_celular_n is null"
data_sql.Refresh
data_sql.Recordset.MoveFirst
Do While Not data_sql.Recordset.EOF
   data_sql.Recordset.Edit
   data_sql.Recordset("cl_cedula_t") = Trim(Str(data_sql.Recordset("cl_cedula"))) & Trim(Str(data_sql.Recordset("cl_codced")))
   data_sql.Recordset("cl_celular_n") = Trim(data_sql.Recordset("cl_dpto"))
   data_sql.Recordset.Update
   data_sql.Recordset.MoveNext
Loop
Form1.MousePointer = 0


MsgBox "Proceso terminado"


End Sub

Private Sub Command29_Click()
Dim Xidel As Integer
Xidel = 1
data_sql.Connect = "ODBC;DSN=sapp;"
'data_sql.RecordSource = "us"
'data_sql.Refresh

data_mdb.Connect = "ODBC;DSN=sapp;"
data_mdb.RecordSource = "medicos"
data_mdb.Refresh

Data1.Connect = "ODBC;DSN=sapp;"
Data1.RecordSource = "meta_tres"
Data1.Refresh


Command23.Enabled = False

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      If IsNull(data_mdb.Recordset("med_socnro")) = False Then
         data_sql.RecordSource = "Select * from us where id =" & data_mdb.Recordset("med_socnro")
         data_sql.Refresh
         If data_sql.Recordset.RecordCount > 0 Then
            Data1.Recordset.AddNew
            Data1.Recordset("m_fecha") = Date
            Data1.Recordset("m_mat") = data_mdb.Recordset("med_cod")
            Data1.Recordset("m_nrofrm") = data_sql.Recordset("documento")
            Data1.Recordset.Update
         End If
      End If
'      data_sql.Recordset.AddNew
'      data_sql.Recordset("id") = data_mdb.Recordset("id")
'      data_sql.Recordset("cedtext") = data_mdb.Recordset("cedtext")
'      data_sql.Recordset("nom1") = data_mdb.Recordset("nom1")
'      data_sql.Recordset("nom2") = data_mdb.Recordset("nom2")
'      data_sql.Recordset("ape1") = data_mdb.Recordset("ape1")
'      data_sql.Recordset("ape2") = data_mdb.Recordset("ape2")
'      data_sql.Recordset("fecnac") = data_mdb.Recordset("fecnac")
'      data_sql.Recordset("fecing") = data_mdb.Recordset("fecing")
'      data_sql.Recordset("sexo") = data_mdb.Recordset("sexo")
'      data_sql.Recordset("estcivil") = data_mdb.Recordset("estcivil")
'      data_sql.Recordset("estado") = data_mdb.Recordset("estado")
'      data_sql.Recordset("fechabaja") = data_mdb.Recordset("fechabaja")
'      data_sql.Recordset("motivo") = data_mdb.Recordset("motivo")
'      data_sql.Recordset("motivod") = data_mdb.Recordset("motivod")
'      data_sql.Recordset("cargo") = data_mdb.Recordset("cargo")
'      data_sql.Recordset("cargod") = data_mdb.Recordset("cargod")
'      data_sql.Recordset("jefe") = data_mdb.Recordset("jefe")
 '     data_sql.Recordset("jefed") = data_mdb.Recordset("jefed")
 '     data_sql.Recordset("hijos") = data_mdb.Recordset("hijos")
 '     data_sql.Recordset("hijoscant") = data_mdb.Recordset("hijoscant")
 '     data_sql.Recordset("nivelest") = data_mdb.Recordset("nivelest")
'      data_sql.Recordset("nivelestd") = data_mdb.Recordset("nivelestd")
'      data_sql.Recordset("tipo") = data_mdb.Recordset("tipo")
'      data_sql.Recordset("tipod") = data_mdb.Recordset("tipod")
'      data_sql.Recordset("profesio") = data_mdb.Recordset("profesio")
'      data_sql.Recordset("contrato") = data_mdb.Recordset("contrato")
'      data_sql.Recordset("contratod") = data_mdb.Recordset("contratod")
'      data_sql.Recordset("foto") = data_mdb.Recordset("foto")
'      data_sql.Recordset("direc") = data_mdb.Recordset("direc")
'      data_sql.Recordset("tel") = data_mdb.Recordset("tel")
'      data_sql.Recordset("nro") = data_mdb.Recordset("nro")
'      data_sql.Recordset("profesion") = data_mdb.Recordset("profesion")
'      data_sql.Recordset("id2") = data_mdb.Recordset("id2")
'      data_sql.Recordset.Update
      data_mdb.Recordset.MoveNext
   Loop
End If

                 
'

MsgBox "Proceso terminado"


End Sub

Private Sub Command3_Click()
Dim Xdesdefec, Xhastafec As String
'Xdesdefec = InputBox("Ingrese desde Fecha:", "Desde")
'Xhastafec = InputBox("Ingrese hasta Fecha:", "Hasta")
data_mdb.DatabaseName = App.Path & "\sapp.mdb"
'data_mdb.RecordSource = "inflla"
'data_mdb.RecordSource = "Select * from llamado where nrolla =" & 70178311
'data_mdb.Refresh

'Data1.DatabaseName = App.Path & "\llamb18.mdb"
'Data1.RecordSource = "llamado2"
'Data1.Refresh
'<> de 50116568

data_sql.DatabaseName = App.Path & "\informes.mdb"
'data_sql.DatabaseName = App.Path & "\sapp.mdb"
data_sql.RecordSource = "Select * from inflla"
'data_sql.RecordSource = "Select * from llamado where nrolla in (50191446,70178620,70178624,70178833)"
data_sql.Refresh
Form1.MousePointer = 11
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_mdb.RecordSource = "Select * from llamado where nrolla =" & data_sql.Recordset("nro")
      data_mdb.Refresh
      If data_mdb.Recordset.RecordCount > 0 Then
          data_mdb.Recordset.Edit
    '      data_mdb.Recordset("nrolla") = data_sql.Recordset("nrolla")
          data_mdb.Recordset("nro") = data_sql.Recordset("nro")
          data_mdb.Recordset("fecha") = data_sql.Recordset("fecha")
          data_mdb.Recordset("hora") = data_sql.Recordset("hora")
          data_mdb.Recordset("usuario") = data_sql.Recordset("usuario")
          data_mdb.Recordset("matric") = data_sql.Recordset("matric")
          data_mdb.Recordset("nombre") = data_sql.Recordset("nombre")
          If IsNull(data_sql.Recordset("edad")) = False Then
             data_mdb.Recordset("edad") = data_sql.Recordset("edad")
          End If
          If IsNull(data_sql.Recordset("unied")) = False Then
             data_mdb.Recordset("unied") = data_sql.Recordset("unied")
          End If
          If IsNull(data_sql.Recordset("categ")) = False Then
             data_mdb.Recordset("categ") = data_sql.Recordset("categ")
          End If
          If IsNull(data_sql.Recordset("nomcat")) = False Then
             data_mdb.Recordset("nomcat") = data_sql.Recordset("nomcat")
          End If
          If IsNull(data_sql.Recordset("ci")) = False Then
             data_mdb.Recordset("ci") = data_sql.Recordset("ci")
          End If
          If IsNull(data_sql.Recordset("direcc")) = False Then
             data_mdb.Recordset("direcc") = data_sql.Recordset("direcc")
          End If
          If IsNull(data_sql.Recordset("telef")) = False Then
             data_mdb.Recordset("telef") = data_sql.Recordset("telef")
          End If
          If IsNull(data_sql.Recordset("referen")) = False Then
             data_mdb.Recordset("referen") = data_sql.Recordset("referen")
          End If
          data_mdb.Recordset("codzon") = data_sql.Recordset("codzon")
          If IsNull(data_sql.Recordset("base")) = False Then
             data_mdb.Recordset("base") = data_sql.Recordset("base")
          End If
          If IsNull(data_sql.Recordset("motcon")) = False Then
             data_mdb.Recordset("motcon") = data_sql.Recordset("motcon")
          End If
          If IsNull(data_sql.Recordset("obsmot")) = False Then
             data_mdb.Recordset("obsmot") = data_sql.Recordset("obsmot")
          End If
          If IsNull(data_sql.Recordset("codmot")) = False Then
             data_mdb.Recordset("codmot") = data_sql.Recordset("codmot")
          End If
          If IsNull(data_sql.Recordset("descol")) = False Then
             data_mdb.Recordset("descol") = data_sql.Recordset("descol")
          End If
          If IsNull(data_sql.Recordset("movilpas")) = False Then
             data_mdb.Recordset("movilpas") = data_sql.Recordset("movilpas")
          End If
          If IsNull(data_sql.Recordset("pend")) = False Then
             data_mdb.Recordset("pend") = data_sql.Recordset("pend")
          End If
          If IsNull(data_sql.Recordset("fecpas")) = False Then
             data_mdb.Recordset("fecpas") = data_sql.Recordset("fecpas")
          End If
          If IsNull(data_sql.Recordset("horpas")) = False Then
             data_mdb.Recordset("horpas") = data_sql.Recordset("horpas")
          End If
          If IsNull(data_sql.Recordset("fecsali")) = False Then
             data_mdb.Recordset("fecsali") = data_sql.Recordset("fecsali")
          End If
          If IsNull(data_sql.Recordset("horsali")) = False Then
             data_mdb.Recordset("horsali") = data_sql.Recordset("horsali")
          End If
          If IsNull(data_sql.Recordset("fec_rea")) = False Then
             data_mdb.Recordset("fec_rea") = data_sql.Recordset("fec_rea")
          End If
          If IsNull(data_sql.Recordset("hor_rea")) = False Then
             data_mdb.Recordset("hor_rea") = data_sql.Recordset("hor_rea")
          End If
          If IsNull(data_sql.Recordset("diag")) = False Then
             data_mdb.Recordset("diag") = data_sql.Recordset("diag")
          End If
          If IsNull(data_sql.Recordset("realiza")) = False Then
             data_mdb.Recordset("realiza") = data_sql.Recordset("realiza")
          End If
          If IsNull(data_sql.Recordset("movil_rea")) = False Then
             data_mdb.Recordset("movil_rea") = data_sql.Recordset("movil_rea")
          End If
          If IsNull(data_sql.Recordset("trasla")) = False Then
             data_mdb.Recordset("trasla") = data_sql.Recordset("trasla")
          End If
          If IsNull(data_sql.Recordset("colormot")) = False Then
             data_mdb.Recordset("colormot") = data_sql.Recordset("colormot")
          End If
          If IsNull(data_sql.Recordset("fec_llega")) = False Then
             data_mdb.Recordset("fec_llega") = data_sql.Recordset("fec_llega")
          End If
          If IsNull(data_sql.Recordset("hor_llega")) = False Then
             data_mdb.Recordset("hor_llega") = data_sql.Recordset("hor_llega")
          End If
          If IsNull(data_sql.Recordset("descol")) = False Then
             data_mdb.Recordset("descol") = data_sql.Recordset("descol")
          End If
          If IsNull(data_sql.Recordset("activo")) = False Then
             data_mdb.Recordset("activo") = data_sql.Recordset("activo")
          End If
          If IsNull(data_sql.Recordset("codmed")) = False Then
             data_mdb.Recordset("codmed") = data_sql.Recordset("codmed")
          End If
          If IsNull(data_sql.Recordset("nommed")) = False Then
             data_mdb.Recordset("nommed") = data_sql.Recordset("nommed")
          End If
          If IsNull(data_sql.Recordset("timdes")) = False Then
             data_mdb.Recordset("timdes") = data_sql.Recordset("timdes")
          End If
          If IsNull(data_sql.Recordset("obs")) = False Then
             data_mdb.Recordset("obs") = data_sql.Recordset("obs")
          End If
          If IsNull(data_sql.Recordset("pasado")) = False Then
             data_mdb.Recordset("pasado") = data_sql.Recordset("pasado")
          End If
          If IsNull(data_sql.Recordset("motmov")) = False Then
             data_mdb.Recordset("motmov") = data_sql.Recordset("motmov")
          End If
          If IsNull(data_sql.Recordset("hsald")) = False Then
             data_mdb.Recordset("hsald") = data_sql.Recordset("hsald")
          End If
          If IsNull(data_sql.Recordset("hllega")) = False Then
             data_mdb.Recordset("hllega") = data_sql.Recordset("hllega")
          End If
          If IsNull(data_sql.Recordset("hzona")) = False Then
             data_mdb.Recordset("hzona") = data_sql.Recordset("hzona")
          End If
          If IsNull(data_sql.Recordset("cancela")) = False Then
             data_mdb.Recordset("cancela") = data_sql.Recordset("cancela")
          End If
          If IsNull(data_sql.Recordset("fec_cance")) = False Then
             data_mdb.Recordset("fec_cance") = data_sql.Recordset("fec_cance")
          End If
          If IsNull(data_sql.Recordset("hor_cance")) = False Then
             data_mdb.Recordset("hor_cance") = data_sql.Recordset("hor_cance")
          End If
          If IsNull(data_sql.Recordset("motcance")) = False Then
             data_mdb.Recordset("motcance") = data_sql.Recordset("motcance")
          End If
          If IsNull(data_sql.Recordset("mes")) = False Then
             data_mdb.Recordset("mes") = data_sql.Recordset("mes")
          End If
          If IsNull(data_sql.Recordset("ano")) = False Then
             data_mdb.Recordset("ano") = data_sql.Recordset("ano")
          End If
          If IsNull(data_sql.Recordset("hh")) = False Then
             data_mdb.Recordset("hh") = data_sql.Recordset("hh")
          End If
          If IsNull(data_sql.Recordset("movtras")) = False Then
             data_mdb.Recordset("movtras") = data_sql.Recordset("movtras")
          End If
          If IsNull(data_sql.Recordset("lugar")) = False Then
             data_mdb.Recordset("lugar") = data_sql.Recordset("lugar")
          End If
          If IsNull(data_sql.Recordset("mm")) = False Then
             data_mdb.Recordset("mm") = data_sql.Recordset("mm")
          End If
          If IsNull(data_sql.Recordset("thh")) = False Then
             data_mdb.Recordset("thh") = data_sql.Recordset("thh")
          End If
          If IsNull(data_sql.Recordset("tmm")) = False Then
             data_mdb.Recordset("tmm") = data_sql.Recordset("tmm")
          End If
          If IsNull(data_sql.Recordset("totdem")) = False Then
             data_mdb.Recordset("totdem") = data_sql.Recordset("totdem")
          End If
          If IsNull(data_sql.Recordset("enfer")) = False Then
             data_mdb.Recordset("enfer") = data_sql.Recordset("enfer")
          End If
          If IsNull(data_sql.Recordset("totend")) = False Then
             data_mdb.Recordset("totend") = data_sql.Recordset("totend")
          End If
          If IsNull(data_sql.Recordset("timsi")) = False Then
             data_mdb.Recordset("timsi") = data_sql.Recordset("timsi")
          End If
          If IsNull(data_sql.Recordset("ncobr")) = False Then
             data_mdb.Recordset("ncobr") = data_sql.Recordset("ncobr")
          End If
          If IsNull(data_sql.Recordset("dcobr")) = False Then
             data_mdb.Recordset("dcobr") = data_sql.Recordset("dcobr")
          End If
            
          data_mdb.Recordset.Update
      End If
      
'      Data1.Recordset.AddNew
        
'      Data1.Recordset.Update
      
      data_sql.Recordset.MoveNext
   Loop
End If

Form1.MousePointer = 0

MsgBox "Proceso terminado"

End Sub

Private Sub Command30_Click()

'data_sql.DatabaseName = App.Path & "\socsapp2.mdb"
data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.RecordSource = "Select * from arq1218 where arqueo ='" & "C" & "' and cob=" & 607
data_sql.Refresh

'data_mdb.DatabaseName = App.Path & "\soc112018.mdb"
data_mdb.Connect = "ODBC;DSN=sappnew;"
'data_mdb.RecordSource = "socyo"
'data_mdb.Refresh

If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_mdb.RecordSource = "Select * from deudas where cliente =" & data_sql.Recordset("matricula") & " and fecha_pago is null and documento =" & data_sql.Recordset("nrorec")
      data_mdb.Refresh
      If data_mdb.Recordset.RecordCount > 0 Then
         data_mdb.Recordset.Edit
         data_mdb.Recordset("fecha_pago") = data_sql.Recordset("fecha")
         data_mdb.Recordset.Update
      Else
'         data_sql.Recordset.Edit
'         data_sql.Recordset("result") = "AL DIA"
'         data_sql.Recordset.Update
      End If
      data_sql.Recordset.MoveNext
   Loop
End If

MsgBox "Terminado"



End Sub

Private Sub Command31_Click()
data_sql.DatabaseName = App.Path & "\estyconvant.mdb"

'data_mdb.DatabaseName = App.Path & "\sapp.mdb"
data_mdb.Connect = "odbc;dsn=sappnew;"
data_mdb.RecordSource = "convenio"
data_mdb.Refresh
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_sql.RecordSource = "Select * from convant where cnv_codigo ='" & data_mdb.Recordset("cnv_codigo") & "'"
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         If data_mdb.Recordset("cnv_precio") <> data_sql.Recordset("cnv_precio") Then
            data_mdb.Recordset.Edit
            data_mdb.Recordset("cnv_precio") = data_sql.Recordset("cnv_precio")
            data_mdb.Recordset.Update
         End If
      End If
'         If IsNull(data_sql.Recordset("uc")) = False Then
'            If IsNull(data_mdb.Recordset("uc")) = False Then
'               If data_mdb.Recordset("uc") <> data_sql.Recordset("uc") Then
'                  data_mdb.Recordset.Edit
'                  data_mdb.Recordset("uc") = data_sql.Recordset("uc")
'                  data_mdb.Recordset.Update
'               End If
'            End If
'         End If
'         If IsNull(data_sql.Recordset("part")) = False Then
'            If IsNull(data_mdb.Recordset("part")) = False Then
'               If data_mdb.Recordset("part") <> data_sql.Recordset("part") Then
'                  data_mdb.Recordset.Edit
'                  data_mdb.Recordset("part") = data_sql.Recordset("part")
'                  data_mdb.Recordset.Update
'               End If
'            End If
'         End If
'         If IsNull(data_sql.Recordset("ucfh")) = False Then
 '           If IsNull(data_mdb.Recordset("ucfh")) = False Then
'               If data_mdb.Recordset("ucfh") <> data_sql.Recordset("ucfh") Then
'                  data_mdb.Recordset.Edit
'                  data_mdb.Recordset("ucfh") = data_sql.Recordset("ucfh")
'                  data_mdb.Recordset.Update
'               End If
'            End If
''         End If
'      End If
      data_mdb.Recordset.MoveNext
   Loop
End If

'data_sql.DatabaseName = App.Path & "\sapp.mdb"
'data_sql.RecordSource = "Select * from cnv_prec where cnv_desde =#" & Format("30/12/2014", "yyyy/mm/dd") & "#"
'data_sql.Refresh
'If data_sql.Recordset.RecordCount > 0 Then
'   data_sql.Recordset.MoveFirst
'   Do While Not data_sql.Recordset.EOF
'      data_sql.Recordset.Delete
'      data_sql.Recordset.MoveNext
'   Loop
'End If


MsgBox "Terminado"

End Sub

Private Sub Command32_Click()

'data_mdb.DatabaseName = App.Path & "\permi.mdb"
'data_mdb.RecordSource = "permi"
'data_mdb.Refresh
'data_mdb.DatabaseName = App.Path & "\deudasu.mdb"
'data_mdb.RecordSource = "select * from ctacte order by cliente"
'data_mdb.Refresh
'data_sql.DatabaseName = App.Path & "\deudasud.mdb"
Form1.MousePointer = 11
data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.RecordSource = "select * from t_fechas"
data_sql.Refresh
data_sql.Recordset.MoveFirst
Do While Not data_sql.Recordset.EOF
   data_sql.Recordset.Edit
   data_sql.Recordset("fecha_cons") = CDate(data_sql.Recordset("fecha"))
   data_sql.Recordset.Update
   data_sql.Recordset.MoveNext
Loop
Form1.MousePointer = 0
MsgBox "Terminado"

'Dim Xmat As Double
'Dim Xcant As Integer
'Xcant = 0
'Xmat = data_mdb.Recordset("cliente")
'data_mdb.Recordset.MoveFirst
'Do While Not data_mdb.Recordset.EOF
'   If data_mdb.Recordset.RecordCount = Xmat Then
'      Xcant = Xcant + 1
'   Else
'      If Xcant > 18 Then
'         data_mdb.Recordset.MovePrevious
'         data_sql.RecordSource = "select * from deudas where cliente =" & data_mdb.Recordset("cliente")
'         data_sql.Refresh
'         data_mdb.Recordset.MoveNext
'         data_sql.Recordset.MoveFirst
'         Do While Not data_sql.Recordset.EOF
'            data_sql.Recordset.Edit
'            data_sql.Recordset("dias") = 1
'            data_sql.Recordset.Update
'            data_sql.Recordset.MoveNext
'         Loop
'      End If
'      Xcant = 1
'   End If
'   Xmat = data_mdb.Recordset("cliente")
'   data_mdb.Recordset.MoveNext
'Loop

'data_sql.Connect = "ODBC;DSN=sappnew;"
'data_sql.RecordSource = "opciones_menu"
'data_sql.Refresh
'data_mdb.Recordset.MoveFirst
'Do While Not data_mdb.Recordset.EOF
'   data_sql.Recordset.AddNew
'   data_sql.Recordset("opcion") = data_mdb.Recordset("opcion")
'   data_sql.Recordset("modulo") = data_mdb.Recordset("modulo")
'   data_sql.Recordset.Update
'   data_mdb.Recordset.MoveNext
'Loop
'Dim Xlaced As Double

'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      If Len(data_mdb.Recordset("ced")) = 7 Then
'         Xlaced = Val(Mid(Trim(data_mdb.Recordset("ced")), 1, 6))
'      Else
'         Xlaced = Val(Mid(Trim(data_mdb.Recordset("ced")), 1, 7))
'      End If
'      data_sql.RecordSource = "select * from clientes where cl_cedula =" & Xlaced
'      data_sql.Refresh
'      If data_sql.Recordset.RecordCount > 0 Then
'         data_mdb.Recordset.Edit
'         data_mdb.Recordset("convenio") = data_sql.Recordset("cl_codconv")
'         data_mdb.Recordset("descconvenio") = data_sql.Recordset("cl_nomconv")
'         data_mdb.Recordset("nombre") = data_sql.Recordset("cl_apellid")
'         data_mdb.Recordset("ingreso") = data_sql.Recordset("cl_fecing")
'         If IsNull(data_sql.Recordset("fecha_baja")) = False Then
'            data_mdb.Recordset("baja") = data_sql.Recordset("fecha_baja")
'         End If
'
'         data_mdb.Recordset.Update
'      End If
'      data_mdb.Recordset.MoveNext
'   Loop
'End If

MsgBox "Terminado"


End Sub

Private Sub Command33_Click()
Dim Xcantsocio, Xcantsocio2, Xcantsocio3, XSocio, Xdias As Integer
Dim Xmat As Double
Xcantsocio = 0
Xcantsocio2 = 0
Xcantsocio3 = 0
XSocio = 0
Data1.DatabaseName = App.Path & "\informes.mdb"
Data1.RecordSource = "infcli"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If
Form1.MousePointer = 11

Data3.Connect = "odbc;dsn=sappnew;"

data_mdb.DatabaseName = App.Path & "\informes.mdb"
data_mdb.RecordSource = "inflla"
data_mdb.Refresh
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_mdb.Recordset.Delete
      data_mdb.Recordset.MoveNext
   Loop
End If

data_sql.Connect = "odbc;dsn=sappnew;"
'data_sql.RecordSource = "select from hc_prescrip where hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# and hc_fecha <=#" & Format("06/07/2020", "yyyy/mm/dd") & "# and hc_tippresd in ('RECETA PACIENTE CRONICO') and hc_codmedica >" & 0 & " and motivo_cance is null order by hc_mat"
'data_sql.Refresh
data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
"hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
"hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
"inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
"hc_prescrip.hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format("10/07/2020", "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('RECETA PACIENTE CRONICO') and hc_prescrip.hc_codmedica >" & 0 & " and cabezal_hcdig.hc_base=" & 18 & " and hc_prescrip.motivo_cance is null order by hc_prescrip.hc_mat"
data_sql.Refresh

If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_mdb.Recordset.AddNew
      data_mdb.Recordset("matric") = data_sql.Recordset("hc_mat")
      data_mdb.Recordset("nombre") = data_sql.Recordset("hc_tippresd")
      data_mdb.Recordset("fecha") = data_sql.Recordset("hc_fecha")
      data_mdb.Recordset("fecpas") = data_sql.Recordset("hc_comfec")
      data_mdb.Recordset("fecsali") = data_sql.Recordset("hc_hastaf")
      data_mdb.Recordset("diag") = data_sql.Recordset("hc_indicanom")
      data_mdb.Recordset("fec_rea") = data_sql.Recordset("hc_fecentrega")
      data_mdb.Recordset("motcon") = data_sql.Recordset("hc_descrip")
      data_mdb.Recordset("movilpas") = data_sql.Recordset("hc_nro")
      data_mdb.Recordset.Update
      data_sql.Recordset.MoveNext
   Loop
   Xcantsocio = 0
   Xcantsocio2 = 0
   Xcantsocio3 = 0
   data_sql.Recordset.MoveFirst
   Xmat = data_sql.Recordset("hc_mat")
   Do While Not data_sql.Recordset.EOF
      If Xmat = data_sql.Recordset("hc_mat") Then
         Xdias = DateDiff("d", data_sql.Recordset("hc_fecha"), data_sql.Recordset("hc_comfec"))
         If Xdias < 30 Then
            Xcantsocio = Xcantsocio + 1
         Else
            If Xdias >= 30 And Xdias < 60 Then
               Xcantsocio2 = Xcantsocio2 + 1
            Else
               Xcantsocio3 = Xcantsocio3 + 1
            End If
         End If
         Xmat = data_sql.Recordset("hc_mat")
         data_sql.Recordset.MoveNext
      Else
         XSocio = XSocio + 1
         data_sql.Recordset.MovePrevious
         Data1.Recordset.AddNew
         Data1.Recordset("cl_codigo") = data_sql.Recordset("hc_mat")
         If Xcantsocio > 0 Then
            Data1.Recordset("cl_codced") = 1
         End If
         If Xcantsocio2 > 0 Then
            Data1.Recordset("cl_cedula") = 1
         End If
         If Xcantsocio3 > 0 Then
            Data1.Recordset("cl_nrovend") = 1
         End If
         Data1.Recordset("cl_direcci") = data_sql.Recordset("hc_tippresd")
         Data1.Recordset("cl_fecing") = data_sql.Recordset("hc_fecha")
         Data1.Recordset("cl_fnac") = data_sql.Recordset("hc_comfec")
         Data1.Recordset("cl_fultmov") = data_sql.Recordset("hc_hastaf")
         Data1.Recordset("cl_apellid") = data_sql.Recordset("hc_indicanom")
         Data1.Recordset("cl_fultvta") = data_sql.Recordset("hc_fecentrega")
         Data1.Recordset("cl_grupo") = data_sql.Recordset("hc_base")
         Data3.RecordSource = "select * from clientes where cl_codigo =" & data_sql.Recordset("hc_mat")
         Data3.Refresh
         If Data3.Recordset.RecordCount > 0 Then
            Data1.Recordset("cl_nombre") = Mid(Data3.Recordset("cl_apellid"), 1, 30)
            Data1.Recordset("cl_codconv") = Data3.Recordset("cl_codconv")
         End If
         Data1.Recordset.Update
         data_sql.Recordset.MoveNext
         Xmat = data_sql.Recordset("hc_mat")
         Xcantsocio = 0
         Xcantsocio2 = 0
         Xcantsocio3 = 0
      End If
   Loop
End If

data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
"hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
"hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
"inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
"hc_prescrip.hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format("10/07/2020", "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('MEDICACION') and hc_prescrip.hc_codmedica >" & 0 & " and cabezal_hcdig.hc_base=" & 18 & " and hc_prescrip.motivo_cance is null order by hc_prescrip.hc_mat"
data_sql.Refresh

If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_mdb.Recordset.AddNew
      data_mdb.Recordset("matric") = data_sql.Recordset("hc_mat")
      data_mdb.Recordset("nombre") = data_sql.Recordset("hc_tippresd")
      data_mdb.Recordset("fecha") = data_sql.Recordset("hc_fecha")
      data_mdb.Recordset("fecpas") = data_sql.Recordset("hc_comfec")
      data_mdb.Recordset("fecsali") = data_sql.Recordset("hc_hastaf")
      data_mdb.Recordset("diag") = data_sql.Recordset("hc_indicanom")
      data_mdb.Recordset("fec_rea") = data_sql.Recordset("hc_fecentrega")
      data_mdb.Recordset("motcon") = data_sql.Recordset("hc_descrip")
      data_mdb.Recordset("movilpas") = data_sql.Recordset("hc_nro")
      data_mdb.Recordset.Update
      data_sql.Recordset.MoveNext
   Loop
   
   data_sql.Recordset.MoveFirst
   Xmat = data_sql.Recordset("hc_mat")
   Do While Not data_sql.Recordset.EOF
      If Xmat = data_sql.Recordset("hc_mat") Then
         Xcantsocio = Xcantsocio + 1
      Else
         XSocio = XSocio + 1
         data_sql.Recordset.MovePrevious
         Data1.Recordset.AddNew
         Data1.Recordset("cl_codigo") = data_sql.Recordset("hc_mat")
         Data1.Recordset("cl_codced") = Xcantsocio
         Data1.Recordset("cl_direcci") = data_sql.Recordset("hc_tippresd")
         Data1.Recordset("cl_fecing") = data_sql.Recordset("hc_fecha")
         Data1.Recordset("cl_apellid") = data_sql.Recordset("hc_indicanom")
         Data1.Recordset("cl_grupo") = data_sql.Recordset("hc_base")
         Data3.RecordSource = "select * from clientes where cl_codigo =" & data_sql.Recordset("hc_mat")
         Data3.Refresh
         If Data3.Recordset.RecordCount > 0 Then
            Data1.Recordset("cl_nombre") = Mid(Data3.Recordset("cl_apellid"), 1, 30)
            Data1.Recordset("cl_codconv") = Data3.Recordset("cl_codconv")
         End If
         Data1.Recordset.Update
         data_sql.Recordset.MoveNext
         Xcantsocio = 1
      End If
      Xmat = data_sql.Recordset("hc_mat")
      data_sql.Recordset.MoveNext
   Loop
End If

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      If IsNull(Data1.Recordset("cl_codced")) = False Then
         If Data1.Recordset("cl_codced") >= 1 Then
            Data1.Recordset.Edit
            Data1.Recordset("cl_atrasoa") = 30
            Data1.Recordset.Update
         End If
      End If
      If IsNull(Data1.Recordset("cl_cedula")) = False Then
         If Data1.Recordset("cl_cedula") >= 1 Then
            Data1.Recordset.Edit
            Data1.Recordset("cl_atrasoa") = 60
            Data1.Recordset.Update
         End If
      End If
      If IsNull(Data1.Recordset("cl_nrovend")) = False Then
         If Data1.Recordset("cl_nrovend") >= 1 Then
            Data1.Recordset.Edit
            Data1.Recordset("cl_atrasoa") = 90
            Data1.Recordset.Update
         End If
      End If
      Data1.Recordset.MoveNext
   Loop
End If
Form1.MousePointer = 0
MsgBox "Terminado"




End Sub

Private Sub Command34_Click()
Data1.DatabaseName = App.Path & "\informes.mdb"
Data1.RecordSource = "infcli"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If
'data_mdb.Connect = "odbc;dsn=sappnew;"
data_mdb.DatabaseName = App.Path & "\simula.mdb"
'data_mdb.RecordSource = "select * from emisim"
'data_mdb.Refresh

'data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.DatabaseName = App.Path & "\simulan.mdb"
data_sql.RecordSource = "select * from emisim"
data_sql.Refresh
Form1.MousePointer = 11

If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_mdb.RecordSource = "select * from emisim where cliente =" & data_sql.Recordset("cliente")
      data_mdb.Refresh
      If data_mdb.Recordset.RecordCount > 0 Then
      Else
         Data1.Recordset.AddNew
         Data1.Recordset("cl_codigo") = data_sql.Recordset("cliente")
         Data1.Recordset("cl_apellid") = data_sql.Recordset("apellidos")
         Data1.Recordset("cl_codconv") = data_sql.Recordset("cod_cnv")
         Data1.Recordset.Update
      End If
      
      data_sql.Recordset.MoveNext
   Loop
End If
data_sql.DatabaseName = ""
data_sql.Connect = "odbc;dsn=sappnew;"

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      data_sql.RecordSource = "select * from clientes where cl_codigo =" & Data1.Recordset("cl_codigo")
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         Data1.Recordset.Edit
         If IsNull(data_sql.Recordset("fecha_baja")) = False Then
            Data1.Recordset("fecha_baja") = data_sql.Recordset("fecha_baja")
            Data1.Recordset("info_debit") = "BAJA"
         Else
            Data1.Recordset("info_debit") = data_sql.Recordset("cl_codconv") & " EMI:" & Trim(Str(data_sql.Recordset("mesproxemi"))) & "/" & Trim(Str(data_sql.Recordset("anoproxemi"))) & " COB:" & data_sql.Recordset("cl_nrocobr")
         End If
         Data1.Recordset.Update
      End If
      Data1.Recordset.MoveNext
      
   Loop
End If

Form1.MousePointer = 0

MsgBox "Terminado"

End Sub

Private Sub Command35_Click()
data_mdb.Connect = ""
data_mdb.DatabaseName = App.Path & "\sapp.mdb"
data_mdb.RecordSource = "resplla"
data_mdb.Refresh

data_sql.DatabaseName = App.Path & "\llamb18.mdb"
data_sql.RecordSource = "llamado2"
data_sql.Refresh
Form1.MousePointer = 11
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_mdb.Recordset.AddNew
      data_mdb.Recordset("nrolla") = data_sql.Recordset("nrolla")
      data_mdb.Recordset("nro") = data_sql.Recordset("nro")
      If IsNull(data_sql.Recordset("fecha")) = False Then
         data_mdb.Recordset("fecha") = data_sql.Recordset("fecha")
      End If
      If IsNull(data_sql.Recordset("hora")) = False Then
         data_mdb.Recordset("hora") = data_sql.Recordset("hora")
      End If
      If IsNull(data_sql.Recordset("usuario")) = False Then
         data_mdb.Recordset("usuario") = data_sql.Recordset("usuario")
      End If
      If IsNull(data_sql.Recordset("matric")) = False Then
         data_mdb.Recordset("matric") = data_sql.Recordset("matric")
      End If
      If IsNull(data_sql.Recordset("nombre")) = False Then
         data_mdb.Recordset("nombre") = data_sql.Recordset("nombre")
      End If
      If IsNull(data_sql.Recordset("edad")) = False Then
         data_mdb.Recordset("edad") = data_sql.Recordset("edad")
      End If
      If IsNull(data_sql.Recordset("unied")) = False Then
         data_mdb.Recordset("unied") = data_sql.Recordset("unied")
      End If
      If IsNull(data_sql.Recordset("categ")) = False Then
         data_mdb.Recordset("categ") = data_sql.Recordset("categ")
      End If
      If IsNull(data_sql.Recordset("nomcat")) = False Then
         data_mdb.Recordset("nomcat") = data_sql.Recordset("nomcat")
      End If
      If IsNull(data_sql.Recordset("ci")) = False Then
         data_mdb.Recordset("ci") = data_sql.Recordset("ci")
      End If
      If IsNull(data_sql.Recordset("direcc")) = False Then
         data_mdb.Recordset("direcc") = data_sql.Recordset("direcc")
      End If
      If IsNull(data_sql.Recordset("telef")) = False Then
         data_mdb.Recordset("telef") = data_sql.Recordset("telef")
      End If
      If IsNull(data_sql.Recordset("referen")) = False Then
         data_mdb.Recordset("referen") = data_sql.Recordset("referen")
      End If
      data_mdb.Recordset("codzon") = data_sql.Recordset("codzon")
      If IsNull(data_sql.Recordset("base")) = False Then
         data_mdb.Recordset("base") = data_sql.Recordset("base")
      End If
      If IsNull(data_sql.Recordset("motcon")) = False Then
         data_mdb.Recordset("motcon") = data_sql.Recordset("motcon")
      End If
      If IsNull(data_sql.Recordset("obsmot")) = False Then
         data_mdb.Recordset("obsmot") = data_sql.Recordset("obsmot")
      End If
      If IsNull(data_sql.Recordset("codmot")) = False Then
         data_mdb.Recordset("codmot") = data_sql.Recordset("codmot")
      End If
      If IsNull(data_sql.Recordset("descol")) = False Then
         data_mdb.Recordset("descol") = data_sql.Recordset("descol")
      End If
      If IsNull(data_sql.Recordset("movilpas")) = False Then
         data_mdb.Recordset("movilpas") = data_sql.Recordset("movilpas")
      End If
      If IsNull(data_sql.Recordset("pend")) = False Then
         data_mdb.Recordset("pend") = data_sql.Recordset("pend")
      End If
      If IsNull(data_sql.Recordset("fecpas")) = False Then
         data_mdb.Recordset("fecpas") = data_sql.Recordset("fecpas")
      End If
      If IsNull(data_sql.Recordset("horpas")) = False Then
         data_mdb.Recordset("horpas") = data_sql.Recordset("horpas")
      End If
      If IsNull(data_sql.Recordset("fecsali")) = False Then
         data_mdb.Recordset("fecsali") = data_sql.Recordset("fecsali")
      End If
      If IsNull(data_sql.Recordset("horsali")) = False Then
         data_mdb.Recordset("horsali") = data_sql.Recordset("horsali")
      End If
      If IsNull(data_sql.Recordset("fec_rea")) = False Then
         data_mdb.Recordset("fec_rea") = data_sql.Recordset("fec_rea")
      End If
      If IsNull(data_sql.Recordset("hor_rea")) = False Then
         data_mdb.Recordset("hor_rea") = data_sql.Recordset("hor_rea")
      End If
      If IsNull(data_sql.Recordset("diag")) = False Then
         data_mdb.Recordset("diag") = data_sql.Recordset("diag")
      End If
      If IsNull(data_sql.Recordset("realiza")) = False Then
         data_mdb.Recordset("realiza") = data_sql.Recordset("realiza")
      End If
      If IsNull(data_sql.Recordset("movil_rea")) = False Then
         data_mdb.Recordset("movil_rea") = data_sql.Recordset("movil_rea")
      End If
      If IsNull(data_sql.Recordset("trasla")) = False Then
         data_mdb.Recordset("trasla") = data_sql.Recordset("trasla")
      End If
      If IsNull(data_sql.Recordset("colormot")) = False Then
         data_mdb.Recordset("colormot") = data_sql.Recordset("colormot")
      End If
      If IsNull(data_sql.Recordset("fec_llega")) = False Then
         data_mdb.Recordset("fec_llega") = data_sql.Recordset("fec_llega")
      End If
      If IsNull(data_sql.Recordset("hor_llega")) = False Then
         data_mdb.Recordset("hor_llega") = data_sql.Recordset("hor_llega")
      End If
      If IsNull(data_sql.Recordset("descol")) = False Then
         data_mdb.Recordset("descol") = data_sql.Recordset("descol")
      End If
      If IsNull(data_sql.Recordset("activo")) = False Then
         data_mdb.Recordset("activo") = data_sql.Recordset("activo")
      End If
      If IsNull(data_sql.Recordset("codmed")) = False Then
         data_mdb.Recordset("codmed") = data_sql.Recordset("codmed")
      End If
      If IsNull(data_sql.Recordset("nommed")) = False Then
         data_mdb.Recordset("nommed") = data_sql.Recordset("nommed")
      End If
      If IsNull(data_sql.Recordset("timdes")) = False Then
         data_mdb.Recordset("timdes") = data_sql.Recordset("timdes")
      End If
      If IsNull(data_sql.Recordset("obs")) = False Then
         data_mdb.Recordset("obs") = data_sql.Recordset("obs")
      End If
      If IsNull(data_sql.Recordset("pasado")) = False Then
         data_mdb.Recordset("pasado") = data_sql.Recordset("pasado")
      End If
      If IsNull(data_sql.Recordset("motmov")) = False Then
         data_mdb.Recordset("motmov") = data_sql.Recordset("motmov")
      End If
      If IsNull(data_sql.Recordset("hsald")) = False Then
         data_mdb.Recordset("hsald") = data_sql.Recordset("hsald")
      End If
      If IsNull(data_sql.Recordset("hllega")) = False Then
         data_mdb.Recordset("hllega") = data_sql.Recordset("hllega")
      End If
      If IsNull(data_sql.Recordset("hzona")) = False Then
         data_mdb.Recordset("hzona") = data_sql.Recordset("hzona")
      End If
      If IsNull(data_sql.Recordset("cancela")) = False Then
         data_mdb.Recordset("cancela") = data_sql.Recordset("cancela")
      End If
      If IsNull(data_sql.Recordset("fec_cance")) = False Then
         data_mdb.Recordset("fec_cance") = data_sql.Recordset("fec_cance")
      End If
      If IsNull(data_sql.Recordset("hor_cance")) = False Then
         data_mdb.Recordset("hor_cance") = data_sql.Recordset("hor_cance")
      End If
      If IsNull(data_sql.Recordset("motcance")) = False Then
         data_mdb.Recordset("motcance") = data_sql.Recordset("motcance")
      End If
      If IsNull(data_sql.Recordset("mes")) = False Then
         data_mdb.Recordset("mes") = data_sql.Recordset("mes")
      End If
      If IsNull(data_sql.Recordset("ano")) = False Then
         data_mdb.Recordset("ano") = data_sql.Recordset("ano")
      End If
      If IsNull(data_sql.Recordset("hh")) = False Then
         data_mdb.Recordset("hh") = data_sql.Recordset("hh")
      End If
      If IsNull(data_sql.Recordset("movtras")) = False Then
         data_mdb.Recordset("movtras") = data_sql.Recordset("movtras")
      End If
      If IsNull(data_sql.Recordset("lugar")) = False Then
         data_mdb.Recordset("lugar") = data_sql.Recordset("lugar")
      End If
      If IsNull(data_sql.Recordset("mm")) = False Then
         data_mdb.Recordset("mm") = data_sql.Recordset("mm")
      End If
      If IsNull(data_sql.Recordset("thh")) = False Then
         data_mdb.Recordset("thh") = data_sql.Recordset("thh")
      End If
      If IsNull(data_sql.Recordset("tmm")) = False Then
         data_mdb.Recordset("tmm") = data_sql.Recordset("tmm")
      End If
      If IsNull(data_sql.Recordset("totdem")) = False Then
         data_mdb.Recordset("totdem") = data_sql.Recordset("totdem")
      End If
      If IsNull(data_sql.Recordset("enfer")) = False Then
         data_mdb.Recordset("enfer") = data_sql.Recordset("enfer")
      End If
      If IsNull(data_sql.Recordset("totend")) = False Then
         data_mdb.Recordset("totend") = data_sql.Recordset("totend")
      End If
      If IsNull(data_sql.Recordset("timsi")) = False Then
         data_mdb.Recordset("timsi") = data_sql.Recordset("timsi")
      End If
      If IsNull(data_sql.Recordset("ncobr")) = False Then
         data_mdb.Recordset("ncobr") = data_sql.Recordset("ncobr")
      End If
      If IsNull(data_sql.Recordset("dcobr")) = False Then
         data_mdb.Recordset("dcobr") = data_sql.Recordset("dcobr")
      End If
      data_mdb.Recordset.Update
      data_sql.Recordset.MoveNext
   Loop
End If
Form1.MousePointer = 0

MsgBox "Terminado"

End Sub

Private Sub Command36_Click()
'data_mdb.DatabaseName = App.Path & "\otrosdesp.mdb"
Dim Xfechatres, Xfechauno As Date
Dim Xlamat As Double
Dim xcantt As Integer

Xfechatres = Date - 3
Xfechauno = Date - 1
xcantt = 0
data_mdb.Connect = "ODBC;DSN=sappnew;"
data_mdb.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfechatres, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfechauno, "yyyy/mm/dd") & "# and matric not in (0) order by matric"
data_mdb.Refresh

Form1.MousePointer = 11

'<> de 50116568
'data_sql.Connect = "ODBC;DSN=sapp;"
data_sql.DatabaseName = App.Path & "\informes.mdb"
data_sql.RecordSource = "Select * from inflla"
data_sql.Refresh
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_sql.Recordset.Delete
      data_sql.Recordset.MoveNext
   Loop
End If

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Xlamat = data_mdb.Recordset("matric")
   Do While Not data_mdb.Recordset.EOF
      If data_mdb.Recordset("matric") = Xlamat Then
         xcantt = xcantt + 1
      Else
         If xcantt >= 4 Then
            data_mdb.Recordset.MovePrevious
            data_sql.Recordset.AddNew
            data_sql.Recordset("matric") = data_mdb.Recordset("matric")
            data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
            data_sql.Recordset("hora") = data_mdb.Recordset("hora")
            data_sql.Recordset("nombre") = data_mdb.Recordset("nombre")
            data_sql.Recordset("categ") = data_mdb.Recordset("categ")
            data_sql.Recordset("edad") = data_mdb.Recordset("edad")
            data_sql.Recordset("codmot") = data_mdb.Recordset("codmot")
            data_sql.Recordset("obsmot") = data_mdb.Recordset("obsmot")
            data_sql.Recordset("nommed") = data_mdb.Recordset("nommed")
            data_sql.Recordset.Update
            data_mdb.Recordset.MoveNext
         End If
         xcantt = 1
      End If
      Xlamat = data_mdb.Recordset("matric")
      data_mdb.Recordset.MoveNext
   Loop
End If


Form1.MousePointer = 0

MsgBox "Proceso terminado"


End Sub

Private Sub Command37_Click()
'Dim Xfecstr, Xlafecanota As String
'Dim ultimo, Xmes1, Xmes2, Xano1, Xano2, Xdiasmes1, Xdiasmes2, Xdiasdif As Integer
'Dim Xfec1, Xfec2 As Date

'   data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha_anota) <=#" & Format(Date, "yyyy/mm/dd") & "# and cod_med =" & Data1.Recordset("cod_med") & " and base =" & Data1.Recordset("base") & " order by cdate(fecha)"
'   data_cabfec.Refresh
'                 data_fechas.Recordset("fecha") = Xfecstr
'data_borrados.Connect = "ODBC;DSN=sappespecial;"
data_sql.DatabaseName = App.Path & "\informes.mdb"

Data1.Connect = "ODBC;DSN=sappespecial;"
'Data1.RecordSource = "select * from emitiq where mat =" & 10133957
'Data1.Refresh
data_mdb.DatabaseName = App.Path & "\informes.mdb"
data_mdb.RecordSource = "select * from infvtas where cod_prod in (2,3)"
data_mdb.Refresh
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      Data1.RecordSource = "Select * from medicos_esp where cod_sapp =" & data_mdb.Recordset("nro_med_a")
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         data_mdb.Recordset.Edit
         data_mdb.Recordset("nom_flia") = Data1.Recordset("esp_med")
         data_mdb.Recordset("nom_prod") = Data1.Recordset("esp_med")
         data_mdb.Recordset("cod_prod") = data_mdb.Recordset("nro_med_a")
         data_mdb.Recordset.Update
      End If
      data_mdb.Recordset.MoveNext
   Loop
End If
data_mdb.RecordSource = "select * from infvtas where cod_prod =" & 190033
data_mdb.Refresh
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_sql.RecordSource = "select * from infvtas where cod_prod =" & 14001 & " and fecha =#" & Format(data_mdb.Recordset("fecha"), "yyyy/mm/dd") & "#"
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.Delete
      End If
      data_mdb.Recordset.MoveNext
   Loop


End If
MsgBox "Terminado"


End Sub

Private Sub Command38_Click()

'data_mdb.DatabaseName = App.Path & "\cedulas.mdb"
'data_mdb.RecordSource = "bajas_ced"
'data_mdb.Refresh

'data_sql.Connect = "odbc;dsn=sappnew;"

'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      data_sql.RecordSource = "Select * from clientes where cl_codigo =" & data_mdb.Recordset("cl_codigo")
'      data_sql.Refresh
'      If data_sql.Recordset.RecordCount > 0 Then
'         If IsNull(data_mdb.Recordset("cl_nro_sup")) = False Then
'            If IsNull(data_sql.Recordset("cl_codced")) = False Then
'                If Val(data_sql.Recordset("cl_codced")) <> Val(data_mdb.Recordset("cl_nro_sup")) Then
'                   data_sql.Recordset.Edit
'                   data_sql.Recordset("cl_codced") = data_mdb.Recordset("cl_nro_sup")
'                   data_sql.Recordset.Update
'                End If
'            End If
'         End If
'      End If
'      data_mdb.Recordset.MoveNext
'   Loop
'End If
data_sql.Connect = "odbc;dsn=sappnew;"

data_mdb.DatabaseName = App.Path & "\mutuales.mdb"
data_mdb.RecordSource = "bmut7"
data_mdb.Refresh
data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   data_sql.RecordSource = "select * from abmsocio where cl_codigo =" & data_mdb.Recordset("cl_codigo") & " and fecha>=#" & Format(data_mdb.Recordset("fecha_baja"), "yyyy/mm/dd") & "# and desc in ('BAJA')"
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
      data_mdb.Recordset.Edit
      data_mdb.Recordset("motivo") = data_sql.Recordset("cl_motivo")
      data_mdb.Recordset.Update
   End If
   data_mdb.Recordset.MoveNext
Loop


data_mdb.RecordSource = "bsapp7"
data_mdb.Refresh
data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   data_sql.RecordSource = "select * from abmsocio where cl_codigo =" & data_mdb.Recordset("cl_codigo") & " and fecha>=#" & Format(data_mdb.Recordset("fecha_baja"), "yyyy/mm/dd") & "# and desc in ('BAJA')"
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
      data_mdb.Recordset.Edit
      data_mdb.Recordset("motivo") = data_sql.Recordset("cl_motivo")
      data_mdb.Recordset.Update
   End If
   data_mdb.Recordset.MoveNext
Loop


'data_mdb.RecordSource = "carne7"
'data_mdb.Refresh
'data_mdb.Recordset.MoveFirst
'Do While Not data_mdb.Recordset.EOF
'   data_sql.RecordSource = "select * from clientes where cl_codigo =" & data_mdb.Recordset("cod_cli")
'   data_sql.Refresh
'   If data_sql.Recordset.RecordCount > 0 Then
'      data_mdb.Recordset.Edit
'      data_mdb.Recordset("cnv_codigo") = data_sql.Recordset("cl_codconv")
'      If IsNull(data_sql.Recordset("cl_dpto")) = False Then
'         data_mdb.Recordset("obs1") = data_sql.Recordset("cl_dpto")
'      Else
'         If IsNull(data_sql.Recordset("cl_telefon")) = False Then
'            data_mdb.Recordset("obs1") = data_sql.Recordset("cl_telefon")
'         End If
''      End If
'      data_mdb.Recordset("obs2") = Trim(Str(data_sql.Recordset("estado")))
'      data_mdb.Recordset.Update
''   End If
'   data_mdb.Recordset.MoveNext
'Loop


data_mdb.RecordSource = "vacunas7"
data_mdb.Refresh
data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   data_sql.RecordSource = "select * from clientes where cl_codigo =" & data_mdb.Recordset("cod_cli")
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
      data_mdb.Recordset.Edit
      data_mdb.Recordset("cnv_codigo") = data_sql.Recordset("cl_codconv")
      If IsNull(data_sql.Recordset("cl_dpto")) = False Then
         data_mdb.Recordset("obs1") = data_sql.Recordset("cl_dpto")
      Else
         If IsNull(data_sql.Recordset("cl_telefon")) = False Then
            data_mdb.Recordset("obs1") = data_sql.Recordset("cl_telefon")
         End If
      End If
      data_mdb.Recordset("ob2") = Trim(Str(data_sql.Recordset("estado")))
      data_mdb.Recordset.Update
   End If
   data_mdb.Recordset.MoveNext
Loop



'If data_mdb.Recordset.RecordCount > 0 Then
'   data_mdb.Recordset.MoveFirst
'   Do While Not data_mdb.Recordset.EOF
'      If IsNull(data_mdb.Recordset("f#_realiza")) = False Then
'         data_sql.RecordSource = "select * from linmmdd where cod_prod =" & 30081 & " and ced_socio =" & Val(data_mdb.Recordset("ced")) & " and fecha =#" & Format(data_mdb.Recordset("f#_realiza"), "yyyy/mm/dd") & "#"
'         data_sql.Refresh
'         If data_sql.Recordset.RecordCount > 0 Then
'            data_mdb.Recordset.Edit
'            data_mdb.Recordset("sapp") = "SI"
'            data_mdb.Recordset.Update
'         Else
'            data_mdb.Recordset.Edit
'            data_mdb.Recordset("sapp") = "NO"
'            data_mdb.Recordset.Update
'         End If
'      Else
'         data_mdb.Recordset.Edit
'         data_mdb.Recordset("sapp") = "NO"
'         data_mdb.Recordset.Update
'      End If
'      data_mdb.Recordset.MoveNext
'   Loop
'End If
   
   
MsgBox "Terminado"

End Sub

Private Sub Command39_Click()
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, xtot, Xlacedu As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo As Long
Dim Xelnrof As Long
Dim Xrruu As Double
Dim Buscarut As String

data_sql.Connect = ""
data_sql.DatabaseName = App.Path & "\udebaj.mdb"
data_sql.RecordSource = "sociostel"
data_sql.Refresh

data_mdb.DatabaseName = App.Path & "\udebaj.mdb"
'data_mdb.RecordSource = "cash21"
'data_mdb.Refresh
''data_sql.DatabaseName = App.Path & "\bajud.mdb"

data_sql.Recordset.MoveFirst
Do While Not data_sql.Recordset.EOF
   If IsNull(data_sql.Recordset("cedula")) = False Then
      data_mdb.RecordSource = "select * from cliude where cl_cedula =" & data_sql.Recordset("cedula")
        data_mdb.Refresh
        If data_mdb.Recordset.RecordCount > 0 Then
           data_sql.Recordset.Edit
           If data_mdb.Recordset("estado") = 2 Then
              data_sql.Recordset("obs") = "BAJA EN UDEMM"
           Else
              data_sql.Recordset("obs") = "ACTIVO EN UDEMM"
           End If
           data_sql.Recordset("convenio") = data_mdb.Recordset("cl_nomconv")
           data_sql.Recordset("fecha_ing") = data_mdb.Recordset("cl_fecing")
           data_sql.Recordset("telefs") = data_mdb.Recordset("cl_telefon")
           data_sql.Recordset("cobrador") = data_mdb.Recordset("cl_nomcobr")
           data_sql.Recordset.Update
        Else
           data_sql.Recordset.Edit
           data_sql.Recordset("obs") = "NO ENCONTRADO"
           data_sql.Recordset.Update
        End If
   Else
           data_sql.Recordset.Edit
           data_sql.Recordset("obs") = "NO ENCONTRADO"
           data_sql.Recordset.Update
   
   End If
   data_sql.Recordset.MoveNext
Loop

MsgBox "Terminado"

End Sub

Private Sub Command4_Click()
Command4.Enabled = False
'''data_sql.DatabaseName = App.Path & "\sapp.mdb"
'''data_sql.RecordSource = "Select * from tesorero where fecha =#" & Format("21/12/2011", "yyyy/mm/dd") & "# and usuario ='" & "MPEREZ" & "'"
'''data_sql.Refresh
'''data_mdb.DatabaseName = App.Path & "\repest.mdb"
'''data_mdb.RecordSource = "estudios"
'''data_mdb.Refresh
'''If data_sql.Recordset.RecordCount > 0 Then
'''   data_sql.Recordset.MoveFirst
'''   Do While Not data_sql.Recordset.EOF
'''      data_sql.Recordset.Edit
'''      data_sql.Recordset("fecha") = CDate("20/12/2011")
'''      data_sql.Recordset.Update
'''      data_sql.Recordset.MoveNext
'''   Loop
'''End If
'Dim Xfacm As String
'Xfacm = InputBox("Ingrese el número de factura a modificar")
'If Xfacm <> "" Then
Dim Xd, Xh As Date
Xd = CDate("24/05/2014")
'Xh = CDate("11/10/2013")

   data_sql.DatabaseName = App.Path & "\sapp.mdb"
   data_sql.RecordSource = "Select * from linmmdd where fecha =#" & Format("22/05/2015", "yyyy/mm/dd") & "# and factura =" & 111180674 & " and dias =" & 0
'   data_sql.RecordSource = "Select * from deudas where cliente =" & 19890777 & " and documento =" & 19946358
   
   data_sql.Refresh
'data_mdb.DatabaseName = App.Path & "\sapp.mdb"
'data_mdb.RecordSource = "respcaja"
'data_mdb.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
'      data_sql.Recordset.MoveFirst
'      Do While Not data_sql.Recordset.EOF
         data_sql.Recordset.Edit
         data_sql.Recordset("vto") = CDate("22/05/2015")
         data_sql.Recordset("dias") = 3
'       data_sql.Recordset("imp_timbre") = 0
         data_sql.Recordset.Update
'         data_sql.Recordset.MoveNext
'      Loop
   End If
'End If
'       data_sql.Recordset("fec_llega") = CDate("01/08/2012")
'       data_sql.Recordset("realizada") = CDate("01/08/2012")
'       data_sql.Recordset.Update
'      data_mdb.Recordset.AddNew
'      data_mdb.Recordset("fecha") = data_sql.Recordset("fecha")
'      data_mdb.Recordset("numero") = data_sql.Recordset("numero")
'      data_mdb.Recordset("nombre") = data_sql.Recordset("nombre")
'      data_mdb.Recordset("movimiento") = data_sql.Recordset("movimiento")
'      data_mdb.Recordset("imp_fact") = data_sql.Recordset("imp_fact")
'      data_mdb.Recordset("nrorub") = data_sql.Recordset("nrorub")
'      data_mdb.Recordset("rubro") = data_sql.Recordset("rubro")
'      data_mdb.Recordset("documento") = data_sql.Recordset("documento")
'      data_mdb.Recordset("observ") = data_sql.Recordset("observ")
'      data_mdb.Recordset("saldo") = data_sql.Recordset("saldo")
'      data_mdb.Recordset("usuario") = data_sql.Recordset("usuario")
'      data_mdb.Recordset("hora") = data_sql.Recordset("hora")
'      data_mdb.Recordset("sys_2") = data_sql.Recordset("sys_2")
'      data_mdb.Recordset("saldo_user") = data_sql.Recordset("saldo_user")
'      data_mdb.Recordset("base") = data_sql.Recordset("base")
'      data_mdb.Recordset("cod_serv") = data_sql.Recordset("cod_serv")
'      data_mdb.Recordset("nom_serv") = data_sql.Recordset("nom_serv")
'      data_mdb.Recordset("cod_socio") = data_sql.Recordset("cod_socio")
'      data_mdb.Recordset("nom_socio") = data_sql.Recordset("nom_socio")
'      data_mdb.Recordset("turno") = data_sql.Recordset("turno")
'      data_mdb.Recordset("caja_mesp") = data_sql.Recordset("caja_mesp")
'      data_mdb.Recordset("caja_anop") = data_sql.Recordset("caja_anop")
'      data_mdb.Recordset("imp_iva") = data_sql.Recordset("imp_iva")
'      data_mdb.Recordset("opiva") = data_sql.Recordset("opiva")
'      data_mdb.Recordset.Update
'      data_sql.Recordset.MoveNext
'   Loop
'End If

Command4.Enabled = True
MsgBox "Proceso terminado"


End Sub

Private Sub Command40_Click()

data_mdb.DatabaseName = App.Path & "\visa.mdb"

'data_mdb.Connect = "odbc;dsn=sappnew;"
data_mdb.RecordSource = "pendvisa"
data_mdb.Refresh

data_sql.Connect = "odbc;dsn=sappnew;"
Data1.Connect = "odbc;dsn=sappnew;"
Dim Xdoc As Integer

data_sql.RecordSource = "select * from arqueo where cob in (514) and arqueo in ('C')"
data_sql.Refresh
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      Data1.RecordSource = "select * from deudas where cliente =" & data_sql.Recordset("matricula") & " and documento =" & data_sql.Recordset("nrorec") & " and fecha_pago is null"
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.Edit
         Data1.Recordset("fecha_pago") = Date
         Data1.Recordset.Update
      End If
      data_sql.Recordset.MoveNext
   Loop
End If

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_sql.RecordSource = "select * from deudas where cliente =" & data_mdb.Recordset("cliente") & " and documento =" & data_mdb.Recordset("documento") & " and fecha_pago is not null"
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.Edit
         data_sql.Recordset("fecha_pago") = Null
         data_sql.Recordset.Update
      End If
      data_sql.RecordSource = "select * from arqueo where matricula =" & data_mdb.Recordset("cliente") & " and nrorec =" & data_mdb.Recordset("documento")
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         data_sql.Recordset.Edit
         data_sql.Recordset("arqueo") = "P"
         data_sql.Recordset("fecha") = Date
         data_sql.Recordset("usuar") = "COMPUTOS"
         data_sql.Recordset.Update
      End If
      data_mdb.Recordset.MoveNext
   Loop
      
End If

MsgBox "Terminado"


End Sub

Private Sub Command41_Click()

data_sql.Connect = "ODBC;DSN=sapp;"
data_sql.RecordSource = "Select * from linmmdd where factura =" & 0
data_sql.Refresh

data_mdb.DatabaseName = App.Path & "\facdesp.mdb"
data_mdb.RecordSource = "lineas2"
data_mdb.Refresh

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
      data_sql.Recordset.AddNew
      data_sql.Recordset("linea") = data_mdb.Recordset("linea")
      data_sql.Recordset("factura") = data_mdb.Recordset("factura")
      data_sql.Recordset("tipo") = data_mdb.Recordset("tipo")
      data_sql.Recordset("realizada") = data_mdb.Recordset("realizada")
      data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
      data_sql.Recordset("cod_cli") = data_mdb.Recordset("cod_cli")
      data_sql.Recordset("nom_cli") = data_mdb.Recordset("nom_cli")
      data_sql.Recordset("convenio") = data_mdb.Recordset("convenio")
      data_sql.Recordset("cod_prod") = data_mdb.Recordset("cod_prod")
      data_sql.Recordset("nom_prod") = data_mdb.Recordset("nom_prod")
      data_sql.Recordset("operador") = data_mdb.Recordset("operador")
      data_sql.Recordset("hora") = data_mdb.Recordset("hora")
      data_sql.Recordset("imp_timbre") = data_mdb.Recordset("imp_timbre") ' sub total de la línea
      data_sql.Recordset("tot_lin") = data_mdb.Recordset("tot_lin") ' total de la linea de la factura
      data_sql.Recordset("valor_iva") = data_mdb.Recordset("pre_civa")
      data_sql.Recordset("base") = data_mdb.Recordset("base")
      data_sql.Recordset("nom_med_a") = data_mdb.Recordset("nom_med_a")
      data_sql.Recordset("rub_cont") = data_mdb.Recordset("rub_cont")
      data_sql.Recordset("nom_flia") = data_mdb.Recordset("nom_flia")
      data_sql.Recordset("pre_civa") = data_mdb.Recordset("pre_civa")
      data_sql.Recordset("reg_cab") = data_mdb.Recordset("reg_cab") '=99
      data_sql.Recordset("servicio") = data_mdb.Recordset("servicio")
      data_sql.Recordset("ced_socio") = data_mdb.Recordset("ced_socio")
      data_sql.Recordset("fact") = data_mdb.Recordset("fact") 'codced
      data_sql.Recordset("moneda") = data_mdb.Recordset("moneda")
      data_sql.Recordset("nro_flia") = data_mdb.Recordset("nro_flia")
      data_sql.Recordset("rub_cont") = data_mdb.Recordset("rub_cont")
      data_sql.Recordset("arancel") = data_mdb.Recordset("arancel")
      data_sql.Recordset("nro_med_a") = data_mdb.Recordset("nro_med_a")
      data_sql.Recordset("precio_est") = data_mdb.Recordset("precio_est")
      data_sql.Recordset("imp_iva") = data_mdb.Recordset("imp_iva")
      data_sql.Recordset("moneda") = data_mdb.Recordset("moneda")
      data_sql.Recordset("tipo_mov") = data_mdb.Recordset("tipo_mov")
      data_sql.Recordset("pendiente") = "T"
      data_sql.Recordset.Update
                                        
      data_sql.RecordSource = "caja"
      data_sql.Refresh
      
      data_sql.Recordset.AddNew
      data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
      data_sql.Recordset("numero") = data_mdb.Recordset("rub_cont")
      data_sql.Recordset("nombre") = "VTAS.CREDITO"
      data_sql.Recordset("moneda") = "$"
      data_sql.Recordset("movimiento") = "INGRESO"
      data_sql.Recordset("imp_fact") = data_mdb.Recordset("tot_lin")
      data_sql.Recordset("documento") = data_mdb.Recordset("factura")
      data_sql.Recordset("observ") = "CREDITO A " & Trim(Str(data_mdb.Recordset("factura")))
      data_sql.Recordset("saldo") = data_mdb.Recordset("tot_lin")
      data_sql.Recordset("usuario") = data_mdb.Recordset("operador")
      data_sql.Recordset("hora") = data_mdb.Recordset("hora")
      data_sql.Recordset("base") = data_mdb.Recordset("base")
      data_sql.Recordset("cod_serv") = data_mdb.Recordset("cod_prod")
      data_sql.Recordset("nom_serv") = Mid(data_mdb.Recordset("nom_prod"), 1, 50)
      data_sql.Recordset("cod_socio") = data_mdb.Recordset("cod_cli")
      data_sql.Recordset("nom_socio") = Mid(data_mdb.Recordset("nom_cli"), 1, 30)
      data_sql.Recordset("imp_iva") = Format(data_mdb.Recordset("imp_iva"), "Standard")
      data_sql.Recordset("opiva") = 1 ' 10% , 0 NO, 2 22%
      data_sql.Recordset.Update
    

      data_sql.RecordSource = "Select * from deudas where cliente =" & data_mdb.Recordset("cod_cli")
      data_sql.Refresh
      data_sql.Recordset.AddNew
      data_sql.Recordset("cod_cnv") = data_mdb.Recordset("convenio")
      data_sql.Recordset("nom_cnv") = "CCOU NO SAPP"
      data_sql.Recordset("cliente") = data_mdb.Recordset("cod_cli")
      data_sql.Recordset("nombre") = Mid(data_mdb.Recordset("nom_cli"), 1, 70)
      data_sql.Recordset("fecha") = data_mdb.Recordset("fecha")
      data_sql.Recordset("tipodoc") = "CRE"
      data_sql.Recordset("nro_superv") = 30
      data_sql.Recordset("documento") = data_mdb.Recordset("factura")
      data_sql.Recordset("tipocta") = "A"
      data_sql.Recordset("importe") = data_mdb.Recordset("tot_lin")
      data_sql.Recordset("moneda") = 1
      data_sql.Recordset("origen") = "E-TICKET NRO." & "A " & " " & data_mdb.Recordset("factura")
      data_sql.Recordset("saldo_cc") = data_mdb.Recordset("tot_lin")
      data_sql.Recordset("mes") = 0
      data_sql.Recordset("ano") = 0
      data_sql.Recordset("estado_cta") = 1
      data_sql.Recordset("tiquet") = 0
      data_sql.Recordset("deudas") = 0
      data_sql.Recordset("total") = data_mdb.Recordset("tot_lin")
      data_sql.Recordset("iva") = data_mdb.Recordset("imp_iva")
      data_sql.Recordset("servi") = 0
      data_sql.Recordset("nro_vende") = 1
      data_sql.Recordset.Update
   
    data_sql.RecordSource = "clirespl"
    data_sql.Refresh
    data_mdb.RecordSource = "cabezados"
    data_mdb.Refresh
    data_sql.Recordset.AddNew
    '           data_cabezal.Recordset("id") = 1
    data_sql.Recordset("cl_tipcli") = "1.0"
    data_sql.Recordset("cl_tipocli") = data_mdb.Recordset("cl_tipocli")
    data_sql.Recordset("cl_socmnro") = data_mdb.Recordset("cl_socmnro")
    data_sql.Recordset("cl_numero") = data_mdb.Recordset("cl_numero")
    data_sql.Recordset("cl_fnac") = data_mdb.Recordset("cl_fnac")
    data_sql.Recordset("fecha_reac") = data_mdb.Recordset("fecha_reac")
    data_sql.Recordset("cl_tj_venc") = data_mdb.Recordset("cl_tj_venc")
    data_sql.Recordset("cl_nrovend") = data_mdb.Recordset("cl_nrovend")
    data_sql.Recordset("cl_forpago") = data_mdb.Recordset("cl_forpago")
    data_sql.Recordset("cl_celular") = data_mdb.Recordset("cl_celular") 'descripcion f.pago
    data_sql.Recordset("fecha_modi") = data_mdb.Recordset("fecha_modi")
    data_sql.Recordset("cl_diacobr") = data_mdb.Recordset("cl_diacobr")
    data_sql.Recordset("cl_nrotarj") = data_mdb.Recordset("cl_nrotarj")
    data_sql.Recordset("cl_tjemi_n") = data_mdb.Recordset("cl_tjemi_n")
    data_sql.Recordset("cl_tjemi_c") = data_mdb.Recordset("cl_tjemi_c")
    data_sql.Recordset("cl_referen") = data_mdb.Recordset("cl_referen")
    data_sql.Recordset("tit_tarj") = data_mdb.Recordset("tit_tarj")
    data_sql.Recordset("cl_nomconv") = data_mdb.Recordset("cl_nomconv")
    'receptor
    data_sql.Recordset("cl_nro_sup") = data_mdb.Recordset("cl_nro_sup")
    data_sql.Recordset("hora_baja") = data_mdb.Recordset("hora_baja")
    data_sql.Recordset("cl_nom_sup") = data_mdb.Recordset("cl_nom_sup")
        'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
    data_sql.Recordset("info_debit") = data_mdb.Recordset("info_debit")
    data_sql.Recordset("cl_direcci") = data_mdb.Recordset("cl_direcci")
    data_sql.Recordset("cl_zona") = data_mdb.Recordset("cl_zona")
    'data_cabezal.Recordset("cl_email") = frm_convenios.t_dpto.Text 'opcional
    data_sql.Recordset("cl_localid") = data_mdb.Recordset("cl_localid") 'opcional
    data_sql.Recordset("cl_codigo") = data_mdb.Recordset("cl_codigo")
    data_sql.Recordset("usu_baja") = data_mdb.Recordset("usu_baja") 'moneda
    data_sql.Recordset("saldo_chc2") = data_mdb.Recordset("saldo_chc2") 'valor dolar
    data_sql.Recordset("saldo_cc") = data_mdb.Recordset("saldo_cc")  'iva minimo
    data_sql.Recordset("saldo_cc2") = data_mdb.Recordset("saldo_cc2") 'iva básico
    data_sql.Recordset("cl_atrasoa") = data_mdb.Recordset("cl_atrasoa") 'subtot iva 22
    data_sql.Recordset("cl_cedula") = data_mdb.Recordset("cl_cedula") 'subtot iva cero
    data_sql.Recordset("saldo_doc2") = data_mdb.Recordset("saldo_doc2")
    data_sql.Recordset("cl_atrasop") = data_mdb.Recordset("cl_atrasop")
    data_sql.Recordset("cl_decuota") = data_mdb.Recordset("cl_decuota")
    data_sql.Recordset("saldo_doc") = data_mdb.Recordset("saldo_doc")
    data_sql.Recordset("cl_grupo") = data_mdb.Recordset("cl_grupo")
    data_sql.Recordset("saldo_chc") = data_mdb.Recordset("saldo_chc")
    data_sql.Recordset("cl_telefon") = data_mdb.Recordset("cl_telefon")
    data_sql.Recordset("cl_nombre") = data_mdb.Recordset("cl_nombre")
    data_sql.Recordset("cl_cuopaga") = data_mdb.Recordset("cl_cuopaga")
    data_sql.Recordset("codmotbaja") = data_mdb.Recordset("codmotbaja")
    data_sql.Recordset("ultanopmut") = data_mdb.Recordset("ultanopmut")
    data_sql.Recordset("cl_fultvta") = data_mdb.Recordset("cl_fultvta")
    data_sql.Recordset("cl_entre") = data_mdb.Recordset("cl_entre")
    data_sql.Recordset("codmotbaja") = data_mdb.Recordset("codmotbaja")
    data_sql.Recordset("ultanopmut") = data_mdb.Recordset("ultanopmut")
    data_sql.Recordset("cl_fultvta") = data_mdb.Recordset("cl_fultvta")
    data_sql.Recordset("cl_entre") = data_mdb.Recordset("cl_entre")
    data_sql.Recordset("cl_fultpag") = data_mdb.Recordset("cl_fultpag")
    data_sql.Recordset("cl_ultmesp") = data_mdb.Recordset("cl_ultmesp")
    data_sql.Recordset("cl_nomvend") = data_mdb.Recordset("cl_nomvend")
    data_sql.Recordset("cl_fax") = data_mdb.Recordset("cl_fax")
    data_sql.Recordset.Update

End If

MsgBox "Terminado"


End Sub

Private Sub Command42_Click()
'On Error GoTo Quepasastock

data_mdb.DatabaseName = App.Path & "\sociospr2.mdb"
'data_mdb.Connect = "odbc;dsn=sappnew;"
data_mdb.RecordSource = "Select * from emispr order by cliente"
data_mdb.Refresh
data_sql.Connect = "odbc;dsn=sappnew;"

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_sql.RecordSource = "Select * from clientes where cl_codigo =" & data_mdb.Recordset("cliente")
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         If IsNull(data_sql.Recordset("fecha_baja")) = False Then
            data_mdb.Recordset.Edit
            data_mdb.Recordset("fecbaja") = data_sql.Recordset("fecha_baja")
            data_mdb.Recordset.Update
         End If
      End If
      data_mdb.Recordset.MoveNext
   Loop
End If
MsgBox "Terminado"

'data_sql.Connect = ""
'data_sql.DatabaseName = App.Path & "\sociospr.mdb"
'data_sql.RecordSource = "emi1pro"
'data_sql.Refresh
''If data_sql.Recordset.RecordCount > 0 Then
'   data_sql.Recordset.MoveFirst
'   Do While Not data_sql.Recordset.EOF
'      data_mdb.RecordSource = "select * from emi12pro where cliente =" & data_sql.Recordset("cliente")
'      data_mdb.Refresh
'      If data_mdb.Recordset.RecordCount > 0 Then
'      Else
'         data_sql.Recordset.Edit
''         data_sql.Recordset("siono") = 1 'no esta
'         data_sql.Recordset.Update
'      End If
'      data_sql.Recordset.MoveNext
'   Loop
'   MsgBox "Terminado"
'End If

End Sub

Private Sub Command43_Click()

End Sub

Private Sub Command5_Click()
Dim Xcant As Integer
Dim Xmat As Long
Dim Xcedstr As String
Xcedstr = ""
Xcant = 0
data_sql.Connect = "odbc;dsn=sappnew;"
data_sql.RecordSource = "select * from clientes where cl_cedula_t is null and cl_cedula not in (0)"
data_sql.Refresh
If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      If IsNull(data_sql.Recordset("cl_cedula")) = False Then
         If IsNull(data_sql.Recordset("cl_codced")) = False Then
            Xcedstr = Trim(Str(data_sql.Recordset("cl_cedula"))) & Trim(Str(data_sql.Recordset("cl_codced")))
         Else
            Xcedstr = ""
         End If
      Else
         Xcedstr = ""
      End If
      If Trim(Xcedstr) <> "" Then
'         If Trim(Xcedstr) = "37725160" Or Trim(Xcedstr) = "18821170" Or Trim(Xcedstr) = "17681520" Or Trim(Xcedstr) = "1658500" Or _
'            Trim(Xcedstr) = "53668196" Then
'         Else
            If IsNull(data_sql.Recordset("cl_cedula_t")) = False Then
               If Trim(data_sql.Recordset("cl_cedula_t")) <> Trim(Xcedstr) Then
                  data_sql.Recordset.Edit
                  data_sql.Recordset("cl_cedula_t") = Trim(Xcedstr)
                  data_sql.Recordset.Update
               End If
            Else
               data_sql.Recordset.Edit
               data_sql.Recordset("cl_cedula_t") = Trim(Xcedstr)
               data_sql.Recordset.Update
            End If
         'End If
      End If
      If IsNull(data_sql.Recordset("cl_dpto")) = False Then
         If UCase(data_sql.Recordset("cl_dpto")) = "NO APLICA" Then
         Else
            If IsNull(data_sql.Recordset("cl_celular_n")) = True Then
               data_sql.Recordset.Edit
               data_sql.Recordset("cl_celular_n") = Trim(data_sql.Recordset("cl_dpto"))
               data_sql.Recordset.Update
            End If
         End If
      End If
      data_sql.Recordset.MoveNext
   Loop
End If
data_sql.Recordset.Close

MsgBox "Proceso Terminado...." & Xcant


End Sub

Private Sub Command6_Click()
End

End Sub

Private Sub Command7_Click()

data_sql.DatabaseName = App.Path & "\estyconvant.mdb"
data_sql.RecordSource = "estudant"
data_sql.Refresh


data_mdb.Connect = "odbc;dsn=sappnew;"
'data_mdb.DatabaseName = App.Path & "\convs.mdb"
'data_mdb.RecordSource = "select * from convenio where cnv_codigo is not null and cnv_precio >" & 0
'data_mdb.Refresh

data_sql.Recordset.MoveFirst
Do While Not data_sql.Recordset.EOF
   data_mdb.RecordSource = "select * from estudios where codest =" & data_sql.Recordset("codest")
   data_mdb.Refresh
   If data_mdb.Recordset.RecordCount > 0 Then
      If data_mdb.Recordset("cons") <> data_sql.Recordset("cons") Then
         data_mdb.Recordset.Edit
         data_mdb.Recordset("cons") = data_sql.Recordset("cons")
         data_mdb.Recordset("uc") = data_sql.Recordset("uc")
         data_mdb.Recordset("ucfh") = data_sql.Recordset("ucfh")
         data_mdb.Recordset("part") = data_sql.Recordset("part")
         data_mdb.Recordset.Update
      End If
   End If
   
   data_mdb.RecordSource = "Select * from Aran_servicios where id_serv =" & data_sql.Recordset("codest") & " and prec_serv >" & 0
   data_mdb.Refresh
   If data_mdb.Recordset.RecordCount > 0 Then
      data_mdb.Recordset.MoveFirst
      Do While Not data_mdb.Recordset.EOF
         If data_mdb.Recordset("prec_serv") <> data_sql.Recordset("cons") Then
            data_mdb.Recordset.Edit
            data_mdb.Recordset("prec_serv") = data_sql.Recordset("cons")
            data_mdb.Recordset.Update
         End If
         data_mdb.Recordset.MoveNext
      Loop
   End If
   
   data_sql.Recordset.MoveNext
Loop


MsgBox "Terminado"

End Sub

Private Sub Command8_Click()
Dim Xfechasta As Date
data_mdb.DatabaseName = ""
data_mdb.Connect = "odbc;dsn=sappnew;"
data_mdb.RecordSource = "emi0220"
data_mdb.Refresh

data_sql.DatabaseName = ""
data_sql.Connect = "odbc;dsn=sappnew;"

Data1.DatabaseName = App.Path & "\informes.mdb"
Data1.RecordSource = "infcli"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If
data_mdb.Recordset.MoveFirst
Do While Not data_mdb.Recordset.EOF
   data_sql.RecordSource = "select * from emi0320 where cliente =" & data_mdb.Recordset("cliente")
   data_sql.Refresh
   If data_sql.Recordset.RecordCount > 0 Then
   Else
      data_sql.RecordSource = "select * from clientes where cl_codigo =" & data_mdb.Recordset("cliente")
      data_sql.Refresh
      If data_sql.Recordset.RecordCount > 0 Then
         Data1.Recordset.AddNew
         Data1.Recordset("cl_codigo") = data_sql.Recordset("cl_codigo")
         Data1.Recordset("cl_apellid") = data_sql.Recordset("cl_apellid")
         Data1.Recordset("cl_direcci") = data_sql.Recordset("cl_direcci")
         Data1.Recordset("cl_dpto") = data_sql.Recordset("cl_dpto")
         Data1.Recordset("cl_telefon") = data_sql.Recordset("cl_telefon")
         Data1.Recordset("cl_codconv") = data_sql.Recordset("cl_codconv")
         Data1.Recordset("cl_nomconv") = data_sql.Recordset("cl_nomconv")
         Data1.Recordset("cl_nrocobr") = data_sql.Recordset("cl_nrocobr")
         Data1.Recordset("cl_nro_sup") = data_mdb.Recordset("nro_cobr")
         Data1.Recordset("cl_zona") = data_sql.Recordset("cl_zona")
         Data1.Recordset("cl_nomvend") = data_mdb.Recordset("cod_cnv")
         Data1.Recordset("fecha_baja") = data_sql.Recordset("fecha_baja")
         Data1.Recordset.Update
      End If
   End If
   data_mdb.Recordset.MoveNext
Loop

Form1.MousePointer = 0
MsgBox "Terminado"

End Sub

Private Sub Command9_Click()
Dim Xmat As Long
Dim Xccant As Long
Dim XcantNro As Long
Dim Laced As String
Laced = ""
Xccant = 0
XcantNro = 59589945

data_sql.Connect = "odbc;dsn=sappnew;"

Data1.Connect = "odbc;dsn=sappper;"

data_mdb.Connect = "odbc;dsn=sappper;"

data_mdb.RecordSource = "select * from cabezal_hcdig where fecha >=#" & Format("08/07/2022", "yyyy/mm/dd") & "# and anterior is not null"
data_mdb.Refresh

'data_sql.RecordSource = "select * from cli_crmdeudas"
'data_sql.Refresh

If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      Data1.RecordSource = "select * from cli_crmdeudas where nrofact =" & Val(data_mdb.Recordset("hc_nro"))
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveFirst
         Do While Not Data1.Recordset.EOF
            data_sql.RecordSource = "select * from cli_crmdeudas where nrofact =" & data_mdb.Recordset("anterior") & " and usuario ='" & Data1.Recordset("usuario") & "'"
            data_sql.Refresh
            If data_sql.Recordset.RecordCount > 0 Then
               data_sql.Recordset.Edit
               data_sql.Recordset("fecha") = Data1.Recordset("fecha")
               data_sql.Recordset("base") = Data1.Recordset("base")
               data_sql.Recordset("hora") = Format(Time, "HH:mm")
               data_sql.Recordset.Update
            Else
               data_sql.Recordset.AddNew
               data_sql.Recordset("id") = XcantNro
               data_sql.Recordset("fecha") = Data1.Recordset("fecha")
               data_sql.Recordset("usuario") = Data1.Recordset("usuario")
               data_sql.Recordset("base") = Data1.Recordset("base")
               data_sql.Recordset("nrofact") = Data1.Recordset("nrofact")
               data_sql.Recordset("obs") = Data1.Recordset("obs")
               data_sql.Recordset("forma_pago") = Data1.Recordset("forma_pago")
               data_sql.Recordset("var1n") = Data1.Recordset("var1n")
               data_sql.Recordset("retirafam") = Data1.Recordset("retirafam")
               data_sql.Recordset.Update
              
               XcantNro = XcantNro + 1
            End If
            Data1.Recordset.MoveNext
         Loop
      End If
      data_mdb.Recordset.MoveNext
   Loop
End If
MsgBox "Terminado"


End Sub



Private Sub Form_Load()
If App.PrevInstance = True Then
   MsgBox "Esta"
   End
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
   laba.Caption = Anios
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
   labm.Caption = Meses
   labd.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   laba.Caption = 0
   labm.Caption = 0
   labd.Caption = 0
End If

End Sub

