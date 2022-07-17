VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtasconv 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Convenio"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frm_vtasconv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   1200
      Top             =   4320
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
      Caption         =   "data_lin"
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
   Begin MSComctlLib.ProgressBar barr 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crc 
      Left            =   4200
      Top             =   4200
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
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6360
      MouseIcon       =   "frm_vtasconv.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasconv.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00FFFFFF&
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
      MouseIcon       =   "frm_vtasconv.frx":0CD6
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasconv.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aceptar"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Datos para informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.Data data_busca 
         Caption         =   "data_busca"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Detalle asistencias NOSAPP"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Incluir solo categorías NOSAPP"
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
         Left            =   3480
         TabIndex        =   14
         Top             =   1680
         Width           =   3375
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   375
         Left            =   2400
         Top             =   0
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
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "data1"
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Informe desde historial"
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
         TabIndex        =   13
         Top             =   2640
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Informe sin detalle"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox tm 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
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
         RecordSource    =   "convenio"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txt_b 
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
         Height          =   405
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "frm_vtasconv.frx":156A
         Height          =   360
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   "cnv_desc"
         BoundColumn     =   "cnv_desc"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   480
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
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
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BASE: (99=TODAS)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Convenio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rango de fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   2400
      Picture         =   "frm_vtasconv.frx":1581
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2700
   End
End
Attribute VB_Name = "frm_vtasconv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub b_acep_Click()


b_acep.Enabled = False
b_canc.Enabled = False
Dim XCol, Xlin, Xnrocan As Integer

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim Xlalinea As String
Dim Xlamatr As Long
XCol = 1
Xlin = 1
Dim Xlafecbien As String

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
Dim Xfecini As Date
Dim Xcantass As Double
Dim XsiCod As Integer
Dim XRealiza, Xvencida As Date
Dim Xrealizatexto As String
Dim Xlafec1 As Date
Xfec1 = Date - 400

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
''      If tm.Text <> "" Then
         If txt_b.Text <> "" Then
            If txt_b.Text = 99 Then
               If DBCombo1.Text = "" Then
                  If Check3.Value = 1 Then
'                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000)"
                     data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod,linmmdd.nom_prod," & _
                     "linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.ced_socio,linmmdd.fact,linmmdd.nro_med_a," & _
                     "linmmdd.nom_med_a,linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.base,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.tot_lin," & _
                     "linmmdd.operador,linmmdd.nom_med_s,linmmdd.tipo,clientes.cl_codigo,clientes.cl_telefon,clientes.cl_cedula,clientes.cl_codced,clientes.cl_zona,clientes.cl_dpto," & _
                     "clientes.cl_decuota,clientes.cl_direcci,clientes.fecha_reac,clientes.cl_codconv from linmmdd inner join clientes on linmmdd.cod_cli=clientes.cl_codigo" & _
                     " where linmmdd.fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And linmmdd.fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "'" & _
                     " and linmmdd.convenio in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','SMINA','UNNSAM','EVNSAM','CCNSAM','CASNSA')" & _
                     " and linmmdd.nro_flia not in (9) and linmmdd.cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,990,802,803,804,805,806) order by linmmdd.cod_cli,linmmdd.fecha"
'10002,10004,10006,10008,10009,10010,10011,10012,10013,10014,10016,10017
'                  data_lla.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre," & _
'                  "llamado.unied,llamado.edad,llamado.matric,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codzon,llamado.obsmot," & _
'                  "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
'                  "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.hora,llamado.ci," & _
'                  "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.fecpas,resplla.mes from llamado " & _
'                  "inner join resplla on llamado.nrolla=resplla.nro where resplla.trasla is not null and " & _
'                  "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.movilpas <>" & 99 & " and llamado.enfer =" & 1 & " order by llamado.fecha"
                     
                     data_lin.Refresh
                  Else
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and tipo <>'" & "NOTA CR" & "' order by convenio"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' order by convenio"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio ='" & tm.Text & "' and tipo <>'" & "NOTA CR" & "'"
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio ='" & tm.Text & "'"
                     data_lin.Refresh
                  End If
               End If
            Else
               If DBCombo1.Text = "" Then
                  If Check3.Value = 1 Then
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' and convenio in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,802,803,804,805,806) order by convenio"
                     data_lin.Refresh
                  Else
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' order by convenio"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text & " order by convenio"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio ='" & tm.Text & "' And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "'"
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio ='" & tm.Text & "' And base =" & txt_b.Text
                     data_lin.Refresh
                  End If
               End If
            End If
            If data_lin.Recordset.RecordCount > 0 Then
               data_lin.Recordset.MoveLast
               barr.Visible = True
               barr.Max = data_lin.Recordset.RecordCount
               barr.Value = 0
               data_lin.Recordset.MoveFirst
               DoEvents
               frm_vtasconv.MousePointer = 11
               If Check3.Value = 1 Then
                  Dim Xobjexel As Excel.Application
                  Dim Xlibexel As Excel.Workbook
                  Dim Xarchexel As New Excel.Worksheet
                  Dim Xlabrir As New Excel.Application
                  Set Xobjexel = New Excel.Application
                  Set Xlibexel = Xobjexel.Workbooks.Add
                  Set Xarchexel = Xlibexel.Worksheets.Add
                  Xarchexel.Name = "VENTASNOSAPP"
                  Xlibexel.SaveAs ("C:\planillas\infvtas.xls")
                  Xarchtex = "C:\planillas\infvtas.xls"
                  Xarchexel.Cells(Xlin, XCol) = "DEPARTAMENTO de TI"
                  XCol = 10
                  Xarchexel.Cells(Xlin, XCol) = "FECHA:" & Format(Date, "dd/mm/yyyy")
                  Xlin = Xlin + 1
                  XCol = 2
                  Xarchexel.Range("A1", "C3").Font.Size = 16
                  Xarchexel.Cells(Xlin, XCol) = "INFORME DE VENTAS POR CONVENIOS NOSAPP  DESDE: " & md.Text & " HASTA: " & mh.Text
                  Xarchexel.Range("B" & Trim(str(Xlin)), "Q" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
                  XCol = 1
                  Xlin = Xlin + 2
                  Xnrocan = Xnrocan + Xlin
                  Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 14
                  Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
                  XCol = XCol + 1
                  Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 14
                  Xarchexel.Cells(Xlin, XCol) = "MUTUALISTA"
                  XCol = XCol + 1
                  Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 14
                  Xarchexel.Cells(Xlin, XCol) = "FECHA CONS.1"
                  XCol = XCol + 1
                  Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 14
                  Xarchexel.Cells(Xlin, XCol) = "FECHA CONS.FIN"
                  XCol = XCol + 1
                  Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
                  Xarchexel.Cells(Xlin, XCol) = "ASIST."
                  XCol = XCol + 1
                  Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 6
                  Xarchexel.Cells(Xlin, XCol) = "BASE"
                  XCol = XCol + 1
                  Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
                  Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
                  XCol = XCol + 1
                  Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 12
                  Xarchexel.Cells(Xlin, XCol) = "OPERADOR"
                  XCol = XCol + 1
                  Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 35
                  Xarchexel.Cells(Xlin, XCol) = "NOMBRES"
                  XCol = XCol + 1
                  Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 12
                  Xarchexel.Cells(Xlin, XCol) = "CEDULA"
                  XCol = XCol + 1
                  Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
                  Xarchexel.Cells(Xlin, XCol) = "CELULAR"
                  XCol = XCol + 1
                  Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 15
                  Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
                  XCol = XCol + 1
                  Xarchexel.Range("M" & Trim(str(Xlin))).ColumnWidth = 15
                  Xarchexel.Cells(Xlin, XCol) = "CARTA MUTUAL"
                  XCol = XCol + 1
                  Xarchexel.Range("N" & Trim(str(Xlin))).ColumnWidth = 15
                  Xarchexel.Cells(Xlin, XCol) = "FECHA"
                  XCol = XCol + 1
                  Xarchexel.Range("O" & Trim(str(Xlin))).ColumnWidth = 30
                  Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
                  XCol = XCol + 1
                  Xarchexel.Range("P" & Trim(str(Xlin))).ColumnWidth = 30
                  Xarchexel.Cells(Xlin, XCol) = "ZONA"
                  XCol = XCol + 1
                  Xarchexel.Range("Q" & Trim(str(Xlin))).ColumnWidth = 15
                  Xarchexel.Cells(Xlin, XCol) = "FACT.CARTA"
                  
                  XCol = 1
                  Xlin = Xlin + 1
                  Xlamatr = data_lin.Recordset("cod_cli")
                  Xnrocan = 0
               End If
               Do While Not data_lin.Recordset.EOF
                  If tm.Text = "CASH" Then
                     'If data_lin.Recordset("nro_flia") = 1 And data_lin.Recordset("cod_prod") <> 10006 And data_lin.Recordset("cod_prod") <> 10008 Then
                        data_inf.Recordset.AddNew
                        data1.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                        data1.Refresh
                        If data1.Recordset.RecordCount > 0 Then
                           data_inf.Recordset("zona") = data1.Recordset("cl_telefon")
                           If IsNull(data1.Recordset("cl_cedula")) = False Then
                              If IsNull(data1.Recordset("cl_codced")) = False Then
                                 data_inf.Recordset("nom_superv") = Trim(str(data1.Recordset("cl_cedula"))) & "-" & Trim(str(data1.Recordset("cl_codced")))
                              Else
                                 data_inf.Recordset("nom_superv") = Trim(str(data1.Recordset("cl_cedula"))) & "-0"
                              End If
                           Else
                              data_inf.Recordset("nom_superv") = "0"
                           End If
                           If IsNull(data1.Recordset("cl_telefon")) = False Then
                              data_inf.Recordset("zona") = data1.Recordset("cl_telefon")
                           End If
                           data_inf.Recordset("nom_medic") = Mid(data1.Recordset("cl_nomconv"), 1, 40)
                        End If
                        data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                        data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                        data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                        data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                        data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                        data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                        data_inf.Recordset("nom_flia") = data_lin.Recordset("nom_flia")
                        data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
'                        Data1.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
'                        Data1.Refresh
''                        If Data1.Recordset.RecordCount > 0 Then
'                           data_inf.Recordset("nom_medic") = Mid(Data1.Recordset("cnv_desc"), 1, 40)
'                        End If
                        data_inf.Recordset("tot_lin") = 0
                        data_inf.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
                        data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                        data_inf.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
                        data_inf.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
                        data_inf.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
                        data_inf.Recordset("base") = data_lin.Recordset("base")
                        data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva")
                        If IsNull(data_lin.Recordset("imp_iva")) = False Then
                           data_inf.Recordset("imp_iva") = Format(data_lin.Recordset("imp_iva"), "Standard")
                        Else
                           data_inf.Recordset("imp_iva") = 0
                        End If
                        If IsNull(data_lin.Recordset("pendiente")) = False Then
                          If data_lin.Recordset("pendiente") = "T" Then
                             data_inf.Recordset("tipo") = "e-Tck " & data_lin.Recordset("tipo")
                          Else
                             If data_lin.Recordset("pendiente") = "F" Then
                                data_inf.Recordset("tipo") = "e-Fct " & data_lin.Recordset("tipo")
                             Else
                                If data_lin.Recordset("pendiente") = "C" Then
                                   data_inf.Recordset("tipo") = "NC e-Tck " & data_lin.Recordset("tipo")
                                   data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                   data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva") * -1
                                   data_inf.Recordset("imp_iva") = Format(data_inf.Recordset("imp_iva"), "Standard")
                                Else
                                   If data_lin.Recordset("pendiente") = "N" Then
                                      data_inf.Recordset("tipo") = "NC e-Fct " & data_lin.Recordset("tipo")
                                      data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                      data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva") * -1
                                      data_inf.Recordset("imp_iva") = Format(data_inf.Recordset("imp_iva"), "Standard")
                                   Else
                                      If data_lin.Recordset("pendiente") = "A" Then
                                         data_inf.Recordset("tipo") = "ND e-Fct " & data_lin.Recordset("tipo")
                                      Else
                                         If data_lin.Recordset("pendiente") = "B" Then
                                            data_inf.Recordset("tipo") = "ND e-Tck " & data_lin.Recordset("tipo")
                                         Else
                                            If data_lin.Recordset("pendiente") = "R" Then
                                               data_inf.Recordset("tipo") = "Dev.Recibo"
                                               data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                            Else
                                               If data_lin.Recordset("pendiente") = "Z" Then
                                                  data_inf.Recordset("tipo") = "Recibo"
                                               Else
                                                  data_inf.Recordset("tipo") = "Registro"
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                        Else
                          data_inf.Recordset("tipo") = data_lin.Recordset("tipo")
                        End If
                        data_inf.Recordset.Update
                     'End If
                  Else
                      If Check3.Value = 1 Then
                         If data_lin.Recordset("cod_cli") = Xlamatr Then
                            If Xcantass = 1 Then
                               Xfecini = data_lin.Recordset("fecha")
                            End If
                            Xcantass = Xcantass + 1
                         Else
                            data_lin.Recordset.MovePrevious
'''aca el error
                            data1.RecordSource = "select * from linmmdd where fecha >='" & Format(Xfec1, "yyyy/mm/dd") & "' and cod_prod in (802,803,804,805,806) and cod_cli =" & data_lin.Recordset("cod_cli")
                            data1.Refresh
                            If data1.Recordset.RecordCount > 0 Then
                               XRealiza = data1.Recordset("fecha")
                               Xrealizatexto = Format(data1.Recordset("fecha"), "dd/mm/yyyy")
                               Xvencida = XRealiza + 150
                            End If
                            If data_lin.Recordset("convenio") = data_lin.Recordset("cl_codconv") Then
                               Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("convenio")
                               XCol = XCol + 1
                               If data_lin.Recordset("convenio") = "SMIN" Or data_lin.Recordset("convenio") = "SMINA" Then
                                  Xarchexel.Cells(Xlin, XCol) = "SMI"
                               Else
                                  If data_lin.Recordset("convenio") = "UNIVS" Or data_lin.Recordset("convenio") = "UNNSAM" Then
                                     Xarchexel.Cells(Xlin, XCol) = "UNIVERSAL"
                                  Else
                                     If data_lin.Recordset("convenio") = "CCNOS" Or data_lin.Recordset("convenio") = "CCNSAM" Then
                                        Xarchexel.Cells(Xlin, XCol) = "CCOU"
                                     Else
                                        If data_lin.Recordset("convenio") = "HEVANO" Or data_lin.Recordset("convenio") = "EVNSAM" Then
                                           Xarchexel.Cells(Xlin, XCol) = "H.EVANGELICO"
                                        Else
                                           If data_lin.Recordset("convenio") = "GANOS" Or data_lin.Recordset("convenio") = "CASANO" Or data_lin.Recordset("convenio") = "CASNSA" Then
                                              Xarchexel.Cells(Xlin, XCol) = "CASA DE GALICIA"
                                           Else
                                              Xarchexel.Cells(Xlin, XCol) = "SIN DATO"
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                               XCol = XCol + 1
                               Xarchexel.Cells(Xlin, XCol) = CDate(Xfecini)
                               XCol = XCol + 1
                               Xlafecbien = Format(data_lin.Recordset("fecha"), "dd/mm/yyyy")
                               Xarchexel.Cells(Xlin, XCol) = "'" & Xlafecbien
                               XCol = XCol + 1
                               Xarchexel.Cells(Xlin, XCol) = Xcantass
                               XCol = XCol + 1
                               Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("base")
                               XCol = XCol + 1
                               Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cod_cli")
                               XCol = XCol + 1
                               Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("operador")
                               XCol = XCol + 1
                               Xarchexel.Cells(Xlin, XCol) = Trim(data_lin.Recordset("nom_cli"))
                               XCol = XCol + 1
                               If IsNull(data_lin.Recordset("cl_cedula")) = False Then
                                  If IsNull(data_lin.Recordset("cl_codced")) = False Then
                                     Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lin.Recordset("cl_cedula"))) & "-" & Trim(str(data_lin.Recordset("cl_codced")))
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lin.Recordset("cl_cedula"))) & "-0"
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "0-0"
                               End If
                               XCol = XCol + 1
                               If IsNull(data_lin.Recordset("cl_dpto")) = False Then
                                  If Trim(data_lin.Recordset("cl_dpto")) <> "" Then
                                     Xarchexel.Cells(Xlin, XCol) = "'" & Trim(data_lin.Recordset("cl_dpto"))
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "Sin Cel"
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "Sin Cel"
                               End If
                               XCol = XCol + 1
                               If IsNull(data_lin.Recordset("cl_telefon")) = False Then
                                  If Trim(data_lin.Recordset("cl_dpto")) <> "" Then
                                     Xarchexel.Cells(Xlin, XCol) = Trim(data_lin.Recordset("cl_telefon"))
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "Sin Tel"
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "Sin Tel"
                               End If
                               XCol = XCol + 1
                               If XsiCod = 1 Then
                                  If Format(Xvencida, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                                     Xarchexel.Cells(Xlin, XCol) = "Firmar Nuevamente"
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "Firmo Carta"
                                  End If
                               Else
                                  If IsNull(data_lin.Recordset("cl_decuota")) = False Then
                                     If data_lin.Recordset("cl_decuota") = 1 Then
                                        Xarchexel.Cells(Xlin, XCol) = "Aviso de Carta"
                                     Else
                                        If data_lin.Recordset("cl_decuota") = 2 Then
                                           Xarchexel.Cells(Xlin, XCol) = "Se recibe Carta"
                                        Else
                                           If data_lin.Recordset("cl_decuota") = 3 Then
                                              Xarchexel.Cells(Xlin, XCol) = "Se Niega"
                                           Else
                                              Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                               XCol = XCol + 1
                               If XsiCod = 1 Then
                                  Xarchexel.Cells(Xlin, XCol) = "'" & Format(XRealiza, "dd/mm/yyyy")
                               Else
                                  If IsNull(data_lin.Recordset("fecha_reac")) = False Then
                                     Xarchexel.Cells(Xlin, XCol) = "'" & Format(data_lin.Recordset("fecha_reac"), "dd/mm/yyyy")
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "No hay fecha"
                                  End If
                               End If
                               XCol = XCol + 1
                               If IsNull(data_lin.Recordset("cl_direcci")) = False Then
                                  If Trim(data_lin.Recordset("cl_direcci")) <> "" Then
                                     Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cl_direcci")
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "Sin dato"
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "Sin dato"
                               End If
                               XCol = XCol + 1
                               If IsNull(data_lin.Recordset("cl_zona")) = False Then
                                  Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cl_zona")
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "Sin dato"
                               End If
                               XCol = XCol + 1
                               If Trim(Xrealizatexto) <> "" Then
                                  Xarchexel.Cells(Xlin, XCol) = "'" & Format(Xrealizatexto, "dd/mm/yyyy")
                               End If
                               XCol = 1
                               Xlin = Xlin + 1
                               Xnrocan = Xnrocan + 1
                               XsiCod = 0
                            End If
                            data_lin.Recordset.MoveNext
                            Xlamatr = data_lin.Recordset("cod_cli")
                            Xcantass = 1
                            Xrealizatexto = ""
                            Xfecini = data_lin.Recordset("fecha")
                         End If
                      Else
                         data_inf.Recordset.AddNew
                         data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                         data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                         data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                         data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                         data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                         data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                         If Check3.Value = 1 Then
                         Else
                            data_inf.Recordset("nom_flia") = data_lin.Recordset("nom_flia")
                         End If
                         data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                         data1.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                         data1.Refresh
                         If data1.Recordset.RecordCount > 0 Then
                            data_inf.Recordset("nom_medic") = Mid(data1.Recordset("cnv_desc"), 1, 40)
                         End If
                         data_inf.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                         data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                         data_inf.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
                         data_inf.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
                         data_inf.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
                         data_inf.Recordset("base") = data_lin.Recordset("base")
                         data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva")
                         data_inf.Recordset("operador") = data_lin.Recordset("operador")
                         
                         If IsNull(data_lin.Recordset("imp_iva")) = False Then
                            data_inf.Recordset("imp_iva") = Format(data_lin.Recordset("imp_iva"), "Standard")
                         Else
                            data_inf.Recordset("imp_iva") = 0
                         End If
                         If IsNull(data_lin.Recordset("pendiente")) = False Then
                            If data_lin.Recordset("pendiente") = "T" Then
                               data_inf.Recordset("tipo") = "e-Tck " & data_lin.Recordset("tipo")
                            Else
                               If data_lin.Recordset("pendiente") = "F" Then
                                  data_inf.Recordset("tipo") = "e-Fct " & data_lin.Recordset("tipo")
                               Else
                                  If data_lin.Recordset("pendiente") = "C" Then
                                     data_inf.Recordset("tipo") = "NC e-Tck " & data_lin.Recordset("tipo")
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                     data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva") * -1
                                     If IsNull(data_inf.Recordset("imp_iva")) = False Then
                                        data_inf.Recordset("imp_iva") = Format(data_inf.Recordset("imp_iva"), "Standard")
                                     Else
                                        data_inf.Recordset("imp_iva") = 0
                                     End If
                                  Else
                                     If data_lin.Recordset("pendiente") = "N" Then
                                        data_inf.Recordset("tipo") = "NC e-Fct " & data_lin.Recordset("tipo")
                                        data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                        data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva") * -1
                                        If IsNull(data_inf.Recordset("imp_iva")) = False Then
                                           data_inf.Recordset("imp_iva") = Format(data_inf.Recordset("imp_iva"), "Standard")
                                        Else
                                           data_inf.Recordset("imp_iva") = 0
                                        End If
                                     Else
                                        If data_lin.Recordset("pendiente") = "A" Then
                                           data_inf.Recordset("tipo") = "ND e-Fct " & data_lin.Recordset("tipo")
                                        Else
                                           If data_lin.Recordset("pendiente") = "B" Then
                                              data_inf.Recordset("tipo") = "ND e-Tck " & data_lin.Recordset("tipo")
                                           Else
                                              If data_lin.Recordset("pendiente") = "R" Then
                                                 data_inf.Recordset("tipo") = "Dev.Recibo"
                                                 data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                              Else
                                                 If data_lin.Recordset("pendiente") = "Z" Then
                                                    data_inf.Recordset("tipo") = "Recibo"
                                                 Else
                                                    data_inf.Recordset("tipo") = "Registro"
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         Else
                            data_inf.Recordset("tipo") = data_lin.Recordset("tipo")
                         End If
                         data_inf.Recordset.Update
                      End If
                  End If
                  data_lin.Recordset.MoveNext
                  barr.Value = barr.Value + 1
               Loop
               frm_vtasconv.MousePointer = 0
               data_inf.RecordSource = "Select * from infvtas order by convenio"
               data_inf.Refresh
               If Check3.Value = 1 Then
                    data_lin.Recordset.MovePrevious
                    data1.RecordSource = "select * from linmmdd where fecha >='" & Format(Xfec1, "yyyy/mm/dd") & "' and cod_prod in (802,803,804,805,806) and cod_cli =" & data_lin.Recordset("cod_cli")
                    data1.Refresh
                    If data1.Recordset.RecordCount > 0 Then
                       XRealiza = data1.Recordset("fecha")
                       Xrealizatexto = Format(data1.Recordset("fecha"), "dd/mm/yyyy")
                       Xvencida = XRealiza + 150
                    End If
                    
                    If data_lin.Recordset("convenio") = data_lin.Recordset("cl_codconv") Then
                       Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("convenio")
                       XCol = XCol + 1
                       If data_lin.Recordset("convenio") = "SMIN" Or data_lin.Recordset("convenio") = "SMINA" Then
                          Xarchexel.Cells(Xlin, XCol) = "SMI"
                       Else
                          If data_lin.Recordset("convenio") = "UNIVS" Or data_lin.Recordset("convenio") = "UNNSAM" Then
                             Xarchexel.Cells(Xlin, XCol) = "UNIVERSAL"
                          Else
                             If data_lin.Recordset("convenio") = "CCNOS" Or data_lin.Recordset("convenio") = "CCNSAM" Then
                                Xarchexel.Cells(Xlin, XCol) = "CCOU"
                             Else
                                If data_lin.Recordset("convenio") = "HEVANO" Or data_lin.Recordset("convenio") = "EVNSAM" Then
                                   Xarchexel.Cells(Xlin, XCol) = "H.EVANGELICO"
                                Else
                                   If data_lin.Recordset("convenio") = "GANOS" Or data_lin.Recordset("convenio") = "CASANO" Or data_lin.Recordset("convenio") = "CASNSA" Then
                                      Xarchexel.Cells(Xlin, XCol) = "CASA DE GALICIA"
                                   Else
                                      Xarchexel.Cells(Xlin, XCol) = "SIN DATO"
                                   End If
                                End If
                             End If
                          End If
                       End If
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = "'" & Format(Xfecini, "dd/mm/yyyy")
                       XCol = XCol + 1
                       Xlafecbien = Format(data_lin.Recordset("fecha"), "dd/mm/yyyy")
                       Xarchexel.Cells(Xlin, XCol) = "'" & Xlafecbien
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Xcantass
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("base")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cod_cli")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("operador")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = Trim(data_lin.Recordset("nom_cli"))
                       XCol = XCol + 1
                       If IsNull(data_lin.Recordset("cl_cedula")) = False Then
                          If IsNull(data_lin.Recordset("cl_codced")) = False Then
                             Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lin.Recordset("cl_cedula"))) & "-" & Trim(str(data_lin.Recordset("cl_codced")))
                          Else
                             Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lin.Recordset("cl_cedula"))) & "-0"
                          End If
                       Else
                          Xarchexel.Cells(Xlin, XCol) = "0-0"
                       End If
                       XCol = XCol + 1
                       If IsNull(data_lin.Recordset("cl_dpto")) = False Then
                          If Trim(data_lin.Recordset("cl_dpto")) <> "" Then
                             Xarchexel.Cells(Xlin, XCol) = "'" & Trim(data_lin.Recordset("cl_dpto"))
                          Else
                             Xarchexel.Cells(Xlin, XCol) = "Sin Cel"
                          End If
                       Else
                          Xarchexel.Cells(Xlin, XCol) = "Sin Cel"
                       End If
                       XCol = XCol + 1
                       If IsNull(data_lin.Recordset("cl_telefon")) = False Then
                          If Trim(data_lin.Recordset("cl_dpto")) <> "" Then
                             Xarchexel.Cells(Xlin, XCol) = Trim(data_lin.Recordset("cl_telefon"))
                          Else
                             Xarchexel.Cells(Xlin, XCol) = "Sin Tel"
                          End If
                       Else
                          Xarchexel.Cells(Xlin, XCol) = "Sin Tel"
                       End If
                       XCol = XCol + 1
                       If XsiCod = 1 Then
                          If Format(Xvencida, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                             Xarchexel.Cells(Xlin, XCol) = "Carta VENCIDA"
                          Else
                             Xarchexel.Cells(Xlin, XCol) = "Firmo Carta"
                          End If
                       Else
                          If IsNull(data_lin.Recordset("cl_decuota")) = False Then
                             If data_lin.Recordset("cl_decuota") = 1 Then
                                Xarchexel.Cells(Xlin, XCol) = "Aviso de Carta"
                             Else
                                If data_lin.Recordset("cl_decuota") = 2 Then
                                   Xarchexel.Cells(Xlin, XCol) = "Se recibe Carta"
                                Else
                                   If data_lin.Recordset("cl_decuota") = 3 Then
                                      Xarchexel.Cells(Xlin, XCol) = "Se Niega"
                                   Else
                                      Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
                                   End If
                                End If
                             End If
                          End If
                       End If
                       XCol = XCol + 1
                       If XsiCod = 1 Then
                          Xarchexel.Cells(Xlin, XCol) = "'" & Format(XRealiza, "dd/mm/yyyy")
                       Else
                          If IsNull(data_lin.Recordset("fecha_reac")) = False Then
                             Xarchexel.Cells(Xlin, XCol) = "'" & Format(data_lin.Recordset("fecha_reac"), "dd/mm/yyyy")
                          Else
                             Xarchexel.Cells(Xlin, XCol) = "No hay fecha"
                          End If
                       End If
                       XCol = XCol + 1
                       If IsNull(data_lin.Recordset("cl_direcci")) = False Then
                          If Trim(data_lin.Recordset("cl_direcci")) <> "" Then
                             Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cl_direcci")
                          Else
                             Xarchexel.Cells(Xlin, XCol) = "Sin dato"
                          End If
                       Else
                          Xarchexel.Cells(Xlin, XCol) = "Sin dato"
                       End If
                       XCol = XCol + 1
                       If IsNull(data_lin.Recordset("cl_zona")) = False Then
                          Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cl_zona")
                       Else
                          Xarchexel.Cells(Xlin, XCol) = "Sin dato"
                       End If
                       XCol = XCol + 1
                       If Trim(Xrealizatexto) <> "" Then
                          Xarchexel.Cells(Xlin, XCol) = "'" & Format(Xrealizatexto, "dd/mm/yyyy")
                       End If
                       XCol = XCol + 1
                       
''''verificar desde acá
                       If data_lin.Recordset("cod_prod") = 10002 Or data_lin.Recordset("cod_prod") = 10004 Or _
                          data_lin.Recordset("cod_prod") = 10006 Or data_lin.Recordset("cod_prod") = 10018 Then
                          If IsNull(data_lin.Recordset("ced_socio")) = False Then
                             
                          End If
                       End If
                       
                       XCol = 1
                       Xlin = Xlin + 1
                       Xnrocan = Xnrocan + 1
                    End If
'                  Xarchexel.Range("B" & Trim(str(4)), "M" & Trim(str(Xlin))).Select
'                  Range(Cells(a, 1), Cells(2, 1)).Sort Key1:=Cells(a, 1)
'                  Xarchexel.Range.Sort
                  XCol = 3
                  Xarchexel.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Trim(str(Xnrocan))
                  
                  Xlibexel.Save
                  Xlibexel.Close
                  Xobjexel.Quit
                  Xlabrir.Workbooks.Open Xarchtex, , False
                  Xlabrir.Visible = True
                  Xlabrir.WindowState = xlMaximized
'                  ShellExecute Me.hwnd, "open", "c:\planillas\Infvtas.xls", "", "", 4
               End If
               barr.Value = 0
               barr.Visible = False
               
'                        Print #1, "JORGE" & vbTab & "FERNANDEZ" & vbTab & "34805844"
'                        Close #1
               If Check3.Value = 1 Then
               
               Else
                    If Check1.Value = 1 Then
                       crc.ReportFileName = App.path & "\infvtasxconn.rpt"
                       crc.ReportTitle = "INFORME DE VENTAS POR CONVENIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                       crc.Action = 1
                    Else
                       If tm.Text = "CASH" Then
                          crc.ReportFileName = App.path & "\infvtasxconc.rpt"
                       Else
                          If tm.Text <> "" Then
                             If Check4.Value = 1 Then
                                data_inf.RecordSource = "select * from infvtas where cod_prod in (999,992,800,881,882,991,993,994,996,997,8000,995,990,802,803,804,805,806)"
                                data_inf.Refresh
                                If data_inf.Recordset.RecordCount > 0 Then
                                   data_inf.Recordset.MoveFirst
                                   Do While Not data_inf.Recordset.EOF
                                      data_inf.Recordset.Delete
                                      data_inf.Recordset.MoveNext
                                   Loop
                                   data_inf.RecordSource = "select * from infvtas"
                                   data_inf.Refresh
                                End If
                                crc.ReportFileName = App.path & "\infvtasxcondet.rpt"
                             Else
                                crc.ReportFileName = App.path & "\infvtasxcon.rpt"
                             End If
                          Else
                             crc.ReportFileName = App.path & "\infvtasxcon.rpt"
                          End If
                       End If
                       crc.ReportTitle = "INFORME DE VENTAS POR CONVENIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                       crc.Action = 1
                    End If
                End If
            Else
               MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
            End If
         Else
            MsgBox "Ingrese Base", vbInformation, "Mensaje"
            txt_b.SetFocus
         End If
      'Else
      '   MsgBox "Código de convenio incorrecto", vbInformation, "Mensaje"
      '   DBCombo1.SetFocus
      'End If
   Else
      MsgBox "Ingrese Fecha", vbInformation, "Mensaje"
      mh.SetFocus
   End If
Else
   MsgBox "Ingrese fecha", vbInformation, "Mensaje"
   md.SetFocus
End If
b_acep.Enabled = True
b_canc.Enabled = True

End Sub

Private Sub b_canc_Click()
Unload Me

End Sub


Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_b.SetFocus
End If

End Sub

Private Sub DBCombo1_LostFocus()

If DBCombo1.Text <> "" Then
   data_med.Recordset.FindFirst "cnv_codigo ='" & DBCombo1.Text & "'"
   If Not data_med.Recordset.NoMatch Then
      tm.Text = data_med.Recordset("cnv_codigo")
      DBCombo1.Text = data_med.Recordset("cnv_desc")
   Else
      data_med.Recordset.FindFirst "cnv_desc ='" & DBCombo1.Text & "'"
      If Not data_med.Recordset.NoMatch Then
         tm.Text = data_med.Recordset("cnv_codigo")
         DBCombo1.Text = data_med.Recordset("cnv_desc")
      Else
         MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
         DBCombo1.SetFocus
      End If
   End If
Else
   tm.Text = ""
   MsgBox "Se emitirán TODOS los convenios"
End If


End Sub

Private Sub Form_Load()
data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med.RecordSource = "select * from convenio order by cnv_desc"
data_med.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
data_lin.ConnectionString = "dsn=" & Xconexrmt
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_busca.Connect = "odbc;dsn=" & Xconexrmt & ";"

data1.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBCombo1.SetFocus
End If

End Sub

Private Sub txt_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub
