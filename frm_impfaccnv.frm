VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_impfaccnv 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "IMPRESION"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lin3 
      Height          =   375
      Left            =   5160
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "data_lin3"
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
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   840
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc data_lincab 
      Height          =   375
      Left            =   4800
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "data_lincab"
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
   Begin VB.TextBox t_mat 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox t_serie 
      Alignment       =   2  'Center
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
      Left            =   3240
      TabIndex        =   7
      Text            =   "A"
      Top             =   720
      Width           =   495
   End
   Begin VB.Data data_imagen 
      Caption         =   "data_imagen"
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_fac 
      Caption         =   "data_fac"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_faccab 
      Caption         =   "data_faccab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "VER Ultimos números"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox t_has 
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
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Número de factura final"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox t_des 
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
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Número de factura de comienzo"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      Picture         =   "frm_impfaccnv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   1680
      Width           =   615
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2640
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
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
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SERIE"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "IMPRESION POR LOTE"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   3840
      Picture         =   "frm_impfaccnv.frx":058A
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frm_impfaccnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
data_faccab.RecordSource = "cabezados"
data_faccab.Refresh
If data_faccab.Recordset.RecordCount > 0 Then
   data_faccab.Recordset.MoveFirst
   Do While Not data_faccab.Recordset.EOF
      data_faccab.Recordset.Delete
      data_faccab.Recordset.MoveNext
   Loop
End If


data_fac.RecordSource = "lineas2"
data_fac.Refresh
If data_fac.Recordset.RecordCount > 0 Then
   data_fac.Recordset.MoveFirst
   Do While Not data_fac.Recordset.EOF
      data_fac.Recordset.Delete
      data_fac.Recordset.MoveNext
   Loop
End If

If t_mat.Text = "" Then
   MsgBox "No ingresó número de matrícula"
   data_lincab.RecordSource = "Select * from clirespl where cl_numero =" & t_des.Text & " and cl_socmnro ='" & t_serie.Text & "'"
   data_lincab.Refresh
Else
   data_lincab.RecordSource = "Select * from clirespl where cl_numero =" & t_des.Text & " and cl_codigo =" & t_mat.Text
   data_lincab.Refresh
End If
If data_lincab.Recordset.RecordCount > 0 Then
   data_faccab.Recordset.AddNew
   data_faccab.Recordset("cl_socmnro") = data_lincab.Recordset("cl_socmnro")
   data_faccab.Recordset("cl_numero") = data_lincab.Recordset("cl_numero")
   data_imagen.RecordSource = "Select * from qr where nrofact =" & t_des.Text & " and serie ='" & t_serie.Text & "'"
   data_imagen.Refresh
   If data_imagen.Recordset.RecordCount > 0 Then
      data_faccab.Recordset("qr") = data_imagen.Recordset("qr")
   End If
   data_faccab.Recordset("cl_tipocli") = data_lincab.Recordset("cl_tipocli")
   data_faccab.Recordset("cl_socmnro") = data_lincab.Recordset("cl_socmnro")
   data_faccab.Recordset("cl_numero") = data_lincab.Recordset("cl_numero")
   data_faccab.Recordset("cl_fnac") = data_lincab.Recordset("cl_fnac")
   data_faccab.Recordset("fecha_reac") = data_lincab.Recordset("fecha_reac")
   data_faccab.Recordset("cl_tj_venc") = data_lincab.Recordset("cl_tj_venc")
   data_faccab.Recordset("cl_nrovend") = data_lincab.Recordset("cl_nrovend")
   data_faccab.Recordset("cl_forpago") = data_lincab.Recordset("cl_forpago")
   data_faccab.Recordset("cl_celular") = data_lincab.Recordset("cl_celular")
   data_faccab.Recordset("fecha_modi") = data_lincab.Recordset("fecha_modi")
   data_faccab.Recordset("cl_diacobr") = data_lincab.Recordset("cl_diacobr")
   data_faccab.Recordset("cl_nrotarj") = data_lincab.Recordset("cl_nrotarj")
   data_faccab.Recordset("cl_tjemi_n") = data_lincab.Recordset("cl_tjemi_n")
   data_faccab.Recordset("cl_tjemi_c") = data_lincab.Recordset("cl_tjemi_c")
   data_faccab.Recordset("cl_referen") = data_lincab.Recordset("cl_referen")
   data_faccab.Recordset("tit_tarj") = data_lincab.Recordset("tit_tarj")
   data_faccab.Recordset("cl_nomconv") = data_lincab.Recordset("cl_nomconv")
   data_faccab.Recordset("cl_nro_sup") = data_lincab.Recordset("cl_nro_sup")
   data_faccab.Recordset("hora_baja") = data_lincab.Recordset("hora_baja")
   data_faccab.Recordset("cl_nom_sup") = data_lincab.Recordset("cl_nom_sup")
   data_faccab.Recordset("info_debit") = data_lincab.Recordset("info_debit")
   data_faccab.Recordset("cl_direcci") = data_lincab.Recordset("cl_direcci")
   data_faccab.Recordset("cl_zona") = data_lincab.Recordset("cl_zona")
   data_faccab.Recordset("cl_localid") = data_lincab.Recordset("cl_localid")
   data_faccab.Recordset("cl_codigo") = data_lincab.Recordset("cl_codigo")
   data_faccab.Recordset("usu_baja") = data_lincab.Recordset("usu_baja")
   data_faccab.Recordset("saldo_chc2") = data_lincab.Recordset("saldo_chc2")
   data_faccab.Recordset("saldo_cc") = data_lincab.Recordset("saldo_cc")
   data_faccab.Recordset("saldo_cc2") = data_lincab.Recordset("saldo_cc2")
   data_faccab.Recordset("cl_atrasoa") = data_lincab.Recordset("cl_atrasoa")
   data_faccab.Recordset("cl_cedula") = data_lincab.Recordset("cl_cedula")
   data_faccab.Recordset("saldo_doc2") = data_lincab.Recordset("saldo_doc2")
   data_faccab.Recordset("cl_atrasop") = data_lincab.Recordset("cl_atrasop")
   data_faccab.Recordset("cl_decuota") = data_lincab.Recordset("cl_decuota")
   data_faccab.Recordset("saldo_doc") = data_lincab.Recordset("saldo_doc")
   data_faccab.Recordset("cl_grupo") = data_lincab.Recordset("cl_grupo")
   data_faccab.Recordset("saldo_chc") = data_lincab.Recordset("saldo_chc")
   data_faccab.Recordset("cl_telefon") = data_lincab.Recordset("cl_telefon")
   data_faccab.Recordset("cl_fultpag") = data_lincab.Recordset("cl_fultpag")
   data_faccab.Recordset("cl_ultmesp") = data_lincab.Recordset("cl_ultmesp")
   data_faccab.Recordset("cl_nomvend") = data_lincab.Recordset("cl_nomvend")
   data_faccab.Recordset("cl_fax") = data_lincab.Recordset("cl_fax")
   data_faccab.Recordset("cl_nombre") = data_lincab.Recordset("cl_nombre")
   data_lin3.RecordSource = "Select * from indica_enfc where idhc =" & t_des.Text & " and in_dosis =" & 1
   data_lin3.Refresh
   If data_lin3.Recordset.RecordCount > 0 Then
      If IsNull(data_lin3.Recordset("in_obs")) = False Then
         data_faccab.Recordset("obsp") = data_lin3.Recordset("in_obs")
      End If
   End If
   data_faccab.Recordset.Update
'fin de cabezal
End If

If t_mat.Text = "" Then
   data_lin.RecordSource = "Select * from linmmdd where factura =" & t_des.Text
Else
   data_lin.RecordSource = "Select * from linmmdd where factura =" & t_des.Text & " and cod_cli =" & t_mat.Text
End If
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
       data_fac.Recordset.AddNew
       data_fac.Recordset("fecha") = data_lin.Recordset("fecha")
       data_fac.Recordset("reg_cab") = data_lin.Recordset("reg_cab")
       data_fac.Recordset("factura") = data_lin.Recordset("factura")
       data_fac.Recordset("moneda") = data_lin.Recordset("moneda")
       data_fac.Recordset("servicio") = data_lin.Recordset("servicio")
       data_fac.Recordset("tipo") = data_lin.Recordset("tipo")
       data_fac.Recordset("realizada") = data_lin.Recordset("realizada")
       data_fac.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
       data_fac.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
       data_fac.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
       data_fac.Recordset("tcambio") = data_lin.Recordset("tcambio")
       data_fac.Recordset("fact") = data_lin.Recordset("fact")
       data_fac.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
       data_fac.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
       data_fac.Recordset("cantidad") = data_lin.Recordset("cantidad")
       data_fac.Recordset("operador") = data_lin.Recordset("operador")
       data_fac.Recordset("hora") = data_lin.Recordset("hora")
       data_fac.Recordset("ruc") = data_lin.Recordset("ruc")
       data_fac.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
       data_fac.Recordset("nom_flia") = data_lin.Recordset("nom_flia")
       data_fac.Recordset("nro_superv") = data_lin.Recordset("nro_superv")
       data_fac.Recordset("nom_superv") = data_lin.Recordset("nom_superv")
       data_fac.Recordset("convenio") = data_lin.Recordset("convenio")
       data_fac.Recordset("unidad") = data_lin.Recordset("unidad")
       data_fac.Recordset("grupo") = data_lin.Recordset("grupo")
       data_fac.Recordset("rub_cont") = data_lin.Recordset("rub_cont")
       data_fac.Recordset("arancel") = data_lin.Recordset("arancel")
       data_fac.Recordset("usa_timbre") = data_lin.Recordset("usa_timbre")
       data_fac.Recordset("imp_timbre") = data_lin.Recordset("imp_timbre")
       data_fac.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
       data_fac.Recordset("rub_nomb") = data_lin.Recordset("rub_nomb")
       data_fac.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
       data_fac.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
'       data_fac.Recordset("nro_med_s") = data_lin.Recordset("nro_med_s")
       data_fac.Recordset("nom_med_s") = data_lin.Recordset("nom_med_s")
       data_fac.Recordset("precio_est") = data_lin.Recordset("precio_est")
       data_fac.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
       data_fac.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
       data_fac.Recordset("base") = data_lin.Recordset("base")
       data_fac.Recordset("imp_iva") = data_lin.Recordset("imp_iva")
       data_fac.Recordset("linea") = data_lin.Recordset("linea")
       data_fac.Recordset("dias") = data_lin.Recordset("dias")
       data_fac.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
       data_fac.Recordset("pre_civa") = data_lin.Recordset("pre_civa")
       data_fac.Recordset("porce_est") = data_lin.Recordset("porce_est")
       data_fac.Recordset("rub_nomb") = data_lin.Recordset("rub_nomb")
       data_fac.Recordset("solicitant") = data_lin.Recordset("solicitant")
       data_lin3.RecordSource = "Select * from indica_enfc where idhc =" & t_des.Text & " and in_hora ='" & t_serie.Text & "' and in_dosis =" & 3 & " and in_uni =" & data_lin.Recordset("linea")
       data_lin3.Refresh
       If data_lin3.Recordset.RecordCount > 0 Then
          If IsNull(data_lin3.Recordset("in_obs")) = False Then
             data_fac.Recordset("obsp") = data_lin3.Recordset("in_obs")
          End If
       End If
       data_fac.Recordset.Update
       data_lin.Recordset.MoveNext
   Loop
End If

data_faccab.RecordSource = "Select * from cabezados"
data_faccab.Refresh

data_fac.RecordSource = "Select * from lineas2"
data_fac.Refresh

If data_fac.Recordset.RecordCount > 0 Then
   cr1.ReportFileName = App.path & "\faccnvnew.rpt"
   cr1.Action = 1
End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xfecha As Date
Xfecha = Date - 30
frm_impfaccnv.MousePointer = 11
'data_lin.DatabaseName = App.path & "\sapp.mdb"
data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xfecha, "yyyy-mm-dd") & "' and base in (101,102) order by fecha DESC"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   t_des.Text = data_lin.Recordset("factura")
   t_has.Text = data_lin.Recordset("factura")
End If
frm_impfaccnv.MousePointer = 0

End Sub

Private Sub Data2_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
data_fac.DatabaseName = App.path & "\factura.mdb"
data_faccab.DatabaseName = App.path & "\factura.mdb"
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.ConnectionString = "DSN=" & Xconexrmt
'data_lincab.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lincab.ConnectionString = "DSN=" & Xconexrmt
data_imagen.DatabaseName = App.path & "\imagen.mdb"

'data_lin3.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin3.ConnectionString = "DSN=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
