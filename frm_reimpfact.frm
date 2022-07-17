VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_reimpfact 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re impresión de Factura"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5715
   Icon            =   "frm_reimpfact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lincab 
      Height          =   375
      Left            =   2880
      Top             =   600
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
   Begin MSAdodcLib.Adodc data_cab 
      Height          =   375
      Left            =   2880
      Top             =   1920
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
      Caption         =   "data_cab"
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
      Left            =   2280
      Top             =   2520
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
   Begin VB.Data data_cablocal 
      Caption         =   "data_cablocal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox t_mat 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   720
      Width           =   2295
   End
   Begin VB.Data data_imagen 
      Caption         =   "data_imagen"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_faccab 
      Caption         =   "data_faccab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2280
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   0   'False
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.TextBox t_serie 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Text            =   "A"
      Top             =   240
      Width           =   375
   End
   Begin VB.Data data_fac 
      Caption         =   "data_fac"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton b_salir 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Picture         =   "frm_reimpfact.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Picture         =   "frm_reimpfact.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Procesar"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txt_nro 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "NRO. SOCIO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "ATENCION!! Solo se pueden re-imprimir las facturas realizadas en el día."
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
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Socio:"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
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
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "NRO.FACTURA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   2760
      Picture         =   "frm_reimpfact.frx":0F56
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "frm_reimpfact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_acep_Click()

''''On Error GoTo Pasaalgoalimp

data_faccab.RecordSource = "cabezados"
data_faccab.Refresh
If data_faccab.Recordset.RecordCount > 0 Then
   data_faccab.Recordset.MoveFirst
   Do While Not data_faccab.Recordset.EOF
      data_faccab.Recordset.Delete
      data_faccab.Recordset.MoveNext
   Loop
End If

data_fac.RecordSource = "lineas"
data_fac.Refresh
If data_fac.Recordset.RecordCount > 0 Then
   data_fac.Recordset.MoveFirst
   Do While Not data_fac.Recordset.EOF
      data_fac.Recordset.Delete
      data_fac.Recordset.MoveNext
   Loop
End If

data_cab.RecordSource = "Select * from clirespl where cl_numero =" & txt_nro.Text & " and cl_socmnro ='" & t_serie.Text & "' and cl_codigo =" & t_mat.Text
data_cab.Refresh
If data_cab.Recordset.RecordCount > 0 Then
   data_faccab.Recordset.AddNew
   data_faccab.Recordset("cl_socmnro") = data_cab.Recordset("cl_socmnro")
   data_faccab.Recordset("cl_numero") = data_cab.Recordset("cl_numero")
   data_imagen.RecordSource = "Select * from qr where nrofact =" & txt_nro.Text & " and serie ='" & t_serie.Text & "'"
   data_imagen.Refresh
   If data_imagen.Recordset.RecordCount > 0 Then
      data_faccab.Recordset("qr") = data_imagen.Recordset("qr")
   End If
   data_faccab.Recordset("cl_tipocli") = data_cab.Recordset("cl_tipocli")
   data_faccab.Recordset("cl_socmnro") = data_cab.Recordset("cl_socmnro")
   data_faccab.Recordset("cl_numero") = data_cab.Recordset("cl_numero")
   data_faccab.Recordset("cl_fnac") = data_cab.Recordset("cl_fnac")
   data_faccab.Recordset("fecha_reac") = data_cab.Recordset("fecha_reac")
   data_faccab.Recordset("cl_tj_venc") = data_cab.Recordset("cl_tj_venc")
   data_faccab.Recordset("cl_nrovend") = data_cab.Recordset("cl_nrovend")
   data_faccab.Recordset("cl_forpago") = data_cab.Recordset("cl_forpago")
   data_faccab.Recordset("cl_celular") = data_cab.Recordset("cl_celular")
   data_faccab.Recordset("fecha_modi") = data_cab.Recordset("fecha_modi")
   data_faccab.Recordset("cl_diacobr") = data_cab.Recordset("cl_diacobr")
   data_faccab.Recordset("cl_nrotarj") = data_cab.Recordset("cl_nrotarj")
   data_faccab.Recordset("cl_tjemi_n") = data_cab.Recordset("cl_tjemi_n")
   data_faccab.Recordset("cl_tjemi_c") = data_cab.Recordset("cl_tjemi_c")
   data_faccab.Recordset("cl_referen") = data_cab.Recordset("cl_referen")
   data_faccab.Recordset("tit_tarj") = data_cab.Recordset("tit_tarj")
   data_faccab.Recordset("cl_nomconv") = data_cab.Recordset("cl_nomconv")
   data_faccab.Recordset("cl_nro_sup") = data_cab.Recordset("cl_nro_sup")
   data_faccab.Recordset("hora_baja") = data_cab.Recordset("hora_baja")
   data_faccab.Recordset("cl_nom_sup") = data_cab.Recordset("cl_nom_sup")
   data_faccab.Recordset("info_debit") = data_cab.Recordset("info_debit")
   data_faccab.Recordset("cl_direcci") = data_cab.Recordset("cl_direcci")
   data_faccab.Recordset("cl_zona") = data_cab.Recordset("cl_zona")
   data_faccab.Recordset("cl_localid") = data_cab.Recordset("cl_localid")
   data_faccab.Recordset("cl_codigo") = data_cab.Recordset("cl_codigo")
   data_faccab.Recordset("usu_baja") = data_cab.Recordset("usu_baja")
   data_faccab.Recordset("saldo_chc2") = data_cab.Recordset("saldo_chc2")
   data_faccab.Recordset("saldo_cc") = data_cab.Recordset("saldo_cc")
   data_faccab.Recordset("saldo_cc2") = data_cab.Recordset("saldo_cc2")
   data_faccab.Recordset("cl_atrasoa") = data_cab.Recordset("cl_atrasoa")
   data_faccab.Recordset("cl_cedula") = data_cab.Recordset("cl_cedula")
   data_faccab.Recordset("saldo_doc2") = data_cab.Recordset("saldo_doc2")
   data_faccab.Recordset("cl_atrasop") = data_cab.Recordset("cl_atrasop")
   data_faccab.Recordset("cl_decuota") = data_cab.Recordset("cl_decuota")
   data_faccab.Recordset("saldo_doc") = data_cab.Recordset("saldo_doc")
   data_faccab.Recordset("cl_grupo") = data_cab.Recordset("cl_grupo")
   data_faccab.Recordset("saldo_chc") = data_cab.Recordset("saldo_chc")
   data_faccab.Recordset("cl_telefon") = data_cab.Recordset("cl_telefon")
   data_faccab.Recordset("cl_fultpag") = data_cab.Recordset("cl_fultpag")
   data_faccab.Recordset("cl_ultmesp") = data_cab.Recordset("cl_ultmesp")
   data_faccab.Recordset("cl_nomvend") = data_cab.Recordset("cl_nomvend")
   data_faccab.Recordset("cl_fax") = data_cab.Recordset("cl_fax")
   data_faccab.Recordset("cl_nombre") = data_cab.Recordset("cl_nombre")
   data_lin.RecordSource = "Select * from linmmdd where factura =" & txt_nro.Text & " and cod_cli =" & t_mat.Text
   data_lin.Refresh
   If data_lin.Recordset.RecordCount > 0 Then
      data_lin.Recordset.MoveFirst
      data_faccab.Recordset("obsp") = "CAJA BASE:" & Trim(str(data_lin.Recordset("base"))) & " Usuario: " & data_lin.Recordset("operador")
   End If
   data_faccab.Recordset.Update
   data_faccab.Refresh
   data_faccab.Recordset.MoveFirst
'fin de cabezal
End If

data_lin.RecordSource = "Select * from linmmdd where factura =" & txt_nro.Text & " and cod_cli =" & t_mat.Text
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
       data_fac.Recordset.AddNew
       data_fac.Recordset("fecha") = data_lin.Recordset("fecha")
       If IsNull(data_cab.Recordset("cl_nombre")) = False Then
          data_fac.Recordset("libro_rub") = data_cab.Recordset("cl_nombre")
       End If
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
       If data_faccab.Recordset.RecordCount > 0 Then
          data_fac.Recordset("cl_nrotarj") = data_faccab.Recordset("cl_nrotarj")
          data_fac.Recordset("cl_referen") = data_faccab.Recordset("cl_referen")
          data_fac.Recordset("cl_tjemi_c") = data_faccab.Recordset("cl_tjemi_c")
          data_fac.Recordset("cl_diacobr") = data_faccab.Recordset("cl_diacobr")
          data_fac.Recordset("cl_telefon") = data_faccab.Recordset("cl_telefon")
          data_fac.Recordset("obsp") = data_faccab.Recordset("obsp")
          data_fac.Recordset("qr") = data_faccab.Recordset("qr")
          data_fac.Recordset("cl_fax") = data_faccab.Recordset("cl_fax")
          data_fac.Recordset("cl_socmnro") = data_faccab.Recordset("cl_socmnro")
          data_fac.Recordset("cl_numero") = data_faccab.Recordset("cl_numero")
          data_fac.Recordset("cl_celular") = data_faccab.Recordset("cl_celular")
          data_fac.Recordset("cl_fnac") = data_faccab.Recordset("cl_fnac")
          data_fac.Recordset("usu_baja") = data_faccab.Recordset("usu_baja")
          data_fac.Recordset("info_debit") = data_faccab.Recordset("info_debit")
          data_fac.Recordset("cl_nrocobr") = data_faccab.Recordset("cl_nrocobr")
          data_fac.Recordset("cl_medflia") = data_faccab.Recordset("cl_medflia")
          data_fac.Recordset("hora_baja") = data_faccab.Recordset("hora_baja")
          data_fac.Recordset("cl_nomcobr") = data_faccab.Recordset("cl_nomcobr")
          data_fac.Recordset("cl_nom_sup") = data_faccab.Recordset("cl_nom_sup")
          data_fac.Recordset("saldo_cc") = data_faccab.Recordset("saldo_cc")
          data_fac.Recordset("saldo_doc") = data_faccab.Recordset("saldo_doc")
          If IsNull(data_faccab.Recordset("cl_fultpag")) = False Then
             data_fac.Recordset("cl_fultpag") = data_faccab.Recordset("cl_fultpag")
          End If
          data_fac.Recordset("cl_nombre") = data_faccab.Recordset("cl_nombre")
       Else
       
       End If
       data_fac.Recordset.Update
       data_lin.Recordset.MoveNext
   Loop
End If

'data_faccab.RecordSource = "Select * from cabezados order by cl_numero"
'data_faccab.Refresh

data_fac.RecordSource = "Select * from lineas order by factura"
data_fac.Refresh

'If data_faccab.Recordset.RecordCount > 0 Then
'   data_faccab.Recordset.MoveFirst
'End If
If data_fac.Recordset.RecordCount > 0 Then
'   data_fac.Recordset.MoveLast
   data_fac.Recordset.MoveFirst
   If data_fac.Recordset("libro_rub") = "RECIBO" Then
      cr1.ReportFileName = App.path & "\infticksapp4.rpt"
      If data_fac.Recordset("cod_prod") = 999 Or data_fac.Recordset("cod_prod") = 997 Then
      Else
         cr1.CopiesToPrinter = 2
      End If
   Else
      cr1.ReportFileName = App.path & "\infticksapp3.rpt"
   End If
   cr1.DiscardSavedData = True
   cr1.Action = 1
End If

b_acep.Enabled = True
b_salir.Enabled = True

'Exit Sub

'Pasaalgoalimp:
'              If Err.Number = 3155 Then
'                 MsgBox "Error al seleccionar datos, verifique números y vuelva a intentar.", vbInformation
'              Else
'                 MsgBox "Error al seleccionar datos, verifique números y vuelva a intentar.", vbInformation
'              End If

End Sub

Private Sub b_salir_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_fac.DatabaseName = App.path & "\factura.mdb"
data_faccab.DatabaseName = App.path & "\factura.mdb"

data_imagen.DatabaseName = App.path & "\imagen.mdb"

data_lin.ConnectionString = "dsn=" & Xconexrmt
data_cab.ConnectionString = "dsn=" & Xconexrmt

'data_fac.RecordSource = "Select * from lineas"
'data_fac.Refresh
'If data_fac.Recordset.RecordCount > 0 Then
'   txt_nro.Text = data_fac.Recordset("factura")
'Else
'   txt_nro.Text = 0
'End If

data_lincab.ConnectionString = "dsn=" & Xconexrmt
data_lincab.RecordSource = "Select * from clirespl where cl_numero =" & 10352
data_lincab.Refresh

data_cablocal.DatabaseName = App.path & "\cablocal.mdb"
data_cablocal.RecordSource = "Select * from cabezados where cl_codced =" & 2
data_cablocal.Refresh

If data_cablocal.Recordset.RecordCount > 0 Then
   data_cablocal.Recordset.MoveFirst
   Do While Not data_cablocal.Recordset.EOF
      data_lincab.Recordset.AddNew
      data_lincab.Recordset("cl_tipcli") = "1.0"
      data_lincab.Recordset("cl_tipocli") = data_cablocal.Recordset("cl_tipocli")
      data_lincab.Recordset("cl_socmnro") = data_cablocal.Recordset("cl_socmnro")
      data_lincab.Recordset("cl_numero") = data_cablocal.Recordset("cl_numero")
      data_lincab.Recordset("cl_fnac") = data_cablocal.Recordset("cl_fnac")
      data_lincab.Recordset("fecha_reac") = data_cablocal.Recordset("fecha_reac")
      data_lincab.Recordset("cl_tj_venc") = data_cablocal.Recordset("cl_tj_venc")
      data_lincab.Recordset("cl_nrovend") = data_cablocal.Recordset("cl_nrovend")
      data_lincab.Recordset("cl_forpago") = data_cablocal.Recordset("cl_forpago")
      data_lincab.Recordset("cl_celular") = data_cablocal.Recordset("cl_celular") 'descripcion f.pago
      data_lincab.Recordset("fecha_modi") = data_cablocal.Recordset("fecha_modi")
      data_lincab.Recordset("cl_diacobr") = data_cablocal.Recordset("cl_diacobr")
      data_lincab.Recordset("cl_nrotarj") = data_cablocal.Recordset("cl_nrotarj")
      data_lincab.Recordset("cl_tjemi_n") = data_cablocal.Recordset("cl_tjemi_n")
      data_lincab.Recordset("cl_tjemi_c") = data_cablocal.Recordset("cl_tjemi_c")
      data_lincab.Recordset("cl_referen") = data_cablocal.Recordset("cl_referen")
      data_lincab.Recordset("tit_tarj") = data_cablocal.Recordset("tit_tarj")
      data_lincab.Recordset("cl_nomconv") = data_cablocal.Recordset("cl_nomconv")
        'receptor
      data_lincab.Recordset("cl_nro_sup") = data_cablocal.Recordset("cl_nro_sup")
      data_lincab.Recordset("hora_baja") = data_cablocal.Recordset("hora_baja")
      data_lincab.Recordset("cl_nom_sup") = data_cablocal.Recordset("cl_nom_sup")
            'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
      data_lincab.Recordset("info_debit") = data_cablocal.Recordset("info_debit")
      data_lincab.Recordset("cl_direcci") = data_cablocal.Recordset("cl_direcci")
      data_lincab.Recordset("cl_zona") = data_cablocal.Recordset("cl_zona")
        'data_cabezal.Recordset("cl_email") = frm_convenios.t_dpto.Text 'opcional
      data_lincab.Recordset("cl_localid") = data_cablocal.Recordset("cl_localid") 'opcional
      data_lincab.Recordset("cl_codigo") = data_cablocal.Recordset("cl_codigo")
      data_lincab.Recordset("usu_baja") = data_cablocal.Recordset("usu_baja") 'moneda
      data_lincab.Recordset("saldo_chc2") = data_cablocal.Recordset("saldo_chc2") 'valor dolar
      data_lincab.Recordset("saldo_cc") = data_cablocal.Recordset("saldo_cc")  'iva minimo
      data_lincab.Recordset("saldo_cc2") = data_cablocal.Recordset("saldo_cc2") 'iva básico
      data_lincab.Recordset("cl_atrasoa") = data_cablocal.Recordset("cl_atrasoa") 'subtot iva 22
      data_lincab.Recordset("cl_cedula") = data_cablocal.Recordset("cl_cedula") 'subtot iva cero
      data_lincab.Recordset("saldo_doc2") = data_cablocal.Recordset("saldo_doc2")
      data_lincab.Recordset("cl_atrasop") = data_cablocal.Recordset("cl_atrasop")
      data_lincab.Recordset("cl_decuota") = data_cablocal.Recordset("cl_decuota")
      data_lincab.Recordset("saldo_doc") = data_cablocal.Recordset("saldo_doc")
      data_lincab.Recordset("cl_grupo") = data_cablocal.Recordset("cl_grupo")
      data_lincab.Recordset("saldo_chc") = data_cablocal.Recordset("saldo_chc")
      data_lincab.Recordset("cl_telefon") = data_cablocal.Recordset("cl_telefon")
      data_lincab.Recordset("cl_nombre") = data_cablocal.Recordset("cl_nombre")
      data_lincab.Recordset("cl_cuopaga") = data_cablocal.Recordset("cl_cuopaga")
      data_lincab.Recordset("codmotbaja") = data_cablocal.Recordset("codmotbaja")
      data_lincab.Recordset("ultanopmut") = data_cablocal.Recordset("ultanopmut")
      data_lincab.Recordset("cl_fultvta") = data_cablocal.Recordset("cl_fultvta")
      data_lincab.Recordset("cl_entre") = data_cablocal.Recordset("cl_entre")
      data_lincab.Recordset("codmotbaja") = data_cablocal.Recordset("codmotbaja")
      data_lincab.Recordset("ultanopmut") = data_cablocal.Recordset("ultanopmut")
      data_lincab.Recordset("cl_fultvta") = data_cablocal.Recordset("cl_fultvta")
      data_lincab.Recordset("cl_entre") = data_cablocal.Recordset("cl_entre")
      data_lincab.Recordset("cl_fultpag") = data_cablocal.Recordset("cl_fultpag")
      data_lincab.Recordset("cl_ultmesp") = data_cablocal.Recordset("cl_ultmesp")
      data_lincab.Recordset("cl_nomvend") = data_cablocal.Recordset("cl_nomvend")
      data_lincab.Recordset("cl_fax") = data_cablocal.Recordset("cl_fax")
      data_lincab.Recordset.Update
        'fin de cabezal
      data_cablocal.Recordset.Edit
      data_cablocal.Recordset("cl_codced") = 1
      data_cablocal.Recordset.Update
      data_cablocal.Recordset.MoveNext
   Loop
End If



End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub

Private Sub t_mat_LostFocus()
If txt_nro.Text <> "" Then
   If t_mat.Text <> "" Then
      data_lin.RecordSource = "Select * from linmmdd where factura =" & txt_nro.Text & " and cod_cli =" & t_mat.Text
      data_lin.Refresh
      If data_lin.Recordset.RecordCount > 0 Then
         Label2.Caption = data_lin.Recordset("nom_cli")
      End If
   Else
      Label2.Caption = ""
   End If
Else
   Label2.Caption = ""
End If

End Sub

Private Sub txt_nro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mat.SetFocus
End If

End Sub

