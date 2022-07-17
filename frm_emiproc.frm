VERSION 5.00
Begin VB.Form frm_emiproc 
   Caption         =   "Emision"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.Data data_ctrolmes 
      Caption         =   "data_ctrolmes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_deu 
      Caption         =   "data_deu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_emicopi 
      Caption         =   "data_emicopi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CARGAR"
      Height          =   1335
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frm_emiproc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_emiproc.MousePointer = 11
If data_emicopi.Recordset.RecordCount > 0 Then
   data_emicopi.Recordset.MoveFirst
   Do While Not data_emicopi.Recordset.EOF
      Data1.Recordset.AddNew
      Data1.Recordset("cliente") = data_emicopi.Recordset("cliente")
      Data1.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
      Data1.Recordset("nom_cnv") = data_emicopi.Recordset("nom_cnv")
      Data1.Recordset("ruc") = data_emicopi.Recordset("ruc")
      Data1.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
      Data1.Recordset("apellidos") = data_emicopi.Recordset("apellidos")
      Data1.Recordset("cedula") = data_emicopi.Recordset("cedula")
      Data1.Recordset("cod") = data_emicopi.Recordset("cod")
      Data1.Recordset("fecha") = data_emicopi.Recordset("fecha")
      Data1.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
      Data1.Recordset("documento") = data_emicopi.Recordset("documento")
      Data1.Recordset("tipo") = data_emicopi.Recordset("tipo")
      Data1.Recordset("importe") = data_emicopi.Recordset("importe")
      Data1.Recordset("debe_haber") = data_emicopi.Recordset("debe_haber")
      Data1.Recordset("moneda") = data_emicopi.Recordset("moneda")
      Data1.Recordset("origen") = data_emicopi.Recordset("origen")
      Data1.Recordset("operador") = data_emicopi.Recordset("operador")
      Data1.Recordset("hora") = data_emicopi.Recordset("hora")
      Data1.Recordset("dir_cli") = data_emicopi.Recordset("dir_cli")
      Data1.Recordset("loc_cli") = data_emicopi.Recordset("loc_cli")
      Data1.Recordset("tel_cli") = data_emicopi.Recordset("tel_cli")
      Data1.Recordset("nro_superv") = data_emicopi.Recordset("nro_superv")
      Data1.Recordset("nom_superv") = data_emicopi.Recordset("nom_superv")
      Data1.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
      Data1.Recordset("nom_vende") = data_emicopi.Recordset("nom_vende")
      Data1.Recordset("grupo") = data_emicopi.Recordset("grupo")
      Data1.Recordset("numero") = data_emicopi.Recordset("numero")
      Data1.Recordset("zona") = data_emicopi.Recordset("zona")
      Data1.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
      Data1.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
      Data1.Recordset("mes") = data_emicopi.Recordset("mes")
      Data1.Recordset("ano") = data_emicopi.Recordset("ano")
      Data1.Recordset("color_rec") = data_emicopi.Recordset("color_rec")
      Data1.Recordset("fecha_ing") = data_emicopi.Recordset("fecha_ing")
      Data1.Recordset("fecha_nac") = data_emicopi.Recordset("fecha_nac")
      Data1.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
      Data1.Recordset("deudas") = data_emicopi.Recordset("deudas")
      Data1.Recordset("servi") = data_emicopi.Recordset("servi")
      Data1.Recordset("iva") = data_emicopi.Recordset("iva")
      Data1.Recordset("total") = data_emicopi.Recordset("total")
      Data1.Recordset.Update
      
      data_deu.Recordset.AddNew
      data_deu.Recordset("cod_cnv") = data_emicopi.Recordset("cod_cnv")
      data_deu.Recordset("nom_cnv") = Mid(data_emicopi.Recordset("nom_cnv"), 1, 20)
      data_deu.Recordset("tipocta") = data_emicopi.Recordset("tipocta")
      data_deu.Recordset("cliente") = data_emicopi.Recordset("cliente")
      data_deu.Recordset("nombre") = data_emicopi.Recordset("apellidos")
      data_deu.Recordset("fecha") = data_emicopi.Recordset("fecha")
      data_deu.Recordset("tipodoc") = data_emicopi.Recordset("tipodoc")
      data_deu.Recordset("documento") = data_emicopi.Recordset("documento")
      data_deu.Recordset("importe") = data_emicopi.Recordset("importe")
      data_deu.Recordset("moneda") = data_emicopi.Recordset("moneda")
      data_deu.Recordset("origen") = "EMISION..." & Trim(Str(data_emicopi.Recordset("mes"))) & "/" & Trim(Str(data_emicopi.Recordset("ano")))
      data_deu.Recordset("nro_vende") = data_emicopi.Recordset("nro_vende")
      data_deu.Recordset("grupo") = data_emicopi.Recordset("grupo")
      data_deu.Recordset("saldo_cc") = 0
      data_deu.Recordset("mes") = data_emicopi.Recordset("mes")
      data_deu.Recordset("ano") = data_emicopi.Recordset("ano")
      data_deu.Recordset("nro_cobr") = data_emicopi.Recordset("nro_cobr")
      data_deu.Recordset("nom_cobr") = data_emicopi.Recordset("nom_cobr")
      data_deu.Recordset("estado_cta") = 1
      data_deu.Recordset("tiquet") = data_emicopi.Recordset("tiquet")
      data_deu.Recordset("deudas") = data_emicopi.Recordset("deudas")
      data_deu.Recordset("total") = data_emicopi.Recordset("total")
      data_deu.Recordset("servi") = data_emicopi.Recordset("servi")
      data_deu.Recordset("iva") = data_emicopi.Recordset("iva")
      data_deu.Recordset("nro_superv") = 50
      data_deu.Recordset.Update
      data_emicopi.Recordset.MoveNext
   Loop
   
'   data_ctrolmes.Refresh
   data_ctrolmes.Recordset.Edit
   data_ctrolmes.Recordset("salidas") = 8
   data_ctrolmes.Recordset("entradas") = 2016
   data_ctrolmes.Recordset.Update

   MsgBox "Terminado", vbInformation
   End
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=sapp;"
Data1.RecordSource = "EMI0816"
Data1.Refresh

data_emicopi.DatabaseName = App.Path & "\emisnueva.mdb"
data_emicopi.RecordSource = "EMI0816"
data_emicopi.Refresh

data_deu.Connect = "odbc;dsn=sapp;"
data_deu.RecordSource = "deudas"
data_deu.Refresh

data_ctrolmes.Connect = "odbc;dsn=sapp;"
data_ctrolmes.RecordSource = "saldos"
data_ctrolmes.Refresh

End Sub
