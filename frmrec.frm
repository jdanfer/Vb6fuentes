VERSION 5.00
Begin VB.Form frmrec 
   Caption         =   "Recibe"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Data data_rec 
      Caption         =   "data_rec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
End
Attribute VB_Name = "frmrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmrec.MousePointer = 11
data_cli.DatabaseName = App.Path & "\sapp.mdb"
data_rec.DatabaseName = App.Path & "\recibido.mdb"
data_rec.RecordSource = "caja"
data_rec.Refresh
data_cli.RecordSource = "caja"
data_cli.Refresh
If data_rec.Recordset.RecordCount > 0 Then
   data_rec.Recordset.MoveFirst
   Do While Not data_rec.Recordset.EOF
      data_cli.Recordset.AddNew
      data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
      data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
      data_cli.Recordset("numero") = data_rec.Recordset("numero")
      data_cli.Recordset("nombre") = data_rec.Recordset("nombre")
      data_cli.Recordset("movimiento") = data_rec.Recordset("movimiento")
      data_cli.Recordset("imp_fact") = data_rec.Recordset("imp_fact")
      data_cli.Recordset("observ") = data_rec.Recordset("observ")
      data_cli.Recordset("saldo") = data_rec.Recordset("saldo")
      data_cli.Recordset("usuario") = data_rec.Recordset("usuario")
      data_cli.Recordset("hora") = data_rec.Recordset("hora")
      data_cli.Recordset("saldo_user") = data_rec.Recordset("saldo_user")
      data_cli.Recordset("base") = data_rec.Recordset("base")
      data_cli.Recordset("cod_serv") = data_rec.Recordset("cod_serv")
      data_cli.Recordset("nom_serv") = data_rec.Recordset("nom_serv")
      data_cli.Recordset("cod_socio") = data_rec.Recordset("cod_socio")
      data_cli.Recordset("nom_socio") = data_rec.Recordset("nom_socio")
      data_cli.Recordset("caja_mesp") = data_rec.Recordset("caja_mesp")
      data_cli.Recordset("caja_anop") = data_rec.Recordset("caja_anop")
      data_cli.Recordset("imp_iva") = data_rec.Recordset("imp_iva")
      data_cli.Recordset("opiva") = data_rec.Recordset("opiva")
      data_cli.Recordset.Update
      data_rec.Recordset.MoveNext
   Loop
End If
data_rec.RecordSource = "linmmdd"
data_rec.Refresh
If data_rec.Recordset.RecordCount > 0 Then
   data_rec.Recordset.MoveFirst
   data_cli.RecordSource = "Select * from linmmdd"
   data_cli.Refresh
   Do While Not data_rec.Recordset.EOF
      data_cli.Recordset.AddNew
      data_cli.Recordset("tipo_mov") = data_rec.Recordset("tipo_mov")
      data_cli.Recordset("factura") = data_rec.Recordset("factura")
      data_cli.Recordset("tipo") = data_rec.Recordset("tipo")
      data_cli.Recordset("realizada") = data_rec.Recordset("realizada")
      data_cli.Recordset("fecha") = data_rec.Recordset("fecha")
      data_cli.Recordset("cod_cli") = data_rec.Recordset("cod_cli")
      data_cli.Recordset("nom_cli") = data_rec.Recordset("nom_cli")
      data_cli.Recordset("cod_prod") = data_rec.Recordset("cod_prod")
      data_cli.Recordset("nom_prod") = data_rec.Recordset("nom_prod")
      data_cli.Recordset("cantidad") = data_rec.Recordset("cantidad")
        data_cli.Recordset("moneda") = data_rec.Recordset("moneda")
        data_cli.Recordset("operador") = data_rec.Recordset("operador")
        data_cli.Recordset("hora") = data_rec.Recordset("hora")
        data_cli.Recordset("nro_flia") = data_rec.Recordset("nro_flia")
        data_cli.Recordset("nom_flia") = data_rec.Recordset("nom_flia")
        data_cli.Recordset("linea") = data_rec.Recordset("linea")
        data_cli.Recordset("convenio") = data_rec.Recordset("convenio")
        data_cli.Recordset("rub_cont") = data_rec.Recordset("rub_cont")
        data_cli.Recordset("usa_timbre") = data_rec.Recordset("usa_timbre")
        data_cli.Recordset("imp_timbre") = data_rec.Recordset("imp_timbre")
        data_cli.Recordset("tot_lin") = data_rec.Recordset("tot_lin")
        data_cli.Recordset("rub_nomb") = data_rec.Recordset("rub_nomb")
        data_cli.Recordset("nro_med_a") = data_rec.Recordset("nro_med_a")
        data_cli.Recordset("nom_med_a") = data_rec.Recordset("nom_med_a")
        data_cli.Recordset("precio_est") = data_rec.Recordset("precio_est")
        data_cli.Recordset("mes_paga") = data_rec.Recordset("mes_paga")
        data_cli.Recordset("ano_paga") = data_rec.Recordset("ano_paga")
        data_cli.Recordset("base") = data_rec.Recordset("base")
        data_cli.Recordset("imp_iva") = data_rec.Recordset("imp_iva")
        data_cli.Recordset("ruc") = data_rec.Recordset("ruc")
      data_cli.Recordset.Update
      data_rec.Recordset.MoveNext
   Loop
End If
frmrec.MousePointer = 0

MsgBox "Fin de la recepción"

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub desde_Validate(Action As Integer, Save As Integer)

End Sub
