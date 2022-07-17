VERSION 5.00
Begin VB.Form frm_sincroreser 
   Caption         =   "Sincroniza reservas"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2400
      Top             =   1440
   End
   Begin VB.Data data_srv 
      Caption         =   "data_srv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_rese 
      Caption         =   "data_rese"
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
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frm_sincroreser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XX As Integer

Private Sub Command1_Click()
Dim Xlaf As Date
Dim Xmodi As Integer
Xmodi = 0
Xlaf = Date
'         data_lla.Recordset("nrolla") = txt_nro.Text
'         data_lla.Recordset("nro") = txt_nro.Text
Timer1.Enabled = False
If XX >= 9 Then
   data_srv.RecordSource = "Select * from abmsocio where fecha =#" & Format(Xlaf, "yyyy/mm/dd") & "#"
   data_srv.Refresh
   If data_srv.Recordset.RecordCount > 0 Then
      data_srv.Recordset.MoveFirst
      Do While Not data_srv.Recordset.EOF
         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_srv.Recordset("cl_codigo")
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            data_rese.RecordSource = "Select * from clientes where cl_codigo =" & data_cli.Recordset("cl_codigo")
            data_rese.Refresh
            If data_rese.Recordset.RecordCount > 0 Then
               data_rese.Recordset.Edit
               If data_rese.Recordset("estado") <> data_cli.Recordset("estado") Then
                  Xmodi = 6
                  data_rese.Recordset("estado") = data_cli.Recordset("estado")
               End If
               If data_rese.Recordset("cl_codconv") <> data_cli.Recordset("cl_codconv") Then
                  Xmodi = 6
                  data_rese.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               End If
               If data_rese.Recordset("cl_apellid") <> data_cli.Recordset("cl_apellid") Then
                  Xmodi = 6
                  data_rese.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               End If
'               data_rese.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
'               data_rese.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
'               data_rese.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
               If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                  If IsNull(data_rese.Recordset("cl_dpto")) = False Then
                     If data_rese.Recordset("cl_dpto") <> data_cli.Recordset("cl_dpto") Then
                        Xmodi = 6
                        data_rese.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     End If
                  Else
                     Xmodi = 6
                     data_rese.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                  End If
               Else
                  If IsNull(data_rese.Recordset("cl_dpto")) = False Then
                     Xmodi = 6
                     data_rese.Recordset("cl_dpto") = Null
                  End If
               End If
'               data_rese.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
'               data_rese.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
'               data_rese.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
               If Xmodi = 6 Then
                  data_rese.Recordset.Update
               End If
               Xmodi = 0
            Else
               data_rese.Recordset.AddNew
               data_rese.Recordset("estado") = data_cli.Recordset("estado")
               data_rese.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_rese.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_rese.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_rese.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_rese.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
               data_rese.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
               data_rese.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
               data_rese.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
               data_rese.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
               data_rese.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
               data_rese.Recordset.Update
            End If
         End If
         data_srv.Recordset.MoveNext
      Loop
   End If
   If Format(Time, "HH:mm") > "20:55" Then
      data_srv.RecordSource = "Select * from deudas where fecha =#" & Format(Xlaf, "yyyy/mm/dd") & "#"
      data_srv.Refresh
      If data_srv.Recordset.RecordCount > 0 Then
         data_srv.Recordset.MoveFirst
         Do While Not data_srv.Recordset.EOF
            data_rese.RecordSource = "Select * from deudas where fecha =#" & Format(Xlaf, "yyyy/mm/dd") & "# and cliente =" & data_srv.Recordset("cliente") & _
            " and documento =" & data_srv.Recordset("documento")
            data_rese.Refresh
            If data_rese.Recordset.RecordCount > 0 Then
            Else
               data_rese.Recordset.AddNew
               data_rese.Recordset("fecha") = data_srv.Recordset("fecha")
               data_rese.Recordset("cliente") = data_srv.Recordset("cliente")
               data_rese.Recordset("mes") = data_srv.Recordset("mes")
               data_rese.Recordset("ano") = data_srv.Recordset("ano")
               data_rese.Recordset("documento") = data_srv.Recordset("documento")
               data_rese.Recordset("total") = data_srv.Recordset("total")
               data_rese.Recordset("fecha_pago") = data_srv.Recordset("fecha_pago")
               data_rese.Recordset("cod_cnv") = data_srv.Recordset("cod_cnv")
               data_rese.Recordset.Update
            End If
            data_srv.Recordset.MoveNext
         Loop
      End If
      data_srv.RecordSource = "Select * from deudas where fecha_pago =#" & Format(Xlaf, "yyyy/mm/dd") & "#"
      data_srv.Refresh
      If data_srv.Recordset.RecordCount > 0 Then
         data_srv.Recordset.MoveFirst
         Do While Not data_srv.Recordset.EOF
            data_rese.RecordSource = "Select * from deudas where cliente =" & data_srv.Recordset("cliente") & _
            " and documento =" & data_srv.Recordset("documento")
            data_rese.Refresh
            If data_rese.Recordset.RecordCount > 0 Then
               If IsNull(data_rese.Recordset("fecha_pago")) = False Then
                  If Format(data_rese.Recordset("fecha_pago"), "yyyy/mm/dd") <> Format(data_srv.Recordset("fecha_pago"), "yyyy/mm/dd") Then
                  
            Else
               data_rese.Recordset.AddNew
               data_rese.Recordset("fecha") = data_srv.Recordset("fecha")
               data_rese.Recordset("cliente") = data_srv.Recordset("cliente")
               data_rese.Recordset("mes") = data_srv.Recordset("mes")
               data_rese.Recordset("ano") = data_srv.Recordset("ano")
               data_rese.Recordset("documento") = data_srv.Recordset("documento")
               data_rese.Recordset("total") = data_srv.Recordset("total")
               data_rese.Recordset("fecha_pago") = data_srv.Recordset("fecha_pago")
               data_rese.Recordset("cod_cnv") = data_srv.Recordset("cod_cnv")
               data_rese.Recordset.Update
            End If
            data_srv.Recordset.MoveNext
         Loop
      End If
   
   
   

Timer1.Enabled = True


End Sub

Private Sub Form_Load()
data_srv.Connect = "odbc;dsn=sappnew;"
data_cli.Connect = "odbc;dsn=sappnew;"
data_rese.Connect = "odbc;dsn=sapplocar;"

End Sub

Private Sub Timer1_Timer()
XX = XX + 1
Command1_Click

End Sub
