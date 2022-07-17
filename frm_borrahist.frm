VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_borrahist 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Respaldar historial"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frm_borrahist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tabla Tesorería"
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
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Data data_respa 
      Caption         =   "data_respa"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_actual 
      Caption         =   "data_actual"
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mfec 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   240
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
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Respalda datos de Caja y líneas de factura.-"
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
      TabIndex        =   4
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "HASTA FECHA...:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frm_borrahist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_borrahist.MousePointer = 11
Command1.Enabled = False
If Check1.value = 1 Then
   Command3_Click
Else
    data_actual.RecordSource = "Select * from linmmdd where fecha <=#" & Format(mfec.Text, "yyyy/mm/dd") & "# order by fecha"
    data_actual.Refresh
    data_respa.RecordSource = "resplin"
    data_respa.Refresh
    If data_actual.Recordset.RecordCount > 0 Then
       data_actual.Recordset.MoveFirst
       Do While Not data_actual.Recordset.EOF
          data_respa.Recordset.AddNew
          data_respa.Recordset("reg_cab") = data_actual.Recordset("reg_cab")
          data_respa.Recordset("tipo_mov") = data_actual.Recordset("tipo_mov")
          data_respa.Recordset("factura") = data_actual.Recordset("factura")
          data_respa.Recordset("tipo") = data_actual.Recordset("tipo")
          data_respa.Recordset("realizada") = data_actual.Recordset("realizada")
          data_respa.Recordset("fecha") = data_actual.Recordset("fecha")
          data_respa.Recordset("vto") = data_actual.Recordset("vto")
          data_respa.Recordset("dias") = data_actual.Recordset("dias")
          data_respa.Recordset("cod_cli") = data_actual.Recordset("cod_cli")
          data_respa.Recordset("nom_cli") = data_actual.Recordset("nom_cli")
          data_respa.Recordset("cod_prod") = data_actual.Recordset("cod_prod")
          data_respa.Recordset("nom_prod") = data_actual.Recordset("nom_prod")
          data_respa.Recordset("univta") = data_actual.Recordset("univta")
          data_respa.Recordset("unidad") = data_actual.Recordset("unidad")
          data_respa.Recordset("cantidad") = data_actual.Recordset("cantidad")
          data_respa.Recordset("moneda") = data_actual.Recordset("moneda")
          data_respa.Recordset("tcambio") = data_actual.Recordset("tcambio")
          data_respa.Recordset("costo_prod") = data_actual.Recordset("costo_prod")
          data_respa.Recordset("margen_prd") = data_actual.Recordset("margen_prd")
          data_respa.Recordset("pre_prod") = data_actual.Recordset("pre_prod")
          data_respa.Recordset("iva") = data_actual.Recordset("iva")
          data_respa.Recordset("valor_iva") = data_actual.Recordset("valor_iva")
          data_respa.Recordset("pre_civa") = data_actual.Recordset("pre_civa")
          data_respa.Recordset("descuento") = data_actual.Recordset("descuento")
          data_respa.Recordset("recargo") = data_actual.Recordset("recargo")
          data_respa.Recordset("operador") = data_actual.Recordset("operador")
          data_respa.Recordset("hora") = data_actual.Recordset("hora")
          data_respa.Recordset("nro_flia") = data_actual.Recordset("nro_flia")
          data_respa.Recordset("nom_flia") = data_actual.Recordset("nom_flia")
          data_respa.Recordset("costo") = data_actual.Recordset("costo")
          data_respa.Recordset("nro_superv") = data_actual.Recordset("nro_superv")
          data_respa.Recordset("nom_superv") = data_actual.Recordset("nom_superv")
          data_respa.Recordset("grupo") = data_actual.Recordset("grupo")
          data_respa.Recordset("numero") = data_actual.Recordset("numero")
          data_respa.Recordset("zona") = data_actual.Recordset("zona")
          data_respa.Recordset("linea") = data_actual.Recordset("linea")
          data_respa.Recordset("convenio") = data_actual.Recordset("convenio")
          data_respa.Recordset("servicio") = data_actual.Recordset("servicio")
          data_respa.Recordset("pendiente") = data_actual.Recordset("pendiente")
          data_respa.Recordset("rub_cont") = data_actual.Recordset("rub_cont")
          data_respa.Recordset("arancel") = data_actual.Recordset("arancel")
          data_respa.Recordset("usa_timbre") = data_actual.Recordset("usa_timbre")
          data_respa.Recordset("imp_timbre") = data_actual.Recordset("imp_timbre")
          data_respa.Recordset("solicitant") = data_actual.Recordset("solicitant")
          data_respa.Recordset("ced_socio") = data_actual.Recordset("ced_socio")
          data_respa.Recordset("tot_lin") = data_actual.Recordset("tot_lin")
          data_respa.Recordset("fact") = data_actual.Recordset("fact")
          data_respa.Recordset("nro_med_s") = data_actual.Recordset("nro_med_s")
          data_respa.Recordset("nom_med_s") = data_actual.Recordset("nom_med_s")
          data_respa.Recordset("rub_nomb") = data_actual.Recordset("rub_nomb")
          data_respa.Recordset("nro_med_a") = data_actual.Recordset("nro_med_a")
          data_respa.Recordset("nom_med_a") = data_actual.Recordset("nom_med_a")
          data_respa.Recordset("precio_est") = data_actual.Recordset("precio_est")
          data_respa.Recordset("porce_est") = data_actual.Recordset("porce_est")
          data_respa.Recordset("mes_paga") = data_actual.Recordset("mes_paga")
          data_respa.Recordset("ano_paga") = data_actual.Recordset("ano_paga")
          data_respa.Recordset("base") = data_actual.Recordset("base")
          data_respa.Recordset("cod_medic") = data_actual.Recordset("cod_medic")
          data_respa.Recordset("nom_medic") = data_actual.Recordset("nom_medic")
          data_respa.Recordset("imp_iva") = data_actual.Recordset("imp_iva")
          data_respa.Recordset("ruc") = data_actual.Recordset("ruc")
          data_respa.Recordset.Update
          data_actual.Recordset.MoveNext
       Loop
       data_actual.Recordset.MoveFirst
       Do While Not data_actual.Recordset.EOF
          data_actual.Recordset.Delete
          data_actual.Recordset.MoveNext
       Loop
    End If
    data_actual.RecordSource = "Select * from caja where fecha <=#" & Format(mfec.Text, "yyyy/mm/dd") & "# order by fecha"
    data_actual.Refresh
    data_respa.RecordSource = "respcaja"
    data_respa.Refresh
    If data_actual.Recordset.RecordCount > 0 Then
       data_actual.Recordset.MoveFirst
       Do While Not data_actual.Recordset.EOF
          data_respa.Recordset.AddNew
          data_respa.Recordset("fecha") = data_actual.Recordset("fecha")
          data_respa.Recordset("numero") = data_actual.Recordset("numero")
          data_respa.Recordset("moneda") = data_actual.Recordset("moneda")
          data_respa.Recordset("nombre") = data_actual.Recordset("nombre")
          data_respa.Recordset("movimiento") = data_actual.Recordset("movimiento")
          data_respa.Recordset("imp_fact") = data_actual.Recordset("imp_fact")
          data_respa.Recordset("nrorub") = data_actual.Recordset("nrorub")
          data_respa.Recordset("rubro") = data_actual.Recordset("rubro")
          data_respa.Recordset("documento") = data_actual.Recordset("documento")
          data_respa.Recordset("observ") = data_actual.Recordset("observ")
          data_respa.Recordset("saldo") = data_actual.Recordset("saldo")
          data_respa.Recordset("usuario") = data_actual.Recordset("usuario")
          data_respa.Recordset("hora") = data_actual.Recordset("hora")
          data_respa.Recordset("saldo_user") = data_actual.Recordset("saldo_user")
          data_respa.Recordset("base") = data_actual.Recordset("base")
          data_respa.Recordset("cod_serv") = data_actual.Recordset("cod_serv")
          data_respa.Recordset("nom_serv") = data_actual.Recordset("nom_serv")
          data_respa.Recordset("cod_socio") = data_actual.Recordset("cod_socio")
          data_respa.Recordset("nom_socio") = data_actual.Recordset("nom_socio")
          data_respa.Recordset("turno") = data_actual.Recordset("turno")
          data_respa.Recordset("caja_mesp") = data_actual.Recordset("caja_mesp")
          data_respa.Recordset("caja_anop") = data_actual.Recordset("caja_anop")
          data_respa.Recordset("imp_iva") = data_actual.Recordset("imp_iva")
          data_respa.Recordset("opiva") = data_actual.Recordset("opiva")
          data_respa.Recordset.Update
          data_actual.Recordset.MoveNext
       Loop
       data_actual.Recordset.MoveFirst
       Do While Not data_actual.Recordset.EOF
          data_actual.Recordset.Delete
          data_actual.Recordset.MoveNext
       Loop
    End If
End If

frm_borrahist.MousePointer = 0
MsgBox "Proceso terminado"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xdesdenr As String
Xdesdenr = InputBox("Ingrese hasta que número de registro borra", "Historial")

data_actual.RecordSource = "Select * from tesorero where nromov <=" & Val(Xdesdenr)
data_actual.Refresh
data_respa.RecordSource = "resptes"
data_respa.Refresh
If data_actual.Recordset.RecordCount > 0 Then
   data_actual.Recordset.MoveFirst
   Do While Not data_actual.Recordset.EOF
      data_respa.Recordset.AddNew
      data_respa.Recordset("nromov") = data_actual.Recordset("nromov")
      data_respa.Recordset("fecha") = data_actual.Recordset("fecha")
      data_respa.Recordset("hora") = data_actual.Recordset("hora")
      data_respa.Recordset("usuario") = data_actual.Recordset("usuario")
      data_respa.Recordset("cod_rub") = data_actual.Recordset("cod_rub")
      data_respa.Recordset("nom_rub") = data_actual.Recordset("nom_rub")
      data_respa.Recordset("moneda") = data_actual.Recordset("moneda")
      data_respa.Recordset("monto") = data_actual.Recordset("monto")
      data_respa.Recordset("obs") = data_actual.Recordset("obs")
      data_respa.Recordset("cod_debe") = data_actual.Recordset("cod_debe")
      data_respa.Recordset("saldos") = data_actual.Recordset("saldos")
      data_respa.Recordset("concep") = data_actual.Recordset("concep")
      data_respa.Recordset("cod_haber") = data_actual.Recordset("cod_haber")
      data_respa.Recordset("saldou") = data_actual.Recordset("saldou")
      data_respa.Recordset("tipoc") = data_actual.Recordset("tipoc")
      data_respa.Recordset("libro") = data_actual.Recordset("libro")
      data_respa.Recordset("iva") = data_actual.Recordset("iva")
      data_respa.Recordset("base") = data_actual.Recordset("base")
      data_respa.Recordset("descon") = data_actual.Recordset("descon")
      data_respa.Recordset("bandera") = data_actual.Recordset("bandera")
      data_respa.Recordset("impiva") = data_actual.Recordset("impiva")
      data_respa.Recordset("tcam") = data_actual.Recordset("tcam")
      data_respa.Recordset.Update
      data_actual.Recordset.MoveNext
   Loop
   data_actual.Recordset.MoveFirst
   Do While Not data_actual.Recordset.EOF
      data_actual.Recordset.Delete
      data_actual.Recordset.MoveNext
   Loop
End If



End Sub

Private Sub Form_Load()
'data_actual.DatabaseName = App.Path & "\sapp.mdb"
data_actual.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_respa.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub
