VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_pedidomedicb 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar pedidos ingresados"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12465
   Icon            =   "frm_pedidomedicb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   12465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      Picture         =   "frm_pedidomedicb.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Asociar número de factura de pago."
      Top             =   5280
      Width           =   375
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5520
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_pedidoimp 
      Caption         =   "data_pedidoimp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data_pedlin 
      Caption         =   "Data_pedlin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir pedido"
      Height          =   495
      Left            =   240
      Picture         =   "frm_pedidomedicb.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pedidos_medic"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_pedidomedicb.frx":109E
      Height          =   4335
      Left            =   360
      OleObjectBlob   =   "frm_pedidomedicb.frx":10B2
      TabIndex        =   1
      Top             =   840
      Width           =   11775
   End
   Begin MSComctlLib.TabStrip tabver 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pendientes"
            Key             =   "a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "En tránsito"
            Key             =   "b"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrados"
            Key             =   "c"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labdescpedido 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Doble click para editar el registro"
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
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   3855
   End
End
Attribute VB_Name = "frm_pedidomedicb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Deseaimprimir As String
Deseaimprimir = MsgBox("Desea imprimir el pedido de: " & Data1.Recordset("nombre") & " ?", vbInformation + vbYesNo, "Pedidos")
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
If Deseaimprimir = vbYes Then
   Command1.Enabled = False
   data_pedidoimp.RecordSource = "select * from pedidos_medic where pedido_nro =" & Data1.Recordset("pedido_nro")
   data_pedidoimp.Refresh
   If data_pedidoimp.Recordset.RecordCount > 0 Then
      data_pedidoimp.Recordset.MoveFirst
      Data_pedlin.RecordSource = "select * from pedidos_mediclin where cod_pedido =" & Data1.Recordset("pedido_nro")
      Data_pedlin.Refresh
      If Data_pedlin.Recordset.RecordCount > 0 Then
         Do While Not Data_pedlin.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = data_pedidoimp.Recordset("matricula")
            data_inf.Recordset("cl_cedula") = data_pedidoimp.Recordset("cedula")
            data_inf.Recordset("cl_codced") = data_pedidoimp.Recordset("codced")
            data_inf.Recordset("cl_apellid") = Mid(data_pedidoimp.Recordset("nombre"), 1, 60)
            data_inf.Recordset("cl_fecing") = data_pedidoimp.Recordset("fecha")
            data_inf.Recordset("cl_celular") = data_pedidoimp.Recordset("hora")
            data_inf.Recordset("cl_nrovend") = data_pedidoimp.Recordset("pedido_nro")
            data_inf.Recordset("cl_socmnom") = data_pedidoimp.Recordset("mutual")
            data_inf.Recordset("info_debit") = data_pedidoimp.Recordset("direcc")
            data_inf.Recordset("cl_zona") = Mid(data_pedidoimp.Recordset("zona"), 1, 25)
            data_inf.Recordset("cl_nombre") = Mid(data_pedidoimp.Recordset("telefs"), 1, 30)
            data_inf.Recordset("cl_localid") = Mid(data_pedidoimp.Recordset("recibe1"), 1, 35)
            data_inf.Recordset("cl_fnac") = data_pedidoimp.Recordset("fec_ent")
            data_inf.Recordset("cl_dpto") = data_pedidoimp.Recordset("hor_ent1")
            data_inf.Recordset("cl_estadoc") = data_pedidoimp.Recordset("hor_ent2")
            data_inf.Recordset("cl_atrasoa") = data_pedidoimp.Recordset("tot_cant")
            data_inf.Recordset("cl_atrasop") = data_pedidoimp.Recordset("tot_pesos")
            data_inf.Recordset("cl_descpag") = data_pedidoimp.Recordset("forma_pago")
            data_inf.Recordset("cl_email") = Mid(data_pedidoimp.Recordset("nom_cadete"), 1, 30)
            data_inf.Recordset("cl_entre") = Mid(Data_pedlin.Recordset("nom_medic"), 1, 80)
            data_inf.Recordset("cl_cantdia") = Data_pedlin.Recordset("cant")
            data_inf.Recordset("cl_cantpag") = Data_pedlin.Recordset("imp_unit")
            data_inf.Recordset("cl_pri_vto") = Data_pedlin.Recordset("tot_imp")
            data_inf.Recordset("cl_nomcobr") = data_pedidoimp.Recordset("recetasen")
            data_inf.Recordset("cl_dircobr") = Mid(data_pedidoimp.Recordset("recetas_obs"), 1, 100)
            data_inf.Recordset.Update
            Data_pedlin.Recordset.MoveNext
         Loop
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_codigo") = data_pedidoimp.Recordset("matricula")
         data_inf.Recordset("cl_cedula") = data_pedidoimp.Recordset("cedula")
         data_inf.Recordset("cl_codced") = data_pedidoimp.Recordset("codced")
         data_inf.Recordset("cl_apellid") = Mid(data_pedidoimp.Recordset("nombre"), 1, 60)
         data_inf.Recordset("cl_fecing") = data_pedidoimp.Recordset("fecha")
         data_inf.Recordset("cl_celular") = data_pedidoimp.Recordset("hora")
         data_inf.Recordset("cl_nrovend") = data_pedidoimp.Recordset("pedido_nro")
         data_inf.Recordset("cl_socmnom") = data_pedidoimp.Recordset("mutual")
         data_inf.Recordset("info_debit") = data_pedidoimp.Recordset("direcc")
         data_inf.Recordset("cl_zona") = Mid(data_pedidoimp.Recordset("zona"), 1, 25)
         data_inf.Recordset("cl_nombre") = Mid(data_pedidoimp.Recordset("telefs"), 1, 30)
         data_inf.Recordset("cl_localid") = Mid(data_pedidoimp.Recordset("recibe1"), 1, 35)
         data_inf.Recordset("cl_fnac") = data_pedidoimp.Recordset("fec_ent")
         data_inf.Recordset("cl_dpto") = data_pedidoimp.Recordset("hor_ent1")
         data_inf.Recordset("cl_estadoc") = data_pedidoimp.Recordset("hor_ent2")
         data_inf.Recordset("cl_atrasoa") = data_pedidoimp.Recordset("tot_cant")
         data_inf.Recordset("cl_atrasop") = data_pedidoimp.Recordset("tot_pesos")
         data_inf.Recordset("cl_descpag") = data_pedidoimp.Recordset("forma_pago")
         data_inf.Recordset("cl_email") = Mid(data_pedidoimp.Recordset("nom_cadete"), 1, 30)
         data_inf.Recordset("cl_cantdia") = 1
         data_inf.Recordset("cl_cantpag") = Devuelve_costoPed()
         data_inf.Recordset("cl_pri_vto") = Devuelve_costoPed()
         If Trim(labdescpedido.Caption) <> "" Then
            data_inf.Recordset("cl_entre") = labdescpedido.Caption
         Else
            data_inf.Recordset("cl_entre") = "PEDIDO A DOMICILIO"
         End If
         '   data_inf.Recordset("cl_cuopaga") = Devuelve_costoPed()
         data_inf.Recordset("cl_nomcobr") = data_pedidoimp.Recordset("recetasen")
         data_inf.Recordset("cl_dircobr") = Mid(data_pedidoimp.Recordset("recetas_obs"), 1, 100)
         data_inf.Recordset.Update
         data_inf.RecordSource = "select * from infcli"
         data_inf.Refresh
         cr1.ReportFileName = App.path & "\pedidosmedic.rpt"
         cr1.Action = 1
         
      Else
         MsgBox "Sin datos de medicamentos."
      End If
   Else
      MsgBox "No se encuentra pedido"
   End If
   Command1.Enabled = True
End If


End Sub

Private Sub Command2_Click()
Dim XnroFact As String
XnroFact = InputBox("Ingrese número de factura:")
If Trim(XnroFact) <> "" Then
   Data_pedlin.RecordSource = "select * from linmmdd where factura =" & Val(XnroFact) & " and nro_flia in (6)"
   Data_pedlin.Refresh
   If Data_pedlin.Recordset.RecordCount > 0 Then
      data_pedidoimp.RecordSource = "select * from pedidos_medic where pedido_nro =" & Data1.Recordset("pedido_nro")
      data_pedidoimp.Refresh
      If data_pedidoimp.Recordset.RecordCount > 0 Then
         data_pedidoimp.Recordset.MoveFirst
         If IsNull(data_pedidoimp.Recordset("nro_factura")) = True Then
            data_pedidoimp.Recordset.Edit
            data_pedidoimp.Recordset("nro_factura") = Data_pedlin.Recordset("factura")
            data_pedidoimp.Recordset("fecha_fac") = Data_pedlin.Recordset("fecha")
            data_pedidoimp.Recordset.Update
            MsgBox "Actualizado!"
         End If
      End If
   Else
      MsgBox "No se encuentra factura, verifique!", vbCritical
   End If
Else
   MsgBox "No ingresó factura"
End If

End Sub

Private Sub DBGrid1_DblClick()
      
frm_pedidomedic.labfec.Caption = Data1.Recordset("fecha")
frm_pedidomedic.labhor.Caption = Data1.Recordset("hora")
frm_pedidomedic.labbase.Caption = Data1.Recordset("base")
frm_pedidomedic.labusua.Caption = Data1.Recordset("usuario")
frm_pedidomedic.labpedido.Caption = Data1.Recordset("pedido_nro")
frm_pedidomedic.t_ced.Text = Data1.Recordset("cedula")
frm_pedidomedic.t_codced.Text = Data1.Recordset("codced")
frm_pedidomedic.t_mat.Text = Data1.Recordset("matricula")
frm_pedidomedic.t_nombre.Text = Data1.Recordset("nombre")
frm_pedidomedic.labcodconv.Caption = Data1.Recordset("codconv")
If IsNull(Data1.Recordset("mutual")) = False Then
   frm_pedidomedic.labmutual.Caption = Data1.Recordset("mutual")
Else
   frm_pedidomedic.labmutual.Caption = ""
End If
If IsNull(Data1.Recordset("recetas_obs")) = False Then
   frm_pedidomedic.t_det.Text = Data1.Recordset("recetas_obs")
Else
   frm_pedidomedic.t_det.Text = ""
End If

frm_pedidomedic.t_direc.Text = Data1.Recordset("direcc")
frm_pedidomedic.cbozona.Text = Data1.Recordset("zona")
frm_pedidomedic.t_telfs.Text = Data1.Recordset("telefs")
If IsNull(Data1.Recordset("correo")) = False Then
   frm_pedidomedic.t_correo.Text = Data1.Recordset("correo")
Else
   frm_pedidomedic.t_correo.Text = ""
End If
If IsNull(Data1.Recordset("recibe1")) = False Then
   frm_pedidomedic.t_recib1.Text = Data1.Recordset("recibe1")
Else
   frm_pedidomedic.t_recib1.Text = ""
End If
If IsNull(Data1.Recordset("fec_ent")) = False Then
   frm_pedidomedic.mfent.Text = Format(Data1.Recordset("fec_ent"), "dd/mm/yyyy")
Else
   frm_pedidomedic.mfent.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("hor_ent1")) = False Then
   frm_pedidomedic.cboh1.Text = Data1.Recordset("hor_ent1")
Else
   frm_pedidomedic.cboh1.Text = "__:__"
End If
If IsNull(Data1.Recordset("hor_ent2")) = False Then
   frm_pedidomedic.cboh2.Text = Data1.Recordset("hor_ent2")
Else
   frm_pedidomedic.cboh2.Text = "__:__"
End If
If IsNull(Data1.Recordset("recetasen")) = False Then
   frm_pedidomedic.cborece.Text = Data1.Recordset("recetasen")
Else
   frm_pedidomedic.cborece.Text = ""
End If
If IsNull(Data1.Recordset("recetascont")) = False Then
   frm_pedidomedic.cborecctrol.Text = Data1.Recordset("recetascont")
Else
   frm_pedidomedic.cborecctrol.Text = ""
End If
frm_pedidomedic.cbofpago.Text = Data1.Recordset("forma_pago")
frm_pedidomedic.cboestado.Text = Data1.Recordset("estado")
If IsNull(Data1.Recordset("fec_estado")) = False Then
   frm_pedidomedic.mfestado.Text = Format(Data1.Recordset("fec_estado"), "dd/mm/yyyy")
Else
   frm_pedidomedic.mfestado.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("hor_estado")) = False Then
   frm_pedidomedic.mhestado.Text = Data1.Recordset("hor_estado")
Else
   frm_pedidomedic.mhestado.Text = "__:__"
End If
If IsNull(Data1.Recordset("usua_estado")) = False Then
   frm_pedidomedic.labusuario.Caption = Data1.Recordset("usua_estado")
Else
   frm_pedidomedic.labusuario.Caption = ""
End If
If IsNull(Data1.Recordset("nom_cadete")) = False Then
   frm_pedidomedic.t_cadete.Text = Data1.Recordset("nom_cadete")
Else
   frm_pedidomedic.t_cadete.Text = ""
End If
If IsNull(Data1.Recordset("nom_recibe")) = False Then
   frm_pedidomedic.t_recib2.Text = Data1.Recordset("nom_recibe")
Else
   frm_pedidomedic.t_recib2.Text = ""
End If
Dim Xcountt As Integer
Xcountt = 1
frm_pedidomedic.labtotcant.Caption = "0"
frm_pedidomedic.labtotp.Caption = "0"
Data_pedlin.RecordSource = "select * from pedidos_mediclin where cod_pedido =" & Data1.Recordset("pedido_nro")
Data_pedlin.Refresh
frm_pedidomedic.ListView1.ListItems.Clear
If Data_pedlin.Recordset.RecordCount > 0 Then
   Data_pedlin.Recordset.MoveFirst
   Do While Not Data_pedlin.Recordset.EOF
      frm_pedidomedic.ListView1.ListItems.Add Xcountt, , Data_pedlin.Recordset("nom_medic")
      frm_pedidomedic.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Data_pedlin.Recordset("cant")
      frm_pedidomedic.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Format(Data_pedlin.Recordset("tot_imp"), "Standard")
      frm_pedidomedic.labtotcant.Caption = Val(frm_pedidomedic.labtotcant.Caption) + Data_pedlin.Recordset("cant")
      frm_pedidomedic.labtotp.Caption = Val(frm_pedidomedic.labtotp.Caption) + Data_pedlin.Recordset("tot_imp")
      Data_pedlin.Recordset.MoveNext
      Xcountt = Xcountt + 1
   Loop
End If

Unload Me


End Sub

Private Sub Form_Load()

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data_pedlin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_pedidoimp.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_inf.DatabaseName = App.path & "\informes.mdb"

tabver.Refresh
'tabver.SelectedItem.Selected = True
If tabver.SelectedItem.index = 1 Then
   DBGrid1.BackColor = &H8080FF
   Data1.RecordSource = "Select * from pedidos_medic where estado ='" & "Pendiente" & "' order by fecha,hora"
   Data1.Refresh
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.BackColor = &H80C0FF
      Data1.RecordSource = "Select * from pedidos_medic where estado ='" & "En tránsito" & "' order by fecha,hora"
      Data1.Refresh
   Else
      If tabver.SelectedItem.index = 3 Then
         DBGrid1.BackColor = &H80FF80
         Data1.RecordSource = "Select * from pedidos_medic where estado in ('Entregado','Cancelado') order by fecha,hora"
         Data1.Refresh
      Else
         Data1.RecordSource = "Select * from pedidos_medic where estado in ('Entregado','Cancelado') order by fecha,hora"
         Data1.Refresh
      End If
   End If
End If

End Sub

Private Sub tabver_Click()
If tabver.SelectedItem.index = 1 Then
   DBGrid1.BackColor = &H8080FF
   Data1.RecordSource = "Select * from pedidos_medic where estado ='" & "Pendiente" & "' order by fecha,hora"
   Data1.Refresh
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.BackColor = &H80C0FF
      Data1.RecordSource = "Select * from pedidos_medic where estado ='" & "En tránsito" & "' order by fecha,hora"
      Data1.Refresh
   Else
      If tabver.SelectedItem.index = 3 Then
         DBGrid1.BackColor = &H80FF80
         Data1.RecordSource = "Select * from pedidos_medic where estado in ('Entregado','Cancelado') order by fecha,hora"
         Data1.Refresh
      Else
         Data1.RecordSource = "Select * from pedidos_medic where estado in ('Entregado','Cancelado') order by fecha,hora"
         Data1.Refresh
      End If
   End If
End If

End Sub

Public Function Devuelve_costoPed() As Integer

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from estudios where codest =" & 60110
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_costoPed = Xrecclii("cons")
   labdescpedido.Caption = Xrecclii("descrip")
Else
   Devuelve_costoPed = 0
   labdescpedido.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

