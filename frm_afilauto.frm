VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_afilauto 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Afiliaciones que requieren autorización"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12015
   Icon            =   "frm_afilauto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   12015
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2055
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Ingrese aquí el texto de la observación. Luego haga doble click sobre el cuadro para cerrar y grabar los datos."
      Top             =   3360
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingresar observación"
      Height          =   495
      Left            =   7560
      Picture         =   "frm_afilauto.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1815
   End
   Begin Crystal.CrystalReport cr2pant 
      Left            =   6240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr1print 
      Left            =   5280
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      DiscardSavedData=   -1  'True
      WindowState     =   1
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
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_afilauto.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Visualizar contrato"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Data data_hist 
      Caption         =   "data_hist"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_afilcons 
      Caption         =   "data_afilcons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_cnvmut 
      Caption         =   "data_cnvmut"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid ms2 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4260
      _Version        =   393216
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5530
      _Version        =   393216
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anular afiliación"
      Height          =   495
      Left            =   5160
      Picture         =   "frm_afilauto.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Autorizar afiliación"
      Height          =   495
      Left            =   9960
      Picture         =   "frm_afilauto.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Doble click para ver la afiliación del cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   4455
   End
End
Attribute VB_Name = "frm_afilauto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Autoriza As String

Autoriza = MsgBox("Desea autorizar la afiliación número: " & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " ?", vbInformation + vbYesNo, "Afiliaciones SAPP")

If Autoriza = vbYes Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
   data_afilcons.Refresh
   If data_afilcons.Recordset.RecordCount > 0 Then
      data_afilcons.Recordset.MoveFirst
      Do While Not data_afilcons.Recordset.EOF
         data_afilcons.Recordset.Edit
         data_afilcons.Recordset("pendiente") = 0
         data_afilcons.Recordset.Update
         data_afilcons.Recordset.MoveNext
      Loop
      data_hist.RecordSource = "select * from afiliaciones_impre"
      data_hist.Refresh
      data_hist.Recordset.AddNew
      data_hist.Recordset("fecha") = Date
      data_hist.Recordset("hora") = Format(Time, "HH:mm")
      data_hist.Recordset("usuario") = WElusuario
      data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
      data_hist.Recordset("nro_afilia") = Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
      data_hist.Recordset("accion") = "AUTORIZACION"
      data_hist.Recordset.Update
      MsgBox "La afiliación pasó a la opción de pendientes.", vbExclamation
      ms2.Clear
      DBGrid1.Clear
      Carga_grid
      
   End If
   
End If

End Sub

Private Sub Command2_Click()
Dim Autoriza As String

Autoriza = MsgBox("Desea ANULAR la afiliación número: " & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " ?", vbInformation + vbYesNo, "Afiliaciones SAPP")

If Autoriza = vbYes Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
   data_afilcons.Refresh
   If data_afilcons.Recordset.RecordCount > 0 Then
      data_afilcons.Recordset.MoveFirst
      Do While Not data_afilcons.Recordset.EOF
         data_afilcons.Recordset.Edit
         data_afilcons.Recordset("pendiente") = 20
         data_afilcons.Recordset("convenio") = "CANCELADO"
         data_afilcons.Recordset.Update
         data_afilcons.Recordset.MoveNext
      Loop
      data_hist.RecordSource = "select * from afiliaciones_impre"
      data_hist.Refresh
      data_hist.Recordset.AddNew
      data_hist.Recordset("fecha") = Date
      data_hist.Recordset("hora") = Format(Time, "HH:mm")
      data_hist.Recordset("usuario") = WElusuario
      data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
      data_hist.Recordset("nro_afilia") = Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
      data_hist.Recordset("accion") = "ANULACION"
      data_hist.Recordset.Update
      MsgBox "La afiliación quedó cancelada y no será visible para cargar al padrón.", vbExclamation
      ms2.Clear
      DBGrid1.Clear
      Carga_grid
      
   End If
   
End If

End Sub

Private Sub Command3_Click()
'Dim Xelnroafil As Long
'Xelnroafil = Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
'MsgBox "NRO:" & Xelnroafil

Dim ImprimeContra As String
ImprimeContra = ""
data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

If Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) > 0 Then
   ImprimeContra = MsgBox("Desea visualizar el contrato?", vbInformation + vbYesNo, "Afiliaciones SAPP")
    If ImprimeContra = vbYes Then
       data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
       data_afilcons.Refresh
       data_afilcons.Recordset.MoveFirst
       Do While Not data_afilcons.Recordset.EOF
          
          Genera_contrato
          data_afilcons.Recordset.MoveNext
       Loop
       data_inf.RecordSource = "select * from infcli order by cl_cantpag"
       data_inf.Refresh
       If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
          If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
             cr2pant.ReportFileName = App.path & "\contrato_debprint.rpt"
          Else
             cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
          End If
       Else
          cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
       End If
       cr2pant.Action = 1
    Else
       MsgBox "Solo puede visualizar el contrato"
    End If
Else
   MsgBox "Falta seleccionar"
   
End If


End Sub

Private Sub Command4_Click()
data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and integra_nro =" & Val(ms2.TextMatrix(ms2.RowSel, 0))
data_afilcons.Refresh
If data_afilcons.Recordset.RecordCount > 0 Then
   If IsNull(data_afilcons.Recordset("obs_noaut")) = False Then
      Text1.Visible = True
      Command4.Enabled = False
      Command1.Enabled = False
      Command2.Enabled = False
      If IsNull(data_afilcons.Recordset("obs_adm")) = False Then
         Text1.Text = data_afilcons.Recordset("obs_adm")
      End If
   Else
      MsgBox "El registro seleccionado está autorizado automáticamente. Seleccione otro integrante.", vbInformation
      
   End If
End If

End Sub

Private Sub DBGrid1_DblClick()
'Xnro = Val(DBGrid1.TextMatrix(flex1.RowSel, 1))
'Xnroh = Val(flex1.TextMatrix(flex1.RowSel, 3))
Dim Xsqlpromos As String
Dim Xreccliis As New ADODB.Recordset
Dim Xcann As Integer

ConectarBD
ConbdSapp.Open
             

Xsqlpromos = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.nom2,afiliaciones_new.ape1,afiliaciones_new.ape2,afiliaciones_new.pendiente," & _
"afiliaciones_new.obs_adm,afiliaciones_new.obs_noaut,afiliaciones_new.matricula,afiliaciones_new.fnac,afiliaciones_new.integra_nro,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.cedula,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
"from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " order by afiliaciones_new.integra_nro"
             
With Xreccliis
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromos, ConbdSapp, , , adCmdText
End With
ms2.Clear
ms2.rows = 2
ms2.Cols = 9
ms2.TextMatrix(0, 0) = "NRO."
ms2.ColWidth(0) = 800
ms2.TextMatrix(0, 1) = "CEDULA"
ms2.ColWidth(1) = 1200
ms2.TextMatrix(0, 2) = "NOMBRES"
ms2.ColWidth(2) = 2000
ms2.TextMatrix(0, 3) = "APELLIDOS"
ms2.ColWidth(3) = 2000
ms2.TextMatrix(0, 4) = "F.NAC."
ms2.ColWidth(4) = 1200
ms2.TextMatrix(0, 5) = "CELULAR"
ms2.ColWidth(5) = 1300
ms2.TextMatrix(0, 6) = "MATRICULA"
ms2.ColWidth(6) = 1300
ms2.TextMatrix(0, 7) = "MOTIVO"
ms2.ColWidth(7) = 3500
ms2.TextMatrix(0, 8) = "OBSERVACION ADM."
ms2.ColWidth(7) = 3500

Xcann = 1

If Xreccliis.RecordCount > 0 Then
   Xreccliis.MoveFirst
   Do While Not Xreccliis.EOF
      ms2.TextMatrix(Xcann, 0) = Xreccliis("integra_nro")
      ms2.TextMatrix(Xcann, 1) = Xreccliis("cedula")
      If IsNull(Xreccliis("nom2")) = False Then
         ms2.TextMatrix(Xcann, 2) = Xreccliis("nom1") & " " & Xreccliis("nom2")
      Else
         ms2.TextMatrix(Xcann, 2) = Xreccliis("nom1")
      End If
      If IsNull(Xreccliis("ape2")) = False Then
         ms2.TextMatrix(Xcann, 3) = Xreccliis("ape1") & " " & Xreccliis("ape2")
      Else
         ms2.TextMatrix(Xcann, 3) = Xreccliis("ape1")
      End If
      ms2.TextMatrix(Xcann, 4) = Xreccliis("fnac")
      ms2.TextMatrix(Xcann, 5) = Xreccliis("celular")
      If IsNull(Xreccliis("matricula")) = False Then
         ms2.TextMatrix(Xcann, 6) = Xreccliis("matricula")
      Else
         ms2.TextMatrix(Xcann, 6) = "0"
      End If
      If IsNull(Xreccliis("obs_noaut")) = False Then
         ms2.TextMatrix(Xcann, 7) = Xreccliis("obs_noaut")
      End If
      If IsNull(Xreccliis("obs_adm")) = False Then
         ms2.TextMatrix(Xcann, 8) = Xreccliis("obs_adm")
      End If
      
      ms2.rows = ms2.rows + 1
      Xreccliis.MoveNext
      Xcann = Xcann + 1
   Loop
End If

Xreccliis.Close
ConbdSapp.Close

End Sub

Private Sub Form_Load()
Carga_grid
data_cnvmut.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hist.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_afilcons.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Public Sub Carga_grid()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xcann As Integer

ConectarBD
ConbdSapp.Open
             
'Data1.ConnectionString = "dsn=sappnew"
Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente," & _
"afiliaciones_new.integra_nro,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
"from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (2) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha"
'Data1.Refresh
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
DBGrid1.rows = 2
DBGrid1.Cols = 7
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1300
DBGrid1.TextMatrix(0, 1) = "Nro.Af."
DBGrid1.ColWidth(1) = 1300
DBGrid1.TextMatrix(0, 2) = "NOMBRE"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "CATEGORIA"
DBGrid1.ColWidth(3) = 1500
DBGrid1.TextMatrix(0, 4) = "CELULAR"
DBGrid1.ColWidth(4) = 1500
DBGrid1.TextMatrix(0, 5) = "TELEFONO"
DBGrid1.ColWidth(5) = 1500
DBGrid1.TextMatrix(0, 6) = "PROMOTOR"
DBGrid1.ColWidth(6) = 1500

Xcann = 1

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      DBGrid1.TextMatrix(Xcann, 0) = Xrecclii("fecha")
      DBGrid1.TextMatrix(Xcann, 1) = Xrecclii("afilia_nro")
      DBGrid1.TextMatrix(Xcann, 2) = Xrecclii("nom1") & " " & Xrecclii("ape1")
      DBGrid1.TextMatrix(Xcann, 3) = Xrecclii("convenio")
      If IsNull(Xrecclii("celular")) = False Then
         DBGrid1.TextMatrix(Xcann, 4) = Xrecclii("celular")
      End If
      If IsNull(Xrecclii("telef")) = False Then
         DBGrid1.TextMatrix(Xcann, 5) = Xrecclii("telef")
      End If
      If IsNull(Xrecclii("nombre")) = False Then
         DBGrid1.TextMatrix(Xcann, 6) = Xrecclii("nombre")
      End If
      DBGrid1.rows = DBGrid1.rows + 1
      Xrecclii.MoveNext
      Xcann = Xcann + 1
   Loop
Else
   DBGrid1.Enabled = False
   ms2.Enabled = False
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Private Sub ms2_DblClick()
Dim Xmataf As Long
Dim Xgpoconv As String
Xgpoconv = ""

If Trim(ms2.TextMatrix(ms2.RowSel, 6)) <> "" Then
   
   Xmataf = Val(ms2.TextMatrix(ms2.RowSel, 6))
   If Xmataf > 0 Then
    frmabm.data_clientes.RecordSource = "select * from clientes where cl_codigo =" & Xmataf
    frmabm.data_clientes.Refresh
    frmabm.txt_ced.Text = frmabm.data_clientes.Recordset("cl_cedula")
    frmabm.txt_ced2.Text = frmabm.data_clientes.Recordset("cl_codced")
    If frmabm.data_clientes.Recordset("estado") = 2 Or frmabm.data_clientes.Recordset("estado") = 3 Then
       frmabm.labestado.Caption = "BAJA"
    Else
       If frmabm.data_clientes.Recordset("fecha_baja") <> "" Then
          frmabm.labestado.Caption = "BAJA"
       Else
          frmabm.labestado.Caption = "ACTIVO"
       End If
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_fultvta")) = False Then
       If IsNull(frmabm.data_clientes.Recordset("cl_tipocli")) = False Then
          If frmabm.data_clientes.Recordset("cl_tipocli") = 1 Or frmabm.data_clientes.Recordset("cl_tipocli") = 2 Then
             frmabm.Image1.Visible = True
          Else
             frmabm.Image1.Visible = False
          End If
       Else
          frmabm.Image1.Visible = False
       End If
    Else
       frmabm.Image1.Visible = False
    End If
    If frmabm.Image1.Visible = False Then
       If IsNull(frmabm.data_clientes.Recordset("cl_fultpag")) = False Then
          frmabm.Image1.Visible = True
       End If
    End If
    frmabm.txt_mat.Caption = frmabm.data_clientes.Recordset("cl_codigo")
    If IsNull(frmabm.data_clientes.Recordset("cl_codconv")) = True Then
       MsgBox "Verifique el convenio", vbCritical, "Mensaje"
       frmabm.txt_codcnv.Text = ""
       Xgpoconv = ""
    Else
       frmabm.txt_codcnv.Text = frmabm.data_clientes.Recordset("cl_codconv")
    End If
    data_cnvmut.RecordSource = "Select * from convenio where cnv_codigo ='" & frmabm.txt_codcnv.Text & "'"
    data_cnvmut.Refresh
    If data_cnvmut.Recordset.RecordCount > 0 Then
       If IsNull(data_cnvmut.Recordset("cnv_entre")) = False Then
          If Trim(data_cnvmut.Recordset("cnv_entre")) <> "" Then
             If Val(data_cnvmut.Recordset("cnv_cuenta")) = Val(frmabm.txt_mat.Caption) Then
                frmabm.t_rs.Text = data_cnvmut.Recordset("cnv_entre")
             Else
                frmabm.t_rs.Text = ""
             End If
          Else
             frmabm.t_rs.Text = ""
          End If
       Else
          frmabm.t_rs.Text = ""
       End If
       If IsNull(data_cnvmut.Recordset("cnv_grupo")) = False Then
          If Trim(data_cnvmut.Recordset("cnv_grupo")) <> "" Then
             Xgpoconv = data_cnvmut.Recordset("cnv_grupo")
          Else
             Xgpoconv = ""
          End If
       Else
          Xgpoconv = ""
       End If
    Else
       Xgpoconv = ""
    End If
        
    frmabm.txt_nomcnv.Enabled = True
    If IsNull(frmabm.data_clientes.Recordset("cl_nomconv")) = True Then
       frmabm.txt_nomcnv.Text = ""
    Else
       frmabm.txt_nomcnv.Text = frmabm.data_clientes.Recordset("cl_nomconv")
    End If
    frmabm.txt_nomcnv.Enabled = False
    If IsNull(frmabm.data_clientes.Recordset("cl_apellid")) = False Then
       frmabm.txt_apellid.Text = frmabm.data_clientes.Recordset("cl_apellid")
    Else
       frmabm.txt_apellid.Text = "NN"
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_ruc")) = False Then
       frmabm.t_otrocnv.Text = frmabm.data_clientes.Recordset("cl_ruc")
    Else
       frmabm.t_otrocnv.Text = ""
    End If
     If IsNull(frmabm.data_clientes.Recordset("cl_cedula")) = True Then
        frmabm.txt_ced.Text = 0
     Else
        frmabm.txt_ced.Text = frmabm.data_clientes.Recordset("cl_cedula")
     End If
     If frmabm.data_clientes.Recordset("cl_codced") <> "" Then
        frmabm.txt_ced2.Text = frmabm.data_clientes.Recordset("cl_codced")
     Else
        frmabm.txt_ced2.Text = 0
     End If
     If IsNull(frmabm.data_clientes.Recordset("cl_fnac")) = False Then
        frmabm.txt_nac.Text = Format(frmabm.data_clientes.Recordset("cl_fnac"), "dd/mm/yyyy")
     Else
        frmabm.txt_nac.Text = "__/__/____"
        frmabm.labedad.Caption = 0
        frmabm.labunie.Caption = 0
        frmabm.labdias.Caption = 0
     End If
     If Not IsDate(frmabm.txt_nac.Text) Then
        '   MsgBox "Digite una fecha válida"
        
     Else
        CalculaEdad (frmabm.txt_nac.Text)
     End If
    If IsNull(frmabm.data_clientes.Recordset("cl_codruta")) = True Then
       frmabm.t_ruta.Text = ""
    Else
       frmabm.t_ruta.Text = frmabm.data_clientes.Recordset("cl_codruta")
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_decuota")) = False Then
       If frmabm.data_clientes.Recordset("cl_decuota") = 1 Then
          frmabm.Option1.Value = True
       Else
          If frmabm.data_clientes.Recordset("cl_decuota") = 2 Then
             frmabm.Option2.Value = True
          Else
             If frmabm.data_clientes.Recordset("cl_decuota") = 3 Then
                frmabm.Option3.Value = True
             Else
                If frmabm.data_clientes.Recordset("cl_decuota") = 4 Then
                   frmabm.Option4.Value = True
                Else
                   If frmabm.data_clientes.Recordset("cl_decuota") = 5 Then
                      frmabm.Option5.Value = True
                   Else
                      frmabm.Option1.Value = False
                      frmabm.Option2.Value = False
                      frmabm.Option3.Value = False
                      frmabm.Option4.Value = False
                      frmabm.Option5.Value = False
                   End If
                End If
             End If
          End If
       End If
    Else
       frmabm.Option1.Value = False
       frmabm.Option2.Value = False
       frmabm.Option3.Value = False
       frmabm.Option4.Value = False
       frmabm.Option5.Value = False
    End If
    If IsNull(frmabm.data_clientes.Recordset("fecha_reac")) = False Then
       frmabm.mfcarta.Text = Format(frmabm.data_clientes.Recordset("fecha_reac"), "dd/mm/yyyy")
    Else
       frmabm.mfcarta.Text = "__/__/____"
    End If
    If IsNull(frmabm.data_clientes.Recordset("saldo_chc2")) = False Then
       frmabm.cbosrv.ListIndex = frmabm.data_clientes.Recordset("saldo_chc2")
    Else
       frmabm.cbosrv.ListIndex = -1
    End If
    If frmabm.data_clientes.Recordset("cl_ultmesp") <> "" Then
       frmabm.labump.Caption = frmabm.data_clientes.Recordset("cl_ultmesp")
    Else
       frmabm.labump.Caption = ""
    End If
    If frmabm.data_clientes.Recordset("cl_ultanop") <> "" Then
       If frmabm.data_clientes.Recordset("cl_ultanop") = 0 Then
          frmabm.labuap.Caption = frmabm.data_clientes.Recordset("cl_ultanop")
       Else
          frmabm.labuap.Caption = "/" + str(frmabm.data_clientes.Recordset("cl_ultanop"))
       End If
    Else
       frmabm.labuap.Caption = ""
    End If
    If frmabm.data_clientes.Recordset("cl_atrasoa") <> "" Then
       frmabm.labatra.Caption = frmabm.data_clientes.Recordset("cl_atrasoa")
    Else
       frmabm.labatra.Caption = ""
    End If
    If frmabm.data_clientes.Recordset("saldo_cc") <> "" Then
       frmabm.labdeudap.Caption = frmabm.data_clientes.Recordset("saldo_cc")
    Else
       frmabm.labdeudap.Caption = ""
    End If
    If frmabm.data_clientes.Recordset("cl_direcci") <> "" Then
       frmabm.txt_direcc1.Text = frmabm.data_clientes.Recordset("cl_direcci")
    Else
       frmabm.txt_direcc1.Text = ""
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_dpto")) = False Then
       frmabm.t_cel.Text = frmabm.data_clientes.Recordset("cl_dpto")
    Else
       frmabm.t_cel.Text = ""
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_referen")) = False Then
       frmabm.t_correo.Text = frmabm.data_clientes.Recordset("cl_referen")
    Else
       frmabm.t_correo.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_entre") <> "" Then
       frmabm.txt_direcc2.Text = frmabm.data_clientes.Recordset("cl_entre")
    Else
       frmabm.txt_direcc2.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_grupo") <> "" Then
       frmabm.txt_codzon.Text = frmabm.data_clientes.Recordset("cl_grupo")
    Else
       frmabm.txt_codzon.Text = 0
    End If
    If frmabm.data_clientes.Recordset("cl_zona") <> "" Then
       frmabm.cbolocalid.Text = frmabm.data_clientes.Recordset("cl_zona")
    Else
       frmabm.cbolocalid.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_sexo") = 2 Then
       frmabm.cbosexo.Text = "FEMENINO"
    Else
       frmabm.cbosexo.Text = "MASCULINO"
    End If
    If frmabm.data_clientes.Recordset("cl_telefon") <> "" Then
       frmabm.txt_telef.Text = frmabm.data_clientes.Recordset("cl_telefon")
    Else
       frmabm.txt_telef.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_dircobr") <> "" Then
       frmabm.txt_dircob.Text = frmabm.data_clientes.Recordset("cl_dircobr")
    Else
       frmabm.txt_dircob.Text = ""
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_nombre")) = False Then
       frmabm.txt_conmut.Text = frmabm.data_clientes.Recordset("cl_nombre")
    Else
       frmabm.txt_conmut.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_socmnom") <> "" Then
       frmabm.cbomutual.Text = frmabm.data_clientes.Recordset("cl_socmnom")
    Else
       frmabm.cbomutual.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_nrosocm") <> "" Then
       frmabm.txt_matmut.Text = frmabm.data_clientes.Recordset("cl_nrosocm")
    Else
       frmabm.txt_matmut.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_fecing") <> "" Then
       frmabm.txt_fecing.Text = Format(frmabm.data_clientes.Recordset("cl_fecing"), "dd/mm/yyyy")
    Else
       frmabm.txt_fecing.Text = "__/__/____"
    End If
    If frmabm.data_clientes.Recordset("fecha_baja") <> "" Then
       frmabm.txt_fecbaj.Text = Format(frmabm.data_clientes.Recordset("fecha_baja"), "dd/mm/yyyy")
    Else
       frmabm.txt_fecbaj.Text = "__/__/____"
    End If
    If IsNull(frmabm.data_clientes.Recordset("idpromos")) = False Then
       frmabm.labidpromo.Caption = frmabm.data_clientes.Recordset("idpromos")
       If Val(frmabm.labidpromo.Caption) > 0 Then
          BuscaPromosId
       Else
          frmabm.cbopromos.Text = ""
       End If
    Else
       frmabm.labidpromo.Caption = 0
       frmabm.cbopromos.Text = ""
    End If
    If IsNull(frmabm.data_clientes.Recordset("mesproxemi")) = False Then
       frmabm.t_pmemi.Text = frmabm.data_clientes.Recordset("mesproxemi")
       frmabm.t_paemi.Text = frmabm.data_clientes.Recordset("anoproxemi")
    Else
       frmabm.t_pmemi.Text = 0
       frmabm.t_paemi.Text = 0
    End If
    
    If frmabm.data_clientes.Recordset("cl_nrovend") <> "" Then
       frmabm.txt_codpro.Text = frmabm.data_clientes.Recordset("cl_nrovend")
    Else
       frmabm.txt_codpro.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_nomvend") <> "" Then
       frmabm.cbonompro.Text = frmabm.data_clientes.Recordset("cl_nomvend")
    Else
       frmabm.cbonompro.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_nrocobr") <> "" Then
       frmabm.txt_codcob.Text = frmabm.data_clientes.Recordset("cl_nrocobr")
    Else
       frmabm.txt_codcob.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_nomcobr") <> "" Then
       frmabm.cbonomcob.Text = frmabm.data_clientes.Recordset("cl_nomcobr")
    Else
       frmabm.cbonomcob.Text = ""
    End If
    Veoladeuda (frmabm.data_clientes.Recordset("cl_codigo"))
              
    If IsNull(frmabm.data_clientes.Recordset("cl_descpag")) = True Then
       frmabm.cbopago.Text = "Abono Mensual"
    Else
       If UCase(frmabm.data_clientes.Recordset("cl_descpag")) = "DEBITO AUTOMATICO" Then
          frmabm.cbopago.Text = "Debito Automatico"
       Else
          frmabm.cbopago.Text = "Abono Mensual"
       End If
    End If
    If frmabm.data_clientes.Recordset("cl_diacobr") <> "" Then
       frmabm.txt_diacob.Text = frmabm.data_clientes.Recordset("cl_diacobr")
    Else
       frmabm.txt_diacob.Text = ""
    End If
    If frmabm.data_clientes.Recordset("tit_tarj") <> "" Then
       frmabm.txt_nomtarj.Text = frmabm.data_clientes.Recordset("tit_tarj")
    Else
       frmabm.txt_nomtarj.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_nrotarj") <> "" Then
       frmabm.txt_nrotarj.Text = frmabm.data_clientes.Recordset("cl_nrotarj")
    Else
       frmabm.txt_nrotarj.Text = ""
    End If
    If frmabm.data_clientes.Recordset("ci_tarj") <> "" Then
       frmabm.txt_cedtarj.Text = frmabm.data_clientes.Recordset("ci_tarj")
    Else
       frmabm.txt_cedtarj.Text = ""
    End If
    If frmabm.data_clientes.Recordset("codcitarj") <> "" Then
       frmabm.txt_codtarj.Text = frmabm.data_clientes.Recordset("codcitarj")
    Else
       frmabm.txt_codtarj.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_tjemi_c") <> "" Then
       frmabm.txt_codemisor.Text = frmabm.data_clientes.Recordset("cl_tjemi_c")
    Else
       frmabm.txt_codemisor.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_tjemi_n") <> "" Then
       frmabm.cbotarj.Text = frmabm.data_clientes.Recordset("cl_tjemi_n")
    Else
       frmabm.cbotarj.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_tj_venc") <> "" Then
       frmabm.txt_vence.Text = Format(frmabm.data_clientes.Recordset("cl_tj_venc"), "dd/mm/yyyy")
    Else
       frmabm.txt_vence.Text = "__/__/____"
    End If
    frmabm.labmr.Caption = ""
    '     Veoladeuda (DBGrid1.TextMatrix(DBGrid1.RowSel, 0))
    If frmabm.cbopromos.Text = "Grupo de 3 o más" Then
       VerPromoCliNew
    Else
       VerPromocion (ms2.TextMatrix(ms2.RowSel, 6))
    End If
    frm_afilauto.Hide
   Else
      MsgBox "Cliente sin ficha, no hay datos para visualizar.", vbExclamation
   End If
Else
    MsgBox "Cliente sin ficha, no hay datos para visualizar.", vbExclamation
End If

'Unload Me

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

Anios = DateDiff("yyyy", CDate(Format(FNaci, "dd/MM/yyyy")), CDate(FAct))
If Month(CDate(FAct)) < Month(CDate(FNaci)) Then
 Anios = Anios - 1
 newmonth = Month(CDate(FAct)) + 12
 Else
 newmonth = Month(CDate(FAct))
 End If
If Month(CDate(FAct)) = Month(CDate(FNaci)) Then
 If Day(CDate(FAct)) < Day(CDate(FNaci)) Then
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

   frmabm.labedad.Caption = Anios
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
   frmabm.labunie.Caption = Meses
   frmabm.labdias.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   frmabm.labedad.Caption = 0
   frmabm.labunie.Caption = 0
   frmabm.labdias.Caption = 0
End If

End Sub
Public Sub Veoladeuda(ByVal Xmatricula As Long)

Dim Xsubt As Double
Dim Xcant As Long
Dim Xmes, Xano As Integer

Xcant = 0
Xsubt = 0
Xmes = 0

ConectarBD
ConbdSapp.Open
Sqlconsulta = "Select * from deudas where cliente =" & Xmatricula & " and fecha_pago is null order by ano,mes"
With Registro1
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open Sqlconsulta, ConbdSapp, , , adCmdText
End With

If Registro1.RecordCount > 0 Then
   Registro1.MoveFirst
   Do While Not Registro1.EOF
      If Registro1("mes") = 0 Then
         Xsubt = Xsubt + Registro1("total")
      Else
         Xsubt = Xsubt + Registro1("total")
         If Xmes = 0 Then
            Xmes = Registro1("mes")
            Xano = Registro1("ano")
         End If
         Xcant = Xcant + 1
      End If
      Registro1.MoveNext
   Loop
   If Xmes = 0 Then
   Else
      If Xmes = 1 Then
         Xano = Xano - 1
         Xmes = 12
      Else
         Xmes = Xmes - 1
      End If
   End If
   frmabm.labump.Caption = Xmes
   frmabm.labuap.Caption = Xano
   frmabm.labatra.Caption = Xcant
   frmabm.labdeudap.Caption = Format(Xsubt, "0.00")
Else
   frmabm.labump.Caption = Month(Date)
   frmabm.labuap.Caption = Year(Date)
   frmabm.labatra.Caption = 0
   frmabm.labdeudap.Caption = 0
End If
Registro1.Close
ConbdSapp.Close


End Sub

Public Sub VerPromocion(ByVal Xmatricula As Long)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & Xmatricula
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   frmabm.Label42.Caption = "Tiene promo X" & Xrecclii.RecordCount
   frmabm.t_ruta.Enabled = False
Else
   frmabm.Label42.Caption = ""
   frmabm.t_ruta.Enabled = True
End If
Xrecclii.Close
ConbdSapp.Close


End Sub
Public Sub BuscaPromosId()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from promocion_gpo where id =" & Val(frmabm.labidpromo.Caption)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   frmabm.cbopromos.Text = Xrecclii("descrip")
Else
   MsgBox "No se encuentra promoción. Verifique!", vbCritical
   frmabm.cbopromos.Text = ""
   frmabm.labidpromo.Caption = 0
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub VerPromoCliNew()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim cedruta As String
If frmabm.txt_ced.Text <> "" Then
   cedruta = Trim(frmabm.txt_ced.Text) & Trim(frmabm.txt_ced2.Text)
Else
   cedruta = "0"
End If
ConectarBD
ConbdSapp.Open
If frmabm.t_ruta.Text = "" Then
   Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & Val(cedruta)
Else
   Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & frmabm.t_ruta.Text
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   frmabm.Label42.Caption = "Grupo: " & Xrecclii.RecordCount + 1
   frmabm.t_ruta.Enabled = False
Else
   frmabm.Label42.Caption = ""
   frmabm.t_ruta.Enabled = True
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Public Sub Genera_contrato()
Dim Direc, Xcontrato As String
Direc = ""

If data_afilcons.Recordset("convenio") = "COMPLEMENTO" Or data_afilcons.Recordset("convenio") = "COMPLEMENTO C.GALICIA" Or data_afilcons.Recordset("convenio") = "AMBULATORIO" Then
   data_inf.DatabaseName = ""
   data_inf.Connect = "odbc;dsn=sappnew;"
   data_inf.RecordSource = "afiliaciones_contratoa"
   data_inf.Refresh
   If IsNull(data_inf.Recordset("descrip")) = False Then
      Xcontrato = data_inf.Recordset("descrip")
   Else
      Xcontrato = "Sin datos"
   End If
   data_inf.Connect = ""

Else
   data_inf.DatabaseName = ""
   data_inf.Connect = "odbc;dsn=sappnew;"
   data_inf.RecordSource = "afiliaciones_contratoe"
   data_inf.Refresh
   If IsNull(data_inf.Recordset("descrip")) = False Then
      Xcontrato = data_inf.Recordset("descrip")
   Else
      Xcontrato = "Sin datos"
   End If
   data_inf.Connect = ""
End If

data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh

If data_afilcons.Recordset.RecordCount > 0 Then
   data_inf.Recordset.AddNew
   data_inf.Recordset("cl_codigo") = data_afilcons.Recordset("afilia_nro")
   If IsNull(data_afilcons.Recordset("catcontrato")) = False Then
      data_inf.Recordset("cl_descpag") = data_afilcons.Recordset("catcontrato")
   Else
      If data_afilcons.Recordset("convenio") = "COMPLEMENTO" Or data_afilcons.Recordset("convenio") = "COMPLEMENTO C.GALICIA" Then
         data_inf.Recordset("cl_descpag") = "AMBULATORIO"
      Else
         data_inf.Recordset("cl_descpag") = data_afilcons.Recordset("convenio")
      End If
   End If
   data_inf.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
   data_inf.Recordset("cl_cantpag") = data_afilcons.Recordset("integra_nro")
   data_inf.Recordset("cl_apellid") = data_afilcons.Recordset("ape1")
   data_inf.Recordset("cl_medflia") = Mid(Devuelve_titular(), 1, 30)
   data_inf.Recordset("tit_tarj") = Mid(Devuelve_titularApe(), 1, 30)
   If IsNull(data_afilcons.Recordset("ape2")) = False Then
      data_inf.Recordset("cl_localid") = Mid(data_afilcons.Recordset("ape2"), 1, 35)
   End If
   data_inf.Recordset("cl_nomvend") = Mid(data_afilcons.Recordset("nom1"), 1, 35)
   If IsNull(data_afilcons.Recordset("nom2")) = False Then
      data_inf.Recordset("cl_nombre") = Mid(data_afilcons.Recordset("nom2"), 1, 30)
   End If
   data_inf.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
   If Len(data_afilcons.Recordset("cedula")) = 7 Then
      data_inf.Recordset("cl_fax") = Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6) & "-" & Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1)
   Else
      data_inf.Recordset("cl_fax") = Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7) & "-" & Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1)
   End If
   If IsNull(data_afilcons.Recordset("telef")) = False Then
      data_inf.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
   End If
   data_inf.Recordset("cl_celular") = data_afilcons.Recordset("celular")
   If IsNull(data_afilcons.Recordset("correo")) = False Then
      data_inf.Recordset("cl_dircobr") = data_afilcons.Recordset("correo")
   End If
   If IsNull(data_afilcons.Recordset("codmut")) = False Then
      data_inf.Recordset("cl_socmnom") = Devuelve_mut()
   End If
   If IsNull(data_afilcons.Recordset("direc2")) = False Then
      Direc = data_afilcons.Recordset("direc1") & " E/" & data_afilcons.Recordset("direc2")
   Else
      Direc = data_afilcons.Recordset("direc1")
   End If
   data_inf.Recordset("cl_direcci") = Mid(Direc, 1, 80)
   If IsNull(data_afilcons.Recordset("manz")) = False Then
      data_inf.Recordset("cl_estadoc") = data_afilcons.Recordset("manz")
   End If
   If IsNull(data_afilcons.Recordset("solar")) = False Then
      data_inf.Recordset("cl_tipcli") = Mid(data_afilcons.Recordset("solar"), 1, 3)
   End If
   If IsNull(data_afilcons.Recordset("nomzona")) = False Then
      data_inf.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
   End If
   data_inf.Recordset("cl_atrasoa") = data_afilcons.Recordset("valorcuota")
   If IsNull(data_afilcons.Recordset("desc_imp")) = False Then
      data_inf.Recordset("cl_seg_vto") = data_afilcons.Recordset("desc_imp")
   Else
      data_inf.Recordset("cl_seg_vto") = 0
   End If
   If IsNull(data_afilcons.Recordset("importe_fin")) = False Then
      data_inf.Recordset("cl_ter_vto") = data_afilcons.Recordset("importe_fin")
   Else
      data_inf.Recordset("cl_ter_vto") = 0
   End If
   
   If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
      data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta"
      data_inf.Recordset("info_debit") = "COBRO POR DÉBITO AUTOMÁTICO:" & chr(13)
      data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "Se adjunta autorización débito automático al final del contrato."
      If IsNull(data_afilcons.Recordset("codvende")) = False Then
         data_inf.Recordset("cl_entre") = Devuelve_vende()
      Else
         data_inf.Recordset("cl_entre") = "Sin promotor"
      End If
      If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
         data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
      End If
      data_inf.Recordset("cl_tipclin") = data_afilcons.Recordset("tarj_sello")
      data_inf.Recordset("cl_email") = Mid(data_afilcons.Recordset("tarj_titular"), 1, 30)
      data_inf.Recordset("cl_nrovend") = data_afilcons.Recordset("tarj_cedtit")
      data_inf.Recordset("cl_forpago") = data_afilcons.Recordset("tarj_codced")
      data_inf.Recordset("cl_nomconv") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 30)
      data_inf.Recordset("cl_nomcobr") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 25)
      data_inf.Recordset("cl_nrotarj") = Mid(data_afilcons.Recordset("tarj_nro"), 1, 20)
      data_inf.Recordset("cl_ultmesp") = data_afilcons.Recordset("tarj_vencmes")
      data_inf.Recordset("cl_ultanop") = data_afilcons.Recordset("tarj_vencanio")
   Else
      If IsNull(data_afilcons.Recordset("debito_brou")) = False Then
         data_inf.Recordset("cl_nom_sup") = "Débito BROU"
         data_inf.Recordset("info_debit") = "CONFIRMA QUE REALIZÓ FORMULARIO PARA DÉBITO BROU?:--->SI" & chr(13)
         data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "NOMBRE DE TITULAR DE LA CUENTA:" & data_afilcons.Recordset("tarj_titular")
         If IsNull(data_afilcons.Recordset("codvende")) = False Then
            data_inf.Recordset("cl_entre") = Devuelve_vende()
         Else
            data_inf.Recordset("cl_entre") = "Sin promotor"
         End If
         If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
            data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
         End If
      Else
         data_inf.Recordset("cl_nom_sup") = "Cobrador a domicilio"
         data_inf.Recordset("info_debit") = "DOMICILIO DE COBRO:" & chr(13)
         If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
            data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & data_afilcons.Recordset("direc_cobro") & chr(13)
            If IsNull(data_afilcons.Recordset("zonacobro")) = False Then
               data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "ZONA: " & data_afilcons.Recordset("zonacobro")
               If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: " & data_afilcons.Recordset("dia_cobro") & " c/mes" & chr(13)
               Else
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: sin datos." & chr(13)
               End If
            Else
               If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: " & data_afilcons.Recordset("dia_cobro") & " c/mes" & chr(13)
               Else
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: sin datos." & chr(13)
               End If
            End If
         Else
            data_inf.Recordset("info_debit") = "Misma dirección." & chr(13)
         End If
         If IsNull(data_afilcons.Recordset("codvende")) = False Then
            data_inf.Recordset("cl_entre") = Devuelve_vende()
         Else
            data_inf.Recordset("cl_entre") = "Sin promotor"
         End If
         If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
            data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
         End If
         
      End If
   End If
   data_inf.Recordset("obsp") = Xcontrato
   data_inf.Recordset.Update
Else
   MsgBox "No hay datos de afiliación para imprimir. Verifique!", vbCritical
   
End If

End Sub

Public Function Devuelve_titular() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_titular = Xrecclii("nom1")
Else
   Devuelve_titular = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function
Public Function Devuelve_titularApe() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_titularApe = Xrecclii("ape1")
Else
   Devuelve_titularApe = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function

Public Function Devuelve_mut() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm where id =" & data_afilcons.Recordset("codmut")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_mut = Xrecclii("ca_nom")
Else
   Devuelve_mut = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

Public Function Devuelve_vende() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from vende_func where idfunc =" & data_afilcons.Recordset("codvende")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_vende = Xrecclii("nombre")
Else
   Devuelve_vende = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function


Private Sub Text1_DblClick()
Dim Xsqlpromos As String
Dim Xreccliis As New ADODB.Recordset
Dim Xcann As Integer

If Text1.Visible = True Then
   If Trim(Text1.Text) <> "" Then
      data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and integra_nro =" & Val(ms2.TextMatrix(ms2.RowSel, 0))
      data_afilcons.Refresh
      If data_afilcons.Recordset.RecordCount > 0 Then
         If IsNull(data_afilcons.Recordset("obs_adm")) = False Then
            If data_afilcons.Recordset("obs_adm") <> Trim(Text1.Text) Then
               data_afilcons.Recordset.Edit
               data_afilcons.Recordset("obs_adm") = Text1.Text
               data_afilcons.Recordset("fec_obsadm") = Date
               data_afilcons.Recordset.Update
            End If
         Else
            data_afilcons.Recordset.Edit
            data_afilcons.Recordset("obs_adm") = Text1.Text
            data_afilcons.Recordset("fec_obsadm") = Date
            data_afilcons.Recordset.Update
         End If
         ms2.Clear
         ConectarBD
         ConbdSapp.Open
         Xsqlpromos = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.nom2,afiliaciones_new.ape1,afiliaciones_new.ape2,afiliaciones_new.pendiente," & _
         "afiliaciones_new.obs_adm,afiliaciones_new.obs_noaut,afiliaciones_new.matricula,afiliaciones_new.fnac,afiliaciones_new.integra_nro,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.cedula,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
         "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " order by afiliaciones_new.integra_nro"
                     
         With Xreccliis
             .CursorLocation = adUseClient
             .CursorType = adOpenKeyset
             .LockType = adLockOptimistic
             .Open Xsqlpromos, ConbdSapp, , , adCmdText
         End With
         ms2.Clear
         ms2.rows = 2
         ms2.Cols = 9
         ms2.TextMatrix(0, 0) = "NRO."
         ms2.ColWidth(0) = 800
         ms2.TextMatrix(0, 1) = "CEDULA"
         ms2.ColWidth(1) = 1200
         ms2.TextMatrix(0, 2) = "NOMBRES"
         ms2.ColWidth(2) = 2000
         ms2.TextMatrix(0, 3) = "APELLIDOS"
         ms2.ColWidth(3) = 2000
         ms2.TextMatrix(0, 4) = "F.NAC."
         ms2.ColWidth(4) = 1200
         ms2.TextMatrix(0, 5) = "CELULAR"
         ms2.ColWidth(5) = 1300
         ms2.TextMatrix(0, 6) = "MATRICULA"
         ms2.ColWidth(6) = 1300
         ms2.TextMatrix(0, 7) = "MOTIVO"
         ms2.ColWidth(7) = 3500
         ms2.TextMatrix(0, 8) = "OBSERVACION ADM."
         ms2.ColWidth(7) = 3500
         Xcann = 1
         If Xreccliis.RecordCount > 0 Then
            Xreccliis.MoveFirst
            Do While Not Xreccliis.EOF
               ms2.TextMatrix(Xcann, 0) = Xreccliis("integra_nro")
               ms2.TextMatrix(Xcann, 1) = Xreccliis("cedula")
               If IsNull(Xreccliis("nom2")) = False Then
                  ms2.TextMatrix(Xcann, 2) = Xreccliis("nom1") & " " & Xreccliis("nom2")
               Else
                  ms2.TextMatrix(Xcann, 2) = Xreccliis("nom1")
               End If
               If IsNull(Xreccliis("ape2")) = False Then
                  ms2.TextMatrix(Xcann, 3) = Xreccliis("ape1") & " " & Xreccliis("ape2")
               Else
                  ms2.TextMatrix(Xcann, 3) = Xreccliis("ape1")
               End If
               ms2.TextMatrix(Xcann, 4) = Xreccliis("fnac")
               ms2.TextMatrix(Xcann, 5) = Xreccliis("celular")
               If IsNull(Xreccliis("matricula")) = False Then
                  ms2.TextMatrix(Xcann, 6) = Xreccliis("matricula")
               Else
                  ms2.TextMatrix(Xcann, 6) = "0"
               End If
               If IsNull(Xreccliis("obs_noaut")) = False Then
                  ms2.TextMatrix(Xcann, 7) = Xreccliis("obs_noaut")
               End If
               If IsNull(Xreccliis("obs_adm")) = False Then
                  ms2.TextMatrix(Xcann, 8) = Xreccliis("obs_adm")
               End If
               ms2.rows = ms2.rows + 1
               Xreccliis.MoveNext
               Xcann = Xcann + 1
            Loop
         End If
         Xreccliis.Close
         ConbdSapp.Close
      Else
         MsgBox "No se encuentra integrante."
      End If
   End If
   Text1.Text = ""
   Text1.Visible = False
   Command4.Enabled = True
   Command1.Enabled = True
   Command2.Enabled = True

End If

End Sub

