VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_afilpend 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Afiliaciones pendientes de modificación por padrón social"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10965
   Icon            =   "frm_afilpend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10965
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_convenios 
      Caption         =   "data_convenios"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data data_abm 
      Caption         =   "data_abm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_nrosoc 
      Caption         =   "data_nrosoc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_clicons 
      Caption         =   "data_clicons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar pago"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C00000&
      Caption         =   "Ver afiliaciones Anuladas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   0
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "Reemplazar datos en ficha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Data data_linmm 
      Caption         =   "data_linmm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_hist 
      Caption         =   "data_hist"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton b_pagos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturar"
      Height          =   495
      Left            =   5280
      Picture         =   "frm_afilpend.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Procesar pago de la afiliación"
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Afiliación procesada"
      Height          =   495
      Left            =   8520
      Picture         =   "frm_afilpend.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Procesar la afiliación como cargada al padrón"
      Top             =   5640
      Width           =   2175
   End
   Begin Crystal.CrystalReport cr2pant 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowControls  =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
   End
   Begin VB.Data data_afilcons 
      Caption         =   "data_afilcons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Crystal.CrystalReport cr1print 
      Left            =   3120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowPrintBtn=   0   'False
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_afilpend.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Visualizar formulario de afiliación."
      Top             =   5880
      Width           =   615
   End
   Begin VB.Data data_cnvmut 
      Caption         =   "data_cnvmut"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid ms2 
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3836
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   10455
      _ExtentX        =   18441
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
   Begin VB.Label labcedtit 
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label labpermiso 
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click para editar en la ficha del socio"
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
      TabIndex        =   0
      Top             =   5640
      Width           =   4815
   End
End
Attribute VB_Name = "frm_afilpend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_pagos_Click()
b_pagos.Enabled = False
data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
data_afilcons.Refresh
If data_afilcons.Recordset.RecordCount > 0 Then
   data_afilcons.Recordset.MoveFirst
   Do While Not data_afilcons.Recordset.EOF
      If IsNull(data_afilcons.Recordset("matricula")) = True Then
         If Len(Trim(data_afilcons.Recordset("cedula"))) = 7 Then
            data_cli.RecordSource = "select * from clientes where cl_cedula =" & Val(Mid(Trim(data_afilcons.Recordset("cedula")), 1, 6))
         Else
            data_cli.RecordSource = "select * from clientes where cl_cedula =" & Val(Mid(Trim(data_afilcons.Recordset("cedula")), 1, 7))
         End If
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            data_afilcons.Recordset.Edit
            data_afilcons.Recordset("matricula") = data_cli.Recordset("cl_codigo")
            data_afilcons.Recordset.Update
         End If
      End If
      Alta_Modif
      data_afilcons.Recordset.MoveNext
   Loop
End If

Xdeb = 12

frm_afilfactura.Show vbModal
Xdeb = 0

End Sub

Private Sub Check2_Click()
Carga_grid

End Sub

Private Sub Command1_Click()
Dim Terminada As String
Dim Terminado2 As String
Terminada = ""

Terminada = MsgBox("Desea procesar la afiliación como VERIFICADA?", vbInformation + vbYesNo, "Afiliaciones SAPP")
If Terminada = vbYes Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and procesa_mod is null"
   data_afilcons.Refresh
   If data_afilcons.Recordset.RecordCount > 0 Then
      MsgBox "Faltan procesar registros de la afiliación al padrón.", vbCritical
      Terminado2 = MsgBox("Desea procesar como VERIFICADA?", vbInformation + vbYesNo, "Afiliaciones")
      If Terminado2 = vbYes Then
        data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
        data_afilcons.Refresh
        If data_afilcons.Recordset.RecordCount > 0 Then
           data_afilcons.Recordset.MoveFirst
           Do While Not data_afilcons.Recordset.EOF
              data_afilcons.Recordset.Edit
              data_afilcons.Recordset("pendiente") = 10
              data_afilcons.Recordset.Update
              data_afilcons.Recordset.MoveNext
           Loop
           MsgBox "Proceso terminado"
           DBGrid1.Clear
           Carga_grid
        End If
      End If
   Else
      data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
      data_afilcons.Refresh
      If data_afilcons.Recordset.RecordCount > 0 Then
         data_afilcons.Recordset.MoveFirst
         Do While Not data_afilcons.Recordset.EOF
            data_afilcons.Recordset.Edit
            data_afilcons.Recordset("pendiente") = 10
            data_afilcons.Recordset.Update
            data_afilcons.Recordset.MoveNext
         Loop
         MsgBox "Proceso terminado"
         DBGrid1.Clear
         Carga_grid
      End If
   End If
End If

End Sub

Private Sub Command2_Click()
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
   ImprimeContra = MsgBox("Desea imprimir contrato?", vbInformation + vbYesNo, "Afiliaciones SAPP")
    If ImprimeContra = vbYes Then
       data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
       data_afilcons.Refresh
       data_afilcons.Recordset.MoveFirst
       Do While Not data_afilcons.Recordset.EOF
          
          Genera_contrato
          data_afilcons.Recordset.Edit
          If IsNull(data_afilcons.Recordset("cant_impre")) = False Then
             data_afilcons.Recordset("cant_impre") = data_afilcons.Recordset("cant_impre") + 1
          Else
             data_afilcons.Recordset("cant_impre") = 1
          End If
          data_afilcons.Recordset.Update
          data_afilcons.Recordset.MoveNext
       Loop
       data_afilcons.Recordset.MovePrevious
       data_hist.RecordSource = "select * from afiliaciones_impre"
       data_hist.Refresh
       data_hist.Recordset.AddNew
       data_hist.Recordset("fecha") = Date
       data_hist.Recordset("hora") = Format(Time, "HH:mm")
       data_hist.Recordset("usuario") = WElusuario
       data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
       data_hist.Recordset("nro_afilia") = data_afilcons.Recordset("afilia_nro")
       data_hist.Recordset("accion") = "IMPRESION"
       data_hist.Recordset.Update
       data_afilcons.Recordset.MoveNext
       data_inf.RecordSource = "select * from infcli order by cl_cantpag"
       data_inf.Refresh
       If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
          If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
             cr1print.ReportFileName = App.path & "\contrato_debprint.rpt"
          Else
             cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
          End If
       Else
          cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
       End If
       cr1print.Action = 1
    Else
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
    End If
Else
   MsgBox "Falta seleccionar"
   
End If

End Sub

Private Sub Command3_Click()
Dim XsiDeseo As String
Dim NrodeFactura As String
XsiDeseo = MsgBox("Desea procesar el pago de esta afiliación?", vbInformation + vbYesNo, "Afiliaciones")
'Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
If XsiDeseo = vbYes Then
   NrodeFactura = InputBox("Ingrese número de factura", "Afiliaciones SAPP")
   If Trim(NrodeFactura) <> "" Then
      data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and integra_nro in (1)"
      data_afilcons.Refresh
      If data_afilcons.Recordset.RecordCount > 0 Then
         data_afilcons.Recordset.MoveFirst
         data_linmm.RecordSource = "select * from linmmdd where cod_prod in (992) and cod_cli =" & data_afilcons.Recordset("matricula") & " and factura =" & Val(NrodeFactura)
         data_linmm.Refresh
         If data_linmm.Recordset.RecordCount > 0 Then
            data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
            data_afilcons.Refresh
            If data_afilcons.Recordset.RecordCount > 0 Then
               data_afilcons.Recordset.MoveFirst
               Do While Not data_afilcons.Recordset.EOF
                  If IsNull(data_afilcons.Recordset("sifact")) = True Then
                     data_afilcons.Recordset.Edit
                     data_afilcons.Recordset("sifact") = 1
                     data_afilcons.Recordset("fecha") = Date
                     data_afilcons.Recordset.Update
                  End If
                  data_afilcons.Recordset.MoveNext
               Loop
               data_hist.RecordSource = "select * from afiliaciones_impre where nro_afilia =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
               data_hist.Refresh
               data_hist.Recordset.AddNew
               data_hist.Recordset("fecha") = Date
               data_hist.Recordset("hora") = Format(Time, "HH:mm")
               data_hist.Recordset("usuario") = WElusuario
               data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
               data_hist.Recordset("nro_afilia") = Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
               data_hist.Recordset("accion") = "PAGO DE AFILIACION"
               data_hist.Recordset.Update
               MsgBox "Pago procesado. Ya puede modificar la afiliación en padrón.", vbInformation, "Afiliaciones SAPP"
            End If
         Else
            MsgBox "No se encuentra la factura, Verifique!", vbCritical, "Afiliaciones SAPP"
         End If
      Else
         MsgBox "No se encuentra titular"
      End If
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
             
''''Grabar matrícula del socio para poder buscar acá

Xsqlpromos = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.nom2,afiliaciones_new.ape1,afiliaciones_new.ape2,afiliaciones_new.pendiente,afiliaciones_new.procesa_mod," & _
"afiliaciones_new.sifact,afiliaciones_new.matricula,afiliaciones_new.fnac,afiliaciones_new.integra_nro,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.cedula,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
"from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and afiliaciones_new.procesa_mod is null order by afiliaciones_new.integra_nro"
             
With Xreccliis
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromos, ConbdSapp, , , adCmdText
End With

ms2.rows = 2
ms2.Cols = 7
ms2.TextMatrix(0, 0) = "NRO."
ms2.ColWidth(0) = 1300
ms2.TextMatrix(0, 1) = "CEDULA"
ms2.ColWidth(1) = 1500
ms2.TextMatrix(0, 2) = "NOMBRES"
ms2.ColWidth(2) = 2500
ms2.TextMatrix(0, 3) = "APELLIDOS"
ms2.ColWidth(3) = 2500
ms2.TextMatrix(0, 4) = "F.NAC."
ms2.ColWidth(4) = 1500
ms2.TextMatrix(0, 5) = "CELULAR"
ms2.ColWidth(4) = 1500
ms2.TextMatrix(0, 6) = "MATRICULA"
ms2.ColWidth(4) = 1500

Xcann = 1
Dim Xpendiente As Integer
Dim Xsinfactura As Integer
Xpendiente = 0
Xsinfactura = 0

If Xreccliis.RecordCount > 0 Then
   Xreccliis.MoveFirst
   Xpendiente = Xreccliis("pendiente")
   If IsNull(Xreccliis("sifact")) = False Then
      Xsinfactura = Xreccliis("sifact")
   End If
   If Xpendiente = 2 Or Xpendiente = 4 Then
      If Xpendiente = 4 Then
         MsgBox "La afiliación está pendiente de facturación.", vbCritical
         If ControlUsuario("b_afil") = 1 Then
            ms2.Enabled = False
         Else
            ms2.Enabled = False
         End If
         Command1.Enabled = False
         b_pagos.Enabled = True
      Else
         MsgBox "La afiliación está pendiente de autorización, no se puede editar los socios", vbCritical
         ms2.Enabled = False
         Command1.Enabled = False
         Command2.Enabled = False
         b_pagos.Enabled = False
      End If
   Else
      If Xsinfactura <> 1 Then
         MsgBox "La afiliación no ha sido facturada, deberá facturar la afiliación para poder ingresar los socios.", vbCritical
         ms2.Enabled = False
         Command1.Enabled = False
         b_pagos.Enabled = True
      Else
         ms2.Enabled = True
         Command1.Enabled = True
         b_pagos.Enabled = False
      End If
   End If
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
      ms2.rows = ms2.rows + 1
      Xreccliis.MoveNext
      Xcann = Xcann + 1
   Loop
End If
If labpermiso.Caption = "0" Then
   Command1.Enabled = False
End If
b_pagos.Enabled = True

Xreccliis.Close
ConbdSapp.Close

End Sub

Private Sub Form_Load()
data_cnvmut.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_afilcons.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hist.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_linmm.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_clicons.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_abm.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_nrosoc.DatabaseName = App.path & "\parse.mdb"
data_nrosoc.RecordSource = "parsec0"
data_nrosoc.Refresh
data_convenios.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_inf.DatabaseName = App.path & "\informes.mdb"

If ControlUsuario("b_afil") = 1 Then
   b_pagos.Enabled = False
   Check1.Enabled = True
   Command1.Enabled = True
   labpermiso.Caption = "1"
   Command3.Visible = True
   Command3.Enabled = True
Else
   b_pagos.Enabled = False
   Check1.Enabled = False
   Command1.Enabled = False
   labpermiso.Caption = "0"
   Command3.Visible = False
   Command3.Enabled = False
End If

Carga_grid


End Sub

Public Sub Carga_grid()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xcann As Integer

ConectarBD
ConbdSapp.Open
             
'Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente," & _
'"afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
'"from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (0) order by afiliaciones_new.fecha"
             
             
'Data1.ConnectionString = "dsn=sappnew"
DBGrid1.Clear
ms2.Clear
If labpermiso.Caption = "1" Then
   If Check2.Value = 1 Then
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.nom1,afiliaciones_new.ape1," & _
      "afiliaciones_new.wusuario,afiliaciones_new.convenio,afiliaciones_new.codcob,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (20) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha"
   Else
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.nom1,afiliaciones_new.ape1," & _
      "afiliaciones_new.wusuario,afiliaciones_new.convenio,afiliaciones_new.codcob,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (0,3,4) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha"
   End If
Else
   If Check2.Value = 1 Then
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.nom1,afiliaciones_new.ape1," & _
      "afiliaciones_new.wusuario,afiliaciones_new.convenio,afiliaciones_new.codcob,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (20) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha"
   Else
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.nom1,afiliaciones_new.ape1," & _
      "afiliaciones_new.wusuario,afiliaciones_new.convenio,afiliaciones_new.codcob,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (0,3,4) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha"
   End If

'   If Check2.Value = 1 Then
'      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.nom1,afiliaciones_new.ape1," & _
'      "afiliaciones_new.wusuario,afiliaciones_new.convenio,afiliaciones_new.codcob,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
'      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (20) and afiliaciones_new.integra_nro in (1) and afiliaciones_new.wusuario ='" & WElusuario & "' order by afiliaciones_new.fecha"
'   Else
'      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.nom1,afiliaciones_new.ape1," & _
'      "afiliaciones_new.wusuario,afiliaciones_new.convenio,afiliaciones_new.codcob,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
'      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (0,3,4) and afiliaciones_new.integra_nro in (1) and afiliaciones_new.wusuario='" & WElusuario & "' order by afiliaciones_new.fecha"
'   End If

End If
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
DBGrid1.ColWidth(1) = 1100
DBGrid1.TextMatrix(0, 2) = "CATEGORIA"
DBGrid1.ColWidth(2) = 2100
DBGrid1.TextMatrix(0, 3) = "TITULAR AFILIACIÓN"
DBGrid1.ColWidth(3) = 3000
DBGrid1.TextMatrix(0, 4) = "PROMOTOR"
DBGrid1.ColWidth(4) = 2800
DBGrid1.TextMatrix(0, 5) = "COBRADOR"
DBGrid1.ColWidth(5) = 1500
DBGrid1.TextMatrix(0, 6) = "USUARIO"
DBGrid1.ColWidth(6) = 1500

Xcann = 1

If Xrecclii.RecordCount > 0 Then
   DBGrid1.Enabled = True
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      DBGrid1.TextMatrix(Xcann, 0) = Xrecclii("fecha")
      DBGrid1.TextMatrix(Xcann, 1) = Xrecclii("afilia_nro")
      DBGrid1.TextMatrix(Xcann, 2) = Xrecclii("convenio")
      DBGrid1.TextMatrix(Xcann, 3) = Xrecclii("ape1") & " " & Xrecclii("nom1")
      DBGrid1.TextMatrix(Xcann, 4) = Xrecclii("nombre")
      DBGrid1.TextMatrix(Xcann, 5) = Xrecclii("codcob")
      DBGrid1.TextMatrix(Xcann, 6) = Xrecclii("wusuario")
      
      DBGrid1.rows = DBGrid1.rows + 1
      Xrecclii.MoveNext
      Xcann = Xcann + 1
   Loop
Else
   DBGrid1.Enabled = False
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Private Sub ms2_DblClick()
Dim Xmataf As Long
Dim Xgpoconv As String
Dim Cl_apellid, Cl_dir, Cl_entre As String

On Error GoTo Alseleccionar

Xgpoconv = ""
Cl_apellid = ""
Cl_dir = ""
Cl_entre = ""

Xmataf = Val(ms2.TextMatrix(ms2.RowSel, 6))

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
If IsNull(frmabm.data_clientes.Recordset("cl_ultmesp")) = False Then
   frmabm.labump.Caption = frmabm.data_clientes.Recordset("cl_ultmesp")
Else
   frmabm.labump.Caption = ""
End If
If IsNull(frmabm.data_clientes.Recordset("cl_ultanop")) = False Then
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

If Check1.Value = 1 Then

End If

'Unload Me
frm_afilpend.Hide

Exit Sub

Alseleccionar:
             If Err.Number = 3157 Then
                MsgBox "Error al seleccionar." & Err.Description
             Else
                MsgBox "Error al seleccionar." & Err.Description
             
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
      data_inf.Recordset("cl_fax") = Mid(data_afilcons.Recordset("cedula"), 1, 6) & "-" & Mid(data_afilcons.Recordset("cedula"), 7, 1)
   Else
      data_inf.Recordset("cl_fax") = Mid(data_afilcons.Recordset("cedula"), 1, 7) & "-" & Mid(data_afilcons.Recordset("cedula"), 8, 1)
   End If
   data_inf.Recordset("cl_tjemi_n") = Devuelve_titularCed()
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


Public Function Devuelve_titularCed() As String
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
   If Len(Xrecclii("cedula")) = 7 Then
      Devuelve_titularCed = Mid(Trim(str(Xrecclii("cedula"))), 1, 6) & "-" & Mid(Trim(str(Xrecclii("cedula"))), 7, 1)
   Else
      Devuelve_titularCed = Mid(Trim(str(Xrecclii("cedula"))), 1, 7) & "-" & Mid(Trim(str(Xrecclii("cedula"))), 8, 1)
   End If
Else
   Devuelve_titularCed = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function

Public Sub Alta_Modif()
Dim Cl_apellid, Cl_entre, Cl_dir, VenceT, MutAfil, VendeAfil As String
Dim Xmatnew As Long
Xmatnew = 0
Cl_apellid = ""
Cl_entre = ""
VenceT = ""
Cl_dir = ""
MutAfil = ""
VendeAfil = ""


If IsNull(data_afilcons.Recordset("matricula")) = True Then
    Xmatnew = data_nrosoc.Recordset("ultimo_soc") + 1
    
    data_nrosoc.Recordset.Edit
    data_nrosoc.Recordset("ultimo_soc") = Xmatnew
    data_nrosoc.Recordset.Update
    
    data_afilcons.Recordset.Edit
    data_afilcons.Recordset("matricula") = Xmatnew
    data_afilcons.Recordset.Update

    data_cli.Recordset.AddNew
    data_cli.Recordset("cl_codigo") = Xmatnew
    data_cli.Recordset("estado") = 1
    If IsNull(data_afilcons.Recordset("catreal")) = False Then
       data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("catreal")
       data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("catrealdes")
    Else
       data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("categ")
       data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("nomcateg")
    End If
    If IsNull(data_afilcons.Recordset("nom2")) = False Then
       If IsNull(data_afilcons.Recordset("ape2")) = False Then
          Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
       Else
          Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
       End If
    Else
       If IsNull(data_afilcons.Recordset("ape2")) = False Then
          Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1")
       Else
          Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
       End If
    End If
    data_cli.Recordset("cl_apellid") = Mid(Cl_apellid, 1, 60)
    If IsNull(data_afilcons.Recordset("manz")) = False Then
       If IsNull(data_afilcons.Recordset("solar")) = False Then
          Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz") & " SL." & data_afilcons.Recordset("solar")
       Else
          Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz")
       End If
    Else
       Cl_dir = data_afilcons.Recordset("direc1")
    End If
    data_cli.Recordset("cl_direcci") = Mid(Cl_dir, 1, 80)
    If IsNull(data_afilcons.Recordset("casa")) = False Then
       If IsNull(data_afilcons.Recordset("direc2")) = False Then
          Cl_entre = data_afilcons.Recordset("casa") & " " & data_afilcons.Recordset("direc2")
       Else
          Cl_entre = data_afilcons.Recordset("casa")
       End If
    Else
       If IsNull(data_afilcons.Recordset("direc2")) = False Then
          Cl_entre = data_afilcons.Recordset("direc2")
       End If
    End If
    If Trim(Cl_entre) <> "" Then
       data_cli.Recordset("cl_entre") = Mid(Cl_entre, 1, 80)
    End If
    data_cli.Recordset("cl_dpto") = data_afilcons.Recordset("celular")
    data_cli.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
    
    data_cli.Recordset("cl_cedula_t") = Trim(data_afilcons.Recordset("cedula"))
    If IsNull(data_afilcons.Recordset("celular")) = False Then
       If data_afilcons.Recordset("celular") = "NO APLICA" Then
       Else
          data_cli.Recordset("cl_celular_n") = Trim(data_afilcons.Recordset("celular"))
       End If
    End If
    
    If IsNull(data_afilcons.Recordset("tipoced")) = False Then
       If data_afilcons.Recordset("tipoced") = 1 Then
          data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
          data_cli.Recordset("cl_codced") = 0
          data_cli.Recordset("cl_tipoced") = 1
       Else
          If Len(data_afilcons.Recordset("cedula")) = 7 Then
             data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
             data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
             data_cli.Recordset("cl_tipoced") = 0
          Else
             If Len(data_afilcons.Recordset("cedula")) = 8 Then
                data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
                data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
                data_cli.Recordset("cl_tipoced") = 0
             Else
                data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
                data_cli.Recordset("cl_codced") = 0
                data_cli.Recordset("cl_tipoced") = 0
             End If
          End If
       End If
    Else
       If Len(data_afilcons.Recordset("cedula")) = 7 Then
          data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
          data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
          data_cli.Recordset("cl_tipoced") = 0
       Else
          If Len(data_afilcons.Recordset("cedula")) = 8 Then
             data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
             data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
             data_cli.Recordset("cl_tipoced") = 0
          Else
             data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
             data_cli.Recordset("cl_codced") = 0
             data_cli.Recordset("cl_tipoced") = 0
          End If
       End If
    End If
    data_cli.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
    data_cli.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
    
    If IsNull(data_afilcons.Recordset("tarj_sello")) = False Then
       data_cli.Recordset("cl_forpago") = 2
       data_cli.Recordset("cl_descpag") = "Debito Automatico"
    Else
       data_cli.Recordset("cl_forpago") = 1
       data_cli.Recordset("cl_descpag") = "Abono Mensual"
    End If
    If IsNull(data_afilcons.Recordset("codpromo")) = False Then
       If data_afilcons.Recordset("codpromo") > 0 Then
          data_cli.Recordset("idpromos") = data_afilcons.Recordset("codpromo")
       End If
    End If
    If Day(Date) >= 25 Then
       If Month(Date) = 12 Then
          data_cli.Recordset("cl_ultmesp") = 1
          data_cli.Recordset("cl_ultanop") = Year(Date) + 1
       Else
          data_cli.Recordset("cl_ultmesp") = Month(Date) + 1
          data_cli.Recordset("cl_ultanop") = Year(Date)
       End If
    Else
       data_cli.Recordset("cl_ultmesp") = Month(Date)
       data_cli.Recordset("cl_ultanop") = Year(Date)
    End If
    If IsNull(data_afilcons.Recordset("codpromo")) = False Then
       If data_afilcons.Recordset("codpromo") = 2 Then
          If Month(Date) = 12 Then
             data_cli.Recordset("cl_ultmesp") = 11
             data_cli.Recordset("cl_ultanop") = Year(Date) + 1
          Else
             data_cli.Recordset("cl_ultmesp") = Month(Date) - 1
             data_cli.Recordset("cl_ultanop") = Year(Date) + 1
          End If
       End If
    End If
    data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
    data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
    If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
       data_cli.Recordset("cl_dircobr") = data_afilcons.Recordset("direc_cobro")
    End If
    If IsNull(data_afilcons.Recordset("tarj_sello")) = True Then
       If IsNull(data_afilcons.Recordset("codcob")) = False Then
          If data_afilcons.Recordset("codcob") = 0 Then
             data_cli.Recordset("cl_nrocobr") = 0
             data_cli.Recordset("cl_nomcobr") = "*TODOS"
          Else
             data_cli.Recordset("cl_nrocobr") = data_afilcons.Recordset("codcob")
             data_cli.Recordset("cl_nomcobr") = Mid(Devuelve_cobradorOk(), 1, 25)
          End If
       Else
          data_cli.Recordset("cl_nrocobr") = 0
          data_cli.Recordset("cl_nomcobr") = "*TODOS"
       End If
    Else
       If data_afilcons.Recordset("tarj_sello") = "OCA CARD" Then
          data_cli.Recordset("cl_nrocobr") = 690
          data_cli.Recordset("cl_nomcobr") = "OCA DEBITO"
       End If
       If data_afilcons.Recordset("tarj_sello") = "VISA" Then
          data_cli.Recordset("cl_nrocobr") = 514
          data_cli.Recordset("cl_nomcobr") = "DEBITO AUTOMATICO VISA"
       End If
       If data_afilcons.Recordset("tarj_sello") = "MASTER CARD" Then
          data_cli.Recordset("cl_nrocobr") = 683
          data_cli.Recordset("cl_nomcobr") = "DEBITO MASTERCARD"
       End If
       If data_afilcons.Recordset("tarj_sello") = "CABAL" Then
          data_cli.Recordset("cl_nrocobr") = 673
          data_cli.Recordset("cl_nomcobr") = "DEBITO CABAL"
       End If
       If data_afilcons.Recordset("tarj_sello") = "DEBITO BROU" Then
          data_cli.Recordset("cl_nrocobr") = 607
          data_cli.Recordset("cl_nomcobr") = "DEBITO BROU"
       End If
       data_cli.Recordset("tit_tarj") = data_afilcons.Recordset("tarj_titular")
       If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
          data_cli.Recordset("cl_nrotarj") = data_afilcons.Recordset("tarj_nro")
       End If
       data_cli.Recordset("ci_tarj") = data_afilcons.Recordset("tarj_cedtit")
       data_cli.Recordset("codcitarj") = data_afilcons.Recordset("tarj_codced")
       data_cli.Recordset("cl_tjemi_c") = data_afilcons.Recordset("tarj_codsello")
       data_cli.Recordset("cl_tjemi_n") = data_afilcons.Recordset("tarj_sello")
       If data_afilcons.Recordset("tarj_vencmes") <> 0 Then
          If data_afilcons.Recordset("tarj_vencmes") > 9 Then
             VenceT = "01/" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
          Else
             VenceT = "01/0" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
          End If
          data_cli.Recordset("cl_tj_venc") = Format(CDate(VenceT), "dd/mm/yyyy")
       End If
       data_cli.Recordset("tarj_domi") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 60)
       data_cli.Recordset("tarj_telef") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 45)
    End If
    If IsNull(data_afilcons.Recordset("codpromo")) = False Then
       If data_afilcons.Recordset("codpromo") = 1 Then
          If data_afilcons.Recordset("integra_nro") <> 1 Then
             Consulta_AfilRuta
             If labcedtit.Caption <> "" Then
                data_cli.Recordset("cl_codruta") = Val(labcedtit.Caption)
             End If
          End If
       End If
    End If
    If Day(Date) >= 25 Then
       If Month(Date) = 11 Or Month(Date) = 12 Then
          If Month(Date) = 11 Then
             data_cli.Recordset("mesproxemi") = 1
             data_cli.Recordset("anoproxemi") = Year(Date) + 1
          Else
             data_cli.Recordset("mesproxemi") = 2
             data_cli.Recordset("anoproxemi") = Year(Date) + 1
          End If
       Else
          data_cli.Recordset("mesproxemi") = Month(Date) + 2
          data_cli.Recordset("anoproxemi") = Year(Date)
       End If
    Else
       If Month(Date) = 12 Then
          data_cli.Recordset("mesproxemi") = 1
          data_cli.Recordset("anoproxemi") = Year(Date) + 1
       Else
          data_cli.Recordset("mesproxemi") = Month(Date) + 1
          data_cli.Recordset("anoproxemi") = Year(Date)
       End If
    End If
    If IsNull(data_afilcons.Recordset("codpromo")) = False Then
       If data_afilcons.Recordset("codpromo") = 2 Then
          If Day(Date) >= 25 Then
             If Month(Date) = 12 Then
                data_cli.Recordset("mesproxemi") = 1
                data_cli.Recordset("anoproxemi") = Year(Date) + 2
             Else
                data_cli.Recordset("mesproxemi") = Month(Date) + 1
                data_cli.Recordset("anoproxemi") = Year(Date) + 1
             End If
          Else
             If Month(Date) = 12 Then
                data_cli.Recordset("mesproxemi") = 1
                data_cli.Recordset("anoproxemi") = Year(Date) + 2
             Else
                data_cli.Recordset("mesproxemi") = Month(Date)
                data_cli.Recordset("anoproxemi") = Year(Date) + 1
             End If
          End If
       End If
    End If
    data_cli.Recordset("cl_referen") = Mid(data_afilcons.Recordset("correo"), 1, 74)
    If IsNull(data_afilcons.Recordset("codzon")) = False Then
       data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
       data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
    Else
       data_cli.Recordset("cl_grupo") = 0
       data_cli.Recordset("cl_zona") = "*TODOS"
    End If
    data_cli.Recordset("cl_sexo") = data_afilcons.Recordset("sexo")
    MutAfil = Devuelve_mut()
    data_cli.Recordset("cl_socmnom") = Mid(MutAfil, 1, 25)
    If IsNull(data_cli.Recordset("fecha_baja")) = False Then
       data_cli.Recordset("fecha_baja") = Null
    End If
    VendeAfil = Devuelve_vende()
    data_cli.Recordset("cl_nrovend") = data_afilcons.Recordset("codvende")
    data_cli.Recordset("cl_nomvend") = Mid(VendeAfil, 1, 35)
    If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
       data_cli.Recordset("cl_diacobr") = Trim(str(data_afilcons.Recordset("dia_cobro"))) & " C/MES"
    End If
    data_cli.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
    data_cli.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
    data_cli.Recordset.Update
    data_abm.RecordSource = "select * from abmsocio where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
    data_abm.Refresh
    data_abm.Recordset.AddNew
    data_abm.Recordset("usuario") = WElusuario
    data_abm.Recordset("fecha") = Date
    data_abm.Recordset("hora") = Format(Time, "HH:mm")
    data_abm.Recordset("cl_codigo") = Xmatnew
    data_abm.Recordset("desc") = "ALTA"
    data_abm.Recordset("cl_motivo") = "ALTA DE FICHA"
    data_abm.Recordset("convenio") = data_afilcons.Recordset("categ")
    data_abm.Recordset("base") = data_nrosoc.Recordset("base")
    data_abm.Recordset.Update
Else
    Xmatnew = data_afilcons.Recordset("matricula")
    data_cli.RecordSource = "select * from clientes where cl_codigo =" & data_afilcons.Recordset("matricula")
    data_cli.Refresh
    If data_cli.Recordset.RecordCount > 0 Then
       If IsNull(data_cli.Recordset("fecha_baja")) = True Then
          data_convenios.RecordSource = "select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "' and cnv_grupo in ('CCOU','SMI','IMPASA','UNIVERSAL','CASA DE GALICIA','H.EVANGELICO')"
          data_convenios.Refresh
          If data_convenios.Recordset.RecordCount = 0 Then
             data_cli.Recordset.Edit
             data_cli.Recordset("estado") = 1
             If IsNull(data_afilcons.Recordset("catreal")) = False Then
                data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("catreal")
                data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("catrealdes")
             Else
                data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("categ")
                data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("nomcateg")
             End If
             If IsNull(data_afilcons.Recordset("nom2")) = False Then
                If IsNull(data_afilcons.Recordset("ape2")) = False Then
                   Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
                Else
                   Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
                End If
             Else
                If IsNull(data_afilcons.Recordset("ape2")) = False Then
                   Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1")
                Else
                   Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
                End If
             End If
             data_cli.Recordset("hora_reac") = Format(Time, "HH:mm")
             data_cli.Recordset("cl_apellid") = Mid(Cl_apellid, 1, 60)
             If IsNull(data_afilcons.Recordset("manz")) = False Then
                If IsNull(data_afilcons.Recordset("solar")) = False Then
                   Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz") & " SL." & data_afilcons.Recordset("solar")
                Else
                   Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz")
                End If
             Else
                Cl_dir = data_afilcons.Recordset("direc1")
             End If
             data_cli.Recordset("cl_direcci") = Mid(Cl_dir, 1, 80)
             If IsNull(data_afilcons.Recordset("casa")) = False Then
                If IsNull(data_afilcons.Recordset("direc2")) = False Then
                   Cl_entre = data_afilcons.Recordset("casa") & " " & data_afilcons.Recordset("direc2")
                Else
                   Cl_entre = data_afilcons.Recordset("casa")
                End If
             Else
                If IsNull(data_afilcons.Recordset("direc2")) = False Then
                   Cl_entre = data_afilcons.Recordset("direc2")
                End If
             End If
             If Trim(Cl_entre) <> "" Then
                data_cli.Recordset("cl_entre") = Mid(Cl_entre, 1, 80)
             End If
             data_cli.Recordset("cl_dpto") = data_afilcons.Recordset("celular")
             data_cli.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
             
             data_cli.Recordset("cl_cedula_t") = Trim(data_afilcons.Recordset("cedula"))
             If IsNull(data_afilcons.Recordset("celular")) = False Then
                If data_afilcons.Recordset("celular") = "NO APLICA" Then
                Else
                   data_cli.Recordset("cl_celular_n") = Trim(data_afilcons.Recordset("celular"))
                End If
             End If
             If Len(data_afilcons.Recordset("cedula")) = 7 Then
                data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
                data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
             Else
                If Len(data_afilcons.Recordset("cedula")) = 8 Then
                   data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
                   data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
                Else
                   data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
                   data_cli.Recordset("cl_codced") = 0
                End If
             End If
             data_cli.Recordset("cl_tipoced") = 0
             data_cli.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
             data_cli.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
       
             If IsNull(data_afilcons.Recordset("tarj_sello")) = False Then
                data_cli.Recordset("cl_forpago") = 2
                data_cli.Recordset("cl_descpag") = "Debito Automatico"
             Else
                data_cli.Recordset("cl_forpago") = 1
                data_cli.Recordset("cl_descpag") = "Abono Mensual"
             End If
             If IsNull(data_afilcons.Recordset("codpromo")) = False Then
                If data_afilcons.Recordset("codpromo") > 0 Then
                   data_cli.Recordset("idpromos") = data_afilcons.Recordset("codpromo")
                End If
             End If
             If Day(Date) >= 25 Then
                If Month(Date) = 12 Then
                   data_cli.Recordset("cl_ultmesp") = 1
                   data_cli.Recordset("cl_ultanop") = Year(Date) + 1
                Else
                   data_cli.Recordset("cl_ultmesp") = Month(Date) + 1
                   data_cli.Recordset("cl_ultanop") = Year(Date)
                End If
             Else
                data_cli.Recordset("cl_ultmesp") = Month(Date)
                data_cli.Recordset("cl_ultanop") = Year(Date)
             End If
             If IsNull(data_afilcons.Recordset("codpromo")) = False Then
                If data_afilcons.Recordset("codpromo") = 2 Then
                   If Month(Date) = 12 Then
                      data_cli.Recordset("cl_ultmesp") = 11
                      data_cli.Recordset("cl_ultanop") = Year(Date) + 1
                   Else
                      data_cli.Recordset("cl_ultmesp") = Month(Date) - 1
                      data_cli.Recordset("cl_ultanop") = Year(Date) + 1
                   End If
                End If
             End If
             data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
             data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
             If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
                data_cli.Recordset("cl_dircobr") = data_afilcons.Recordset("direc_cobro")
             End If
             If IsNull(data_afilcons.Recordset("tarj_sello")) = True Then
                If IsNull(data_afilcons.Recordset("codcob")) = False Then
                   If data_afilcons.Recordset("codcob") = 0 Then
                      data_cli.Recordset("cl_nrocobr") = 0
                      data_cli.Recordset("cl_nomcobr") = "*TODOS"
                   Else
                      data_cli.Recordset("cl_nrocobr") = data_afilcons.Recordset("codcob")
                      data_cli.Recordset("cl_nomcobr") = Mid(Devuelve_cobradorOk(), 1, 25)
                   End If
                Else
                   data_cli.Recordset("cl_nrocobr") = 0
                   data_cli.Recordset("cl_nomcobr") = "*TODOS"
                End If
             Else
                If data_afilcons.Recordset("tarj_sello") = "OCA CARD" Then
                   data_cli.Recordset("cl_nrocobr") = 690
                   data_cli.Recordset("cl_nomcobr") = "OCA DEBITO"
                End If
                If data_afilcons.Recordset("tarj_sello") = "VISA" Then
                   data_cli.Recordset("cl_nrocobr") = 514
                   data_cli.Recordset("cl_nomcobr") = "DEBITO AUTOMATICO VISA"
                End If
                If data_afilcons.Recordset("tarj_sello") = "MASTER CARD" Then
                   data_cli.Recordset("cl_nrocobr") = 683
                   data_cli.Recordset("cl_nomcobr") = "DEBITO MASTERCARD"
                End If
                If data_afilcons.Recordset("tarj_sello") = "CABAL" Then
                   data_cli.Recordset("cl_nrocobr") = 673
                   data_cli.Recordset("cl_nomcobr") = "DEBITO CABAL"
                End If
                If data_afilcons.Recordset("tarj_sello") = "DEBITO BROU" Then
                   data_cli.Recordset("cl_nrocobr") = 607
                   data_cli.Recordset("cl_nomcobr") = "DEBITO BROU"
                End If
                data_cli.Recordset("tit_tarj") = data_afilcons.Recordset("tarj_titular")
                If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
                   data_cli.Recordset("cl_nrotarj") = data_afilcons.Recordset("tarj_nro")
                End If
                data_cli.Recordset("ci_tarj") = data_afilcons.Recordset("tarj_cedtit")
                data_cli.Recordset("codcitarj") = data_afilcons.Recordset("tarj_codced")
                data_cli.Recordset("cl_tjemi_c") = data_afilcons.Recordset("tarj_codsello")
                data_cli.Recordset("cl_tjemi_n") = data_afilcons.Recordset("tarj_sello")
                If data_afilcons.Recordset("tarj_vencmes") <> 0 Then
                   If data_afilcons.Recordset("tarj_vencmes") > 9 Then
                      VenceT = "01/" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
                   Else
                      VenceT = "01/0" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
                   End If
                   data_cli.Recordset("cl_tj_venc") = Format(CDate(VenceT), "dd/mm/yyyy")
                End If
                data_cli.Recordset("tarj_domi") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 60)
                data_cli.Recordset("tarj_telef") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 45)
             End If
             If IsNull(data_afilcons.Recordset("codpromo")) = False Then
                If data_afilcons.Recordset("codpromo") = 1 Then
                   If data_afilcons.Recordset("integra_nro") <> 1 Then
                      Consulta_AfilRuta
                      If labcedtit.Caption <> "" Then
                         data_cli.Recordset("cl_codruta") = Val(labcedtit.Caption)
                      End If
                   End If
                End If
             End If
             If Day(Date) >= 25 Then
                If Month(Date) = 11 Or Month(Date) = 12 Then
                   If Month(Date) = 11 Then
                      data_cli.Recordset("mesproxemi") = 1
                      data_cli.Recordset("anoproxemi") = Year(Date) + 1
                   Else
                      data_cli.Recordset("mesproxemi") = 2
                      data_cli.Recordset("anoproxemi") = Year(Date) + 1
                   End If
                Else
                   data_cli.Recordset("mesproxemi") = Month(Date) + 2
                   data_cli.Recordset("anoproxemi") = Year(Date)
                End If
             Else
                If Month(Date) = 12 Then
                   data_cli.Recordset("mesproxemi") = 1
                   data_cli.Recordset("anoproxemi") = Year(Date) + 1
                Else
                   data_cli.Recordset("mesproxemi") = Month(Date) + 1
                   data_cli.Recordset("anoproxemi") = Year(Date)
                End If
             End If
             If IsNull(data_afilcons.Recordset("codpromo")) = False Then
                If data_afilcons.Recordset("codpromo") = 2 Then
                   If Day(Date) >= 25 Then
                      If Month(Date) = 12 Then
                         data_cli.Recordset("mesproxemi") = 1
                         data_cli.Recordset("anoproxemi") = Year(Date) + 2
                      Else
                         data_cli.Recordset("mesproxemi") = Month(Date) + 1
                         data_cli.Recordset("anoproxemi") = Year(Date) + 1
                      End If
                   Else
                      If Month(Date) = 12 Then
                         data_cli.Recordset("mesproxemi") = 1
                         data_cli.Recordset("anoproxemi") = Year(Date) + 2
                      Else
                         data_cli.Recordset("mesproxemi") = Month(Date)
                         data_cli.Recordset("anoproxemi") = Year(Date) + 1
                      End If
                   End If
                End If
             End If
             data_cli.Recordset("cl_referen") = Mid(data_afilcons.Recordset("correo"), 1, 74)
             If IsNull(data_afilcons.Recordset("codzon")) = False Then
                data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
                data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
             Else
                data_cli.Recordset("cl_grupo") = 0
                data_cli.Recordset("cl_zona") = "*TODOS"
             End If
             data_cli.Recordset("cl_sexo") = data_afilcons.Recordset("sexo")
             MutAfil = Devuelve_mut()
             data_cli.Recordset("cl_socmnom") = Mid(MutAfil, 1, 25)
             If IsNull(data_cli.Recordset("fecha_baja")) = False Then
                data_cli.Recordset("fecha_baja") = Null
             End If
             VendeAfil = Devuelve_vende()
             data_cli.Recordset("cl_nrovend") = data_afilcons.Recordset("codvende")
             data_cli.Recordset("cl_nomvend") = Mid(VendeAfil, 1, 35)
             If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
                data_cli.Recordset("cl_diacobr") = Trim(str(data_afilcons.Recordset("dia_cobro"))) & " C/MES"
             End If
             data_cli.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
             data_cli.Recordset.Update
             data_abm.RecordSource = "select * from abmsocio where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
             data_abm.Refresh
             data_abm.Recordset.AddNew
             data_abm.Recordset("usuario") = WElusuario
             data_abm.Recordset("fecha") = Date
             data_abm.Recordset("hora") = Format(Time, "HH:mm")
             data_abm.Recordset("cl_codigo") = Xmatnew
             data_abm.Recordset("desc") = "MODIF"
             data_abm.Recordset("cl_motivo") = "MODIF DE FICHA"
             data_abm.Recordset("convenio") = data_afilcons.Recordset("categ")
             data_abm.Recordset("base") = data_nrosoc.Recordset("base")
             data_abm.Recordset.Update
          End If
       Else
          data_cli.Recordset.Edit
          data_cli.Recordset("fecha_baja") = Null
          data_cli.Recordset("estado") = 1
       
            If IsNull(data_afilcons.Recordset("catreal")) = False Then
               data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("catreal")
               data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("catrealdes")
            Else
               data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("categ")
               data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("nomcateg")
            End If
            If IsNull(data_afilcons.Recordset("nom2")) = False Then
               If IsNull(data_afilcons.Recordset("ape2")) = False Then
                  Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
               Else
                  Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
               End If
            Else
               If IsNull(data_afilcons.Recordset("ape2")) = False Then
                  Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1")
               Else
                  Cl_apellid = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
               End If
            End If
            data_cli.Recordset("hora_reac") = Format(Time, "HH:mm")
          data_cli.Recordset("cl_apellid") = Mid(Cl_apellid, 1, 60)
          If IsNull(data_afilcons.Recordset("manz")) = False Then
             If IsNull(data_afilcons.Recordset("solar")) = False Then
                Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz") & " SL." & data_afilcons.Recordset("solar")
             Else
                Cl_dir = data_afilcons.Recordset("direc1") & " MZ." & data_afilcons.Recordset("manz")
             End If
          Else
             Cl_dir = data_afilcons.Recordset("direc1")
          End If
          data_cli.Recordset("cl_direcci") = Mid(Cl_dir, 1, 80)
          If IsNull(data_afilcons.Recordset("casa")) = False Then
             If IsNull(data_afilcons.Recordset("direc2")) = False Then
                Cl_entre = data_afilcons.Recordset("casa") & " " & data_afilcons.Recordset("direc2")
             Else
                Cl_entre = data_afilcons.Recordset("casa")
             End If
          Else
             If IsNull(data_afilcons.Recordset("direc2")) = False Then
                Cl_entre = data_afilcons.Recordset("direc2")
             End If
          End If
          If Trim(Cl_entre) <> "" Then
             data_cli.Recordset("cl_entre") = Mid(Cl_entre, 1, 80)
          End If
          data_cli.Recordset("cl_dpto") = data_afilcons.Recordset("celular")
          data_cli.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
          If Len(data_afilcons.Recordset("cedula")) = 7 Then
             data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
             data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
          Else
             If Len(data_afilcons.Recordset("cedula")) = 8 Then
                data_cli.Recordset("cl_cedula") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
                data_cli.Recordset("cl_codced") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
             Else
                data_cli.Recordset("cl_cedula") = Val(Trim(str(data_afilcons.Recordset("cedula"))))
                data_cli.Recordset("cl_codced") = 0
             End If
          End If
          data_cli.Recordset("cl_tipoced") = 0
          data_cli.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
          data_cli.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
    
          If IsNull(data_afilcons.Recordset("tarj_sello")) = False Then
             data_cli.Recordset("cl_forpago") = 2
             data_cli.Recordset("cl_descpag") = "Debito Automatico"
          Else
             data_cli.Recordset("cl_forpago") = 1
             data_cli.Recordset("cl_descpag") = "Abono Mensual"
          End If
          If IsNull(data_afilcons.Recordset("codpromo")) = False Then
             If data_afilcons.Recordset("codpromo") > 0 Then
                data_cli.Recordset("idpromos") = data_afilcons.Recordset("codpromo")
             End If
          End If
          If Day(Date) >= 25 Then
             If Month(Date) = 12 Then
                data_cli.Recordset("cl_ultmesp") = 1
                data_cli.Recordset("cl_ultanop") = Year(Date) + 1
             Else
                data_cli.Recordset("cl_ultmesp") = Month(Date) + 1
                data_cli.Recordset("cl_ultanop") = Year(Date)
             End If
          Else
             data_cli.Recordset("cl_ultmesp") = Month(Date)
             data_cli.Recordset("cl_ultanop") = Year(Date)
          End If
          If IsNull(data_afilcons.Recordset("codpromo")) = False Then
             If data_afilcons.Recordset("codpromo") = 2 Then
                If Month(Date) = 12 Then
                   data_cli.Recordset("cl_ultmesp") = 11
                   data_cli.Recordset("cl_ultanop") = Year(Date) + 1
                Else
                   data_cli.Recordset("cl_ultmesp") = Month(Date) - 1
                   data_cli.Recordset("cl_ultanop") = Year(Date) + 1
                End If
             End If
          End If
          data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
          data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
          If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
             data_cli.Recordset("cl_dircobr") = data_afilcons.Recordset("direc_cobro")
          End If
          If IsNull(data_afilcons.Recordset("tarj_sello")) = True Then
             If IsNull(data_afilcons.Recordset("codcob")) = False Then
                If data_afilcons.Recordset("codcob") = 0 Then
                   data_cli.Recordset("cl_nrocobr") = 0
                   data_cli.Recordset("cl_nomcobr") = "*TODOS"
                Else
                   data_cli.Recordset("cl_nrocobr") = data_afilcons.Recordset("codcob")
                   data_cli.Recordset("cl_nomcobr") = Mid(Devuelve_cobradorOk(), 1, 25)
                End If
             Else
                data_cli.Recordset("cl_nrocobr") = 0
                data_cli.Recordset("cl_nomcobr") = "*TODOS"
             End If
          Else
             If data_afilcons.Recordset("tarj_sello") = "OCA CARD" Then
                data_cli.Recordset("cl_nrocobr") = 690
                data_cli.Recordset("cl_nomcobr") = "OCA DEBITO"
             End If
             If data_afilcons.Recordset("tarj_sello") = "VISA" Then
                data_cli.Recordset("cl_nrocobr") = 514
                data_cli.Recordset("cl_nomcobr") = "DEBITO AUTOMATICO VISA"
             End If
             If data_afilcons.Recordset("tarj_sello") = "MASTER CARD" Then
                data_cli.Recordset("cl_nrocobr") = 683
                data_cli.Recordset("cl_nomcobr") = "DEBITO MASTERCARD"
             End If
             If data_afilcons.Recordset("tarj_sello") = "CABAL" Then
                data_cli.Recordset("cl_nrocobr") = 673
                data_cli.Recordset("cl_nomcobr") = "DEBITO CABAL"
             End If
             If data_afilcons.Recordset("tarj_sello") = "DEBITO BROU" Then
                data_cli.Recordset("cl_nrocobr") = 607
                data_cli.Recordset("cl_nomcobr") = "DEBITO BROU"
             End If
             data_cli.Recordset("tit_tarj") = data_afilcons.Recordset("tarj_titular")
             If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
                data_cli.Recordset("cl_nrotarj") = data_afilcons.Recordset("tarj_nro")
             End If
             data_cli.Recordset("ci_tarj") = data_afilcons.Recordset("tarj_cedtit")
             data_cli.Recordset("codcitarj") = data_afilcons.Recordset("tarj_codced")
             data_cli.Recordset("cl_tjemi_c") = data_afilcons.Recordset("tarj_codsello")
             data_cli.Recordset("cl_tjemi_n") = data_afilcons.Recordset("tarj_sello")
             If data_afilcons.Recordset("tarj_vencmes") <> 0 Then
                If data_afilcons.Recordset("tarj_vencmes") > 9 Then
                   VenceT = "01/" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
                Else
                   VenceT = "01/0" & data_afilcons.Recordset("tarj_vencmes") & "/" & data_afilcons.Recordset("tarj_vencanio")
                End If
                data_cli.Recordset("cl_tj_venc") = Format(CDate(VenceT), "dd/mm/yyyy")
             End If
             data_cli.Recordset("tarj_domi") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 60)
             data_cli.Recordset("tarj_telef") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 45)
          End If
          If IsNull(data_afilcons.Recordset("codpromo")) = False Then
             If data_afilcons.Recordset("codpromo") = 1 Then
                If data_afilcons.Recordset("integra_nro") <> 1 Then
                   Consulta_AfilRuta
                   If labcedtit.Caption <> "" Then
                      data_cli.Recordset("cl_codruta") = Val(labcedtit.Caption)
                   End If
                End If
             End If
          End If
          If Day(Date) >= 25 Then
             If Month(Date) = 11 Or Month(Date) = 12 Then
                If Month(Date) = 11 Then
                   data_cli.Recordset("mesproxemi") = 1
                   data_cli.Recordset("anoproxemi") = Year(Date) + 1
                Else
                   data_cli.Recordset("mesproxemi") = 2
                   data_cli.Recordset("anoproxemi") = Year(Date) + 1
                End If
             Else
                data_cli.Recordset("mesproxemi") = Month(Date) + 2
                data_cli.Recordset("anoproxemi") = Year(Date)
             End If
          Else
             If Month(Date) = 12 Then
                data_cli.Recordset("mesproxemi") = 1
                data_cli.Recordset("anoproxemi") = Year(Date) + 1
             Else
                data_cli.Recordset("mesproxemi") = Month(Date) + 1
                data_cli.Recordset("anoproxemi") = Year(Date)
             End If
          End If
          If IsNull(data_afilcons.Recordset("codpromo")) = False Then
             If data_afilcons.Recordset("codpromo") = 2 Then
                If Day(Date) >= 25 Then
                   If Month(Date) = 12 Then
                      data_cli.Recordset("mesproxemi") = 1
                      data_cli.Recordset("anoproxemi") = Year(Date) + 2
                   Else
                      data_cli.Recordset("mesproxemi") = Month(Date) + 1
                      data_cli.Recordset("anoproxemi") = Year(Date) + 1
                   End If
                Else
                   If Month(Date) = 12 Then
                      data_cli.Recordset("mesproxemi") = 1
                      data_cli.Recordset("anoproxemi") = Year(Date) + 2
                   Else
                      data_cli.Recordset("mesproxemi") = Month(Date)
                      data_cli.Recordset("anoproxemi") = Year(Date) + 1
                   End If
                End If
             End If
          End If
          data_cli.Recordset("cl_referen") = Mid(data_afilcons.Recordset("correo"), 1, 74)
          If IsNull(data_afilcons.Recordset("codzon")) = False Then
             data_cli.Recordset("cl_grupo") = data_afilcons.Recordset("codzon")
             data_cli.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
          Else
             data_cli.Recordset("cl_grupo") = 0
             data_cli.Recordset("cl_zona") = "*TODOS"
          End If
          data_cli.Recordset("cl_sexo") = data_afilcons.Recordset("sexo")
          MutAfil = Devuelve_mut()
          data_cli.Recordset("cl_socmnom") = Mid(MutAfil, 1, 25)
          If IsNull(data_cli.Recordset("fecha_baja")) = False Then
             data_cli.Recordset("fecha_baja") = Null
          End If
          VendeAfil = Devuelve_vende()
          data_cli.Recordset("cl_nrovend") = data_afilcons.Recordset("codvende")
          data_cli.Recordset("cl_nomvend") = Mid(VendeAfil, 1, 35)
          If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
             data_cli.Recordset("cl_diacobr") = Trim(str(data_afilcons.Recordset("dia_cobro"))) & " C/MES"
          End If
          data_cli.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
          data_cli.Recordset.Update
          data_abm.RecordSource = "select * from abmsocio where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
          data_abm.Refresh
          data_abm.Recordset.AddNew
          data_abm.Recordset("usuario") = WElusuario
          data_abm.Recordset("fecha") = Date
          data_abm.Recordset("hora") = Format(Time, "HH:mm")
          data_abm.Recordset("cl_codigo") = Xmatnew
          data_abm.Recordset("desc") = "MODIF"
          data_abm.Recordset("cl_motivo") = "MODIF DE FICHA"
          data_abm.Recordset("convenio") = data_afilcons.Recordset("categ")
          data_abm.Recordset("base") = data_nrosoc.Recordset("base")
          data_abm.Recordset.Update
       End If
    Else
       MsgBox "No se encuentra matrícula para actualizar.", vbCritical
    End If
End If

End Sub

Public Sub Consulta_AfilRuta()
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
   Xrecclii.MoveFirst
   labcedtit.Caption = Xrecclii("cedula")
Else
   labcedtit.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Function Devuelve_cobradorOk() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from cobrador where cb_numero =" & data_afilcons.Recordset("codcob")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_cobradorOk = Xrecclii("cb_nombre")
Else
   Devuelve_cobradorOk = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

