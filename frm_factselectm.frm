VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_factselectm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de pedido medicación a facturar"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9720
   Icon            =   "frm_factselectm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_ctrolcant 
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data data_pedidos 
      Caption         =   "data_pedidos"
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
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      Picture         =   "frm_factselectm.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Registrar selección de pedido como CANCELADO"
      Top             =   4680
      Width           =   495
   End
   Begin VB.Data data_codcaja 
      Caption         =   "data_codcaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton b_borrauna 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frm_factselectm.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Borrar una línea o todo"
      Top             =   7560
      Width           =   375
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Realizar factura"
      Height          =   495
      Left            =   3240
      Picture         =   "frm_factselectm.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Vuelve a la pantalla anterior para terminar la factura"
      Top             =   7560
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_factselectm.frx":1628
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "frm_factselectm.frx":163F
      TabIndex        =   6
      Top             =   5160
      Width           =   9135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingrese el importe correspondiente que aplicará a cada registro seleccionado x 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   9135
      Begin VB.TextBox t_canti 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4440
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox t_importe 
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
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   7320
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox t_codigo 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Cantidad:"
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
         Height          =   375
         Left            =   3360
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Importe:"
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
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Código:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frm_factselectm.frx":26D2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Marcar/Desmarcar todos"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      Picture         =   "frm_factselectm.frx":2C5C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   4680
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Base"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cédula"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CodMed"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Medicamento"
         Object.Width           =   6421
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Id"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label labveriftot 
      Height          =   375
      Left            =   960
      TabIndex        =   25
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labunload 
      Height          =   255
      Left            =   8400
      TabIndex        =   20
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labvaltimbre 
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labtipodoc 
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labnomprod 
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label labcodced 
      Height          =   255
      Left            =   9120
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labconvenio 
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labmatric 
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labcedula 
      Height          =   255
      Left            =   7440
      TabIndex        =   12
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Seleccione la medicación a facturar. Ingrese el importe y luego presione el botón de Aceptar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Pedidos pendientes de facturar del socio seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frm_factselectm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_borrauna_Click()
Dim Deseaborrar As String
Deseaborrar = MsgBox("Desea borrar todos los registros?", vbInformation + vbYesNo, "Facturación")

If Deseaborrar = vbYes Then
   frm_factura.labtot.Caption = 0
   frm_factura.Label8.Caption = 0
   If data_lin.Recordset.RecordCount > 0 Then
      data_lin.Recordset.MoveFirst
      Do While Not data_lin.Recordset.EOF
         data_lin.Recordset.Delete
         data_lin.Recordset.MoveNext
      Loop
   End If
   data_lin.Refresh
   data_lindbgri.RecordSource = "select * from lineas"
   data_lindbgri.Refresh
   data_lindbgri.RecordSource = ""
End If

End Sub

Private Sub b_cancela_Click()

Dim XNroPedido As Long

Dim Xcantveces, Xind, Xcandelin As Integer

XNroPedido = 0

If WElusuario = "COMPUTOS" Then
    If Verificar_seleccion() = 1 Then
       frm_factselectm.MousePointer = 11
       For Xind = 1 To ListView1.ListItems.count
           ListView1.ListItems(Xind).Selected = True
           If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
              XNroPedido = Val(ListView1.SelectedItem.ListSubItems(6).Text)
              data_pedidos.RecordSource = "select * from pedidos_facturar where id =" & XNroPedido
              data_pedidos.Refresh
              If data_pedidos.Recordset.RecordCount > 0 Then
                 data_pedidos.Recordset.Edit
                 data_pedidos.Recordset("fecha_fact") = Date
                 data_pedidos.Recordset("nro_factura") = 1987
                 data_pedidos.Recordset("usuario") = WElusuario
                 data_pedidos.Recordset.Update
              End If
           End If
       Next Xind
       frm_factselectm.MousePointer = 0
       MsgBox "Proceso terminado."
    Else
       MsgBox "Verifique si hay datos seleccionados.", vbCritical, "Facturar"
    End If
Else
    MsgBox "Usuario no habilitado", vbCritical
End If
'Exit Sub

'Vererror:
'         If Err.Number = 3421 Then
'            MsgBox "Verifique datos ingresados", vbCritical, "Mensaje"
'            If DBCombo1.Enabled = True Then
'               DBCombo1.SetFocus
'            End If
'         Else
'            MsgBox "Hay un error en los datos de la factura ANOTE ERROR Y ENVIE A COMPUTOS!" & Err.Description, vbCritical, "Mensaje"
'            If DBCombo1.Enabled = True Then
'               DBCombo1.SetFocus
'            End If
'         End If



End Sub

Private Sub Command1_Click()

Dim Xrub As Long
Dim Xiva, Xvaltimme, Xtotimporte, Xtotimportesintim As Double

Dim XValtim As Long
Dim Xlafenaci As String
Dim Xnomedica, Xlin997, Xcuotabase, Xquerr As Integer
Dim Xnroform As String
Dim Xestco, Xestfl As Long
Dim Xmotivoref As String
Dim Xtelcmt As String
Dim XcantidadMed, XcantidadMed2 As Integer
Dim XimporteUnit As Double
Dim X As Integer

Dim Xcantveces, Xind, Xcandelin As Integer

XimporteUnit = 0
Xcantveces = 0
Xcandelin = 0
Xtotimporte = 0
Xtotimportesintim = 0
XcantidadMed = 0
XcantidadMed2 = 0
X = 0

Command1.Enabled = False
'''''''''''''''''''''''''*******************MULTIPLICAR POR CANTIDADES******************
If Trim(t_importe.Text) = "" Then
   t_importe.Text = 0
End If

If t_codigo.Text = 60103 Or t_codigo.Text = 60107 Then
   If labtipodoc.Caption = "E-TICKET" Then
      Xcantveces = 1
   End If
Else
   If t_codigo.Text = 60106 Or t_codigo.Text = 60108 Then
      If labtipodoc.Caption = "RECIBO" Then
         Xcantveces = 1
      End If
   End If
End If
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveLast
   Xcandelin = data_lin.Recordset.RecordCount
   data_lin.Recordset.MoveFirst
End If

If Verificar_seleccion() = 1 And Trim(t_codigo.Text) <> "" And Xcantveces = 1 Then
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
          XcantidadMed = 1
          If Trim(t_canti.Text) = "" Then
             XcantidadMed2 = Val(ListView1.SelectedItem.ListSubItems(5).Text)
          Else
             XcantidadMed2 = Val(t_canti.Text)
          End If
          X = 1
          If XcantidadMed2 > 0 Then
             For X = 1 To XcantidadMed2
''''             Do While XcantidadMed2 <= x
                Xvaltimme = 0
                XValtim = 0
                Xcandelin = Xcandelin + 1
                frm_factura.data_lineas.Recordset.AddNew
                frm_factura.data_lineas.Recordset("numero") = t_codigo.Text 'codigo serv ??
                frm_factura.data_lineas.Recordset("reg_cab") = 99
                frm_factura.data_lineas.Recordset("factura") = 0
                frm_factura.data_lineas.Recordset("nro_pedido") = Val(ListView1.SelectedItem.ListSubItems(6).Text)
                If t_importe.Text > 0 Then
                    frm_factura.data_lineas.Recordset("tipo_mov") = 2 'tipo de iva (indic de fact)
                Else
                    frm_factura.data_lineas.Recordset("tipo_mov") = 5
                End If
                frm_factura.data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                frm_factura.data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                frm_factura.data_lineas.Recordset("cod_cli") = Val(labmatric.Caption)
                frm_factura.data_lineas.Recordset("nom_cli") = Mid(frm_factura.labnomb.Caption, 1, 30)
                frm_factura.data_lineas.Recordset("cod_prod") = t_codigo.Text
                frm_factura.data_lineas.Recordset("nom_prod") = labnomprod.Caption
                frm_factura.data_lineas.Recordset("cantidad") = XcantidadMed
                frm_factura.data_lineas.Recordset("moneda") = "SR" 'Serie
                If frm_factura.txt_rut.Visible = True Then
                   If Trim(frm_factura.txt_rut.Text) <> "" Then
                      frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
                   End If
                End If
                If t_codigo.Text = 60106 Or t_codigo.Text = 60108 Then
                   frm_factura.data_lineas.Recordset("tipo_mov") = 1
                End If
                frm_factura.data_lineas.Recordset("cod_medic") = Val(ListView1.SelectedItem.ListSubItems(3).Text)
                frm_factura.data_lineas.Recordset("idtablapres") = 0
                frm_factura.data_lineas.Recordset("nom_medic") = Mid(ListView1.SelectedItem.ListSubItems(4).Text, 1, 50)
                frm_factura.data_lineas.Recordset("dias") = 0
                frm_factura.data_lineas.Recordset("operador") = WElusuario
                frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
                frm_factura.data_lineas.Recordset("nro_flia") = 6
                frm_factura.data_lineas.Recordset("nom_flia") = "MEDICAMENTOS"
                frm_factura.data_lineas.Recordset("convenio") = labconvenio.Caption
                frm_factura.data_lineas.Recordset("ced_socio") = Val(labcedula.Caption)
                frm_factura.data_lineas.Recordset("fact") = Val(labcodced.Caption)
                If Xfpago = 2 Then
                   frm_factura.data_lineas.Recordset("rub_cont") = data_parsec.Recordset("srvcrd")
                   Xrub = data_parsec.Recordset("srvcrd")
                Else
                   frm_factura.data_lineas.Recordset("rub_cont") = data_parsec.Recordset("srvcnt")
                   Xrub = data_parsec.Recordset("srvcnt")
                End If
                data_codcaja.RecordSource = "select * from cod_caja where numero =" & Xrub
                data_codcaja.Refresh
                If data_codcaja.Recordset.RecordCount > 0 Then
                   frm_factura.data_lineas.Recordset("rub_nomb") = data_codcaja.Recordset("nombre")
                End If
                If t_codigo.Text = 60106 Then
                   frm_factura.data_lineas.Recordset("rub_cont") = 211397
                   frm_factura.data_lineas.Recordset("rub_nomb") = "M. SMI"
                End If
                If t_codigo.Text = 60105 Then
                   frm_factura.data_lineas.Recordset("rub_cont") = 211397
                   frm_factura.data_lineas.Recordset("rub_nomb") = "M. SMI"
                End If
                If t_codigo.Text = 60108 Then
                   frm_factura.data_lineas.Recordset("rub_cont") = 211302
                   frm_factura.data_lineas.Recordset("rub_nomb") = "M.UNIVERSAL"
                End If
                If (t_codigo.Text = 60107 Or t_codigo.Text = 60103) And labtipodoc.Caption = "E-TICKET" Then
                   If t_importe.Text > 0 Then
                      If Trim(frm_factura.labtimme.Caption) = "" Then
                         Xvaltimme = Val(labvaltimbre.Caption) * XcantidadMed
                         frm_factura.labtimme.Caption = Val(Xvaltimme)
                      Else
                         Xvaltimme = Val(labvaltimbre.Caption) * XcantidadMed
                         frm_factura.labtimme.Caption = Val(frm_factura.labtimme.Caption) + Val(Xvaltimme)
                      End If
                   End If
                   Xtotimporte = t_importe.Text * XcantidadMed
                   If CDbl(Xtotimporte) >= CDbl(Xvaltimme) Then
                      XimporteUnit = Xtotimporte - Xvaltimme
                      XimporteUnit = XimporteUnit / XcantidadMed
                      frm_factura.data_lineas.Recordset("arancel") = Format(XimporteUnit, "Standard")
                      frm_factura.data_lineas.Recordset("tot_lin") = Format(Xtotimporte - Xvaltimme, "Standard")
                      Xtotimportesintim = Xtotimporte - Xvaltimme
                   Else
                      MsgBox "Importe menor al timbre, VERIFIQUE!!!", vbCritical
                      Unload Me
                      Exit Sub
                   End If
                Else
                   Xtotimporte = t_importe.Text * XcantidadMed
                   frm_factura.data_lineas.Recordset("arancel") = Format(Xtotimporte, "Standard")
                   frm_factura.data_lineas.Recordset("tot_lin") = Format(Xtotimporte, "Standard")
                End If
                Xiva = Xtotimportesintim / 1.1
                Xiva = Xiva * 0.1
                frm_factura.data_lineas.Recordset("imp_iva") = Format(Xiva, "Standard")
                If frm_factura.Label8.Caption = "" Then
                   frm_factura.Label8.Caption = Format(Xiva, "Standard")
                Else
                   frm_factura.Label8.Caption = CDbl(frm_factura.Label8.Caption) + Xiva
                   frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
                End If
                If t_codigo.Text = 60106 Or t_codigo.Text = 60108 Then
                   frm_factura.data_lineas.Recordset("imp_iva") = 0
                   Xiva = 0
                   frm_factura.Label8.Caption = 0
                   frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
                End If
                If (t_codigo.Text = 60107 Or t_codigo.Text = 60103) And labtipodoc.Caption = "E-TICKET" Then
                   If Xtotimporte > 0 Then
                      If Val(Xtotimporte) >= Val(Xvaltimme) Then
                         frm_factura.data_lineas.Recordset("precio_est") = Format(Xtotimporte - Xvaltimme, "Standard")
                      Else
                         frm_factura.data_lineas.Recordset("precio_est") = Format(Xtotimporte, "Standard")
                      End If
                   Else
                      frm_factura.data_lineas.Recordset("precio_est") = Format(Xtotimporte, "Standard")
                   End If
                Else
                   frm_factura.data_lineas.Recordset("precio_est") = Format(Xtotimporte, "Standard")
                End If
                frm_factura.data_lineas.Recordset("porce_est") = 0
                frm_factura.data_lineas.Recordset("base") = data_parsec.Recordset("base")
                If frm_factura.labtot.Caption = "" Then
                Else
                   If (t_codigo.Text = 60107 Or t_codigo.Text = 60103) And labtipodoc.Caption = "E-TICKET" Then
                      If Xtotimporte > 0 Then
                         If Val(Xtotimporte) >= Val(Xvaltimme) Then
                            If Xiva = 0 Then
                               frm_factura.data_lineas.Recordset("costo") = Xtotimporte - Xvaltimme
                            Else
                               frm_factura.data_lineas.Recordset("costo") = Xtotimporte - Xvaltimme - Xiva
                            End If
                         Else
                            If Xiva = 0 Then
                               frm_factura.data_lineas.Recordset("costo") = Xtotimporte
                            Else
                               frm_factura.data_lineas.Recordset("costo") = Xtotimporte - Xiva
                            End If
                         End If
                      Else
                         If Xiva = 0 Then
                            frm_factura.data_lineas.Recordset("costo") = Xtotimporte
                         Else
                            frm_factura.data_lineas.Recordset("costo") = Xtotimporte - Xiva
                         End If
                      End If
                   Else
                      If Xiva = 0 Then
                         frm_factura.data_lineas.Recordset("costo") = Xtotimporte
                      Else
                         frm_factura.data_lineas.Recordset("costo") = Xtotimporte - Xiva
                      End If
                   End If
                End If
                If Xfpago = 2 Then
                   frm_factura.data_lineas.Recordset("tipo") = "CREDITO"
                Else
                   frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
                End If
                If frm_factura.labtot.Caption <> "" Then
                   frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + Val(Xtotimporte)
                   frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
                Else
                   frm_factura.labtot.Caption = Format(Xtotimporte, "Standard")
                   frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
                End If
                frm_factura.data_lineas.Recordset("linea") = Xcandelin
                frm_factura.data_lineas.Recordset("libro_rub") = labtipodoc.Caption ' tipo de documento (Ej.e-ticket)
                frm_factura.data_lineas.Recordset("in_unid") = "INT1"
                frm_factura.data_lineas.Recordset.Update
                frm_factura.data_lineas.Refresh
                data_lin.Refresh
                If (t_codigo.Text = 60107 Or t_codigo.Text = 60103) And Xtotimporte > 0 And labtipodoc.Caption = "E-TICKET" Then
                    Xcandelin = Xcandelin + 1
                    frm_factura.data_lineas.Recordset.AddNew
                    frm_factura.data_lineas.Recordset("reg_cab") = 0
                    frm_factura.data_lineas.Recordset("factura") = 0
                    frm_factura.data_lineas.Recordset("tipo_mov") = 1
                    frm_factura.data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
                    frm_factura.data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                    frm_factura.data_lineas.Recordset("cod_cli") = Val(labmatric.Caption)
                    frm_factura.data_lineas.Recordset("nom_cli") = Mid(frm_factura.labnomb.Caption, 1, 30)
                    frm_factura.data_lineas.Recordset("cod_prod") = 990
                    frm_factura.data_lineas.Recordset("nom_prod") = "TIMBRES PROFESIONAL M"
                    frm_factura.data_lineas.Recordset("usa_timbre") = "M"
                    frm_factura.data_lineas.Recordset("moneda") = "SR" 'Serie
                    If frm_factura.txt_rut.Visible = True Then
                       frm_factura.data_lineas.Recordset("ruc") = txt_rut.Text
                    End If
                    frm_factura.data_lineas.Recordset("cantidad") = XcantidadMed
                    frm_factura.data_lineas.Recordset("operador") = WElusuario
                    frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
                    frm_factura.data_lineas.Recordset("nro_flia") = 8
                    frm_factura.data_lineas.Recordset("nom_flia") = "OTROS SERVICIOS"
                    frm_factura.data_lineas.Recordset("convenio") = labconvenio.Caption
                    frm_factura.data_lineas.Recordset("rub_cont") = 213076
                    frm_factura.data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
                    frm_factura.data_lineas.Recordset("arancel") = Val(labvaltimbre.Caption)
                    frm_factura.data_lineas.Recordset("tot_lin") = Xvaltimme
                    frm_factura.data_lineas.Recordset("precio_est") = Xvaltimme
                    frm_factura.data_lineas.Recordset("porce_est") = 0
                    frm_factura.data_lineas.Recordset("base") = data_parsec.Recordset("base")
                    If Xfpago = 2 Then
                       frm_factura.data_lineas.Recordset("tipo") = "CREDITO"
                    Else
                       frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
                    End If
                    frm_factura.data_lineas.Recordset("linea") = Xcandelin
                    frm_factura.data_lineas.Recordset("libro_rub") = labtipodoc.Caption ' tipo de documento (Ej.e-ticket)
                    frm_factura.data_lineas.Recordset("in_unid") = "INT1"
                    frm_factura.data_lineas.Recordset.Update
                    frm_factura.data_lineas.Refresh
                    data_lin.Refresh
                    frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + Xvaltimme
                    t_canti.Text = ""
                End If
             Next X
          Else
             MsgBox "Falta dato de cantidad de medicación, VERIFIQUE!!", vbCritical, "FACTURAR"
          End If
       End If
   Next Xind
Else
   MsgBox "Verifique datos de facturación.", vbCritical, "Facturar"
End If
Command1.Enabled = True
'Exit Sub

'Vererror:
'         If Err.Number = 3421 Then
'            MsgBox "Verifique datos ingresados", vbCritical, "Mensaje"
'            If DBCombo1.Enabled = True Then
'               DBCombo1.SetFocus
'            End If
'         Else
'            MsgBox "Hay un error en los datos de la factura ANOTE ERROR Y ENVIE A COMPUTOS!" & Err.Description, vbCritical, "Mensaje"
'            If DBCombo1.Enabled = True Then
'               DBCombo1.SetFocus
'            End If
'         End If

End Sub

Private Sub Command2_Click()
Dim Xind As Integer

For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = False
    Else
       ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
    End If
Next Xind

End Sub

Private Sub Command3_Click()
Verifica_totalesCant
If labveriftot.Caption = "OK" Then
    If frm_factura.data_lineas.Recordset.RecordCount > 0 Then
       frm_factura.data_lineas.Recordset.MoveFirst
       frm_factura.labtot.Caption = 0
       frm_factura.Label8.Caption = 0
       Do While Not frm_factura.data_lineas.Recordset.EOF
          frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + frm_factura.data_lineas.Recordset("tot_lin")
          frm_factura.Label8.Caption = Val(frm_factura.Label8.Caption) + frm_factura.data_lineas.Recordset("imp_iva")
          frm_factura.data_lineas.Recordset.MoveNext
       Loop
       frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
       frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
    End If
    frm_factura.data_lindbgri.RecordSource = "select * from lineas"
    frm_factura.data_lindbgri.Refresh
    frm_factura.data_lindbgri.RecordSource = ""
    frm_factura.btn_graba.Enabled = False
    'frm_factura.btn_fin.SetFocus
    
    labunload.Caption = "1"
    
    Unload Me
Else
    MsgBox "Hay un error en las cantidades de la factura con relación al pedido, VERIFIQUE!", vbCritical
    
End If

End Sub

Private Sub Form_Load()
labunload.Caption = "2"

data_parsec.DatabaseName = App.path & "\parse.mdb"
data_parsec.RecordSource = "parsec0"
data_parsec.Refresh

data_codcaja.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_pedidos.Connect = "odbc;dsn=" & Xconexrmt & ";"

'data_codcaja.RecordSource = "cod_caja"
'data_codcaja.Refresh

data_lin.DatabaseName = App.path & "\factura.mdb"
data_lin.RecordSource = "select * from lineas"
data_lin.Refresh

labmatric.Caption = frm_factura.labmatri.Caption

If Trim(labmatric.Caption) = "" Then
   MsgBox "No hay matrícula seleccionada. VERIFIQUE!!", vbCritical
   Command1.Enabled = False
   Command3.Enabled = False
Else
   Retorna_cedula
   Carga_grid
'''   Codigo_facturacion
   t_codigo.Text = Val(frm_factura.Label5.Caption)
   labnomprod.Caption = frm_factura.DBCombo1.Text
   frm_factura.txt_precio.Text = ""
   frm_factura.Label5.Caption = ""
   frm_factura.DBCombo1.Text = ""
   labtipodoc.Caption = frm_factura.Label7.Caption
   t_importe.Text = 0
   Devuelve_valores
   Contar_totales
End If

End Sub


Public Sub Carga_grid()
Dim Xcount As Long
Dim a, b, c, d, e, f, g, h, i, j, k As String
a = "a"
b = "b"
c = "c"
d = "d"
e = "e"
f = "f"
g = "g"
h = "h"
i = "i"
j = "j"
k = "k"
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset


If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD

ConbdSapp.Open
             
Xsqlpromo = "Select * from pedidos_facturar where matricula =" & Val(labcedula.Caption) & " and fecha_fact is null and cantidad >" & 0
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xcount = 1
ListView1.ListItems.Clear

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("fecha")) = False Then
         ListView1.ListItems.Add Xcount, , Format(Xrecclii("fecha"), "dd/mm/yyyy")
      Else
         ListView1.ListItems.Add Xcount, , " "
      End If
      If IsNull(Xrecclii("base")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_parsec.Recordset("base")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , " "
      End If
      If IsNull(Xrecclii("matricula")) = True Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN DATOS"
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("matricula")
      End If
      If IsNull(Xrecclii("cod_medicacion")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("cod_medicacion")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
      End If
      If IsNull(Xrecclii("nom_medicacion")) = True Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("nom_medicacion")
      End If
      If IsNull(Xrecclii("cantidad")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("cantidad")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
      End If
      If IsNull(Xrecclii("id")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("id")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
      End If
      Xrecclii.MoveNext
      Xcount = Xcount + 1
   Loop
Else
    MsgBox "No existe pedido de medicación.", vbInformation, "Pedidos"
End If

Xrecclii.Close
ConbdSapp.Close


End Sub


Public Function Verificar_seleccion() As Integer
Dim Xind, Xsihay As Integer
Xsihay = 0
For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       Xsihay = 1
    End If
Next Xind
Verificar_seleccion = Xsihay


End Function

Public Sub Codigo_facturacion()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from convenio where cnv_codigo ='" & labconvenio.Caption & "'"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cnv_grupo")) = False Then
      If Xrecclii("cnv_grupo") = "CCOU" Then
         t_codigo.Text = 60103
         labnomprod.Caption = "CCOU M."
      Else
         If Xrecclii("cnv_grupo") = "H.EVANGELICO" Then
            t_codigo.Text = 60107
            labnomprod.Caption = "EVANGELICO M."
         Else
            If Xrecclii("cnv_grupo") = "SMI" Then
               t_codigo.Text = 60106
               labnomprod.Caption = "SMI M."
            Else
               If Xrecclii("cnv_grupo") = "UNIVERSAL" Then
                  t_codigo.Text = 60108
                  labnomprod.Caption = "UNIVERSAL M."
               Else
                  t_codigo.Text = ""
                  labnomprod.Caption = ""
               End If
            End If
         End If
      End If
   Else
      t_codigo.Text = ""
      labnomprod.Caption = ""
   End If
Else
   t_codigo.Text = ""
   labnomprod.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Retorna_cedula()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from clientes where cl_codigo =" & labmatric.Caption

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labcodced.Caption = Xrecclii("cl_codced")
   labcedula.Caption = Xrecclii("cl_cedula")
   labconvenio.Caption = Xrecclii("cl_codconv")
Else
   labcodced.Caption = ""
   labcedula.Caption = ""
   labconvenio.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Devuelve_valores()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xdescuento As Double
Dim XPorcen As Integer
Dim XPrec As Integer
Xdescuento = 0
XPorcen = 0
XPrec = 0
If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from estudios where codest =" & 990

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labvaltimbre.Caption = Xrecclii("cons")
Else
   labvaltimbre.Caption = ""
End If

Xrecclii.Close

Xsqlpromo = "Select * from estudios where codest =" & t_codigo.Text

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   t_importe.Text = Xrecclii("cons")
Else
   t_importe.Text = 0
End If
Xrecclii.Close

Xsqlpromo = "Select * from Aran_servicios where id_gpo =" & Xop1 & " and id_serv =" & t_codigo.Text & " and prec_serv is not null"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   XPrec = Xrecclii("prec_serv")
   XPorcen = Xrecclii("por_serv")
   If XPrec > 0 Then
      t_importe.Text = Xrecclii("prec_serv")
   Else
      If XPorcen = 100 Then
'         t_importe.Text = 0
      Else
         If XPorcen <> 0 Then
            Xdescuento = Xrecclii("por_serv") * Val(t_importe.Text) / 100
            t_importe.Text = t_importe.Text - Xdescuento
         End If
      End If
   End If
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Val(labunload.Caption) = 2 Then
   frm_factura.labtot.Caption = 0
   frm_factura.Label8.Caption = 0
   If data_lin.Recordset.RecordCount > 0 Then
      data_lin.Recordset.MoveFirst
      Do While Not data_lin.Recordset.EOF
         data_lin.Recordset.Delete
         data_lin.Recordset.MoveNext
      Loop
   End If
End If
End Sub

Public Sub Contar_totales()
Dim CantidadTot As Integer
CantidadTot = 0
For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       CantidadTot = CantidadTot + Val(ListView1.SelectedItem.ListSubItems(5).Text)
    End If
    ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = False

Next Xind
t_ctrolcant.Text = CantidadTot

End Sub

Public Sub Verifica_totalesCant()
Dim VerifCantidad As Integer
VerifCantidad = 0
If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       If data_lin.Recordset("cod_prod") = 60103 Or data_lin.Recordset("cod_prod") = 60106 Or _
          data_lin.Recordset("cod_prod") = 60107 Or data_lin.Recordset("cod_prod") = 60108 Then
          VerifCantidad = VerifCantidad + 1
       End If
       data_lin.Recordset.MoveNext
    Loop
    
    If Val(t_ctrolcant.Text) = Val(VerifCantidad) Then
       labveriftot.Caption = "OK"
    Else
       labveriftot.Caption = "ERR"
    End If
    data_lin.Recordset.MoveFirst
End If

End Sub

