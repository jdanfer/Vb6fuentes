VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_factselectmd 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de pedido medicación a facturar"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9720
   Icon            =   "frm_factselectmd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_pedidos 
      Caption         =   "data_pedidos"
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
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
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
         TabIndex        =   9
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
         TabIndex        =   7
         Top             =   360
         Width           =   1455
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
         TabIndex        =   8
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
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      Picture         =   "frm_factselectmd.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Marcar/Desmarcar todos"
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8640
      Picture         =   "frm_factselectmd.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   4680
      Width           =   735
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
   Begin VB.Label labunload 
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labvaltimbre 
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labtipodoc 
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labnomprod 
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label labcodced 
      Height          =   255
      Left            =   8160
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labconvenio 
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labmatric 
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labcedula 
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   4920
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
Attribute VB_Name = "frm_factselectmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




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
Dim XcantidadMed As Integer
Dim XimporteUnit As Double

Dim Xind, Xcandelin As Integer
Dim Xcount As Integer

XimporteUnit = 0
Xcandelin = 0
Xtotimporte = 0
Xtotimportesintim = 0

Command1.Enabled = False

Xcount = frm_pedidomedic.ListView1.ListItems.count + 1

If Trim(t_importe.Text) = "" Then
   t_importe.Text = 0
End If
frm_pedidomedic.Data_pedlin.RecordSource = "select * from pedidos_mediclin where cod_pedido =" & Val(frm_pedidomedic.labpedido.Caption)
frm_pedidomedic.Data_pedlin.Refresh

frm_pedidomedic.ListView1.ListItems.Clear

If Verificar_seleccion() = 1 And Trim(t_codigo.Text) <> "" Then
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
          XcantidadMed = Val(ListView1.SelectedItem.ListSubItems(5).Text)
          If XcantidadMed > 0 Then
             frm_pedidomedic.ListView1.ListItems.Add Xcount, , ListView1.SelectedItem.ListSubItems(4).Text
             frm_pedidomedic.ListView1.ListItems.Item(Xcount).ListSubItems.Add , , XcantidadMed
             frm_pedidomedic.ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(t_importe.Text, "Standard")
             XimporteUnit = Val(t_importe.Text) * XcantidadMed
             Xcandelin = Xcandelin + XcantidadMed
             Xtotimporte = Xtotimporte + XimporteUnit
             frm_pedidomedic.labtotcant.Caption = Xcandelin
             frm_pedidomedic.labtotp.Caption = Val(Xtotimporte)
             frm_pedidomedic.Data_pedlin.Recordset.AddNew
             frm_pedidomedic.Data_pedlin.Recordset("cod_pedido") = Val(frm_pedidomedic.labpedido.Caption)
             frm_pedidomedic.Data_pedlin.Recordset("nom_medic") = Trim(ListView1.SelectedItem.ListSubItems(4).Text)
             frm_pedidomedic.Data_pedlin.Recordset("cant") = Val(XcantidadMed)
             frm_pedidomedic.Data_pedlin.Recordset("imp_unit") = Val(t_importe.Text)
             frm_pedidomedic.Data_pedlin.Recordset("tot_imp") = Val(XimporteUnit)
             frm_pedidomedic.Data_pedlin.Recordset.Update
          Else
             MsgBox "Falta dato de cantidad de medicación, VERIFIQUE!!", vbCritical, "FACTURAR"
          End If
       End If
   Next Xind
   Command1.Enabled = True
   Unload Me

Else
   MsgBox "Verifique datos y selección.", vbCritical, "Pedidos"
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

If Trim(frm_pedidomedic.t_mat.Text) <> "" Then
   labmatric.Caption = frm_pedidomedic.t_mat.Text
Else
   labmatric.Caption = ""
End If

If Trim(labmatric.Caption) = "" Then
   MsgBox "No hay matrícula seleccionada. VERIFIQUE!!", vbCritical
   Command1.Enabled = False
Else
   Retorna_cedula
   Carga_grid
   Codigo_facturacion
   t_importe.Text = 0
   Devuelve_valores
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
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("base")
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

ConbdSapp.Close

End Sub

