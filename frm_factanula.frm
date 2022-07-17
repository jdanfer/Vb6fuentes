VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_factanula 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas anuladas del sistema SAPP"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9285
   Icon            =   "frm_factanula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Factura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Servicio"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fec.Anula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Usuario Anula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tipo Fact"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frm_factanula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Xcount As Long
Dim a, b, c, d, e, f, g, h, i, j As String

data1.Connect = "odbc;dsn=" & Xconexrmt & ";"


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
Xcount = 1
ListView1.ListItems.Clear
frm_factanula.MousePointer = 11

data1.RecordSource = "select * from lin_anula where cod_cli =" & frm_estcnv.Label2.Caption
data1.Refresh
If data1.Recordset.RecordCount <> 0 Then
   data1.Recordset.MoveFirst
    Do While Not data1.Recordset.EOF
       If IsNull(data1.Recordset("factura")) = False Then
          ListView1.ListItems.Add Xcount, , data1.Recordset("factura")
       Else
          ListView1.ListItems.Add Xcount, , " "
       End If
       If IsNull(data1.Recordset("fecha")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("fecha")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , " "
       End If
       If IsNull(data1.Recordset("tot_lin")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("tot_lin")
       End If
       If IsNull(data1.Recordset("nom_prod")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("nom_prod")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data1.Recordset("fec_a")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "s/d"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("fec_a")
       End If
       If IsNull(data1.Recordset("usua_a")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("usua_a")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data1.Recordset("pendiente")) = False Then
          If data1.Recordset("pendiente") = "F" Then
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Factura"
          Else
             If data1.Recordset("pendiente") = "T" Then
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Ticket"
             Else
                If data1.Recordset("pendiente") = "N" Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NC e-Factura"
                Else
                   If data1.Recordset("pendiente") = "C" Then
                      ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NC e-Ticket"
                   Else
                      ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Factura"
                   End If
                End If
             End If
          End If
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Factura"
       End If
       data1.Recordset.MoveNext
       Xcount = Xcount + 1
    Loop
    frm_factanula.MousePointer = 0

Else
    MsgBox "No existe historial", vbInformation, "Ver historial"
End If
frm_factanula.MousePointer = 0

End Sub
