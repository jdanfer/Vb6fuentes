VERSION 5.00
Begin VB.Form frm_pasoemi 
   Caption         =   "frm_pasoemi"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_emitiqok 
      Caption         =   "data_emitiqok"
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data data_emitiqmal 
      Caption         =   "data_emitiqmal"
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
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data data_rectiq 
      Caption         =   "data_rectiq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "acepto"
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frm_pasoemi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
data_emi.DatabaseName = App.Path & "\emisiones.mdb"
data_emi.RecordSource = "EMI1108"
data_emi.Refresh
data_emi.Recordset.MoveFirst
data_rectiq.DatabaseName = App.Path & "\env_tiq.mdb"
data_rectiq.RecordSource = "emiserv"
data_rectiq.Refresh

data_emitiqmal.DatabaseName = App.Path & "\emimal.mdb"
data_emitiqmal.RecordSource = "EMI1108"
data_emitiqmal.Refresh

data_emitiqok.DatabaseName = App.Path & "\emitiqok.mdb"
data_emitiqok.RecordSource = "EMI1108"
data_emitiqok.Refresh

Dim XIVA As Double
Dim TOT As Double
data_emitiqok.Recordset.MoveFirst
Do While Not data_emitiqok.Recordset.EOF
   TOT = data_emitiqok.Recordset("importe") + data_emitiqok.Recordset("deudas")
   XIVA = TOT / 1.1
   XIVA = XIVA * 0.1
   data_emitiqok.Recordset.Edit
   data_emitiqok.Recordset("iva") = Format(XIVA, "Standard")
   TOT = data_emitiqok.Recordset("importe") + data_emitiqok.Recordset("deudas") + data_emitiqok.Recordset("tiquet") + data_emitiqok.Recordset("servi")
   data_emitiqok.Recordset("total") = TOT
   data_emitiqok.Recordset.Update
   data_emitiqok.Recordset.MoveNext
Loop
'data_emitiqok.Recordset.MoveFirst
'Do While Not data_emitiqok.Recordset.EOF
'   data_emi.Recordset.FindFirst "cliente =" & data_emitiqok.Recordset("cliente")
'   If Not data_emi.Recordset.NoMatch Then
'      data_emi.Recordset.Delete
'   End If
'   data_emitiqok.Recordset.MoveNext
   
'Loop

'Do While Not data_emi.Recordset.EOF
'   data_emi.Recordset.Edit
'   data_emi.Recordset("tiquet") = 0
'   data_emi.Recordset("servi") = 0
'   data_emi.Recordset.Update
'   data_emi.Recordset.MoveNext
'Loop
'MsgBox "Fin"
'data_rectiq.Recordset.MoveFirst
'Do While Not data_rectiq.Recordset.EOF
'   data_emi.Recordset.FindFirst "cliente =" & data_rectiq.Recordset("mat")
'   If Not data_emi.Recordset.NoMatch Then
'      data_emi.Recordset.Edit
'      data_emi.Recordset("servi") = data_emi.Recordset("servi") + data_rectiq.Recordset("imp")
'      data_emi.Recordset("total") = data_emision.Recordset("total") + data_rectiq.Recordset("imp")
'      data_emi.Recordset.Update
'   End If
'   data_rectiq.Recordset.MoveNext
'Loop
'MsgBox "Fin"
'data_emitiqok.Recordset.MoveFirst
'Do While Not data_emitiqok.Recordset.EOF
'   data_emitiqmal.Recordset.FindFirst "cliente =" & data_emitiqok.Recordset("cliente")
'   If Not data_emitiqmal.Recordset.NoMatch Then
'      data_emitiqmal.Recordset.Delete
'   End If
'   data_emitiqok.Recordset.MoveNext
   
'Loop
'data_emitiqmal.Recordset.MoveFirst
'Do While Not data_emitiqmal.Recordset.EOF
'   data_emitiqmal.Recordset.Edit
'   data_emitiqmal.Recordset("tiquet") = 0
'   data_emitiqmal.Recordset("servi") = 0
'   data_emitiqmal.Recordset.Update
'   data_emitiqmal.Recordset.MoveNext
'Loop



MsgBox "FIN"

End Sub
