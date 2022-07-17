VERSION 5.00
Begin VB.Form frm_arreglo 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Arreglo"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data clientes 
      Caption         =   "clientes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_cliamod 
      Caption         =   "data_cliamod"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar..."
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar..."
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "frm_arreglo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_arreglo.MousePointer = 11
Command1.Enabled = False
data_cliamod.Recordset.MoveFirst
Do While Not data_cliamod.Recordset.EOF
   clientes.Recordset.FindFirst "cl_codigo =" & data_cliamod.Recordset("cl_codigo")
   If Not clientes.Recordset.NoMatch Then
      clientes.Recordset.Edit
      clientes.Recordset("cl_apellid") = data_cliamod.Recordset("cl_apellid")
      clientes.Recordset.Update
   End If
   data_cliamod.Recordset.MoveNext
Loop
frm_arreglo.MousePointer = 0
MsgBox "Terminado..."

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Form_Load()
data_cliamod.DatabaseName = App.Path & "\clib1.mdb"
data_cliamod.RecordSource = "clib1"
data_cliamod.Refresh
clientes.DatabaseName = App.Path & "\sapp.mdb"
clientes.RecordSource = "clientes"
clientes.Refresh

End Sub
