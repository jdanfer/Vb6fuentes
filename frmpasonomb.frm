VERSION 5.00
Begin VB.Form frmpasonomb 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Paso nombres"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_clientes 
      Caption         =   "data_clientes"
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.Data data_climod 
      Caption         =   "data_climod"
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
      Top             =   2160
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar..."
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar..."
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmpasonomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmpasonomb.MousePointer = 11
Command1.Enabled = False
data_climod.Recordset.MoveFirst

Do While Not data_climod.Recordset.EOF
   data_clientes.Recordset.FindFirst "cl_codigo =" & data_climod.Recordset("cl_codigo")
   If Not data_clientes.Recordset.NoMatch Then
      data_clientes.Recordset.Edit
      data_clientes.Recordset("cl_apellid") = Mid(data_climod.Recordset("cl_apellid"), 1, 32) + " " + Mid(data_climod.Recordset("cl_nombre"), 1, 27)
      data_clientes.Recordset("cl_email") = Mid(data_climod.Recordset("cl_apellid"), 1, 30)
      data_clientes.Recordset("cl_nombre") = Mid(data_climod.Recordset("cl_nombre"), 1, 30)
      data_clientes.Recordset.Update
   End If
   data_climod.Recordset.MoveNext
Loop
frmpasonomb.MousePointer = 0
MsgBox "Proceso terminado"
End

End Sub

Private Sub Form_Load()
data_climod.DatabaseName = App.Path & "\sappnomb.mdb"
data_climod.RecordSource = "clientes"
data_climod.Refresh
data_clientes.DatabaseName = App.Path & "\sapp.mdb"
data_clientes.RecordSource = "clientes"
data_clientes.Refresh


End Sub
