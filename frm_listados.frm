VERSION 5.00
Begin VB.Form frm_listados 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Socios eliminados"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "FECHA"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "USUARIO"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "CEDULA"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "NOMBRES"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "MATRICULA"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm_listados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\borrado.mdb"
Data1.RecordSource = "select * from infcli order by cl_codigo"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      List1.AddItem Data1.Recordset("cl_codigo") & " " & Data1.Recordset("cl_apellid") & " " & Data1.Recordset("cl_cedula") & " " & Data1.Recordset("cl_nombre") & " " & Data1.Recordset("cl_fultpag")
      Data1.Recordset.MoveNext
   Loop
End If

End Sub
