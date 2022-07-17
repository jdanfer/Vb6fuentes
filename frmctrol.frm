VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "D:\sapprespaldo\controlemi.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "emis"
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\sapprespaldo\emisiones.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EMI0408"
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\sapprespaldo\emisiones.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EMI0208"
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.MousePointer = 11
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
   Data2.Recordset.FindFirst "cliente =" & Data1.Recordset("cliente")
   If Not Data2.Recordset.NoMatch Then
   
   Else
      Data3.Recordset.AddNew
      Data3.Recordset("matricula") = Data1.Recordset("cliente")
      Data3.Recordset("nombre") = Mid(Data1.Recordset("apellidos"), 1, 50)
      Data3.Recordset("cnv") = Data1.Recordset("cod_cnv")
      Data3.Recordset("cobr") = Data1.Recordset("nro_cobr")
      
      Data3.Recordset.Update
   End If
   Data1.Recordset.MoveNext
Loop
MsgBox "fin"
Form1.MousePointer = 0

End Sub
