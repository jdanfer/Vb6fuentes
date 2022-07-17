VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xu, Xa As String
Dim Xf As Date
Dim Xcon, Xm As Long
Xu = "JFERNAN"
Xa = "C"
Xf = CDate("27/04/2009")
Command1.Enabled = False
Data1.DatabaseName = App.Path & "\sapp.mdb"
'Data1.RecordSource = "Select * from arqueo where arqueo ='" & Xa & "' and usuario ='" & Xu & "'"
Data1.RecordSource = "Select * from arqueo where mes =" & 4 & " and ano =" & 2009 & " order by matricula"
Data1.Refresh
''Data1.Recordset.FindFirst "matricula =" & 6026770 & " and fecha =#" & Format(Xf, "yyyy/mm/dd") & "# and usuar ='" & Xu & "' and mes =" & 4 & " And ano =" & 2009
''If Not Data1.Recordset.NoMatch Then
''   Data1.Recordset.Delete
''End If
Data1.Recordset.MoveFirst
Xm = 0
Xu = Data1.Recordset("cat")
Do While Not Data1.Recordset.EOF
   If Data1.Recordset("matricula") = Xm And Data1.Recordset("cat") = Xu Then
      If Data1.Recordset("arqueo") = "C" Then
         Data1.Recordset.Edit
         Data1.Recordset("cob") = 151
         Data1.Recordset.Update
      End If
   End If
   Xm = Data1.Recordset("matricula")
   Xu = Data1.Recordset("cat")
   Data1.Recordset.MoveNext
   
Loop
MsgBox "Terminado"

End Sub
