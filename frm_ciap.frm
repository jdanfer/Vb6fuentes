VERSION 5.00
Begin VB.Form frm_ciap 
   BackColor       =   &H00808080&
   Caption         =   "CIAP"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   6255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "frm_ciap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Command1_Click

End Sub

Private Sub Command1_Click()

Data1.Recordset.FindFirst "desc ='" & Combo1.Text & "'"
List1.Clear

If Not Data1.Recordset.NoMatch Then
   Dim Xtexd, Xtexh As String
   Xtexd = Data1.Recordset("cod") + "000"
   Xtexh = Data1.Recordset("cod") + "999"
   Data2.RecordSource = "Select * from ciap2 where code >='" & Xtexd & "' and code <='" & Xtexh & "'"
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      Data2.Recordset.MoveFirst
      Do While Not Data2.Recordset.EOF
         List1.AddItem Data2.Recordset("code") & "--" & Data2.Recordset("text")
         Data2.Recordset.MoveNext
      Loop
   End If
End If
List1.SetFocus
List1.ListIndex = 0

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\diagnost.mdb"
Data1.RecordSource = "capciap"
Data1.Refresh

Data2.DatabaseName = App.Path & "\diagnost.mdb"

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Combo1.AddItem Data1.Recordset("desc")
      Data1.Recordset.MoveNext
   Loop
End If

End Sub
