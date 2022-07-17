VERSION 5.00
Begin VB.Form frm_carga 
   BackColor       =   &H00808080&
   Caption         =   "Diagnosticos"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "CIAP2"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   840
      TabIndex        =   2
      Top             =   3840
      Width           =   8295
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   8415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   8535
   End
End
Attribute VB_Name = "frm_carga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Command1_Click

End Sub

Private Sub Command1_Click()
Data1.Recordset.FindFirst "cap ='" & Combo1.Text & "'"
List1.Clear

If Not Data1.Recordset.NoMatch Then
   Data2.RecordSource = "Select * from cie10 where cod >='" & Data1.Recordset("desde") & "' and cod <='" & Data1.Recordset("hasta") & "'"
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      Data2.Recordset.MoveFirst
      Do While Not Data2.Recordset.EOF
         If Len(Data2.Recordset("cod")) <= 3 Then
            List1.AddItem Data2.Recordset("cod") & "--" & Data2.Recordset("desc")
         End If
         Data2.Recordset.MoveNext
      Loop
   End If
End If
List1.SetFocus
List1.ListIndex = 0

'       If List1.List(List1.ListIndex) =
End Sub

Private Sub Command2_Click()
frm_ciap.Show vbModal

End Sub

Private Sub Form_Load()
Data2.DatabaseName = App.Path & "\diagnost.mdb"

Data1.DatabaseName = App.Path & "\diagnost.mdb"
Data1.RecordSource = "capcie"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Combo1.AddItem Data1.Recordset("cap")
      Data1.Recordset.MoveNext
   Loop
End If


End Sub

Private Sub List1_Click()
Dim Xtextd, Xtexth As String
Xtextd = Mid(List1.List(List1.ListIndex), 1, 3)
Xtextd = Xtextd + "0"
Xtexth = Mid(List1.List(List1.ListIndex), 1, 3) + "9"

List2.Clear

Data2.RecordSource = "Select * from cie10 where cod >='" & Xtextd & "' and cod <='" & Xtexth & "'"
Data2.Refresh
If Data2.Recordset.RecordCount > 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      If Len(Data2.Recordset("cod")) > 3 Then
         List2.AddItem Data2.Recordset("cod") & "--" & Data2.Recordset("desc")
      End If
      Data2.Recordset.MoveNext
   Loop
Else
   List2.AddItem List1.List(List1.ListIndex)
End If

'       If List1.List(List1.ListIndex) =
End Sub
