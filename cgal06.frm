VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Data cgal06 
      Caption         =   "cgal06"
      Connect         =   "Access"
      DatabaseName    =   "D:\sapprespaldo\cgal.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CGAL06"
      Top             =   1440
      Width           =   2700
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\sapprespaldo\cgal.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CGAL"
      Top             =   480
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.MousePointer = 11
'Data1.Recordset.MoveFirst
'Do While Not Data1.Recordset.EOF
'   Data1.Recordset.Edit
'   Data1.Recordset("cednum") = Val(Data1.Recordset("ced2"))
'   Data1.Recordset.Update
'   Data1.Recordset.MoveNext
'Loop
cgal06.Recordset.MoveFirst
Do While Not cgal06.Recordset.EOF
   Data1.Recordset.FindFirst "ci =" & cgal06.Recordset("cednum")
   If Not Data1.Recordset.NoMatch Then
      cgal06.Recordset.Edit
      cgal06.Recordset("nota") = ""
      cgal06.Recordset.Update
   Else
      cgal06.Recordset.Edit
      cgal06.Recordset("nota") = "NO ENCONTRADO"
      cgal06.Recordset.Update
   End If
   cgal06.Recordset.MoveNext
Loop
Form1.MousePointer = 0
MsgBox "Terminado"

End Sub
