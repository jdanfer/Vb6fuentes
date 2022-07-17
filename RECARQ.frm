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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Data DATA_ARQ 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
    If IsNull(Data1.Recordset("nrorec")) = False Then
       DATA_ARQ.Recordset.FindFirst "matricula =" & Data1.Recordset("matricula") & " and nrorec =" & Data1.Recordset("nrorec")
       If Not DATA_ARQ.Recordset.NoMatch Then
          DATA_ARQ.Recordset.Edit
          DATA_ARQ.Recordset("arqueo") = Data1.Recordset("arqueo")
          DATA_ARQ.Recordset("fecha") = Data1.Recordset("fecha")
          DATA_ARQ.Recordset("usuar") = Mid(Data1.Recordset("usuar"), 1, 10)
          DATA_ARQ.Recordset.Update
       Else
          DATA_ARQ.Recordset.AddNew
         DATA_ARQ.Recordset("matricula") = Data1.Recordset("matricula")
         DATA_ARQ.Recordset("nombre") = Mid(Data1.Recordset("nombre"), 1, 30)
         DATA_ARQ.Recordset("mes") = Data1.Recordset("mes")
         DATA_ARQ.Recordset("ano") = Data1.Recordset("ano")
         DATA_ARQ.Recordset("color") = Data1.Recordset("color")
         DATA_ARQ.Recordset("cat") = Mid(Data1.Recordset("cat"), 1, 6)
         DATA_ARQ.Recordset("nomcat") = Mid(Data1.Recordset("nomcat"), 1, 25)
         DATA_ARQ.Recordset("arqueo") = Data1.Recordset("arqueo")
         DATA_ARQ.Recordset("importe") = Data1.Recordset("importe")
         DATA_ARQ.Recordset("fecha") = Data1.Recordset("fecha")
         DATA_ARQ.Recordset("nrorec") = Data1.Recordset("nrorec")
         DATA_ARQ.Recordset("moneda") = Data1.Recordset("moneda")
         DATA_ARQ.Recordset("usuar") = Mid(Data1.Recordset("usuar"), 1, 10)
         DATA_ARQ.Recordset("cob") = Data1.Recordset("cob")
         DATA_ARQ.Recordset("nomcob") = Mid(Data1.Recordset("nomcob"), 1, 20)
         DATA_ARQ.Recordset("codzon") = Data1.Recordset("codzon")
         DATA_ARQ.Recordset("codpro") = Data1.Recordset("codpro")
         DATA_ARQ.Recordset("codsup") = Data1.Recordset("codsup")
         DATA_ARQ.Recordset("tiquet") = Data1.Recordset("tiquet")
         DATA_ARQ.Recordset("total") = Data1.Recordset("total")
         DATA_ARQ.Recordset("varia") = Data1.Recordset("varia")
         DATA_ARQ.Recordset.Update
       End If
    End If
    Data1.Recordset.MoveNext
   Loop
End If
MsgBox "Finalizado"


End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\env_arq.mdb"
Data1.RecordSource = "env_arq"
Data1.Refresh
DATA_ARQ.DatabaseName = App.Path & "\sapp.mdb"
DATA_ARQ.RecordSource = "arqueo"
DATA_ARQ.Refresh

End Sub
