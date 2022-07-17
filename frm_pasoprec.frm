VERSION 5.00
Begin VB.Form frm_pasoprec 
   Caption         =   "Paso precios"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar"
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar..."
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Data data_nuevos 
      Caption         =   "data_nuevos"
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
      Top             =   840
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Data data_est 
      Caption         =   "data_est"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3180
   End
End
Attribute VB_Name = "frm_pasoprec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_pasoprec.MousePointer = 11
Command1.Enabled = False
If data_nuevos.Recordset.RecordCount > 0 Then
   data_nuevos.Recordset.MoveFirst
   Do While Not data_nuevos.Recordset.EOF
      data_est.Recordset.FindFirst "codest =" & data_nuevos.Recordset("codest")
      If Not data_est.Recordset.NoMatch Then
         If IsNull(data_nuevos.Recordset("cons")) = False Then
            If IsNull(data_nuevos.Recordset("uc")) = False Then
               data_est.Recordset.Edit
               data_est.Recordset("cons") = data_nuevos.Recordset("cons")
               data_est.Recordset("uc") = data_nuevos.Recordset("uc")
               data_est.Recordset("part") = data_nuevos.Recordset("uc")
               data_est.Recordset("moneda") = data_nuevos.Recordset("moneda")
               data_est.Recordset.Update
            End If
         End If
      Else
         data_est.Recordset.AddNew
         data_est.Recordset("codest") = data_nuevos.Recordset("codest")
         data_est.Recordset("flia") = data_nuevos.Recordset("flia")
         data_est.Recordset("nomflia") = data_nuevos.Recordset("nomflia")
         data_est.Recordset("descrip") = data_nuevos.Recordset("descrip")
         data_est.Recordset("moneda") = data_nuevos.Recordset("moneda")
         data_est.Recordset("cons") = data_nuevos.Recordset("cons")
         data_est.Recordset("uc") = data_nuevos.Recordset("uc")
         data_est.Recordset("part") = data_nuevos.Recordset("uc")
         data_est.Recordset.Update
      End If
      data_nuevos.Recordset.MoveNext
   Loop
End If
MsgBox "Proceso finalizado"
frm_pasoprec.MousePointer = 0

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Form_Load()
data_est.DatabaseName = App.Path & "\sapp.mdb"
data_est.RecordSource = "estudios"
data_est.Refresh
data_nuevos.DatabaseName = App.Path & "\estudios.mdb"
data_nuevos.RecordSource = "estudios"
data_nuevos.Refresh

End Sub
