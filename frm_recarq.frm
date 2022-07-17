VERSION 5.00
Begin VB.Form frm_recarq 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibe datos de arqueo"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "env_arq"
      Top             =   2880
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
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "arqueo"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frm_recarq.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frm_recarq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Dir(App.Path & "\Env_arq.mdb") <> "" Then
   If Data2.Recordset.RecordCount > 0 Then
      frm_recarq.MousePointer = 11
      If frm_menu.data_parse.Recordset("base") <> 3 Then
         Data2.Recordset.MoveFirst
         If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
               Data1.Recordset.Delete
               Data1.Recordset.MoveNext
            Loop
         End If
      End If
      Do While Not Data2.Recordset.EOF
         If IsNull(Data2.Recordset("nrorec")) = False Then
            Data1.Recordset.FindFirst "matricula =" & Data2.Recordset("matricula") & " and nrorec =" & Data2.Recordset("nrorec")
            If Not Data1.Recordset.NoMatch Then
               Data1.Recordset.Edit
               Data1.Recordset("arqueo") = Data2.Recordset("arqueo")
               Data1.Recordset("fecha") = Data2.Recordset("fecha")
               Data1.Recordset("usuar") = Mid(Data2.Recordset("usuar"), 1, 10)
               Data1.Recordset.Update
               Data2.Recordset.MoveNext
            Else
               Data1.Recordset.AddNew
                Data1.Recordset("matricula") = Data2.Recordset("matricula")
                Data1.Recordset("nombre") = Mid(Data2.Recordset("nombre"), 1, 30)
                Data1.Recordset("mes") = Data2.Recordset("mes")
                Data1.Recordset("ano") = Data2.Recordset("ano")
                Data1.Recordset("color") = Data2.Recordset("color")
                Data1.Recordset("cat") = Mid(Data2.Recordset("cat"), 1, 6)
                Data1.Recordset("nomcat") = Mid(Data2.Recordset("nomcat"), 1, 25)
                Data1.Recordset("arqueo") = Data2.Recordset("arqueo")
                Data1.Recordset("importe") = Data2.Recordset("importe")
                Data1.Recordset("fecha") = Data2.Recordset("fecha")
                Data1.Recordset("nrorec") = Data2.Recordset("nrorec")
                Data1.Recordset("moneda") = Data2.Recordset("moneda")
                Data1.Recordset("usuar") = Mid(Data2.Recordset("usuar"), 1, 10)
                Data1.Recordset("cob") = Data2.Recordset("cob")
                Data1.Recordset("nomcob") = Mid(Data2.Recordset("nomcob"), 1, 20)
                Data1.Recordset("codzon") = Data2.Recordset("codzon")
                Data1.Recordset("codpro") = Data2.Recordset("codpro")
                Data1.Recordset("codsup") = Data2.Recordset("codsup")
                Data1.Recordset("tiquet") = Data2.Recordset("tiquet")
                Data1.Recordset("total") = Data2.Recordset("total")
                Data1.Recordset("varia") = Data2.Recordset("varia")
                Data1.Recordset.Update
                Data2.Recordset.MoveNext
           End If
        Else
           Data2.Recordset.MoveNext
        End If
      Loop
      frm_recarq.MousePointer = 0
      MsgBox "Proceso terminado", vbInformation, "Mensaje"
      Unload Me
   Else
      MsgBox "El archivo está vacío", vbCritical, "Mensaje"
   End If
Else
   MsgBox "El archivo no existe", vbCritical, "Mensaje"
End If

   
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\sapp.mdb"
If Dir(App.Path & "\Env_arq.mdb") <> "" Then
   Data2.DatabaseName = App.Path & "\env_arq.mdb"
   Data1.Refresh
   Data2.Refresh
Else
   MsgBox "No existe archivo de arqueo", vbCritical, "Mensaje"
End If

End Sub
