VERSION 5.00
Begin VB.Form env_arq 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar Arqueo"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   Icon            =   "env_arq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"env_arq.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "env_arq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Dir(App.Path & "\Env_arq.mdb") <> "" Then
   If Data2.Recordset.RecordCount > 0 Then
      env_arq.MousePointer = 11
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveFirst
         Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
         Loop
      End If
      Data2.Recordset.MoveFirst
      Do While Not Data2.Recordset.EOF
         If frm_menu.data_parse.Recordset("base") <> 3 Then
            If Data2.Recordset("arqueo") = "P" Or Data2.Recordset("arqueo") = "B" Or Data2.Recordset("arqueo") = "D" Then
                Data1.Recordset.AddNew
                Data1.Recordset("matricula") = Data2.Recordset("matricula")
                Data1.Recordset("nombre") = Mid(Data2.Recordset("nombre"), 1, 30)
                Data1.Recordset("mes") = Data2.Recordset("mes")
                Data1.Recordset("ano") = Data2.Recordset("ano")
                Data1.Recordset("color") = Data2.Recordset("color")
                Data1.Recordset("cat") = Mid(Data2.Recordset("cat"), 1, 6)
                Data1.Recordset("nomcat") = Mid(Data2.Recordset("nomcat"), 1, 30)
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
            Else
                Data2.Recordset.MoveNext
            End If
         Else
            Data1.Recordset.AddNew
            Data1.Recordset("matricula") = Data2.Recordset("matricula")
            Data1.Recordset("nombre") = Mid(Data2.Recordset("nombre"), 1, 30)
            Data1.Recordset("mes") = Data2.Recordset("mes")
            Data1.Recordset("ano") = Data2.Recordset("ano")
            Data1.Recordset("color") = Data2.Recordset("color")
            Data1.Recordset("cat") = Mid(Data2.Recordset("cat"), 1, 6)
            Data1.Recordset("nomcat") = Mid(Data2.Recordset("nomcat"), 1, 30)
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
      Loop
      env_arq.MousePointer = 0
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
Data1.DatabaseName = App.Path & "\env_arq.mdb"
Data1.RecordSource = "env_arq"
Data1.Refresh
Data2.DatabaseName = App.Path & "\sapp.mdb"
Data2.RecordSource = "arqueo"
Data2.Refresh

End Sub
