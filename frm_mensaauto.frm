VERSION 5.00
Begin VB.Form frm_mensaauto 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Mensaje"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   Icon            =   "frm_mensaauto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      Picture         =   "frm_mensaauto.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frm_mensaauto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.path & "\informes.mdb"
Data1.RecordSource = "infcli"
Data1.Refresh
If XMensaFertilab = 9 Then
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         Data1.Recordset.Delete
         Data1.Recordset.MoveNext
      Loop
   End If
   Data1.Recordset.AddNew
   Data1.Recordset("info_debit") = "ATENCION!! Paciente debe concurrir al laboratorio Fertilab a realizar el Examen."
   Data1.Recordset.Update
   Data1.Refresh
   Label1.Caption = Data1.Recordset("info_debit")
Else
    If Data1.Recordset.RecordCount > 0 Then
       If IsNull(Data1.Recordset("info_debit")) = False Then
          Label1.Caption = Data1.Recordset("info_debit")
       Else
          Label1.Caption = ""
       End If
    End If
End If

End Sub

