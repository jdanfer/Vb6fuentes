VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_busentre 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar y modificar entregas de cobradores"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "frm_busentre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar registro seleccionado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6480
      Picture         =   "frm_busentre.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_busentre.frx":09CC
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "frm_busentre.frx":09E0
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cobradores ordenados por número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   3000
      Picture         =   "frm_busentre.frx":156F
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "frm_busentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Dim Xsiborra As String
Xsiborra = MsgBox("Desea borrar el registro seleccionado?", vbYesNo + vbInformation, "Entregas")
If Xsiborra = vbYes Then
   Data1.Recordset.Delete
   Data1.Refresh
End If
DBGrid1.SetFocus

End Sub

Private Sub DBGrid1_DblClick()
frm_entrega.data_ent.Recordset.FindFirst "cobrador =" & Data1.Recordset("cobrador")
If Not frm_entrega.data_ent.Recordset.NoMatch Then
   frm_entrega.txt_cob.Text = frm_entrega.data_ent.Recordset("cobrador")
   frm_entrega.txt_imp.Text = frm_entrega.data_ent.Recordset("pesos")
   frm_entrega.Label2.Caption = frm_entrega.data_ent.Recordset("nombre")

End If
Unload Me

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from entregas order by cobrador"
Data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub
