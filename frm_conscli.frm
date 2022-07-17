VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_conscli 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de clientes..."
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_conscli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   ""
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      Picture         =   "frm_conscli.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   3480
      Width           =   495
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_conscli.frx":09CC
      Height          =   2775
      Left            =   240
      OleObjectBlob   =   "frm_conscli.frx":09E3
      TabIndex        =   3
      Top             =   720
      Width           =   8415
   End
   Begin VB.TextBox t_cons 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_conscli.frx":1556
      Left            =   2520
      List            =   "frm_conscli.frx":1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click para seleccionar."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Consultar por..."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   4560
      Picture         =   "frm_conscli.frx":1576
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frm_conscli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_reggasto.labcli.Caption = data_cli.Recordset("nombre")
frm_reggasto.t_cli.Text = data_cli.Recordset("id")
Unload Me

End Sub

Private Sub Form_Load()
'data_cli.DatabaseName = App.Path & "\" & Trim(Xlabdd)
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_cons_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo1.ListIndex = 0 Then
      data_cli.RecordSource = "Select * from clieco where nombre >='" & t_cons.Text & "' order by nombre"
      data_cli.Refresh
      DBGrid1.SetFocus
   Else
      If Combo1.ListIndex = 1 Then
         If IsNumeric(t_cons.Text) = True Then
            data_cli.RecordSource = "select top 70, * from clieco where id >=" & t_cons.Text & " order by id"
            data_cli.Refresh
         End If
         DBGrid1.SetFocus
      End If
   End If
End If

End Sub
