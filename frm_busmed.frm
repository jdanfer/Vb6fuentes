VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_busmed 
   BackColor       =   &H00C000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar médico"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_busmed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7950
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_busmed.frx":0442
      Height          =   2655
      Left            =   240
      OleObjectBlob   =   "frm_busmed.frx":0456
      TabIndex        =   3
      Top             =   600
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      Picture         =   "frm_busmed.frx":0FD9
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   3240
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox t_bus 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Doble click para seleccionar registro"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "BUSCAR POR NOMBRE:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4680
      Picture         =   "frm_busmed.frx":1563
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "frm_busmed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_largador.txt_codmedtra.Text = data1.Recordset("med_cod")
MsgBox "El médico seleccionado es: " & data1.Recordset("med_nombre")
Unload Me
End Sub

Private Sub Form_Load()
'Data1.DatabaseName = App.Path & "\sapp.mdb"
data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data1.RecordSource = "Select * from medicos order by med_cod"
data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub

Private Sub t_bus_Change()
data1.RecordSource = "Select * from medicos where med_nombre >='" & t_bus.Text & "' order by med_nombre"
data1.Refresh

End Sub

Private Sub t_bus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub
