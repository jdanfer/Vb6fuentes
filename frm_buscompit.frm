VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_buscompit 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar ITEMS"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_buscompit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7725
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "stock"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6840
      Picture         =   "frm_buscompit.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3000
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscompit.frx":09CC
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frm_buscompit.frx":09E0
      TabIndex        =   2
      Top             =   600
      Width           =   7335
   End
   Begin VB.TextBox t_desbus 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click selecciona."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DESCRIPCION:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   4680
      Picture         =   "frm_buscompit.frx":13B3
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frm_buscompit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(data1.Recordset("id")) = False Then
   frm_compsto.t_codprod.Text = data1.Recordset("id")
End If
If IsNull(data1.Recordset("descrip")) = False Then
   frm_compsto.labdesc.Caption = data1.Recordset("descrip")
   frm_compsto.t_pre.Text = data1.Recordset("preuni")

End If
Unload Me

End Sub

Private Sub Form_Load()
data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub t_desbus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
      data1.RecordSource = "Select * from stock where descrip >='" & t_desbus.Text & "' and grupo =" & 3 & " order by descrip"
   Else
      data1.RecordSource = "Select * from stock where descrip >='" & t_desbus.Text & "' and grupo not in (3) order by descrip"
   End If
   data1.Refresh
   DBGrid1.SetFocus
End If

End Sub
