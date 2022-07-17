VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_buscondesp 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar convenios..."
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_buscondesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8160
      Picture         =   "frm_buscondesp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   3840
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscondesp.frx":09CC
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "frm_buscondesp.frx":09E0
      TabIndex        =   3
      Top             =   840
      Width           =   8655
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Por código"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Por descripción"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   5400
      Picture         =   "frm_buscondesp.frx":1573
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "frm_buscondesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
If Xwesvtas = 9 Then
   If IsNull(Data1.Recordset("cnv_codigo")) = False Then
      frm_vtasxgpo.Text1.Text = Data1.Recordset("cnv_codigo")
   Else
      frm_vtasxgpo.Text1.Text = "A"
   End If
   If IsNull(Data1.Recordset("cnv_desc")) = False Then
      frm_vtasxgpo.Text2.Text = Data1.Recordset("cnv_desc")
   Else
      frm_vtasxgpo.Text2.Text = "A"
   End If
Else
   If IsNull(Data1.Recordset("cnv_codigo")) = False Then
      frm_infdesp2.Text1.Text = Data1.Recordset("cnv_codigo")
   Else
      frm_infdesp2.Text1.Text = "A"
   End If
   If IsNull(Data1.Recordset("cnv_desc")) = False Then
      frm_infdesp2.Text2.Text = Data1.Recordset("cnv_desc")
   Else
      frm_infdesp2.Text2.Text = "A"
   End If
End If
Xwesvtas = 0

Unload Me

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.RecordSource = "select top 70, * from convenio order by cnv_desc"
'Data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Option1.Value = True Then
       Data1.RecordSource = "select top 70, * from convenio where cnv_desc >='" & Text1.Text & "' order by cnv_desc"
       Data1.Refresh
    Else
       Data1.RecordSource = "select top 70, * from convenio where cnv_codigo >='" & Text1.Text & "' order by cnv_codigo"
       Data1.Refresh
    End If
   DBGrid1.SetFocus
End If

End Sub
