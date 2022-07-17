VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_buscasocat 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar datos..."
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10560
   Icon            =   "frm_buscasocat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   10560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9600
      Picture         =   "frm_buscasocat.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   3960
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscasocat.frx":09CC
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "frm_buscasocat.frx":09E3
      TabIndex        =   3
      Top             =   600
      Width           =   10095
   End
   Begin VB.TextBox TXT_BUSCA 
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
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_buscasocat.frx":1A6E
      Left            =   2280
      List            =   "frm_buscasocat.frx":1A7B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Con doble click selecciona el registro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Buscar por..."
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
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   7680
      Picture         =   "frm_buscasocat.frx":1A9C
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "frm_buscasocat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TXT_BUSCA.SetFocus
End If

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_atsocio.txt_cliente.Text = data_cli.Recordset("cl_codigo")
frm_atsocio.txt_nomb.Text = data_cli.Recordset("cl_apellid")
frm_atsocio.txt_codconv.Text = data_cli.Recordset("cl_codconv")
frm_atsocio.txt_desconv.Text = data_cli.Recordset("cl_nomconv")
If IsNull(data_cli.Recordset("cl_cedula")) = False Then
   frm_atsocio.txt_ced.Text = data_cli.Recordset("cl_cedula")
Else
   frm_atsocio.txt_ced.Text = 0
End If
If IsNull(data_cli.Recordset("cl_codced")) = False Then
   frm_atsocio.txt_codced.Text = data_cli.Recordset("cl_codced")
Else
   frm_atsocio.txt_codced.Text = 0
End If
If IsNull(data_cli.Recordset("cl_fecing")) = False Then
   frm_atsocio.ming.Text = Format(data_cli.Recordset("cl_fecing"), "dd/mm/yyyy")
Else
   frm_atsocio.ming.Text = "__/__/____"
End If
If IsNull(data_cli.Recordset("cl_fnac")) = False Then
   frm_atsocio.mnac.Text = Format(data_cli.Recordset("cl_fnac"), "dd/mm/yyyy")
Else
   frm_atsocio.mnac.Text = "__/__/____"
End If
If IsNull(data_cli.Recordset("cl_telefon")) = False Then
   frm_atsocio.txt_telef.Text = data_cli.Recordset("cl_telefon")
Else
   frm_atsocio.txt_telef.Text = ""
End If
Unload Me

End Sub

Private Sub Form_Load()
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
Combo1.ListIndex = 0

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub TXT_BUSCA_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If Combo1.ListIndex = 0 Then
      data_cli.RecordSource = "Select top 80 * from clientes where cl_apellid >='" & TXT_BUSCA.Text & "' order by cl_apellid"
      data_cli.Refresh
   Else
      If Combo1.ListIndex = 1 Then
         data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Val(TXT_BUSCA.Text) & " order by cl_cedula"
         data_cli.Refresh
      Else
         If Combo1.ListIndex = 2 Then
            data_cli.RecordSource = "Select * from clientes where cl_telefon ='" & TXT_BUSCA.Text & "' order by cl_telefon"
            data_cli.Refresh
         End If
      End If
   End If
   DBGrid1.SetFocus
End If

End Sub
