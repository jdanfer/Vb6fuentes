VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_enferm 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enfermería"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_enferm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Seleccionar..."
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_enferm.frx":0442
      Height          =   2055
      Left            =   240
      OleObjectBlob   =   "frm_enferm.frx":0456
      TabIndex        =   11
      Top             =   2520
      Width           =   5655
   End
   Begin VB.CommandButton b_eli 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_enferm.frx":0E29
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_enferm.frx":13B3
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_enferm.frx":193D
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1080
      Picture         =   "frm_enferm.frx":1EC7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_enferm.frx":2451
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox t_tel 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      MaxLength       =   35
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox t_nomb 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox t_cod 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Telef:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nombre:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CODIGO:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frm_enferm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
t_cod.Enabled = True
t_nomb.Enabled = True
t_tel.Enabled = True
t_cod.Text = ""
t_nomb.Text = ""
t_tel.Text = ""
If Data1.Recordset.RecordCount > 0 Then
   Data1.RecordSource = "Select * from enferm order by id"
   Data1.Refresh
   Data1.Recordset.MoveLast
   t_cod.Text = Data1.Recordset("id") + 1
Else
   t_cod.Text = 15
End If
t_cod.SetFocus
XAlta = 1
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_eli.Enabled = False


End Sub

Private Sub b_cance_Click()
b_alta.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cance.Enabled = False
b_eli.Enabled = True

End Sub

Private Sub b_eli_Click()
Data1.Recordset.Delete
Data1.Refresh
t_cod.Enabled = True
t_nomb.Enabled = True
t_tel.Enabled = True
t_cod.Text = ""
t_nomb.Text = ""
t_tel.Text = ""
t_cod.Enabled = False
t_nomb.Enabled = False
t_tel.Enabled = False

End Sub

Private Sub b_graba_Click()
If t_cod.Text <> "" Then
   If XAlta = 1 Then
      Data1.Recordset.AddNew
      Data1.Recordset("id") = t_cod.Text
      Data1.Recordset("nomb") = t_nomb.Text
      Data1.Recordset("telef") = t_tel.Text
      Data1.Recordset.Update
      XAlta = 0
      Data1.Refresh
      t_cod.Text = ""
      t_nomb.Text = ""
      t_tel.Text = ""
      t_cod.Enabled = False
      t_nomb.Enabled = False
      t_tel.Enabled = False
      DBGrid1.SetFocus
   Else
      Data1.Recordset.Edit
      Data1.Recordset("id") = t_cod.Text
      Data1.Recordset("nomb") = t_nomb.Text
      Data1.Recordset("telef") = t_tel.Text
      Data1.Recordset.Update
      XAlta = 0
      Data1.Refresh
      t_cod.Text = ""
      t_nomb.Text = ""
      t_tel.Text = ""
      t_cod.Enabled = False
      t_nomb.Enabled = False
      t_tel.Enabled = False
      DBGrid1.SetFocus
   End If

End If
b_alta.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cance.Enabled = False
b_eli.Enabled = True
Data1.RecordSource = "Select * from enferm order by nomb"
Data1.Refresh

End Sub

Private Sub b_modif_Click()
t_cod.Enabled = True
t_nomb.Enabled = True
t_tel.Enabled = True
t_cod.SetFocus
XAlta = 0
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_eli.Enabled = False


End Sub

Private Sub Command1_Click()
If IsNull(Data1.Recordset("nomb")) = False Then
   frm_opsdesp.t_enf(0).Text = Data1.Recordset("nomb")
Else
   frm_opsdesp.t_enf(0).Text = ""
End If
If IsNull(Data1.Recordset("id")) = False Then
   frm_opsdesp.labenf.Caption = Data1.Recordset("id")
Else
   frm_opsdesp.labenf.Caption = 0
End If
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(Data1.Recordset("id")) = False Then
   t_cod.Text = Data1.Recordset("id")
Else
   t_cod.Text = ""
End If
If IsNull(Data1.Recordset("nomb")) = False Then
   t_nomb.Text = Data1.Recordset("nomb")
Else
   t_nomb.Text = ""
End If
If IsNull(Data1.Recordset("telef")) = False Then
   t_tel.Text = Data1.Recordset("telef")
Else
   t_tel.Text = ""
End If

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\enferm.mdb"
Data1.RecordSource = "Select * from enferm order by nomb"
Data1.Refresh

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nomb.SetFocus
End If

End Sub

Private Sub t_nomb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_tel.SetFocus
End If

End Sub
