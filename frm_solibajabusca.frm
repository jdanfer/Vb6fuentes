VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_solibajabusca 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar solicitudes"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   Icon            =   "frm_solibajabusca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "Solo pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_solibajabusca.frx":058A
      Height          =   2775
      Left            =   240
      OleObjectBlob   =   "frm_solibajabusca.frx":059E
      TabIndex        =   4
      Top             =   600
      Width           =   8535
   End
   Begin MSMask.MaskEdBox mf 
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_solibajabusca.frx":1475
      Left            =   2160
      List            =   "frm_solibajabusca.frx":1482
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Doble click para seleccionar el registro."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frm_solibajabusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
   Text1.Visible = False
   mf.Visible = True
Else
   Text1.Visible = True
   mf.Visible = False
End If


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Text1.Visible = True Then
      Text1.SetFocus
   Else
      mf.SetFocus
   End If
End If

End Sub

Private Sub DBGrid1_DblClick()
frm_solicitudbaja.t_cod.Text = Data1.Recordset("id")
frm_solicitudbaja.mfec.Text = Data1.Recordset("fecha")
frm_solicitudbaja.mhora.Text = Data1.Recordset("hora")
frm_solicitudbaja.t_base.Text = Data1.Recordset("base")
frm_solicitudbaja.labusu.Caption = Data1.Recordset("usuario")
If IsNull(Data1.Recordset("matricula")) = False Then
   frm_solicitudbaja.t_mat.Text = Data1.Recordset("matricula")
Else
   frm_solicitudbaja.t_mat.Text = 0
End If
If IsNull(Data1.Recordset("cedula")) = False Then
   frm_solicitudbaja.t_ced.Text = Data1.Recordset("cedula")
Else
   frm_solicitudbaja.t_ced.Text = 0
End If
If IsNull(Data1.Recordset("codced")) = False Then
   frm_solicitudbaja.t_cced.Text = Data1.Recordset("codced")
Else
   frm_solicitudbaja.t_cced.Text = 0
End If
If IsNull(Data1.Recordset("convenio")) = False Then
   frm_solicitudbaja.t_conv.Text = Data1.Recordset("convenio")
Else
   frm_solicitudbaja.t_conv.Text = "S/D"
End If
If IsNull(Data1.Recordset("nombre")) = False Then
   frm_solicitudbaja.t_nom.Text = Data1.Recordset("nombre")
Else
   frm_solicitudbaja.t_nom.Text = "NN"
End If
If IsNull(Data1.Recordset("telefono")) = False Then
   frm_solicitudbaja.t_tel.Text = Data1.Recordset("telefono")
Else
   frm_solicitudbaja.t_tel.Text = "0"
End If
If IsNull(Data1.Recordset("celular")) = False Then
   frm_solicitudbaja.t_cel.Text = Data1.Recordset("celular")
Else
   frm_solicitudbaja.t_cel.Text = "9"
End If
If IsNull(Data1.Recordset("hora1id")) = False Then
   frm_solicitudbaja.cbodes.ListIndex = Data1.Recordset("hora1id")
Else
   frm_solicitudbaja.cbodes.ListIndex = -1
End If
If IsNull(Data1.Recordset("hora2id")) = False Then
   frm_solicitudbaja.cbohas.ListIndex = Data1.Recordset("hora2id")
Else
   frm_solicitudbaja.cbohas.ListIndex = -1
End If
If IsNull(Data1.Recordset("origid")) = False Then
   frm_solicitudbaja.cboorig.ListIndex = Data1.Recordset("origid")
Else
   frm_solicitudbaja.cboorig.ListIndex = -1
End If
If IsNull(Data1.Recordset("motid")) = False Then
   frm_solicitudbaja.cbomot.ListIndex = Data1.Recordset("motid")
Else
   frm_solicitudbaja.cbomot.ListIndex = -1
End If
If IsNull(Data1.Recordset("otrotel")) = False Then
   frm_solicitudbaja.t_otro.Text = Data1.Recordset("otrotel")
Else
   frm_solicitudbaja.t_otro.Text = ""
End If
If IsNull(Data1.Recordset("fechafin")) = False Then
   frm_solicitudbaja.mffin.Text = Data1.Recordset("fechafin")
Else
   frm_solicitudbaja.mffin.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("contrato")) = False Then
   frm_solicitudbaja.Check2.Value = Data1.Recordset("contrato")
Else
   frm_solicitudbaja.Check2.Value = 0
End If
If IsNull(Data1.Recordset("resulid")) = False Then
   frm_solicitudbaja.cborec.ListIndex = Data1.Recordset("resulid")
Else
   frm_solicitudbaja.cborec.ListIndex = -1
End If
Data2.RecordSource = "select * from solbaja_acc where idid =" & Data1.Recordset("id")
Data2.Refresh
frm_solicitudbaja.t_accion.Text = ""
If Data2.Recordset.RecordCount > 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      If Trim(frm_solicitudbaja.t_accion.Text) = "" Then
         frm_solicitudbaja.t_accion.Text = Data2.Recordset("fecha") & "--" & Data2.Recordset("hora") & "--" & Data2.Recordset("usuario") & "--" & Data2.Recordset("accion")
      Else
         frm_solicitudbaja.t_accion.Text = frm_solicitudbaja.t_accion.Text & Data2.Recordset("fecha") & "--" & Data2.Recordset("hora") & "--" & Data2.Recordset("usuario") & "--" & Data2.Recordset("accion")
      End If
      frm_solicitudbaja.t_accion.Text = frm_solicitudbaja.t_accion.Text & Chr(13) & Chr(10) & "-----------------------------------------------------" & Chr(13) & Chr(10)
      Data2.Recordset.MoveNext
   Loop
End If
If IsNull(Data1.Recordset("terminado")) = False Then
   frm_solicitudbaja.Check1.Value = Data1.Recordset("terminado")
Else
   frm_solicitudbaja.Check1.Value = 0
End If

Unload Me

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from solic_bajas order by fecha DESC"
Data1.Refresh

End Sub

Private Sub mf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If mf.Text <> "__/__/____" Then
      If Check1.Value = 1 Then
         Data1.RecordSource = "Select * from solic_bajas where fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "# and terminado is not (1)"
      Else
         Data1.RecordSource = "Select * from solic_bajas where fecha >=#" & Format(mf.Text, "yyyy/mm/dd") & "#"
      End If
      Data1.Refresh
      DBGrid1.SetFocus
   Else
      If Check1.Value = 1 Then
         Data1.RecordSource = "select * from solic_bajas where terminado is not (1) order by fecha DESC"
      Else
         Data1.RecordSource = "select * from solic_bajas order by fecha DESC"
      End If
      Data1.Refresh
      DBGrid1.SetFocus
   End If
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Text1.Text <> "" Then
      If Combo1.ListIndex = 1 Then
         Data1.RecordSource = "select * from solic_bajas where matricula =" & Text1.Text
         Data1.Refresh
      Else
         Data1.RecordSource = "select * from solic_bajas where cedula =" & Text1.Text
         Data1.Refresh
      End If
   Else
      If Check1.Value = 1 Then
         Data1.RecordSource = "select * from solic_bajas where terminado is not (1) order by fecha DESC"
      Else
         Data1.RecordSource = "select * from solic_bajas order by fecha DESC"
      End If
      Data1.Refresh
   End If
   DBGrid1.SetFocus
End If

End Sub
