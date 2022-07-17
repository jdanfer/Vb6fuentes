VERSION 5.00
Begin VB.Form frm_entrega 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas de cobradores"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "frm_entrega.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6195
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_otracom 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Data data_cobotra 
      Caption         =   "data_cobotra"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton b_bus 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      Picture         =   "frm_entrega.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   495
   End
   Begin VB.Data data_ent 
      Caption         =   "data_ent"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ENTREGAS"
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_cob 
      Caption         =   "data_cob"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cobrador"
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      Picture         =   "frm_entrega.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton b_ant 
      BackColor       =   &H00808080&
      Caption         =   "---Anterior"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton b_sig 
      BackColor       =   &H00808080&
      Caption         =   "Siguiente---"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      Picture         =   "frm_entrega.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton b_grab 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      Picture         =   "frm_entrega.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      Picture         =   "frm_entrega.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton b_nue 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      Picture         =   "frm_entrega.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_imp 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txt_cob 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Total Otras comisiones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Total entregado $."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Cobrador:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   720
      Picture         =   "frm_entrega.frx":257E
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "frm_entrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_ant_Click()
If data_ent.Recordset.BOF = True Then
   MsgBox "Comienzo de archivo", vbInformation, "Mensaje"
Else
   data_ent.Recordset.MovePrevious
   If Not data_ent.Recordset.BOF Then
      txt_cob.Text = data_ent.Recordset("cobrador")
      Label2.Caption = data_ent.Recordset("nombre")
      txt_imp.Text = Format(data_ent.Recordset("pesos"), "Standard")
      If txt_cob.Text <> "" Then
         data_cobotra.RecordSource = "Select * from cobrador where cb_numero =" & txt_cob.Text
         data_cobotra.Refresh
         If data_cobotra.Recordset.RecordCount > 0 Then
            If IsNull(data_cobotra.Recordset("cb_recpes")) = False Then
               t_otracom.Text = Format(data_cobotra.Recordset("cb_recpes"), "Standard")
            Else
               t_otracom.Text = 0
            End If
         Else
            t_otracom.Text = 0
         End If
      End If
   End If
End If

End Sub

Private Sub b_bus_Click()
frm_busentre.Show vbModal

End Sub

Private Sub b_canc_Click()
If XAlta = 1 Then
   data_ent.Recordset.CancelUpdate
   txt_cob.Text = ""
   txt_imp.Text = ""
   t_otracom.Text = ""
   txt_cob.Enabled = False
   txt_imp.Enabled = False
   t_otracom.Enabled = False
   b_nue.Enabled = True
   b_mod.Enabled = True
   b_grab.Enabled = False
   b_canc.Enabled = False
   txt_cob.Text = data_ent.Recordset("cobrador")
   Label2.Caption = data_ent.Recordset("nombre")
   txt_imp.Text = Format(data_ent.Recordset("pesos"), "Standard")
   If t_otracom.Text <> "" Then
      t_otracom.Text = Format(t_otracom.Text, "Standard")
   End If
Else
   txt_cob.Text = ""
   txt_imp.Text = ""
   t_otracom.Text = ""
   txt_cob.Enabled = False
   txt_imp.Enabled = False
   t_otracom.Enabled = False
   b_nue.Enabled = True
   b_mod.Enabled = True
   b_grab.Enabled = False
   b_canc.Enabled = False
   txt_cob.Text = data_ent.Recordset("cobrador")
   Label2.Caption = data_ent.Recordset("nombre")
   txt_imp.Text = Format(data_ent.Recordset("pesos"), "Standard")
   If t_otracom.Text <> "" Then
      t_otracom.Text = Format(t_otracom.Text, "Standard")
   End If

End If

End Sub

Private Sub b_grab_Click()

If XAlta = 1 Then
   data_ent.Recordset.AddNew
   data_ent.Recordset("cobrador") = txt_cob.Text
   data_ent.Recordset("nombre") = Label2.Caption
   data_ent.Recordset("pesos") = txt_imp.Text
   data_ent.Recordset.Update
   txt_cob.Enabled = False
   txt_imp.Enabled = False
   b_nue.Enabled = True
   b_mod.Enabled = True
   b_grab.Enabled = False
   b_canc.Enabled = False
   txt_cob.Text = data_ent.Recordset("cobrador")
   Label2.Caption = data_ent.Recordset("nombre")
   txt_imp.Text = Format(data_ent.Recordset("pesos"), "Standard")
   If t_otracom.Text <> "" Then
      data_cobotra.RecordSource = "Select * from cobrador where cb_numero =" & txt_cob.Text
      data_cobotra.Refresh
      If data_cobotra.Recordset.RecordCount > 0 Then
         If IsNull(data_cobotra.Recordset("cb_recpes")) = False Then
            If Format(data_cobotra.Recordset("cb_recpes"), "Standard") <> Format(t_otracom.Text, "Standard") Then
               data_cobotra.Recordset.Edit
               data_cobotra.Recordset("cb_recpes") = Format(t_otracom.Text, "Standard")
               data_cobotra.Recordset.Update
               t_otracom.Text = Format(t_otracom.Text, "Standard")
            End If
         Else
            data_cobotra.Recordset.Edit
            data_cobotra.Recordset("cb_recpes") = Format(t_otracom.Text, "Standard")
            data_cobotra.Recordset.Update
            t_otracom.Text = Format(t_otracom.Text, "Standard")
         End If
      End If
   End If
   t_otracom.Enabled = False
Else
   data_ent.Recordset.FindFirst "cobrador =" & txt_cob.Text
   If Not data_ent.Recordset.NoMatch Then
      If txt_imp.Text <> "" Then
         If Format(data_ent.Recordset("pesos"), "Standard") <> Format(txt_imp.Text, "Standard") Then
            data_ent.Recordset.Edit
            data_ent.Recordset("cobrador") = txt_cob.Text
            data_ent.Recordset("nombre") = Label2.Caption
            data_ent.Recordset("pesos") = txt_imp.Text
            data_ent.Recordset.Update
         End If
      End If
      If t_otracom.Text <> "" Then
         data_cobotra.RecordSource = "Select * from cobrador where cb_numero =" & txt_cob.Text
         data_cobotra.Refresh
         If data_cobotra.Recordset.RecordCount > 0 Then
            If IsNull(data_cobotra.Recordset("cb_recpes")) = False Then
               If Format(data_cobotra.Recordset("cb_recpes"), "Standard") <> Format(t_otracom.Text, "Standard") Then
                  data_cobotra.Recordset.Edit
                  data_cobotra.Recordset("cb_recpes") = Format(t_otracom.Text, "Standard")
                  data_cobotra.Recordset.Update
                  t_otracom.Text = Format(t_otracom.Text, "Standard")
               End If
            Else
               data_cobotra.Recordset.Edit
               data_cobotra.Recordset("cb_recpes") = Format(t_otracom.Text, "Standard")
               data_cobotra.Recordset.Update
               t_otracom.Text = Format(t_otracom.Text, "Standard")
            End If
         End If
      End If
      t_otracom.Enabled = False
   Else
      MsgBox "Atención! no se pudo grabar el registro"
   End If
   txt_cob.Enabled = False
   txt_imp.Enabled = False
   b_nue.Enabled = True
   b_mod.Enabled = True
   b_grab.Enabled = False
   b_canc.Enabled = False
   txt_cob.Text = data_ent.Recordset("cobrador")
   Label2.Caption = data_ent.Recordset("nombre")
   txt_imp.Text = Format(data_ent.Recordset("pesos"), "Standard")

End If
End Sub

Private Sub b_mod_Click()
txt_cob.Enabled = False
txt_imp.Enabled = True
t_otracom.Enabled = True
b_nue.Enabled = False
b_mod.Enabled = False
b_grab.Enabled = True
b_canc.Enabled = True
XAlta = 0
txt_imp.SetFocus

End Sub

Private Sub b_nue_Click()
txt_cob.Enabled = True
txt_imp.Enabled = True
t_otracom.Enabled = True
b_nue.Enabled = False
b_mod.Enabled = False
b_grab.Enabled = True
b_canc.Enabled = True
data_ent.Recordset.AddNew
XAlta = 1
txt_imp.Text = ""
txt_cob.Text = ""
t_otracom.Text = ""
txt_cob.SetFocus

End Sub

Private Sub b_sig_Click()
If data_ent.Recordset.EOF = True Then
   MsgBox "Final de archivo", vbInformation, "Mensaje"
   
Else
   data_ent.Recordset.MoveNext
   If Not data_ent.Recordset.EOF Then
      txt_cob.Text = data_ent.Recordset("cobrador")
      Label2.Caption = data_ent.Recordset("nombre")
      txt_imp.Text = Format(data_ent.Recordset("pesos"), "Standard")
      If txt_cob.Text <> "" Then
         data_cobotra.RecordSource = "Select * from cobrador where cb_numero =" & txt_cob.Text
         data_cobotra.Refresh
         If data_cobotra.Recordset.RecordCount > 0 Then
            If IsNull(data_cobotra.Recordset("cb_recpes")) = False Then
               t_otracom.Text = Format(data_cobotra.Recordset("cb_recpes"), "Standard")
            Else
               t_otracom.Text = 0
            End If
         Else
            t_otracom.Text = 0
         End If
      End If
   End If
End If

End Sub

Private Sub Form_Initialize()
data_ent.Recordset.MoveLast
txt_cob.Text = data_ent.Recordset("cobrador")
Label2.Caption = data_ent.Recordset("nombre")
txt_imp.Text = Format(data_ent.Recordset("pesos"), "Standard")
If txt_cob.Text <> "" Then
   data_cobotra.RecordSource = "Select * from cobrador where cb_numero =" & txt_cob.Text
   data_cobotra.Refresh
   If data_cobotra.Recordset.RecordCount > 0 Then
      If IsNull(data_cobotra.Recordset("cb_recpes")) = False Then
         t_otracom.Text = Format(data_cobotra.Recordset("cb_recpes"), "Standard")
      Else
         t_otracom.Text = 0
      End If
   Else
      t_otracom.Text = 0
   End If
End If


End Sub

Private Sub Form_Load()
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.Refresh
data_ent.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ent.RecordSource = "Select * from entregas order by cobrador"
data_ent.Refresh
'cb_recpes
data_cobotra.Connect = "odbc;dsn=" & Xconexrmt & ";"


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_cob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_imp.SetFocus
End If
End Sub

Private Sub txt_cob_LostFocus()
If txt_cob.Text <> "" Then
   data_cob.Recordset.FindFirst "cb_numero =" & txt_cob.Text
   If Not data_cob.Recordset.NoMatch Then
      Label2.Caption = data_cob.Recordset("cb_nombre")
   Else
      MsgBox "No encontrado", vbInformation, "Mensaje"
      txt_cob.SetFocus
   End If
   
Else
   MsgBox "No encontrado", vbInformation, "Mensaje"
   txt_cob.SetFocus
End If

End Sub

Private Sub txt_imp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_grab.SetFocus
End If

End Sub

Private Sub txt_imp_LostFocus()
txt_imp.Text = Format(txt_imp.Text, "Standard")

End Sub
