VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_abmcli 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso datos de cliente"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_abmcliu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   3240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      Picture         =   "frm_abmcliu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Informe de clientes"
      Top             =   2880
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_abmcliu.frx":09CC
      Height          =   2055
      Left            =   240
      OleObjectBlob   =   "frm_abmcliu.frx":09E0
      TabIndex        =   12
      Top             =   3480
      Width           =   7215
   End
   Begin VB.CommandButton b_elim 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      Picture         =   "frm_abmcliu.frx":1553
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      Picture         =   "frm_abmcliu.frx":1ADD
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      Picture         =   "frm_abmcliu.frx":2067
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      Picture         =   "frm_abmcliu.frx":25F1
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_abmcliu.frx":2B7B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos..."
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.TextBox t_base 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox t_nom 
         Height          =   375
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox t_cod 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "BASE:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "NOMBRE:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "CODIGO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   6480
      Picture         =   "frm_abmcliu.frx":3105
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "frm_abmcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()

End Sub

Private Sub b_alta_Click()
Frame1.Enabled = True
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_elim.Enabled = False
t_cod.SetFocus
Data1.RecordSource = "Select * from clieco order by id"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
   t_cod.Text = Data1.Recordset("id") + 1
Else
   t_cod.Text = 1
End If
XAlta = 1

End Sub

Private Sub b_cance_Click()
      XAlta = 0
      t_cod.Text = ""
      t_nom.Text = ""
      t_base.Text = ""
      Frame1.Enabled = False
      b_alta.Enabled = True
      b_modif.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_elim.Enabled = True

End Sub

Private Sub b_elim_Click()
Dim Xelmen As String
Xelmen = MsgBox("Desea eliminar el registro seleccionado? " & t_nom.Text, vbInformation + vbYesNo, "SAPP")
If Xelmen = vbYes Then
   If t_cod.Text <> "" Then
      Data1.Recordset.Delete
      Data1.Refresh
   End If
End If
XAlta = 0
t_cod.Text = ""
t_nom.Text = ""
t_base.Text = ""
Frame1.Enabled = False
b_alta.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cance.Enabled = False
b_elim.Enabled = True


End Sub

Private Sub b_graba_Click()
If XAlta = 1 Then
   If t_cod.Text <> "" Then
      Data1.Recordset.AddNew
      Data1.Recordset("id") = t_cod.Text
      If t_nom.Text = "" Then
         t_nom.Text = "S/N"
      End If
      Data1.Recordset("nombre") = t_nom.Text
      If t_base.Text = "" Then
         t_base.Text = 0
      End If
      Data1.Recordset("base") = t_base.Text
      Data1.Recordset.Update
      Data1.Refresh
      XAlta = 0
      t_cod.Text = ""
      t_nom.Text = ""
      t_base.Text = ""
      Frame1.Enabled = False
      b_alta.Enabled = True
      b_modif.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_elim.Enabled = True
   Else
      MsgBox "Ingrese código"
   End If
Else
   Data1.Recordset.Edit
'   Data1.Recordset("id") = t_cod.Text
   If t_nom.Text = "" Then
      t_nom.Text = "NN"
   End If
   Data1.Recordset("nombre") = t_nom.Text
   If t_base.Text = "" Then
      t_base.Text = 0
   End If
   Data1.Recordset("base") = t_base.Text
   Data1.Recordset.Update
   Data1.Refresh
   XAlta = 0
   t_cod.Text = ""
   t_nom.Text = ""
   t_base.Text = ""
   Frame1.Enabled = False
   b_alta.Enabled = True
   b_modif.Enabled = True
   b_graba.Enabled = False
   b_cance.Enabled = False
   b_elim.Enabled = True

End If

End Sub

Private Sub b_imp_Click()
Data2.DatabaseName = App.path & "\informes.mdb"
Data2.RecordSource = "inflla"
Data2.Refresh
If Data2.Recordset.RecordCount > 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      Data2.Recordset.Delete
      Data2.Recordset.MoveNext
   Loop
End If
Data2.Refresh
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
   Data2.Recordset.AddNew
   Data2.Recordset("nombre") = Data1.Recordset("nombre")
   Data2.Recordset("matric") = Data1.Recordset("id")
   Data2.Recordset("edad") = Data1.Recordset("base")
   Data2.Recordset.Update
   Data1.Recordset.MoveNext
Loop
Data2.Refresh

cr1.ReportFileName = App.path & "\infclistok.rpt"
cr1.Action = 1


End Sub

Private Sub b_modif_Click()
If t_cod.Text = "" Then
   MsgBox "No existe código"
Else
   Frame1.Enabled = True
   b_alta.Enabled = False
   b_modif.Enabled = False
   b_graba.Enabled = True
   b_cance.Enabled = True
   b_elim.Enabled = False
   t_cod.SetFocus
   XAlta = 0
End If

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(Data1.Recordset("id")) = False Then
   t_cod.Text = Data1.Recordset("id")
Else
   t_cod.Text = ""
End If
If IsNull(Data1.Recordset("nombre")) = False Then
   t_nom.Text = Data1.Recordset("nombre")
Else
   t_nom.Text = ""
End If
If IsNull(Data1.Recordset("base")) = False Then
   t_base.Text = Data1.Recordset("base")
Else
   t_base.Text = ""
End If

End Sub

Private Sub Form_Load()
'Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "clieco"
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

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_base.SetFocus
End If

End Sub
