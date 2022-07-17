VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_rubrec 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros Caja Recepción"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frm_rubrec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton b_elim 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      Picture         =   "frm_rubrec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Eliminar registro seleccionado"
      Top             =   2640
      Width           =   495
   End
   Begin VB.ComboBox cbocon 
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
      ItemData        =   "frm_rubrec.frx":09CC
      Left            =   2160
      List            =   "frm_rubrec.frx":09D6
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox cbomon 
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
      ItemData        =   "frm_rubrec.frx":09EB
      Left            =   2160
      List            =   "frm_rubrec.frx":09F5
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1320
      Width           =   3375
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sociedad"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_rubrec.frx":0A1E
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frm_rubrec.frx":0A35
      TabIndex        =   12
      Top             =   3960
      Width           =   6855
   End
   Begin VB.TextBox txt_bcob 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Data data_cob 
      Caption         =   "data_cob"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "COD_CAJA"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton bimp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_rubrec.frx":1410
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Informes"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_rubrec.frx":199A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Buscar"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      Picture         =   "frm_rubrec.frx":1F24
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancelar acción"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "frm_rubrec.frx":24AE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Modificar datos"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_rubrec.frx":2A38
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bnuevo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "frm_rubrec.frx":2FC2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo registro"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txt_nomcob 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txt_nrocob 
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Moneda:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Concepto:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nombre a buscar:"
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
      TabIndex        =   10
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Descripción:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Código:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   5280
      Picture         =   "frm_rubrec.frx":354C
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1935
   End
End
Attribute VB_Name = "frm_rubrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_elim_Click()
Dim Loborra As String
Loborra = MsgBox(WElusuario & " desea eliminar el registro seleccionado?", vbInformation + vbYesNo, "Eliminar registro")
If Loborra = vbYes Then
   data_cob.Recordset.Delete
   data_cob.Refresh
   MsgBox "Registro Eliminado!"
End If

End Sub

Private Sub bbusca_Click()
txt_bcob.Enabled = True
DBGrid1.Enabled = True
txt_bcob.SetFocus

End Sub

Private Sub bcance_Click()
If XAcnv = 1 Then
   data_cob.Recordset.CancelUpdate
   igualcob
   XAcnv = 0
   desh
Else
   igualcob
   XAcnv = 0
   desh
End If
bgraba.Enabled = False
bcance.Enabled = False
bmodif.Enabled = True
bbusca.Enabled = True
bimp.Enabled = True
bnuevo.Enabled = True

End Sub

Private Sub bgraba_Click()
If txt_nrocob.Text <> "" Then
   If txt_nrocob.Text <> 0 Then
         If XAcnv = 1 Then
            data_cob.Recordset("numero") = txt_nrocob.Text
            data_cob.Recordset("nombre") = txt_nomcob.Text
            If cbomon.ListIndex = 1 Then
               data_cob.Recordset("moneda") = "U$S"
            Else
               data_cob.Recordset("moneda") = "$"
            End If
            data_cob.Recordset("movimiento") = cbocon.Text
            data_cob.Recordset.Update
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            b_elim.Enabled = True
            bnuevo.Enabled = True
            desh
         Else
            If data_cob.Recordset("numero") <> txt_nrocob.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("numero") = txt_nrocob.Text
               data_cob.Recordset.Update
            End If
            If data_cob.Recordset("nombre") <> txt_nomcob.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("nombre") = txt_nomcob.Text
               data_cob.Recordset.Update
            End If
            If cbomon.ListIndex = 1 Then
               If data_cob.Recordset("moneda") <> "U$S" Then
                  data_cob.Recordset.Edit
                  data_cob.Recordset("moneda") = "U$S"
                  data_cob.Recordset.Update
               End If
            Else
               If data_cob.Recordset("moneda") <> "$" Then
                  data_cob.Recordset.Edit
                  data_cob.Recordset("moneda") = "$"
                  data_cob.Recordset.Update
               End If
            End If
            If data_cob.Recordset("movimiento") <> cbocon.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("movimiento") = cbocon.Text
               data_cob.Recordset.Update
            End If
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            b_elim.Enabled = True
            bnuevo.Enabled = True
            desh
         End If
   Else
      MsgBox "No ingresó datos", vbCritical, "Mensaje"
      txt_nrocob.SetFocus
   End If
Else
   MsgBox "No ingresó datos", vbCritical, "Mensaje"
   txt_nrocob.SetFocus
End If

End Sub

Private Sub bimp_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
If data_cob.Recordset.BOF = False Then
   data_cob.Recordset.MoveFirst
End If

Do While Not data_cob.Recordset.EOF
    frm_rubrec.MousePointer = 11
    bimp.Enabled = False
    data_inf.Recordset.AddNew
    data_inf.Recordset("cod_cli") = data_cob.Recordset("numero")
    data_inf.Recordset("nom_cli") = Mid(data_cob.Recordset("nombre"), 1, 30)
    data_inf.Recordset("hora") = data_cob.Recordset("moneda")
    data_inf.Recordset("nom_flia") = data_cob.Recordset("movimiento")
    data_inf.Recordset.Update
    data_cob.Recordset.MoveNext
Loop
bimp.Enabled = True
frm_rubrec.MousePointer = 0
data_inf.RecordSource = "Select * from infvtas"
data_inf.Refresh
cr1.ReportFileName = App.Path & "\infrubrosr.rpt"
cr1.Action = 1

End Sub

Private Sub bmodif_Click()
XAcnv = 0
hab
txt_nrocob.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
b_elim.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False

End Sub

Private Sub bnuevo_Click()
XAcnv = 1
hab
txt_nrocob.Text = ""
txt_nomcob.Text = ""
cbomon.ListIndex = 0
cbocon.ListIndex = 0
txt_nrocob.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
b_elim.Enabled = False

data_cob.Recordset.AddNew

End Sub

Private Sub cbocon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub cbomon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocon.SetFocus
End If

End Sub

Private Sub DBGrid1_DblClick()
    If IsNull(data_cob.Recordset("numero")) = False Then
       txt_nrocob.Text = data_cob.Recordset("numero")
    Else
       txt_nrocob.Text = ""
    End If
    If IsNull(data_cob.Recordset("nombre")) = False Then
       txt_nomcob.Text = data_cob.Recordset("nombre")
    Else
       txt_nomcob.Text = ""
    End If
    If IsNull(data_cob.Recordset("moneda")) = False Then
       If data_cob.Recordset("moneda") = "U$S" Then
          cbomon.ListIndex = 1
       Else
          cbomon.ListIndex = 0
       End If
    Else
       cbomon.ListIndex = 0
    End If
    If IsNull(data_cob.Recordset("movimiento")) = False Then
       If data_cob.Recordset("movimiento") = "EGRESO" Then
          cbocon.ListIndex = 1
       Else
          cbocon.ListIndex = 0
       End If
    Else
       cbocon.ListIndex = 0
    End If
txt_bcob.Enabled = False
DBGrid1.Enabled = False
bmodif.SetFocus


End Sub

Public Function hab()
txt_nrocob.Enabled = True
txt_nomcob.Enabled = True
cbomon.Enabled = True
cbocon.Enabled = True
End Function

Public Function desh()
txt_nrocob.Enabled = False
txt_nomcob.Enabled = False
cbomon.Enabled = False
cbocon.Enabled = False

End Function

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "cod_caja"
Data1.Refresh
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.RecordSource = "cod_caja"
data_cob.Refresh
data_inf.DatabaseName = App.Path & "\informes.mdb"

igualcob

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_bcob_Change()
data_cob.RecordSource = "select * from cod_caja where nombre >='" & txt_bcob.Text & "' order by nombre"
data_cob.Refresh

End Sub

Private Sub txt_bcob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   KeyAscii = 0
   DBGrid1.SetFocus
End If

End Sub

Public Function igualcob()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("numero")) = False Then
   txt_nrocob.Text = data_cob.Recordset("numero")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("moneda")) = False Then
   If data_cob.Recordset("moneda") = "U$S" Then
      cbomon.ListIndex = 1
   Else
      cbomon.ListIndex = 0
   End If
Else
   cbomon.ListIndex = 0
End If
If IsNull(data_cob.Recordset("movimiento")) = False Then
   If data_cob.Recordset("movimiento") = "EGRESO" Then
      cbocon.ListIndex = 1
   Else
      cbocon.ListIndex = 0
   End If
Else
   cbocon.ListIndex = 0
End If

End Function


Private Sub txt_espec_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txt_nomcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomon.SetFocus
End If

End Sub

Private Sub txt_nrocob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomcob.SetFocus
End If

End Sub

Private Sub txt_nrocob_LostFocus()
If XAcnv = 1 Then
   Data1.Recordset.FindFirst "numero =" & txt_nrocob.Text
   If Not Data1.Recordset.NoMatch Then
      MsgBox "Ya existe este número", vbCritical, "Mensaje"
   End If
End If

End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub
