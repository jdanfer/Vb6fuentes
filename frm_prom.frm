VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_prom 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promotores"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frm_prom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport CrystalReport1 
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
      Height          =   300
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "vendedor"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_prom.frx":0442
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frm_prom.frx":0459
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
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "vendedor"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton bimp 
      Height          =   495
      Left            =   4320
      Picture         =   "frm_prom.frx":0E3C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      Height          =   495
      Left            =   3480
      Picture         =   "frm_prom.frx":127E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_prom.frx":16C0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      Height          =   495
      Left            =   1800
      Picture         =   "frm_prom.frx":1B02
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Modificar datos"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_prom.frx":1F44
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton bnuevo 
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
      Picture         =   "frm_prom.frx":2386
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Top             =   1080
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
      Width           =   1455
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
      Caption         =   "Nombre:"
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
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nro.Promotor:"
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
      Height          =   1815
      Left            =   5160
      Picture         =   "frm_prom.frx":27C8
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "frm_prom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
            data_cob.Recordset("vn_numero") = txt_nrocob.Text
            data_cob.Recordset("vn_nombre") = txt_nomcob.Text
            data_cob.Recordset.Update
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            desh
         Else
            data_cob.Recordset.Edit
            data_cob.Recordset("vn_numero") = txt_nrocob.Text
            data_cob.Recordset("vn_nombre") = txt_nomcob.Text
            data_cob.Recordset.Update
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            desh
         End If
   Else
      MsgBox "No ingresó cobrador", vbCritical, "Cobradores"
      txt_nrocob.SetFocus
   End If
Else
   MsgBox "No ingresó cobrador", vbCritical, "Cobradores"
   txt_nrocob.SetFocus
End If

End Sub

Private Sub bimp_Click()
CrystalReport1.Action = 1

End Sub

Private Sub bmodif_Click()
XAcnv = 0
hab
txt_nrocob.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False

End Sub

Private Sub bnuevo_Click()
XAcnv = 1
hab
txt_nrocob.Text = ""
txt_nomcob.Text = ""
txt_nrocob.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
data_cob.Recordset.AddNew

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNull(data_cob.Recordset("vn_numero")) = False Then
       txt_nrocob.Text = data_cob.Recordset("vn_numero")
    Else
       txt_nrocob.Text = ""
    End If
    If IsNull(data_cob.Recordset("vn_nombre")) = False Then
       txt_nomcob.Text = data_cob.Recordset("vn_nombre")
    Else
       txt_nomcob.Text = ""
    End If
End If
txt_bcob.Enabled = False
DBGrid1.Enabled = False
bmodif.SetFocus

End Sub

Private Sub Form_Initialize()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("vn_numero")) = False Then
   txt_nrocob.Text = data_cob.Recordset("vn_numero")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("vn_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("vn_nombre")
Else
   txt_nomcob.Text = ""
End If

End Sub

Public Function hab()
txt_nrocob.Enabled = True
txt_nomcob.Enabled = True

End Function

Public Function desh()
txt_nrocob.Enabled = False
txt_nomcob.Enabled = False

End Function

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
CrystalReport1.ReportFileName = App.Path & "\promotor.rpt"

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
data_cob.RecordSource = "select * from vendedor where vn_nombre >='" & txt_bcob.Text & "' order by vn_nombre"
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
If IsNull(data_cob.Recordset("vn_numero")) = False Then
   txt_nrocob.Text = data_cob.Recordset("vn_numero")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("vn_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("vn_nombre")
Else
   txt_nomcob.Text = ""
End If

End Function


Private Sub txt_nomcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub txt_nrocob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomcob.SetFocus
End If

End Sub

Private Sub txt_nrocob_LostFocus()
If XAcnv = 1 Then
   Data1.Recordset.FindFirst "vn_numero =" & txt_nrocob.Text
   If Not Data1.Recordset.NoMatch Then
      MsgBox "Ya existe este número de promotor", vbCritical, "Promotores"
   End If
End If

End Sub
