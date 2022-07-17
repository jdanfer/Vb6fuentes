VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_mutu 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutualistas"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frm_mutu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_espec 
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
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox txt_tel 
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
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\sapp\sociedad.rpt"
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
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sociedad"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_mutu.frx":0442
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frm_mutu.frx":0459
      TabIndex        =   14
      Top             =   3840
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
      TabIndex        =   13
      Top             =   3480
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
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sociedad"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton bimp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_mutu.frx":0E38
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Informes"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_mutu.frx":13C2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_mutu.frx":194C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar acción"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "frm_mutu.frx":1ED6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Modificar datos"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_mutu.frx":2460
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Guardar datos"
      Top             =   2520
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
      Picture         =   "frm_mutu.frx":29EA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo registro"
      Top             =   2520
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
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Dirección:"
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
      TabIndex        =   16
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Teléfonos:"
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
      TabIndex        =   15
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
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   3240
      Y2              =   3240
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
      Height          =   1695
      Left            =   5160
      Picture         =   "frm_mutu.frx":2F74
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "frm_mutu"
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
            data_cob.Recordset("soc_nro") = txt_nrocob.Text
            data_cob.Recordset("soc_nombre") = txt_nomcob.Text
            data_cob.Recordset("soc_dir") = txt_espec.Text
            data_cob.Recordset("soc_tel") = txt_tel.Text
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
            data_cob.Recordset("soc_nro") = txt_nrocob.Text
            data_cob.Recordset("soc_nombre") = txt_nomcob.Text
            data_cob.Recordset("soc_dir") = txt_espec.Text
            data_cob.Recordset("soc_tel") = txt_tel.Text
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
      MsgBox "No ingresó datos", vbCritical, "Mensaje"
      txt_nrocob.SetFocus
   End If
Else
   MsgBox "No ingresó datos", vbCritical, "Mensaje"
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
txt_tel.Text = ""
txt_espec.Text = ""
txt_nrocob.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
data_cob.RecordSource = "Select * from sociedad order by SOC_NRO DESC"
data_cob.Refresh
data_cob.Recordset.MoveFirst
txt_nrocob.Text = data_cob.Recordset("SOC_NRO") + 1

data_cob.Recordset.AddNew

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNull(data_cob.Recordset("soc_nro")) = False Then
       txt_nrocob.Text = data_cob.Recordset("soc_nro")
    Else
       txt_nrocob.Text = ""
    End If
    If IsNull(data_cob.Recordset("soc_nombre")) = False Then
       txt_nomcob.Text = data_cob.Recordset("soc_nombre")
    Else
       txt_nomcob.Text = ""
    End If
    If IsNull(data_cob.Recordset("soc_dir")) = False Then
       txt_espec.Text = data_cob.Recordset("soc_dir")
    Else
       txt_espec.Text = ""
    End If
    If IsNull(data_cob.Recordset("soc_tel")) = False Then
       txt_tel.Text = data_cob.Recordset("soc_tel")
    Else
       txt_tel.Text = ""
    End If
End If
txt_bcob.Enabled = False
DBGrid1.Enabled = False
bmodif.SetFocus

End Sub

Private Sub Form_Initialize()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("soc_nro")) = False Then
   txt_nrocob.Text = data_cob.Recordset("soc_nro")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("soc_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("soc_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("soc_dir")) = False Then
   txt_espec.Text = data_cob.Recordset("soc_dir")
Else
   txt_espec.Text = ""
End If
If IsNull(data_cob.Recordset("soc_tel")) = False Then
   txt_tel.Text = data_cob.Recordset("soc_tel")
Else
   txt_tel.Text = ""
End If

End Sub

Public Function hab()
txt_nrocob.Enabled = True
txt_nomcob.Enabled = True
txt_tel.Enabled = True
txt_espec.Enabled = True
End Function

Public Function desh()
txt_nrocob.Enabled = False
txt_nomcob.Enabled = False
txt_tel.Enabled = False
txt_espec.Enabled = False
End Function

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
CrystalReport1.ReportFileName = App.Path & "\sociedad.rpt"

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
data_cob.RecordSource = "select * from sociedad where soc_nombre >='" & txt_bcob.Text & "' order by soc_nombre"
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
If IsNull(data_cob.Recordset("soc_nro")) = False Then
   txt_nrocob.Text = data_cob.Recordset("soc_nro")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("soc_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("soc_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("soc_dir")) = False Then
   txt_espec.Text = data_cob.Recordset("soc_dir")
Else
   txt_espec.Text = ""
End If
If IsNull(data_cob.Recordset("soc_tel")) = False Then
   txt_tel.Text = data_cob.Recordset("soc_tel")
Else
   txt_tel.Text = ""
End If
End Function


Private Sub txt_espec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_tel.SetFocus
End If

End Sub

Private Sub txt_nomcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_espec.SetFocus
End If

End Sub

Private Sub txt_nrocob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomcob.SetFocus
End If

End Sub

Private Sub txt_nrocob_LostFocus()
If XAcnv = 1 Then
   Data1.Recordset.FindFirst "soc_nro =" & txt_nrocob.Text
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
