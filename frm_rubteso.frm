VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_rubteso 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros tesorería"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frm_rubteso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
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
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton b_elim 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      Picture         =   "frm_rubteso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txt_lib 
      Alignment       =   2  'Center
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
      Left            =   5520
      MaxLength       =   1
      TabIndex        =   22
      Top             =   1680
      Width           =   615
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
      ItemData        =   "frm_rubteso.frx":09CC
      Left            =   2160
      List            =   "frm_rubteso.frx":09D6
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2520
      Width           =   3015
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
      ItemData        =   "frm_rubteso.frx":09FF
      Left            =   2160
      List            =   "frm_rubteso.frx":0A09
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txt_deb 
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
      Width           =   1815
   End
   Begin VB.TextBox txt_hab 
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
      Top             =   1680
      Width           =   1815
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6720
      Top             =   2400
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
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RUBTESO"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_rubteso.frx":0A1E
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "frm_rubteso.frx":0A35
      TabIndex        =   14
      Top             =   4440
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
      Left            =   2160
      TabIndex        =   13
      Top             =   4080
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
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RUBTESO"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton bimp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_rubteso.frx":1410
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Informes"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_rubteso.frx":199A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      Picture         =   "frm_rubteso.frx":1F24
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar acción"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "frm_rubteso.frx":24AE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Modificar datos"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_rubteso.frx":2A38
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar"
      Top             =   3240
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
      Picture         =   "frm_rubteso.frx":2FC2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo registro"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txt_nom 
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
   Begin VB.TextBox txt_nrorub 
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
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      Caption         =   "LIBRO:"
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
      Left            =   4200
      TabIndex        =   21
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Al Debe:"
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
      TabIndex        =   16
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Al Haber:"
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
      TabIndex        =   15
      Top             =   1680
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
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   5400
      Picture         =   "frm_rubteso.frx":354C
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frm_rubteso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_elim_Click()
Dim Loborra As String
Loborra = MsgBox(WElusuario & " desea eliminar el registro seleccionado " & data_cob.Recordset("codigo") & " ?", vbInformation + vbYesNo, "Eliminar registro")
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
'   data_cob.Recordset.CancelUpdate
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
If txt_nrorub.Text <> "" Then
   If txt_nrorub.Text <> 0 Then
         If XAcnv = 1 Then
            data_cob.Recordset.AddNew
            data_cob.Recordset("codigo") = txt_nrorub.Text
            data_cob.Recordset("nombre") = txt_nom.Text
            data_cob.Recordset("debe") = txt_deb.Text
            data_cob.Recordset("haber") = txt_hab.Text
            If cbocon.ListIndex = 0 Then
               data_cob.Recordset("es") = "E"
               data_cob.Recordset("concep") = "ENTRADA"
            Else
               If cbocon.ListIndex = 1 Then
                  data_cob.Recordset("es") = "S"
                  data_cob.Recordset("concep") = "SALIDA"
               Else
                  data_cob.Recordset("es") = "E"
                  data_cob.Recordset("concep") = "ENTRADA"
               End If
            End If
            If cbomon.ListIndex = 0 Then
               data_cob.Recordset("moneda") = 1
            Else
               If cbomon.ListIndex = 1 Then
                  data_cob.Recordset("moneda") = 2
               Else
                  data_cob.Recordset("moneda") = 1
               End If
            End If
            data_cob.Recordset("libro") = txt_lib.Text
            data_cob.Recordset.Update
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            b_elim.Enabled = True
            desh
         Else
            If data_cob.Recordset("codigo") <> txt_nrorub.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("codigo") = txt_nrorub.Text
               data_cob.Recordset.Update
            End If
            If data_cob.Recordset("nombre") <> txt_nom.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("nombre") = txt_nom.Text
               data_cob.Recordset.Update
            End If
            If data_cob.Recordset("debe") <> txt_deb.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("debe") = txt_deb.Text
               data_cob.Recordset.Update
            End If
            If data_cob.Recordset("haber") <> txt_hab.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("haber") = txt_hab.Text
               data_cob.Recordset.Update
            End If
            If cbocon.ListIndex = 0 Then
               If data_cob.Recordset("es") <> "E" Then
                  data_cob.Recordset.Edit
                  data_cob.Recordset("es") = "E"
                  data_cob.Recordset("concep") = "ENTRADA"
                  data_cob.Recordset.Update
               End If
            Else
               If cbocon.ListIndex = 1 Then
                  If data_cob.Recordset("es") <> "S" Then
                     data_cob.Recordset.Edit
                     data_cob.Recordset("es") = "S"
                     data_cob.Recordset("concep") = "SALIDA"
                     data_cob.Recordset.Update
                  End If
               End If
            End If
            If cbomon.ListIndex = 0 Then
               If data_cob.Recordset("moneda") <> 1 Then
                  data_cob.Recordset.Edit
                  data_cob.Recordset("moneda") = 1
                  data_cob.Recordset.Update
               End If
            Else
               If cbomon.ListIndex = 1 Then
                  If data_cob.Recordset("moneda") <> 2 Then
                     data_cob.Recordset.Edit
                     data_cob.Recordset("moneda") = 2
                     data_cob.Recordset.Update
                  End If
               End If
            End If
            If data_cob.Recordset("libro") <> txt_lib.Text Then
               data_cob.Recordset.Edit
               data_cob.Recordset("libro") = txt_lib.Text
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
      txt_nrorub.SetFocus
   End If
Else
   MsgBox "No ingresó datos", vbCritical, "Mensaje"
   txt_nrorub.SetFocus
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
    frm_rubteso.MousePointer = 11
    bimp.Enabled = False
    data_inf.Recordset.AddNew
    data_inf.Recordset("cod_cli") = data_cob.Recordset("codigo")
    data_inf.Recordset("nom_cli") = Mid(data_cob.Recordset("nombre"), 1, 30)
    If data_cob.Recordset("moneda") = 1 Then
       data_inf.Recordset("hora") = "$."
    Else
       data_inf.Recordset("hora") = "U$."
    End If
    data_inf.Recordset("nom_flia") = data_cob.Recordset("libro")
    data_inf.Recordset.Update
    data_cob.Recordset.MoveNext
Loop
bimp.Enabled = True
frm_rubteso.MousePointer = 0
data_inf.RecordSource = "Select * from infvtas"
data_inf.Refresh
cr1.ReportFileName = App.Path & "\infrubrosr.rpt"
cr1.Action = 1

End Sub

Private Sub bmodif_Click()
XAcnv = 0
hab
txt_nrorub.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
b_elim.Enabled = False


End Sub

Private Sub bnuevo_Click()
XAcnv = 1
txt_nrorub.Text = ""
txt_nom.Text = ""
txt_deb.Text = ""
txt_hab.Text = ""
txt_lib.Text = ""
cbocon.ListIndex = 0
cbomon.ListIndex = 0
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
b_elim.Enabled = False

hab
txt_nrorub.SetFocus

End Sub

Private Sub cbocon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbomon.SetFocus
   cbomon.ListIndex = 0
End If

End Sub

Private Sub cbomon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub DBGrid1_DblClick()
'If KeyAscii = 13 Then
    If IsNull(data_cob.Recordset("codigo")) = False Then
       txt_nrorub.Text = data_cob.Recordset("codigo")
    Else
       txt_nrorub.Text = ""
    End If
    If IsNull(data_cob.Recordset("nombre")) = False Then
       txt_nom.Text = data_cob.Recordset("nombre")
    Else
       txt_nom.Text = ""
    End If
    If IsNull(data_cob.Recordset("debe")) = False Then
       txt_deb.Text = data_cob.Recordset("debe")
    Else
       txt_deb.Text = ""
    End If
    If IsNull(data_cob.Recordset("haber")) = False Then
       txt_hab.Text = data_cob.Recordset("haber")
    Else
       txt_hab.Text = ""
    End If
    If IsNull(data_cob.Recordset("libro")) = False Then
       txt_lib.Text = data_cob.Recordset("libro")
    Else
       txt_lib.Text = ""
    End If
    If IsNull(data_cob.Recordset("es")) = False Then
       If data_cob.Recordset("es") = "E" Then
          cbocon.ListIndex = 0
       Else
          cbocon.ListIndex = 1
       End If
    Else
       cbocon.ListIndex = 0
    End If
    If IsNull(data_cob.Recordset("moneda")) = False Then
       If data_cob.Recordset("moneda") = 2 Then
          cbomon.ListIndex = 1
       Else
          cbomon.ListIndex = 0
       End If
    Else
       cbomon.ListIndex = 0
    End If

'End If
txt_bcob.Enabled = False
DBGrid1.Enabled = False
bmodif.SetFocus

End Sub

Public Function hab()
txt_nrorub.Enabled = True
txt_nom.Enabled = True
txt_deb.Enabled = True
txt_hab.Enabled = True
txt_lib.Enabled = True
cbocon.Enabled = True
cbomon.Enabled = True

End Function

Public Function desh()
txt_nrorub.Enabled = False
txt_nom.Enabled = False
txt_deb.Enabled = False
txt_hab.Enabled = False
txt_lib.Enabled = False
cbocon.Enabled = False
cbomon.Enabled = False

End Function

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "rubteso"
Data1.Refresh
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.RecordSource = "rubteso"
data_cob.Refresh
igualcob
data_inf.DatabaseName = App.Path & "\informes.mdb"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub txt_bcob_Change()
data_cob.RecordSource = "select * from rubteso where nombre >='" & txt_bcob.Text & "' order by nombre"
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
If IsNull(data_cob.Recordset("codigo")) = False Then
   txt_nrorub.Text = data_cob.Recordset("codigo")
Else
   txt_nrorub.Text = ""
End If
If IsNull(data_cob.Recordset("nombre")) = False Then
   txt_nom.Text = data_cob.Recordset("nombre")
Else
   txt_nom.Text = ""
End If
If IsNull(data_cob.Recordset("debe")) = False Then
   txt_deb.Text = data_cob.Recordset("debe")
Else
   txt_deb.Text = ""
End If
If IsNull(data_cob.Recordset("haber")) = False Then
   txt_hab.Text = data_cob.Recordset("haber")
Else
   txt_hab.Text = ""
End If
If IsNull(data_cob.Recordset("libro")) = False Then
   txt_lib.Text = data_cob.Recordset("libro")
Else
   txt_lib.Text = ""
End If
If IsNull(data_cob.Recordset("es")) = False Then
   If data_cob.Recordset("es") = "S" Then
      cbocon.ListIndex = 1
   Else
      cbocon.ListIndex = 0
   End If
End If
If IsNull(data_cob.Recordset("moneda")) = False Then
   If data_cob.Recordset("moneda") = 2 Then
      cbomon.ListIndex = 1
   Else
      cbomon.ListIndex = 0
   End If
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

End Sub

Private Sub txt_nrocob_LostFocus()

End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub txt_deb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_hab.SetFocus
End If

End Sub

Private Sub txt_hab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_lib.SetFocus
End If

End Sub

Private Sub txt_lib_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cbocon.ListIndex = 0
   cbocon.SetFocus
End If

End Sub

Private Sub txt_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_deb.SetFocus
End If

End Sub

Private Sub txt_nrorub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nom.SetFocus
End If

End Sub

Private Sub txt_nrorub_LostFocus()
If XAcnv = 1 Then
   If txt_nrorub.Text <> "" Then
      data_cob.RecordSource = "Select * from rubteso where codigo =" & txt_nrorub.Text
      data_cob.Refresh
      If data_cob.Recordset.RecordCount > 0 Then
         MsgBox "Ya existe el código", vbCritical, "Mensaje"
         txt_nrorub.SetFocus
      End If
   End If
   data_cob.RecordSource = "select * from rubteso"
   data_cob.Refresh
End If


End Sub
