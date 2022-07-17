VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pasdeuda 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasar deudas para emisión"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8805
   Icon            =   "frm_pasdeuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8805
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   4920
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_cli"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton b_elim 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_pasdeuda.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Eliminar el registro seleccionado"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Data data_busdeu 
      Caption         =   "data_busdeu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4800
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_infemi 
      Caption         =   "data_infemi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton b_genemi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generar Provisorios"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton b_bort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eliminar TODO"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton b_ant 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton b_sig 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Data data_deu 
      Caption         =   "data_deu"
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton b_selec 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_pasdeuda.frx":09CC
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frm_pasdeuda.frx":09E3
      TabIndex        =   10
      Top             =   3240
      Width           =   8535
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      Picture         =   "frm_pasdeuda.frx":170A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Informes"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      MaskColor       =   &H00C0FFC0&
      Picture         =   "frm_pasdeuda.frx":1C94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancelar acción"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_pasdeuda.frx":221E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar datos"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Picture         =   "frm_pasdeuda.frx":27A8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Editar registro"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton b_nuev 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frm_pasdeuda.frx":2D32
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Alta de registro"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txt_imp 
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
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txt_mat 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8760
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "IMPORTE:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MATRICULA:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   4080
      Picture         =   "frm_pasdeuda.frx":32BC
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2175
   End
End
Attribute VB_Name = "frm_pasdeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_bort_Click()
Dim Xsiborro As String
Xsiborro = MsgBox("Desea BORRAR todos las deudas?", vbYesNo + vbInformation, "Mensaje")
If Xsiborro = vbYes Then
   If data_deu.Recordset.RecordCount > 0 Then
      data_deu.Recordset.MoveFirst
      Do While Not data_deu.Recordset.EOF
         data_deu.Recordset.Delete
         data_deu.Recordset.MoveNext
      Loop
      MsgBox "Terminado"
   End If
End If

End Sub

Private Sub b_canc_Click()
If XAlta = 1 Then
   XAlta = 0
   b_nuev.Enabled = True
   b_graba.Enabled = False
   b_mod.Enabled = True
   b_canc.Enabled = False
   b_selec.Enabled = True
   b_sig.Enabled = True
   b_ant.Enabled = True
   data_deu.Recordset.CancelUpdate
   txt_mat.Enabled = False
   txt_imp.Enabled = False
   data_deu.Recordset.MoveLast
   txt_mat.Text = data_deu.Recordset("mat")
   txt_imp.Text = data_deu.Recordset("imp")
   Label2.Caption = data_deu.Recordset("nombre")
Else
   XAlta = 0
   b_nuev.Enabled = True
   b_graba.Enabled = False
   b_mod.Enabled = True
   b_canc.Enabled = False
   b_selec.Enabled = True
   b_sig.Enabled = True
   b_ant.Enabled = True
   txt_mat.Enabled = False
   txt_imp.Enabled = False
   txt_mat.Text = data_deu.Recordset("mat")
   txt_imp.Text = data_deu.Recordset("imp")
   Label2.Caption = data_deu.Recordset("nombre")
End If
End Sub

Private Sub b_elim_Click()
Dim Xsieliono As String
Xsieliono = MsgBox("Desea eliminar el registro de " & data_deu.Recordset("mat") & "?", vbInformation + vbYesNo)
If Xsieliono = vbYes Then
   data_deu.Recordset.Delete
   data_deu.Refresh
End If

End Sub

Private Sub b_genemi_Click()
Dim Xnomdeemi As String
Dim Xmes, Xano As Integer
Xmes = Month(Date)
Xano = Year(Date)

Xnomdeemi = "EMI"
If Xmes > 9 Then
   Xnomdeemi = Xnomdeemi + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
Else
   Xnomdeemi = Xnomdeemi + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
End If
frm_pasdeuda.MousePointer = 11
data_emi.RecordSource = "Select * from " & Xnomdeemi
data_emi.Refresh
If data_infemi.Recordset.RecordCount > 0 Then
   data_infemi.Recordset.MoveFirst
   Do While Not data_infemi.Recordset.EOF
      data_infemi.Recordset.Delete
      data_infemi.Recordset.MoveNext
   Loop
End If
If data_deu.Recordset.RecordCount > 0 Then
   Do While Not data_deu.Recordset.EOF
      data_emi.Recordset.FindFirst "cliente =" & data_deu.Recordset("mat")
      If Not data_emi.Recordset.NoMatch Then
         data_deu.Recordset.MoveNext
      Else
         data_cli.RecordSource = "select * from clientes where cl_codigo =" & data_deu.Recordset("mat")
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            data_infemi.Recordset.AddNew
            data_infemi.Recordset("cod_cnv") = data_cli.Recordset("cl_codconv")
            data_infemi.Recordset("nom_cnv") = Mid(data_cli.Recordset("cl_nomconv"), 1, 30)
            data_infemi.Recordset("cliente") = data_cli.Recordset("cl_codigo")
            data_infemi.Recordset("apellidos") = Mid(data_cli.Recordset("cl_apellid"), 1, 40)
            data_infemi.Recordset("cedula") = data_cli.Recordset("cl_cedula")
            data_infemi.Recordset("dir_cli") = Mid(data_cli.Recordset("cl_direcci"), 1, 80)
            data_infemi.Recordset("loc_cli") = Mid(data_cli.Recordset("cl_zona"), 1, 20)
            If IsNull(data_cli.Recordset("cl_telefon")) = False Then
               If data_cli.Recordset("cl_telefon") <> "" Then
                  data_infemi.Recordset("tel_cli") = Mid(data_cli.Recordset("cl_telefon"), 1, 15)
               Else
                  data_infemi.Recordset("tel_cli") = "SIN T"
               End If
            Else
               data_infemi.Recordset("tel_cli") = "SIN T"
            End If
            data_infemi.Recordset("grupo") = data_cli.Recordset("cl_grupo")
            data_infemi.Recordset("fecha_ing") = data_cli.Recordset("cl_fecing")
            data_infemi.Recordset("documento") = 0
            data_infemi.Recordset("importe") = 0
            data_infemi.Recordset("nro_cobr") = data_cli.Recordset("cl_nrocobr")
            data_infemi.Recordset("nom_cobr") = Mid(data_cli.Recordset("cl_nomcobr"), 1, 20)
            data_infemi.Recordset("mes") = Xmes
            data_infemi.Recordset("ano") = Xano
            data_infemi.Recordset("color_rec") = "A"
            data_infemi.Recordset("tiquet") = 0
            data_infemi.Recordset("servi") = 0
            data_infemi.Recordset("deudas") = data_deu.Recordset("imp")
            data_infemi.Recordset("iva") = 0
            data_infemi.Recordset("total") = data_deu.Recordset("imp")
            data_infemi.Recordset.Update
         End If
         data_deu.Recordset.MoveNext
      End If
   Loop
   CrystalReport1.ReportFileName = App.path & "\rspsappp.rpt"
   CrystalReport1.Action = 1
End If
frm_pasdeuda.MousePointer = 0

End Sub

Private Sub b_graba_Click()
If XAlta = 1 Then
   If txt_mat.Text <> "" Then
      data_deu.Recordset("mat") = txt_mat.Text
      If txt_imp.Text <> "" Then
         data_deu.Recordset("imp") = txt_imp.Text
      Else
         data_deu.Recordset("imp") = 0
      End If
      data_deu.Recordset("nombre") = Label2.Caption
      data_deu.Recordset("fecha") = Date
      data_deu.Recordset.Update
      XAlta = 0
      b_nuev.Enabled = True
      b_graba.Enabled = False
      b_mod.Enabled = True
      b_canc.Enabled = False
      b_selec.Enabled = True
      b_sig.Enabled = True
      b_ant.Enabled = True
      txt_mat.Enabled = False
      txt_imp.Enabled = False
      data_deu.Refresh
      data_deu.Recordset.MoveLast
      txt_mat.Text = data_deu.Recordset("mat")
      txt_imp.Text = data_deu.Recordset("imp")
      Label2.Caption = data_deu.Recordset("nombre")
   Else
      MsgBox "No se ingresó matrícula", vbCritical, "Mensaje"
      txt_mat.SetFocus
   End If
Else
   If txt_mat.Text <> "" Then
      data_deu.Recordset.Edit
      data_deu.Recordset("mat") = txt_mat.Text
      If txt_imp.Text <> "" Then
         data_deu.Recordset("imp") = txt_imp.Text
      Else
         data_deu.Recordset("imp") = 0
      End If
      data_deu.Recordset("nombre") = Label2.Caption
      data_deu.Recordset("fecha") = Date
      data_deu.Recordset.Update
      XAlta = 0
      b_nuev.Enabled = True
      b_graba.Enabled = False
      b_mod.Enabled = True
      b_canc.Enabled = False
      b_selec.Enabled = True
      b_sig.Enabled = True
      b_ant.Enabled = True
      txt_mat.Enabled = False
      txt_imp.Enabled = False
      txt_mat.Text = data_deu.Recordset("mat")
      txt_imp.Text = data_deu.Recordset("imp")
      Label2.Caption = data_deu.Recordset("nombre")
   Else
      MsgBox "No se ingresó matrícula", vbCritical, "Mensaje"
      txt_mat.SetFocus
   End If
End If
End Sub

Private Sub b_mod_Click()
XAlta = 0
b_nuev.Enabled = False
b_graba.Enabled = True
b_mod.Enabled = False
b_canc.Enabled = True
b_selec.Enabled = False
b_sig.Enabled = False
b_ant.Enabled = False
txt_mat.Enabled = True
txt_imp.Enabled = True
txt_mat.SetFocus

End Sub

Private Sub b_nuev_Click()
XAlta = 1
b_nuev.Enabled = False
b_graba.Enabled = True
b_mod.Enabled = False
b_canc.Enabled = True
b_selec.Enabled = False
b_sig.Enabled = False
b_ant.Enabled = False
data_deu.Recordset.AddNew
txt_mat.Enabled = True
txt_imp.Enabled = True
txt_mat.SetFocus
txt_mat.Text = ""
txt_imp.Text = ""
Label2.Caption = ""

End Sub

Private Sub b_selec_Click()
txt_mat.Text = data_deu.Recordset("mat")
Label2.Caption = data_deu.Recordset("nombre")
txt_imp.Text = Format(data_deu.Recordset("imp"), "Standard")

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Initialize()
If data_deu.Recordset.RecordCount > 0 Then
   data_deu.Recordset.MoveLast
   txt_mat.Text = data_deu.Recordset("mat")
   txt_imp.Text = data_deu.Recordset("imp")
   Label2.Caption = data_deu.Recordset("nombre")
End If

End Sub

Private Sub Form_Load()
data_cli.ConnectionString = "dsn=" & Xconexrmt

data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deu.RecordSource = "emitiq"
data_deu.Refresh

data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_infemi.DatabaseName = App.path & "\informes.mdb"
data_infemi.RecordSource = "infemirec"
data_infemi.Refresh
data_busdeu.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_imp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub txt_imp_LostFocus()
If txt_imp.Text <> "" Then
   txt_imp.Text = Format(txt_imp.Text, "Standard")
End If

End Sub

Private Sub txt_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_imp.SetFocus
End If

End Sub

Private Sub txt_mat_LostFocus()
If txt_mat.Text <> "" Then
'   data_cli.Recordset.FindFirst "cl_codigo =" & txt_mat.Text
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & txt_mat.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      Label2.Caption = data_cli.Recordset("cl_apellid")
   Else
      Label2.Caption = ""
      MsgBox "No encontrado", vbCritical, "Mensaje"
      txt_mat.SetFocus
   End If
   If XAlta = 1 Then
      data_busdeu.RecordSource = "Select * from emitiq where mat =" & txt_mat.Text
      data_busdeu.Refresh
      If data_busdeu.Recordset.RecordCount > 0 Then
         MsgBox "Ya existe una deuda registrada a ésta matrícula, modifique la existente!", vbInformation
         b_canc_Click
      End If
   End If
End If

End Sub
