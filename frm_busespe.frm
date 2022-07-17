VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_busespe 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Buscar..."
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7435
      _Version        =   393216
      BackColorSel    =   -2147483631
      ForeColorSel    =   192
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton b_cerrar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cerrar"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frm_busespe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_cerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
'DBGrid1
If Xcolesp > 0 Then
    With msf1
     .Row = Xcolesp
     .TopRow = msf1.Row
     .RowSel = msf1.Row
     .Col = 0
     .ColSel = msf1.Cols - 1
    End With
    
End If

End Sub


Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\sapp.mdb"
If WElusuario = "JFERNAN" Or WElusuario = "CLAUDIA" Or WElusuario = "GFERNANDEZ" Or WElusuario = "MARIAROSA" Or frm_menu.data_parse.Recordset("base") = 15 Or _
   WElusuario = "MPEREZ" Or WElusuario = "JONATHAN" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "AACUÑA" Or WElusuario = "MSANCHEZ" Or WElusuario = "GUSTAVO" Or WElusuario = "MIKAELA" Then
   Data1.RecordSource = "Select * from lineas order by base,hora"
   Data1.Refresh
Else
   Data1.RecordSource = "Select * from lineas where base =" & frm_espec.Data1.Recordset("base")
   Data1.Refresh
End If
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
   Data1.Recordset.MoveFirst
End If
Dim Xc As Long
Xc = 0
msf1.Rows = Data1.Recordset.RecordCount + 1
msf1.Cols = 7

msf1.TextMatrix(Xc, 0) = "BASE"
msf1.TextMatrix(Xc, 1) = "CODIGO"
msf1.TextMatrix(Xc, 2) = "DESCRIPCION"
msf1.ColWidth(2) = 4500
msf1.TextMatrix(Xc, 3) = "HH"
msf1.ColWidth(3) = 450
msf1.TextMatrix(Xc, 4) = "MM"
msf1.ColWidth(4) = 450
msf1.TextMatrix(Xc, 5) = "ESPERA"
msf1.TextMatrix(Xc, 6) = "CANT.P"
Xc = Xc + 1

Do While Not Data1.Recordset.EOF
   msf1.TextMatrix(Xc, 0) = Data1.Recordset("base")
   msf1.TextMatrix(Xc, 1) = Data1.Recordset("hora")
   msf1.TextMatrix(Xc, 2) = Data1.Recordset("nom_medic")
   msf1.TextMatrix(Xc, 3) = Data1.Recordset("convenio")
   msf1.TextMatrix(Xc, 4) = Data1.Recordset("moneda")
   msf1.TextMatrix(Xc, 5) = Data1.Recordset("mes_paga")
   msf1.TextMatrix(Xc, 6) = Data1.Recordset("imp_iva")

   Data1.Recordset.MoveNext
   Xc = Xc + 1
Loop
msf1.Row = 1
msf1.Col = 0
If Xcolesp > 0 Then
   Command1.Visible = True
   Command1_Click
   Command1.Visible = False
    
End If

End Sub

Private Sub msf1_DblClick()
frm_espec.data_espec.Recordset.FindFirst "base =" & msf1.TextMatrix(msf1.RowSel, 0) & " And hora ='" & msf1.TextMatrix(msf1.RowSel, 1) & "'"

If Not frm_espec.data_espec.Recordset.NoMatch Then
   frm_espec.txt_base.Text = frm_espec.data_espec.Recordset("base")
   frm_espec.txt_cod.Text = frm_espec.data_espec.Recordset("hora")
   frm_espec.txt_desc.Text = frm_espec.data_espec.Recordset("nom_medic")
   frm_espec.txt_hh.Text = frm_espec.data_espec.Recordset("convenio")
   frm_espec.txt_mm.Text = frm_espec.data_espec.Recordset("moneda")
   frm_espec.txt_mmpp.Text = Int(frm_espec.data_espec.Recordset("cod_medic"))
   frm_espec.txt_cantp.Text = Int(frm_espec.data_espec.Recordset("imp_iva"))
   frm_espec.txt_espera.Text = Int(frm_espec.data_espec.Recordset("mes_paga"))
   If IsNull(frm_espec.data_espec.Recordset("factura")) = False Then
      frm_espec.t_hh.Text = Int(frm_espec.data_espec.Recordset("factura"))
   Else
      frm_espec.t_hh.Text = 0
   End If
   If IsNull(frm_espec.data_espec.Recordset("cod_cli")) = False Then
      frm_espec.t_mh.Text = Int(frm_espec.data_espec.Recordset("cod_cli"))
   Else
      frm_espec.t_mh.Text = 0
   End If
   frm_espec.Check1.value = frm_espec.data_espec.Recordset("reg_cab")
   If IsNull(frm_espec.data_espec.Recordset("cod_prod")) = False Then
      frm_espec.t_cmed.Text = frm_espec.data_espec.Recordset("cod_prod")
   Else
      frm_espec.t_cmed.Text = 0
      frm_espec.cbomed.Text = ""
   End If
   If frm_espec.t_cmed.Text <> 0 Then
      frm_espec.data_medicos.RecordSource = "Select * from medicos where med_cod =" & frm_espec.t_cmed.Text
      frm_espec.data_medicos.Refresh
      If frm_espec.data_medicos.Recordset.RecordCount > 0 Then
         frm_espec.cbomed.Text = frm_espec.data_medicos.Recordset("med_nombre")
      Else
         frm_espec.cbomed.Text = ""
      End If
   Else
      frm_espec.cbomed.Text = ""
   End If
   Xcolesp = msf1.Row

'   DBGrid1.Refresh
'   DBGrid1.GetBookmark
'   Xcolesp = DBGrid1.Bookmark
'  MsgBox "ES el:" & Xcolesp
   
Else
   MsgBox "Atención no se encontró", vbCritical, "Mensaje"
End If
Unload Me

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   msf1_DblClick
End If

End Sub
