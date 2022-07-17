VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_abmm 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Modificaciones"
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   7710
   Icon            =   "frm_abmm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_afilcons 
      Caption         =   "data_afilcons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_deu 
      Caption         =   "data_deu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
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
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_guardam 
      Caption         =   "data_guardam"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton btn_ace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
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
      Picture         =   "frm_abmm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Data data_abmmm 
      Caption         =   "data_abmmm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM motivos WHERE MC_NUMERO>=""C00"" AND MC_NUMERO<=""C50"""
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBCtls.DBCombo cbomod 
      Bindings        =   "frm_abmm.frx":09CC
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "MC_DESC"
      Text            =   ""
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TIPO DE MODIFICACION:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2280
      Picture         =   "frm_abmm.frx":09E5
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frm_abmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_ace_Click()
Dim Xelmensajecob As String
Dim Xnohaymas As Integer
Xnohaymas = 0

frmabm.data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & frmabm.txt_mat.Caption
frmabm.data_abm.Refresh

frmabm.data_abm.Recordset.AddNew
frmabm.data_abm.Recordset("cl_codigo") = frmabm.txt_mat.Caption
frmabm.data_abm.Recordset("cl_motivo") = cbomod.Text
frmabm.data_abm.Recordset("desc") = "MODIF"
frmabm.data_abm.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
frmabm.data_abm.Recordset("hora") = Format(Time, "HH:mm")
frmabm.data_abm.Recordset("usuario") = WElusuario
frmabm.data_abm.Recordset("convenio") = frmabm.txt_codcnv.Text
frmabm.data_abm.Recordset("base") = frmabm.data_parsec.Recordset("base")
frmabm.data_abm.Recordset.Update
If cbomod.Text = "DATOS AFILIACION" Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where matricula =" & Val(frmabm.txt_mat.Caption) & " and pendiente in (0)"
   data_afilcons.Refresh
   If data_afilcons.Recordset.RecordCount > 0 Then
      If IsNull(data_afilcons.Recordset("procesa_mod")) = True Then
         data_afilcons.Recordset.Edit
         data_afilcons.Recordset("procesa_mod") = 11
         data_afilcons.Recordset.Update
      End If
   Else
      MsgBox "No se encontró afiliación pendiente. Verifique!", vbInformation
   End If

End If
   
If cbomod.Text = "CAMBIO DE COBRADOR" Then
   Xelmensajecob = MsgBox("Desea cambiar las facturas pendientes al nuevo cobrador?", vbInformation + vbYesNo)
   If Xelmensajecob = vbYes Then
      data_deu.RecordSource = "Select * from deudas where cliente =" & frmabm.txt_mat.Caption & " and fecha_pago is null and mes >" & 0
      data_deu.Refresh
      If data_deu.Recordset.RecordCount > 0 Then
         data_deu.Recordset.MoveFirst
         Do While Not data_deu.Recordset.EOF
            If IsNull(data_deu.Recordset("nro_cobr")) = False Then
               If data_deu.Recordset("nro_cobr") <> frmabm.txt_codcob.Text Then
                  data_deu.Recordset.Edit
                  data_deu.Recordset("nro_cobr") = frmabm.txt_codcob.Text
                  data_deu.Recordset("nom_cobr") = Mid(frmabm.cbonomcob.Text, 1, 20)
                  data_deu.Recordset.Update
               End If
                data_arq.RecordSource = "Select * from arqueo where matricula =" & Val(frmabm.txt_mat.Caption) & " and nrorec =" & data_deu.Recordset("documento")
                data_arq.Refresh
                If data_arq.Recordset.RecordCount > 0 Then
                   data_arq.Recordset.MoveFirst
                   Do While Not data_arq.Recordset.EOF
                      If data_arq.Recordset("cob") <> Val(frmabm.txt_codcob.Text) Then
                         data_arq.Recordset.Edit
                         data_arq.Recordset("cob") = Val(frmabm.txt_codcob.Text)
                         data_arq.Recordset("nomcob") = Mid(frmabm.cbonomcob.Text, 1, 35)
                         data_arq.Recordset("arqueo") = "E"
                         data_arq.Recordset.Update
                      End If
                      data_arq.Recordset.MoveNext
                   Loop
                Else
                   MsgBox "ATENCION!! El recibo número: " & Trim(str(data_deu.Recordset("documento"))) & " NO se encuentra en arqueo. VERIFIQUE!!", vbCritical
'                      data_arq.Recordset.AddNew
'                      data_arq.Recordset("matricula") = data_deu.Recordset("cliente")
'                      data_arq.Recordset("nombre") = Mid(data_deu.Recordset("nombre"), 1, 40)
'                      data_arq.Recordset("mes") = data_deu.Recordset("mes")
'                      data_arq.Recordset("ano") = data_deu.Recordset("ano")
'                      data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_deu.Recordset("cod_cnv") & "'"
'                      data_conv.Refresh
'                      If data_conv.Recordset.RecordCount > 0 Then
'                         data_arq.Recordset("color") = data_conv.Recordset("cnv_colrec")
'                      Else
'                         data_arq.Recordset("color") = "M"
'                      End If
'                      data_arq.Recordset("cat") = data_deu.Recordset("cod_cnv")
'                      data_arq.Recordset("nomcat") = data_deu.Recordset("nom_cnv")
'                      data_arq.Recordset("arqueo") = "E"
'                      data_arq.Recordset("importe") = data_deu.Recordset("importe")
'                      data_arq.Recordset("fecha") = Date
'                      data_arq.Recordset("nrorec") = data_deu.Recordset("documento")
'                      data_arq.Recordset("usuar") = WElusuario
'                      data_arq.Recordset("moneda") = data_deu.Recordset("moneda")
'                      data_arq.Recordset("cob") = Val(frmabm.txt_codcob.Text)
'                      data_arq.Recordset("nomcob") = Mid(frmabm.cbonomcob.Text, 1, 35)
'                      data_arq.Recordset("codzon") = data_deu.Recordset("grupo")
'                      data_arq.Recordset("codpro") = data_deu.Recordset("nro_vende")
'                      data_arq.Recordset("codsup") = 1
'                      data_arq.Recordset("tiquet") = data_deu.Recordset("tiquet")
''                      data_arq.Recordset("total") = data_deu.Recordset("total")
'                      data_arq.Recordset("varia") = 0
'                      data_arq.Recordset("iva") = data_deu.Recordset("iva")
'                      data_arq.Recordset("deudas") = data_deu.Recordset("deudas")
'                      data_arq.Recordset("servi") = data_deu.Recordset("servi")
'                      data_arq.Recordset.Update
                End If
            
            End If
            data_deu.Recordset.MoveNext
         Loop
      Else
         data_arq.RecordSource = "Select * from arqueo where matricula =" & Val(frmabm.txt_mat.Caption)
         data_arq.Refresh
         If data_arq.Recordset.RecordCount > 0 Then
            data_arq.Recordset.MoveFirst
            Do While Not data_arq.Recordset.EOF
               If data_arq.Recordset("cob") <> Val(frmabm.txt_codcob.Text) Then
                  data_arq.Recordset.Edit
                  data_arq.Recordset("cob") = Val(frmabm.txt_codcob.Text)
                  data_arq.Recordset("nomcob") = Mid(frmabm.cbonomcob.Text, 1, 35)
                  data_arq.Recordset("arqueo") = "E"
                  data_arq.Recordset.Update
               End If
               data_arq.Recordset.MoveNext
            Loop
         End If
      End If
   End If
End If
Unload Me


End Sub

Private Sub cbomod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btn_ace.SetFocus
End If

End Sub

Private Sub Form_Initialize()
data_abmmm.Recordset.MoveFirst
cbomod.Text = data_abmmm.Recordset("mc_desc")

End Sub

Private Sub Form_Load()
data_abmmm.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_afilcons.Connect = "odbc;dsn=" & Xconexrmt & ";"

'data_abmmm.RecordSource = "MOTIVOS"
data_abmmm.Refresh
data_guardam.Connect = "odbc;dsn=" & Xconexrmt & ";"
'SelectLimit 10
'data_guardam.RecordSource = "select top 10, * from abmsocio"
'data_guardam.Refresh
'SelectLimit 0
data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_abmm.Hide

End Sub
