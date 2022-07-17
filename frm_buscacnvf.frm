VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_buscacnvf 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar datos de convenios"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_buscacnvf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   9000
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscacnvf.frx":0442
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frm_buscacnvf.frx":0456
      TabIndex        =   3
      Top             =   960
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_buscacnvf.frx":118D
      Left            =   2280
      List            =   "frm_buscacnvf.frx":119A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar por:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frm_buscacnvf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text1.SetFocus
End If

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(Data1.Recordset("cl_codigo")) = False Then
   frm_abmconve.data_cli.RecordSource = "Select * from clifact where cl_codigo =" & Data1.Recordset("cl_codigo")
   frm_abmconve.data_cli.Refresh
   If frm_abmconve.data_cli.Recordset.RecordCount > 0 Then
      If IsNull(Data1.Recordset("cl_codigo")) = False Then
         frm_abmconve.t_cod.Text = Data1.Recordset("cl_codigo")
      Else
         frm_abmconve.t_cod.Text = 0
      End If
      If IsNull(Data1.Recordset("cl_apellid")) = False Then
         frm_abmconve.t_nom.Text = Data1.Recordset("cl_apellid")
      Else
         frm_abmconve.t_nom.Text = ""
      End If
      If IsNull(Data1.Recordset("cl_nombre")) = False Then
         frm_abmconve.t_razon.Text = Data1.Recordset("cl_nombre")
      Else
         frm_abmconve.t_razon.Text = ""
      End If
      If IsNull(Data1.Recordset("cl_fecing")) = False Then
         frm_abmconve.mfing.Text = Data1.Recordset("cl_fecing")
      Else
         frm_abmconve.mfing.Text = "__/__/____"
      End If
      If IsNull(Data1.Recordset("cl_ultmesp")) = False Then
         frm_abmconve.labultmes.Caption = Data1.Recordset("cl_ultmesp")
      Else
         frm_abmconve.labultmes.Caption = 0
      End If
      If IsNull(Data1.Recordset("cl_ultanop")) = False Then
         frm_abmconve.labultano.Caption = Data1.Recordset("cl_ultanop")
      Else
         frm_abmconve.labultano.Caption = 0
      End If
      If IsNull(Data1.Recordset("cl_nom_sup")) = False Then
         frm_abmconve.t_rut.Text = Data1.Recordset("cl_nom_sup")
      Else
         frm_abmconve.t_rut.Text = 0
      End If
      If IsNull(Data1.Recordset("cl_direcci")) = False Then
         frm_abmconve.t_dir.Text = Data1.Recordset("cl_direcci")
      Else
         frm_abmconve.t_dir.Text = ""
      End If
      If IsNull(Data1.Recordset("cl_email")) = False Then
         frm_abmconve.t_correo.Text = Data1.Recordset("cl_email")
      Else
         frm_abmconve.t_correo.Text = ""
      End If
      If IsNull(Data1.Recordset("cl_telefon")) = False Then
         frm_abmconve.t_tel.Text = Data1.Recordset("cl_telefon")
      Else
         frm_abmconve.t_tel.Text = ""
      End If
      If IsNull(Data1.Recordset("cl_zona")) = False Then
         frm_abmconve.cbozon.Text = Data1.Recordset("cl_zona")
         frm_abmconve.t_codz.Text = Data1.Recordset("cl_grupo")
      Else
         frm_abmconve.cbozon.ListIndex = -1
         frm_abmconve.t_codz.Text = 0
      End If
      If IsNull(Data1.Recordset("cl_cuopaga")) = False Then
         frm_abmconve.t_imp.Text = Data1.Recordset("cl_cuopaga")
      Else
         frm_abmconve.t_imp.Text = 0
      End If
      If IsNull(Data1.Recordset("derechos")) = False Then
         frm_abmconve.t_der.Text = Data1.Recordset("derechos")
      Else
         frm_abmconve.t_der.Text = ""
      End If
      If IsNull(Data1.Recordset("cl_nrocobr")) = False Then
         frm_abmconve.t_codcob.Text = Data1.Recordset("cl_nrocobr")
         frm_abmconve.cbocob.Text = Data1.Recordset("cl_nomcobr")
      Else
         frm_abmconve.t_codcob.Text = 0
         frm_abmconve.cbocob.ListIndex = -1
      End If
      If IsNull(Data1.Recordset("cl_nrovend")) = False Then
         frm_abmconve.t_codpro.Text = Data1.Recordset("cl_nrovend")
         frm_abmconve.cbopro.Text = Data1.Recordset("cl_nomvend")
      Else
         frm_abmconve.t_codpro.Text = 0
         frm_abmconve.cbopro.ListIndex = -1
      End If
      If IsNull(Data1.Recordset("obsfact")) = False Then
         frm_abmconve.t_datos.Text = Data1.Recordset("obsfact")
      Else
         frm_abmconve.t_datos.Text = ""
      End If
      If IsNull(Data1.Recordset("observa")) = False Then
         frm_abmconve.t_obs.Text = Data1.Recordset("observa")
      Else
         frm_abmconve.t_obs.Text = ""
      End If
   End If
End If
Unload Me



End Sub

Private Sub Form_Load()
frm_buscacnvf.MousePointer = 11
Data1.Connect = "ODBC;DSN=facturacion;"
Data1.RecordSource = "Select * from clifact order by cl_apellid"
Data1.Refresh
frm_buscacnvf.MousePointer = 0

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
frm_buscacnvf.MousePointer = 11
If KeyAscii = 13 Then
   If Text1.Text = "" Then
   Else
      If Combo1.ListIndex = 0 Then
         Data1.RecordSource = "Select * from clifact where cl_apellid >='" & Text1.Text & "' order by cl_apellid"
         Data1.Refresh
      Else
         If Combo1.ListIndex = 1 Then
            Data1.RecordSource = "Select * from clifact where cl_nombre >='" & Text1.Text & "' order by cl_nombre"
            Data1.Refresh
         Else
            If Combo1.ListIndex = 2 Then
               Data1.RecordSource = "Select * from clifact where cl_telefon >='" & Text1.Text & "' order by cl_telefon"
               Data1.Refresh
            Else
               Data1.RecordSource = "Select * from clifact where cl_apellid >='" & Text1.Text & "' order by cl_apellid"
               Data1.Refresh
            End If
         End If
      End If
   End If
   DBGrid1.SetFocus
End If
frm_buscacnvf.MousePointer = 0

End Sub
