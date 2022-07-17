VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_buscaatsoc 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar datos de solicitudes"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "frm_buscaatsoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   11400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar solo en PENDIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar solo en TERMINADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Data data1 
      Caption         =   "data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10440
      Picture         =   "frm_buscaatsoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10200
      Picture         =   "frm_buscaatsoc.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cerrar"
      Top             =   4440
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscaatsoc.frx":0F56
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "frm_buscaatsoc.frx":0F6A
      TabIndex        =   5
      Top             =   600
      Width           =   10935
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSMask.MaskEdBox mfh 
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mfd 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H008080FF&
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
      ItemData        =   "frm_buscaatsoc.frx":2195
      Left            =   2400
      List            =   "frm_buscaatsoc.frx":21A2
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "BUSCAR POR:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   7440
      Picture         =   "frm_buscaatsoc.frx":21BE
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "frm_buscaatsoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
   mfd.Visible = True
   mfh.Visible = True
   Text1.Visible = False
   mfd.SetFocus
Else
   If Combo1.ListIndex = 1 Then
      mfd.Visible = False
      mfh.Visible = False
      Text1.Visible = True
      Text1.SetFocus
   Else
      If Combo1.ListIndex = 2 Then
         mfd.Visible = False
         mfh.Visible = False
         Text1.Visible = True
         Text1.SetFocus
      Else
         mfd.Visible = False
         mfh.Visible = False
         Text1.Visible = False
      End If
   End If
End If
      
End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
'Data1.DatabaseName = ""
'Data1.Connect = "ODBC;DSN=sappat;"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

If Combo1.ListIndex = 0 Then
   If mfd.Text = "__/__/____" Then
   Else
      If mfh.Text = "__/__/____" Then
      Else
         If Check1.Value = 1 Then
            Data1.RecordSource = "select top 70, * from ingresosat where at_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and at_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and at_estado in (1,2) order by at_fecha"
            Data1.Refresh
         Else
            If Check2.Value = 1 Then
               Data1.RecordSource = "select top 70, * from ingresosat where at_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and at_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and at_estado =" & 0 & " order by at_fecha"
               Data1.Refresh
            Else
               Data1.RecordSource = "select top 70, * from ingresosat where at_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and at_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# order by at_fecha"
               Data1.Refresh
            End If
         End If
      End If
   End If
Else
   If Combo1.ListIndex = 1 Then
      If Text1.Text = "" Then
         Text1.Text = 0
      End If
      If Check1.Value = 1 Then
         Data1.RecordSource = "Select * from ingresosat where at_ced =" & Text1.Text & " and at_estado in(1,2) order by at_ced"
         Data1.Refresh
      Else
         If Check2.Value = 1 Then
            Data1.RecordSource = "Select * from ingresosat where at_ced =" & Text1.Text & " and at_estado =" & 0 & " order by at_ced"
            Data1.Refresh
         Else
            Data1.RecordSource = "Select * from ingresosat where at_ced =" & Text1.Text & " order by at_ced"
            Data1.Refresh
         End If
      End If
   Else
      If Combo1.ListIndex = 2 Then
         If Text1.Text = "" Then
            Text1.Text = "A"
         End If
         If Check1.Value = 1 Then
            Data1.RecordSource = "select top 70, * from ingresosat where at_nomb >='" & Text1.Text & "' and at_estado in(1,2) order by at_nomb"
            Data1.Refresh
         Else
            If Check2.Value = 1 Then
               Data1.RecordSource = "select top 70, * from ingresosat where at_nomb >='" & Text1.Text & "' and at_estado =" & 0 & " order by at_nomb"
               Data1.Refresh
            Else
               Data1.RecordSource = "select top 70, * from ingresosat where at_nomb >='" & Text1.Text & "' order by at_nomb"
               Data1.Refresh
            End If
         End If
      End If
   End If
End If

End Sub



Private Sub DBGrid1_DblClick()

If IsNull(Data1.Recordset("at_cliente")) = False Then
   frm_atsocio.txt_cliente.Text = Data1.Recordset("at_cliente")
Else
   frm_atsocio.txt_cliente.Text = 0
End If
If IsNull(Data1.Recordset("at_nomb")) = False Then
   frm_atsocio.txt_nomb.Text = Data1.Recordset("at_nomb")
Else
   frm_atsocio.txt_nomb.Text = ""
End If
If IsNull(Data1.Recordset("at_codconv")) = False Then
   frm_atsocio.txt_codconv.Text = Data1.Recordset("at_codconv")
Else
   frm_atsocio.txt_codconv.Text = ""
End If
If IsNull(Data1.Recordset("at_nomconv")) = False Then
   frm_atsocio.txt_desconv.Text = Data1.Recordset("at_nomconv")
Else
   frm_atsocio.txt_desconv.Text = ""
End If
If IsNull(Data1.Recordset("at_via")) = False Then
   frm_atsocio.Combo2.ListIndex = Data1.Recordset("at_via")
Else
   frm_atsocio.Combo2.ListIndex = -1
End If

If IsNull(Data1.Recordset("at_ced")) = False Then
   frm_atsocio.txt_ced.Text = Data1.Recordset("at_ced")
Else
   frm_atsocio.txt_ced.Text = 0
End If
If IsNull(Data1.Recordset("at_codced")) = False Then
   frm_atsocio.txt_codced.Text = Data1.Recordset("at_codced")
Else
   frm_atsocio.txt_codced.Text = 0
End If
If IsNull(Data1.Recordset("at_ing")) = False Then
   frm_atsocio.ming.Text = Format(Data1.Recordset("at_ing"), "dd/mm/yyyy")
Else
   frm_atsocio.ming.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("at_fecrec")) = False Then
   frm_atsocio.mfecrec.Text = Format(Data1.Recordset("at_fecrec"), "dd/mm/yyyy")
Else
   frm_atsocio.mfecrec.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("at_nac")) = False Then
   frm_atsocio.mnac.Text = Format(Data1.Recordset("at_nac"), "dd/mm/yyyy")
Else
   frm_atsocio.mnac.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("at_telef")) = False Then
   frm_atsocio.txt_telef.Text = Data1.Recordset("at_telef")
Else
   frm_atsocio.txt_telef.Text = ""
End If
If IsNull(Data1.Recordset("at_nro")) = False Then
   frm_atsocio.txt_nro.Text = Data1.Recordset("at_nro")
Else
   frm_atsocio.txt_nro.Text = 9999999
End If
If IsNull(Data1.Recordset("at_fecha")) = False Then
   frm_atsocio.mfecha.Text = Data1.Recordset("at_fecha")
Else
   frm_atsocio.mfecha.Text = Date
End If
If IsNull(Data1.Recordset("at_hora")) = False Then
   frm_atsocio.mhora.Text = Data1.Recordset("at_hora")
Else
   frm_atsocio.mhora.Text = Format(Time, "HH:mm:ss")
End If
If IsNull(Data1.Recordset("at_usuario")) = False Then
   frm_atsocio.txt_usua.Text = Data1.Recordset("at_usuario")
Else
   frm_atsocio.txt_usua.Text = "YO"
End If
If IsNull(Data1.Recordset("at_categ")) = False Then
   frm_atsocio.cbodet.ListIndex = Data1.Recordset("at_categ")
Else
   frm_atsocio.cbodet.ListIndex = 0
End If
If IsNull(Data1.Recordset("at_detal")) = False Then
   frm_atsocio.txt_det.Text = Data1.Recordset("at_detal")
Else
   frm_atsocio.txt_det.Text = ""
End If
If IsNull(Data1.Recordset("at_estado")) = False Then
   frm_atsocio.cboest.ListIndex = Data1.Recordset("at_estado")
Else
   frm_atsocio.cboest.ListIndex = -1
End If
If IsNull(Data1.Recordset("at_fecfin")) = False Then
   frm_atsocio.mfecfin.Text = Data1.Recordset("at_fecfin")
Else
   frm_atsocio.mfecfin.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("at_horfin")) = False Then
   frm_atsocio.mhorfin.Text = Data1.Recordset("at_horfin")
Else
   frm_atsocio.mhorfin.Text = "__:__:__"
End If
If IsNull(Data1.Recordset("at_usufin")) = False Then
   frm_atsocio.txt_usufin.Text = Data1.Recordset("at_usufin")
Else
   frm_atsocio.txt_usufin.Text = ""
End If
If IsNull(Data1.Recordset("at_confor")) = False Then
   frm_atsocio.cboconf.ListIndex = Data1.Recordset("at_confor")
Else
   frm_atsocio.cboconf.ListIndex = -1
End If

If IsNull(Data1.Recordset("at_motiind")) = False Then
   frm_atsocio.Combo1.ListIndex = Data1.Recordset("at_motiind")
Else
   frm_atsocio.Combo1.ListIndex = -1
End If

Unload Me


End Sub

Private Sub Form_Load()
Dim Xlaf As Date
Xlaf = Date - 120
'Data1.DatabaseName = ""
'Data1.Connect "ODBC;DSN=sappat;"
'data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from ingresosat where at_fecha >=#" & Format(Xlaf, "yyyy/mm/dd") & "# order by at_fecha DESC"
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))

End Sub
