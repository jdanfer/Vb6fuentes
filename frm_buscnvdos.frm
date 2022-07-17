VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_buscnvdos 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Buscar..."
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9900
   Icon            =   "frm_buscnvdos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C00000&
      Caption         =   "RAZON SOCIAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FF0000&
      Caption         =   "Buscar por Razón social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   840
      Width           =   3615
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF0000&
      Caption         =   "Incluir convenios ocultos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   4680
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   7440
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
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
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5741
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "Incluir convenios baja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton bcerra 
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
      Left            =   9120
      Picture         =   "frm_buscnvdos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txt_desc 
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5415
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "DESCRIPCION"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "CODIGO..."
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
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Para seleccionar presione ENTER"
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
      TabIndex        =   5
      Top             =   4440
      Width           =   9015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   9840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BUSCAR POR..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   6720
      Picture         =   "frm_buscnvdos.frx":09CC
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "frm_buscnvdos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcerra_Click()
frm_buscnvdos.Hide

End Sub

Private Sub Check2_Click()
If WElusuario = "COMPUTOS" Or XWeltipoU = "ADMINISTRADOR" Then
Else
   MsgBox "Usuario no autorizado a cambiar esta opción"
   Check2.Value = 0
End If

End Sub

Private Sub DBGrid1_DblClick()
   frm_convenios.data_conv.Recordset.FindFirst "cnv_codigo ='" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0) & "'"
   If Not frm_convenios.data_conv.Recordset.NoMatch Then
        If IsNull(frm_convenios.data_conv.Recordset("cnv_codigo")) = False Then
           frm_convenios.txt_cod.Text = frm_convenios.data_conv.Recordset("cnv_codigo")
        Else
           frm_convenios.txt_cod.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_desc")) = False Then
           frm_convenios.txt_desc.Text = frm_convenios.data_conv.Recordset("cnv_desc")
        Else
           frm_convenios.txt_desc.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_entre")) = False Then
           frm_convenios.t_razon.Text = frm_convenios.data_conv.Recordset("cnv_entre")
        Else
           frm_convenios.t_razon.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_correoe")) = False Then
           frm_convenios.t_email.Text = frm_convenios.data_conv.Recordset("cnv_correoe")
        Else
           frm_convenios.t_email.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_sindeuda")) = False Then
           frm_convenios.chdeuda.Value = frm_convenios.data_conv.Recordset("cnv_sindeuda")
        Else
           frm_convenios.chdeuda.Value = 0
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_gpoafilia")) = False Then
           frm_convenios.chafilia.Value = frm_convenios.data_conv.Recordset("cnv_gpoafilia")
        Else
           frm_convenios.chafilia.Value = 0
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_pmserv")) = False Then
           frm_convenios.cbofact.ListIndex = frm_convenios.data_conv.Recordset("cnv_pmserv")
        Else
           frm_convenios.cbofact.ListIndex = -1
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_umpago")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_umpago") > 2 Then
              frm_convenios.choculta.Value = 0
           Else
              frm_convenios.choculta.Value = frm_convenios.data_conv.Recordset("cnv_umpago")
           End If
        Else
           frm_convenios.choculta.Value = 0
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_uapago")) = False Then
           frm_convenios.t_rub.Text = frm_convenios.data_conv.Recordset("cnv_uapago")
        Else
           frm_convenios.t_rub.Text = ""
        End If
        
        If IsNull(frm_convenios.data_conv.Recordset("cnv_motbaj")) = False Then
           frm_convenios.t_der.Text = frm_convenios.data_conv.Recordset("cnv_motbaj")
        Else
           frm_convenios.t_der.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_email")) = False Then
           frm_convenios.t_nrocompra.Text = frm_convenios.data_conv.Recordset("cnv_email")
        Else
           frm_convenios.t_nrocompra.Text = ""
        End If
        
        If IsNull(frm_convenios.data_conv.Recordset("cnv_paserv")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_paserv") = 0 Then
              frm_convenios.cbovenc.ListIndex = 0
           Else
              If frm_convenios.data_conv.Recordset("cnv_paserv") = 15 Then
                 frm_convenios.cbovenc.ListIndex = 1
              Else
                 If frm_convenios.data_conv.Recordset("cnv_paserv") = 30 Then
                    frm_convenios.cbovenc.ListIndex = 2
                 Else
                    If frm_convenios.data_conv.Recordset("cnv_paserv") = 60 Then
                       frm_convenios.cbovenc.ListIndex = 3
                    Else
                       If frm_convenios.data_conv.Recordset("cnv_paserv") = 90 Then
                          frm_convenios.cbovenc.ListIndex = 4
                       Else
                          If frm_convenios.data_conv.Recordset("cnv_paserv") = 120 Then
                             frm_convenios.cbovenc.ListIndex = 5
                          Else
                             frm_convenios.cbovenc.ListIndex = -1
                          End If
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_convenios.cbovenc.ListIndex = -1
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_ruc")) = False Then
           frm_convenios.txt_ruc.Text = frm_convenios.data_conv.Recordset("cnv_ruc")
        Else
           frm_convenios.txt_ruc.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_direcc")) = False Then
           If IsNull(frm_convenios.data_conv.Recordset("cnv_entre")) = False Then
              frm_convenios.txt_direc.Text = frm_convenios.data_conv.Recordset("cnv_direcc")
           Else
              frm_convenios.txt_direc.Text = frm_convenios.data_conv.Recordset("cnv_direcc")
           End If
        Else
           frm_convenios.txt_direc.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_local")) = False Then
           frm_convenios.txt_localid.Text = frm_convenios.data_conv.Recordset("cnv_local")
        Else
           frm_convenios.txt_localid.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_tel")) = False Then
           frm_convenios.txt_tel.Text = frm_convenios.data_conv.Recordset("cnv_tel")
        Else
           frm_convenios.txt_tel.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_desde")) = False Then
           frm_convenios.vdesde.Text = Format(frm_convenios.data_conv.Recordset("cnv_desde"), "dd/mm/yyyy")
        Else
           frm_convenios.vdesde.Text = "__/__/____"
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_hasta")) = False Then
           frm_convenios.vhasta.Text = Format(frm_convenios.data_conv.Recordset("cnv_hasta"), "dd/mm/yyyy")
        Else
           frm_convenios.vhasta.Text = "__/__/____"
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_codmon")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_codmon") = 1 Then
              frm_convenios.cbomon.ListIndex = 0
           Else
              If frm_convenios.data_conv.Recordset("cnv_codmon") = 2 Then
                 frm_convenios.cbomon.ListIndex = 1
              Else
                 frm_convenios.cbomon.Text = ""
              End If
           End If
        Else
           frm_convenios.cbomon.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_colrec")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_colrec") = "R" Then
              frm_convenios.cbocolrec.ListIndex = 0
           Else
              If frm_convenios.data_conv.Recordset("cnv_colrec") = "A" Then
                 frm_convenios.cbocolrec.ListIndex = 1
              Else
                 If frm_convenios.data_conv.Recordset("cnv_colrec") = "M" Then
                    frm_convenios.cbocolrec.ListIndex = 2
                 Else
                    If frm_convenios.data_conv.Recordset("cnv_colrec") = "V" Then
                       frm_convenios.cbocolrec.ListIndex = 3
                    Else
                       frm_convenios.cbocolrec.Text = ""
                    End If
                 End If
              End If
           End If
        Else
           frm_convenios.cbocolrec.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_precio")) = False Then
           frm_convenios.txt_precio.Text = frm_convenios.data_conv.Recordset("cnv_precio")
        Else
           frm_convenios.txt_precio.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_emite")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_emite") = "SI" Then
              frm_convenios.cbosirec.ListIndex = 1
           Else
              If frm_convenios.data_conv.Recordset("cnv_emite") = "NO" Then
                 frm_convenios.cbosirec.ListIndex = 0
              Else
                 frm_convenios.cbosirec.ListIndex = 0
              End If
           End If
        Else
           frm_convenios.cbosirec.ListIndex = 0
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_cuenta")) = False Then
           frm_convenios.txt_cuenta.Text = frm_convenios.data_conv.Recordset("cnv_cuenta")
        Else
           frm_convenios.txt_cuenta.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_alta")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_alta") = "SI" Then
              frm_convenios.cboaltasi.ListIndex = 0
           Else
              frm_convenios.cboaltasi.ListIndex = 1
           End If
        Else
           frm_convenios.cboaltasi.ListIndex = 1
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_cant_r")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_cant_r") = 1 Then
              frm_convenios.opunosolo.Value = True
           Else
              If frm_convenios.data_conv.Recordset("cnv_cant_r") = 2 Then
                 frm_convenios.optodos.Value = True
              Else
                 frm_convenios.opunosolo.Value = True
              End If
           End If
        Else
           frm_convenios.opunosolo.Value = True
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_grupo")) = False Then
           If frm_convenios.data_conv.Recordset("cnv_grupo") <> "" Then
              frm_convenios.cbomut.Text = frm_convenios.data_conv.Recordset("cnv_grupo")
           Else
              frm_convenios.cbomut.Text = ""
           End If
        Else
           frm_convenios.cbomut.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_fbaja")) = False Then
           frm_convenios.fbaja.Text = Format(frm_convenios.data_conv.Recordset("cnv_fbaja"), "dd/mm/yyyy")
        Else
           frm_convenios.fbaja.Text = "__/__/____"
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_ctrato")) = False Then
           frm_convenios.txt_obs.Text = frm_convenios.data_conv.Recordset("cnv_ctrato")
        Else
           frm_convenios.txt_obs.Text = ""
        End If
      
        If IsNull(frm_convenios.data_conv.Recordset("cnv_cantcons")) = False Then
           frm_convenios.t_cantlla.Text = frm_convenios.data_conv.Recordset("cnv_cantcons")
        Else
           frm_convenios.t_cantlla.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_preccons")) = False Then
           frm_convenios.t_implla.Text = frm_convenios.data_conv.Recordset("cnv_preccons")
        Else
           frm_convenios.t_implla.Text = ""
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_menanio")) = False Then
           frm_convenios.cbomesanio.ListIndex = frm_convenios.data_conv.Recordset("cnv_menanio")
        Else
           frm_convenios.cbomesanio.ListIndex = -1
        End If
        If IsNull(frm_convenios.data_conv.Recordset("cnv_grupoap")) = False Then
           frm_convenios.cbogrupoap.Text = frm_convenios.data_conv.Recordset("cnv_grupoap")
        Else
           frm_convenios.cbogrupoap.Text = ""
        End If
            
        If IsNull(frm_convenios.data_conv.Recordset("cnv_sald")) = False Then
           frm_convenios.chtimbre.Value = frm_convenios.data_conv.Recordset("cnv_sald")
        Else
           frm_convenios.chtimbre.Value = 0
        End If
            
        If IsNull(frm_convenios.data_conv.Recordset("cnv_aran")) = True Then
           frm_convenios.Combo1.ListIndex = -1
           frm_convenios.Combo1.Text = ""
        Else
           If frm_convenios.data_conv.Recordset("cnv_aran") = 0 Then
              frm_convenios.Combo1.ListIndex = -1
              frm_convenios.Combo1.Text = ""
           Else
              frm_convenios.data_aranc.RecordSource = "Select * from Aran_grupos where id =" & frm_convenios.data_conv.Recordset("cnv_aran")
              frm_convenios.data_aranc.Refresh
              If frm_convenios.data_aranc.Recordset.RecordCount > 0 Then
                 frm_convenios.Combo1.Text = frm_convenios.data_aranc.Recordset("desc_gpo")
              Else
                 frm_convenios.Combo1.ListIndex = -1
                 frm_convenios.Combo1.Text = ""
              End If
           End If
        End If
      
      frm_buscnvdos.Hide
   Else
      MsgBox "error en la búsqueda", vbCritical, "Buscar..."
      DBGrid1.SetFocus
   End If

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Load()
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.RecordSource = "convenio"
'Data1.Refresh
Data1.ConnectionString = "dsn=" & Xconexrmt
DBGrid1.rows = 2
DBGrid1.Cols = 4
DBGrid1.TextMatrix(0, 0) = "CODIGO"
DBGrid1.ColWidth(0) = 1900
DBGrid1.TextMatrix(0, 1) = "DESCRIPCION"
DBGrid1.ColWidth(1) = 6900
DBGrid1.TextMatrix(0, 2) = "RAZON SOCIAL"
DBGrid1.ColWidth(2) = 6900
DBGrid1.TextMatrix(0, 3) = "DIRECCION"
DBGrid1.ColWidth(3) = 5900

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_desc.SetFocus
End If

End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_desc.SetFocus
End If

End Sub

Private Sub txt_desc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))
If KeyAscii = 13 Then
    If Option1.Value = True Then
       If Check1.Value = 1 Then
          If Check2.Value = 1 Then
             Data1.RecordSource = "select * from convenio where cnv_codigo >='" & txt_desc.Text & "' order by cnv_codigo"
             Data1.Refresh
          Else
             Data1.RecordSource = "select * from convenio where cnv_codigo >='" & txt_desc.Text & "' and cnv_umpago not in (1) order by cnv_codigo"
             Data1.Refresh
          End If
       Else
          If Check2.Value = 1 Then
             Data1.RecordSource = "select * from convenio where cnv_codigo >='" & txt_desc.Text & "' and cnv_fbaja is null order by cnv_codigo"
             Data1.Refresh
          Else
             Data1.RecordSource = "select * from convenio where cnv_codigo >='" & txt_desc.Text & "' and cnv_fbaja is null and cnv_umpago not in (1) order by cnv_codigo"
             Data1.Refresh
          End If
       End If
    Else
       If Option2.Value = True Then
          If Check1.Value = 1 Then
             If Check2.Value = 1 Then
                Data1.RecordSource = "select * from convenio where cnv_desc >='" & txt_desc.Text & "' order by cnv_desc"
                Data1.Refresh
             Else
                Data1.RecordSource = "select * from convenio where cnv_desc >='" & txt_desc.Text & "' and cnv_umpago not in (1) order by cnv_desc"
                Data1.Refresh
             End If
          Else
             If Check2.Value = 1 Then
                Data1.RecordSource = "select * from convenio where cnv_desc >='" & txt_desc.Text & "' and cnv_fbaja is null order by cnv_desc"
                Data1.Refresh
             Else
                Data1.RecordSource = "select * from convenio where cnv_desc >='" & txt_desc.Text & "' and cnv_fbaja is null and cnv_umpago not in (1) order by cnv_desc"
                Data1.Refresh
             End If
          End If
       Else
          If Option3.Value = True Then
             If Check1.Value = 1 Then
                If Check2.Value = 1 Then
                   Data1.RecordSource = "select * from convenio where cnv_entre >='" & txt_desc.Text & "' order by cnv_entre"
                   Data1.Refresh
                Else
                   Data1.RecordSource = "select * from convenio where cnv_entre >='" & txt_desc.Text & "' and cnv_umpago not in (1) order by cnv_entre"
                   Data1.Refresh
                End If
             Else
                If Check2.Value = 1 Then
                   Data1.RecordSource = "select * from convenio where cnv_entre >='" & txt_desc.Text & "' and cnv_fbaja is null order by cnv_entre"
                   Data1.Refresh
                Else
                   Data1.RecordSource = "select * from convenio where cnv_entre >='" & txt_desc.Text & "' and cnv_fbaja is null and cnv_umpago not in (1) order by cnv_entre"
                   Data1.Refresh
                End If
             End If
          End If
       End If
    End If
    DBGrid1.rows = 2
    DBGrid1.Cols = 4
    DBGrid1.TextMatrix(0, 0) = "CODIGO"
    DBGrid1.ColWidth(0) = 1900
    DBGrid1.TextMatrix(0, 1) = "DESCRIPCION"
    DBGrid1.ColWidth(1) = 6900
    DBGrid1.TextMatrix(0, 2) = "RAZON SOCIAL"
    DBGrid1.ColWidth(2) = 6900
    DBGrid1.TextMatrix(0, 3) = "DIRECCION"
    DBGrid1.ColWidth(3) = 5900
    
    Dim Xcann As Integer
    Xcann = 1
    If Data1.Recordset.RecordCount > 0 Then
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
           If IsNull(Data1.Recordset("cnv_codigo")) = False Then
              DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("cnv_codigo")
           End If
           If IsNull(Data1.Recordset("cnv_desc")) = False Then
              DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("cnv_desc")
           End If
           If IsNull(Data1.Recordset("cnv_entre")) = False Then
              DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("cnv_entre")
           End If
           If IsNull(Data1.Recordset("cnv_direcc")) = False Then
              DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("cnv_direcc")
           End If
           DBGrid1.rows = DBGrid1.rows + 1
           Data1.Recordset.MoveNext
           Xcann = Xcann + 1
        Loop
    End If
    Data1.Recordset.Close
    
   DBGrid1.SetFocus
End If

End Sub
