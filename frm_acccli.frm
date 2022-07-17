VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_acccli 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de acciones ante Clientes"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_acccli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   8160
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de solicitud"
      Enabled         =   0   'False
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9255
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_acccli.frx":0442
         Left            =   1800
         List            =   "frm_acccli.frx":044F
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox t_conv 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   1920
         Width           =   5655
      End
      Begin VB.TextBox t_nom 
         Height          =   360
         Left            =   1800
         TabIndex        =   13
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox t_codced 
         Height          =   375
         Left            =   8520
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6960
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox t_mat 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1800
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_acccli.frx":0479
         Left            =   1800
         List            =   "frm_acccli.frx":048C
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txt_det 
         Height          =   1200
         Left            =   1800
         MaxLength       =   130
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3120
         Width           =   6015
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   8040
         TabIndex        =   9
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   5280
         TabIndex        =   15
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Conformidad del cliente:"
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "Convenio:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Documento:"
         Height          =   255
         Left            =   5280
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Matrícula:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Nro."
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Hora:"
         Height          =   255
         Left            =   6960
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Acción:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Más detalles:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   1575
      End
   End
   Begin VB.Data data_numera 
      Caption         =   "data_numera"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_acccli.frx":04BF
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frm_acccli.frx":04D3
      TabIndex        =   7
      Top             =   6480
      Width           =   9255
   End
   Begin VB.Data data_reg 
      Caption         =   "data_reg"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   6600
      Picture         =   "frm_acccli.frx":13A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Solicitudes ingresadas en un rango de fecha..."
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton b_bus 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   5520
      Picture         =   "frm_acccli.frx":17E4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar..."
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton b_eli 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   4440
      Picture         =   "frm_acccli.frx":1C26
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar registro seleccionado"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   735
      Left            =   3360
      Picture         =   "frm_acccli.frx":2068
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar..."
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton b_gra 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   735
      Left            =   2280
      Picture         =   "frm_acccli.frx":24AA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Grabar datos"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   1200
      Picture         =   "frm_acccli.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Modificar datos"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton b_nue 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   120
      Picture         =   "frm_acccli.frx":2D2E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo registro"
      Top             =   5640
      Width           =   855
   End
End
Attribute VB_Name = "frm_acccli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_bus_Click()
If XWeltipoU = "ADMINISTRADOR" Then
    Dim Xingfec As String
    Xingfec = InputBox("Ingrese a partir de que fecha:")
    If Xingfec <> "" Then
       Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 2 & " And cl_fnac >=#" & Format(Xingfec, "yyyy/mm/dd") & "# order by cl_fnac"
       Data1.Refresh
       DBGrid1.SetFocus
    Else
       Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 2 & " order by cl_fnac"
       Data1.Refresh
       DBGrid1.SetFocus
    End If
Else
    MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_can_Click()
If XAlta = 1 Then
   data_reg.Recordset.CancelUpdate
   XAlta = 0
   Frame1.Enabled = False
   habboton
Else
   XAlta = 0
   Frame1.Enabled = False
   habboton
End If
borracua
igualadat
   
End Sub

Private Sub b_eli_Click()
Dim Xqueborra As String
If XWeltipoU = "ADMINISTRADOR" Then
    Xqueborra = MsgBox("Desea borrar el registro seleccionado?", vbInformation + vbYesNo, "SAPP")
    If Xqueborra = vbYes Then
       data_reg.Recordset.FindFirst "cl_codigo =" & txt_nro.Text
       If Not data_reg.Recordset.NoMatch Then
          data_reg.Recordset.Delete
          data_reg.Refresh
          borracua
          igualadat
       End If
    End If
Else
    MsgBox "Usuario no habilitado"
End If

End Sub

Private Sub b_gra_Click()
If XAlta = 1 Then
   If mfd.Text <> "__/__/____" And mhd.Text <> "__:__" And t_nom.Text <> "" Then
      data_reg.Recordset("cl_codigo") = txt_nro.Text
      data_reg.Recordset("cl_fnac") = mfd.Text
      data_reg.Recordset("cl_ruc") = mhd.Text
      If t_mat.Text <> "" Then
      Else
         t_mat.Text = 0
      End If
      If t_ced.Text = "" Then
         t_ced.Text = 0
      End If
      If t_codced.Text = "" Then
         t_codced.Text = 0
      End If
      If t_conv.Text = "" Then
         t_conv.Text = "SC"
      End If
      data_reg.Recordset("cl_schqmn") = t_mat.Text
      data_reg.Recordset("cl_cantdia") = t_ced.Text
      data_reg.Recordset("cl_etiquet") = t_codced.Text
      data_reg.Recordset("cl_nomvend") = Mid(t_nom.Text, 1, 45)
      data_reg.Recordset("cl_descpag") = Mid(t_conv.Text, 1, 45)
      data_reg.Recordset("val1") = Combo1.ListIndex
      data_reg.Recordset("val2") = Combo2.ListIndex
      data_reg.Recordset("info_debit") = txt_det.Text
      data_reg.Recordset("cl_tipocli") = 2
      data_reg.Recordset("cl_tipoclin") = WElusuario
      data_reg.Recordset.Update
      data_reg.Refresh
      Data1.Refresh
      data_reg.Recordset.MoveLast
      XAlta = 0
      Frame1.Enabled = False
      borracua
      habboton
      igualadat
   Else
      MsgBox "El registro no se grabó porque falta fecha y hora"
   End If
Else
   data_reg.Recordset.Edit
'   data_reg.Recordset("cl_codigo") = txt_nro.Text
'   data_reg.Recordset("cl_fnac") = mfd.Text
'   data_reg.Recordset("cl_ruc") = mhd.Text
    If t_mat.Text <> "" Then
    Else
       t_mat.Text = 0
    End If
    If t_ced.Text = "" Then
       t_ced.Text = 0
    End If
    If t_codced.Text = "" Then
       t_codced.Text = 0
    End If
    If t_conv.Text = "" Then
       t_conv.Text = "SC"
    End If
    data_reg.Recordset("cl_schqmn") = t_mat.Text
    data_reg.Recordset("cl_cantdia") = t_ced.Text
    data_reg.Recordset("cl_etiquet") = t_codced.Text
    data_reg.Recordset("cl_nomvend") = Mid(t_nom.Text, 1, 45)
    data_reg.Recordset("cl_descpag") = Mid(t_conv.Text, 1, 45)
    data_reg.Recordset("val1") = Combo1.ListIndex
    data_reg.Recordset("val2") = Combo2.ListIndex
    data_reg.Recordset("info_debit") = txt_det.Text
    data_reg.Recordset("cl_tipocli") = 2
   data_reg.Recordset.Update
   Data1.Refresh
   XAlta = 0
   Frame1.Enabled = False
   habboton
   borracua
   igualadat

End If

End Sub

Private Sub b_imp_Click()
''   frm_infreg.Show vbModal
Dim Xxd, Xxh As String
Xxd = InputBox("Ingrese DESDE QUE FECHA:")
Xxh = InputBox("Ingrese HASTA QUE FECHA:")
If Xxd <> "" And Xxh <> "" Then
   data_inf.RecordSource = "infcli"
   data_inf.Refresh
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         data_inf.Recordset.Delete
         data_inf.Recordset.MoveNext
      Loop
   End If
   data_reg.RecordSource = "Select * from solaudito where cl_tipocli =" & 2 & " And cl_fnac >=#" & Format(CDate(Xxd), "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(CDate(Xxh), "yyyy/mm/dd") & "#"
   data_reg.Refresh
   If data_reg.Recordset.RecordCount > 0 Then
      data_reg.Recordset.MoveFirst
      Do While Not data_reg.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_fnac") = data_reg.Recordset("cl_fnac")
         data_inf.Recordset("cl_ruc") = data_reg.Recordset("cl_ruc")
         data_inf.Recordset("info_debit") = data_reg.Recordset("info_debit")
         data_inf.Recordset("cl_descpag") = Mid(data_reg.Recordset("cl_descpag"), 1, 25)
         data_inf.Recordset("cl_nrovend") = data_reg.Recordset("cl_schqmn")
         data_inf.Recordset("cl_nomvend") = Mid(data_reg.Recordset("cl_nomvend"), 1, 25)
         data_inf.Recordset("cl_cedula") = data_reg.Recordset("cl_cantdia")
         data_inf.Recordset("cl_codced") = data_reg.Recordset("cl_etiquet")
         data_inf.Recordset("cl_nom_sup") = Mid(data_reg.Recordset("cl_tipoclin"), 1, 25)
         data_inf.Recordset("cl_nro_sup") = data_reg.Recordset("val1")
         If IsNull(data_reg.Recordset("val1")) = False Then
            If data_reg.Recordset("val1") = 0 Then
               data_inf.Recordset("cl_zona") = "SEGUIMIENTO"
            Else
               If data_reg.Recordset("val1") = 1 Then
                  data_inf.Recordset("cl_zona") = "PROMOCION"
               Else
                  If data_reg.Recordset("val1") = 2 Then
                     data_inf.Recordset("cl_zona") = "OBSEQUIO"
                  Else
                     If data_reg.Recordset("val1") = 3 Then
                        data_inf.Recordset("cl_zona") = "CARTA"
                     Else
                        data_inf.Recordset("cl_zona") = "OTRO"
                     End If
                  End If
               End If
            End If
         End If
         data_inf.Recordset("cl_nrocobr") = data_reg.Recordset("val2")
         If IsNull(data_reg.Recordset("val2")) = False Then
            If data_reg.Recordset("val2") = 0 Then
               data_inf.Recordset("cl_nomcobr") = "CONFORME"
            Else
               If data_reg.Recordset("val2") = 1 Then
                  data_inf.Recordset("cl_nomcobr") = "NO CONFORME"
               Else
                  data_inf.Recordset("cl_nomcobr") = "SIN REGISTRAR"
               End If
            End If
         End If
         data_inf.Recordset.Update
         data_reg.Recordset.MoveNext
      Loop
      data_inf.RecordSource = "Select * from infcli"
      data_inf.Refresh
      cr1.ReportFileName = App.Path & "\infregacc.rpt"
      cr1.ReportTitle = "REGISTRO DE ACCIONES ANTE CLIENTES DESDE: " & mfdd.Text & " HASTA: " & mfhh.Text
      cr1.Action = 1
         
   End If
Else
   MsgBox "No ingresó Fecha"
End If

End Sub

Private Sub b_mod_Click()
XAlta = 0
If txt_usua.Text = WElusuario Or WElusuario = "JFERNAN" Then
    Frame1.Enabled = True
    mfd.SetFocus
    If XWeltipoU = "ADMINISTRADOR" Then
       b_imp.Enabled = True
       mfh.Enabled = True
       mhh.Enabled = True
       txt_obs.Enabled = True
    Else
       b_imp.Enabled = False
       mfh.Enabled = False
       mhh.Enabled = False
       txt_obs.Enabled = False
    End If
    If mfh.Text = "__/__/____" Then
       Check4.Enabled = False
       mfconf.Enabled = False
       Combo2.Enabled = False
    Else
       Check4.Enabled = True
       mfconf.Enabled = True
       Combo2.Enabled = True
    End If
    
    desboton
Else
   MsgBox "No es el usuario creador de la tarea", vbCritical
End If

End Sub

Private Sub b_nue_Click()
XAlta = 1
borracua
txt_nro.Text = ""
txt_nro.Text = data_numera.Recordset("exento") + 1
data_reg.Recordset.AddNew
Frame1.Enabled = True
'mfd.SetFocus
mfd.Text = Format(Date, "dd/mm/yyyy")
mhd.Text = Format(Time, "HH:mm")
t_mat.SetFocus

data_numera.Recordset.Edit
data_numera.Recordset("exento") = txt_nro.Text
data_numera.Recordset.Update
b_imp.Enabled = True
desboton

End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_det.SetFocus
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Data2_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBGrid1_DblClick()
data_reg.Recordset.FindFirst "cl_codigo =" & Data1.Recordset("cl_codigo")
If Not data_reg.Recordset.NoMatch Then
   borracua
   igualadat
End If

End Sub

Private Sub Form_Load()
data_reg.DatabaseName = App.Path & "\sapp.mdb"
If WElusuario = "JFERNAN" Or WElusuario = "BDD" Or WElusuario = "SPEREZ" Then
   data_reg.RecordSource = "Select * from solaudito where cl_tipocli =" & 2 & " order by cl_codigo DESC"
Else
   data_reg.RecordSource = "Select * from solaudito where cl_tipocli =" & 2 & " order by cl_codigo DESC"
End If
data_reg.Refresh
igualadat
Data1.DatabaseName = App.Path & "\sapp.mdb"
If WElusuario = "JFERNAN" Or WElusuario = "BDD" Or WElusuario = "SPEREZ" Then
   Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 2 & " order by cl_fnac DESC"
Else
   Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 2 & " order by cl_fnac DESC"
End If
Data1.Refresh
data_numera.DatabaseName = App.Path & "\parse.mdb"
data_numera.RecordSource = "parsec0"
data_numera.Refresh
data_inf.DatabaseName = App.Path & "\informes.mdb"

End Sub

Public Function desboton()
b_nue.Enabled = False
b_mod.Enabled = False
b_gra.Enabled = True
b_can.Enabled = True
b_eli.Enabled = False
b_bus.Enabled = False
b_imp.Enabled = False
DBGrid1.Enabled = False

End Function

Public Function habboton()
b_nue.Enabled = True
b_mod.Enabled = True
b_gra.Enabled = False
b_can.Enabled = False
b_eli.Enabled = True
b_bus.Enabled = True
b_imp.Enabled = True
DBGrid1.Enabled = True

End Function

Private Sub mfconf_GotFocus()
mfconf.Text = Format(Date, "dd/mm/yyyy")
mhconf.Text = Format(Time, "HH:mm")


End Sub

Private Sub mfconf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhconf.SetFocus
End If

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhd.SetFocus
End If

End Sub

Private Sub mhconf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub mhd_GotFocus()
mhd.Text = Format(Time, "HH:mm")

End Sub

Private Sub mhd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub txt_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_usua.SetFocus
End If

End Sub


Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codced.SetFocus
End If

End Sub

Private Sub t_codced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_conv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_ced.SetFocus
End If

End Sub

Private Sub t_mat_LostFocus()
If t_mat.Text <> "" Then
   If t_mat.Text <> 0 Then
      data_cli.DatabaseName = App.Path & "\sapp.mdb"
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         t_ced.Text = data_cli.Recordset("cl_cedula")
         t_codced.Text = data_cli.Recordset("cl_codced")
         t_nom.Text = data_cli.Recordset("cl_apellid")
         t_conv.Text = data_cli.Recordset("cl_nomconv")
      End If
    End If
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_conv.SetFocus
End If

End Sub

Private Sub txt_det_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Public Function igualadat()

If data_reg.Recordset.RecordCount > 0 Then
   txt_nro.Text = data_reg.Recordset("cl_codigo")
   If IsNull(data_reg.Recordset("cl_fnac")) = False Then
      If data_reg.Recordset("cl_fnac") <> "" Then
         mfd.Text = Format(data_reg.Recordset("cl_fnac"), "dd/mm/yyyy")
      Else
         mfd.Text = "__/__/____"
      End If
   Else
      mfd.Text = "__/__/____"
   End If
   If IsNull(data_reg.Recordset("cl_ruc")) = False Then
      mhd.Text = data_reg.Recordset("cl_ruc")
   Else
      mhd.Text = "__:__"
   End If
   If IsNull(data_reg.Recordset("cl_schqmn")) = False Then
      t_mat.Text = data_reg.Recordset("cl_schqmn")
   Else
      t_mat.Text = 0
   End If
   If IsNull(data_reg.Recordset("cl_cantdia")) = False Then
      t_ced.Text = data_reg.Recordset("cl_cantdia")
   Else
      t_ced.Text = 0
   End If
   If IsNull(data_reg.Recordset("cl_nomvend")) = False Then
      t_nom.Text = data_reg.Recordset("cl_nomvend")
   Else
      t_nom.Text = ""
   End If
   If IsNull(data_reg.Recordset("cl_descpag")) = False Then
      t_conv.Text = data_reg.Recordset("cl_descpag")
   Else
      t_conv.Text = ""
   End If
   If IsNull(data_reg.Recordset("val1")) = False Then
      Combo1.ListIndex = data_reg.Recordset("val1")
   Else
      Combo1.ListIndex = -1
   End If
   If IsNull(data_reg.Recordset("info_debit")) = False Then
      txt_det.Text = data_reg.Recordset("info_debit")
   Else
      txt_det.Text = ""
   End If
   If IsNull(data_reg.Recordset("val2")) = False Then
      Combo2.ListIndex = data_reg.Recordset("val2")
   Else
      Combo2.ListIndex = -1
   End If
   
End If
        
End Function

Public Function borracua()
txt_nro.Text = ""
mfd.Text = "__/__/____"
mhd.Text = "__:__"
Combo1.ListIndex = -1
txt_det.Text = ""
Combo2.ListIndex = -1
t_mat.Text = ""
t_ced.Text = ""
t_codced.Text = ""
t_nom.Text = ""
t_conv.Text = ""

End Function
