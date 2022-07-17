VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_solmejoras 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de solicitudes de Mejora"
   ClientHeight    =   8205
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
   Icon            =   "frm_regsolmejoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver únicamente solicitudes sin terminar"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6000
      Width           =   5175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_regsolmejoras.frx":0442
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frm_regsolmejoras.frx":0456
      TabIndex        =   24
      Top             =   6240
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
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_regsolmejoras.frx":1181
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Informes..."
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_bus 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_regsolmejoras.frx":170B
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Buscar..."
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_eli 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_regsolmejoras.frx":1C95
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Eliminar registro seleccionado"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_regsolmejoras.frx":221F
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancelar..."
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_gra 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_regsolmejoras.frx":27A9
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Grabar datos"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      Picture         =   "frm_regsolmejoras.frx":2D33
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Modificar datos"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton b_nue 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_regsolmejoras.frx":32BD
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo registro"
      Top             =   5280
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Datos de solicitud"
      Enabled         =   0   'False
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Doble click para EDITAR cuadro de DESCRIPCION"
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Solicitud terminada"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   4920
         Width           =   2775
      End
      Begin VB.TextBox t_obs 
         Height          =   855
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   3960
         Width           =   6495
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         ItemData        =   "frm_regsolmejoras.frx":3847
         Left            =   2400
         List            =   "frm_regsolmejoras.frx":3854
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3480
         Width           =   2895
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   7920
         TabIndex        =   16
         Top             =   3480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_usua 
         Height          =   375
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txt_base 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txt_det 
         Height          =   1200
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1200
         Width           =   7095
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_regsolmejoras.frx":387A
         Left            =   1800
         List            =   "frm_regsolmejoras.frx":3890
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   3975
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   8040
         TabIndex        =   6
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
         TabIndex        =   4
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
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label labusfin 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   34
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   33
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "DESCRIPCION DE LA ACCION:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF8080&
         Caption         =   "Acción tomada:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Control por parte de Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3120
         Width           =   3255
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF8080&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   9240
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Base:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Solicitante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "DESCRIPCION:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Mejora en:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Hora:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Nro."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   6480
      Picture         =   "frm_regsolmejoras.frx":38DD
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "frm_solmejoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_bus_Click()
If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "AGUILLEN" Or WElusuario = "DARIOH" Then
    Dim Xingfec As String
    Xingfec = InputBox("Ingrese a partir de que fecha:")
    If Xingfec <> "" Then
       Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " and cl_fnac >=#" & Format(Xingfec, "yyyy/mm/dd") & "# order by cl_fnac"
       Data1.Refresh
       DBGrid1.SetFocus
    Else
       Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " order by cl_fnac"
       Data1.Refresh
       DBGrid1.SetFocus
    End If
Else
    MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_can_Click()
If WElusuario = "ENRIQUE" Then
   XAlta = 0
   Frame1.Enabled = False
   b_can.Enabled = False
Else
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
End If
borracua
igualadat
   
End Sub

Private Sub b_eli_Click()
Dim Xqueborra As String
If WElusuario = "JFERNAN" Then
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
   If mfd.Text <> "__/__/____" And mhd.Text <> "__:__" Then
      data_reg.Recordset("cl_codigo") = txt_nro.Text
      data_reg.Recordset("cl_fnac") = mfd.Text
      data_reg.Recordset("cl_ruc") = mhd.Text
'      data_reg.Recordset("cl_descpag") = Combo1.Text
      data_reg.Recordset("cl_etiquet") = Combo1.ListIndex
      data_reg.Recordset("info_debit") = txt_det.Text
      data_reg.Recordset("cl_cantdia") = txt_base.Text
      data_reg.Recordset("cl_nomvend") = txt_usua.Text
      data_reg.Recordset("cl_schqmn") = Combo3.ListIndex
      If mfh.Text <> "__/__/____" Then
         data_reg.Recordset("cl_fecing") = mfh.Text
         data_reg.Recordset("cl_tipoclin") = mhh.Text
         data_reg.Recordset("cl_dircobr") = t_obs.Text
      End If
      data_reg.Recordset("val1") = Check1.Value
      data_reg.Recordset("cl_descpag") = labusfin.Caption
      data_reg.Recordset("cl_tipocli") = 1
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
   data_reg.Recordset("cl_fnac") = mfd.Text
   data_reg.Recordset("cl_ruc") = mhd.Text
'     data_reg.Recordset("cl_descpag") = Combo1.Text
   data_reg.Recordset("cl_etiquet") = Combo1.ListIndex
   data_reg.Recordset("info_debit") = txt_det.Text
   data_reg.Recordset("cl_cantdia") = txt_base.Text
   data_reg.Recordset("cl_nomvend") = txt_usua.Text
   data_reg.Recordset("cl_schqmn") = Combo3.ListIndex
   If mfh.Text <> "__/__/____" Then
      data_reg.Recordset("cl_fecing") = mfh.Text
      data_reg.Recordset("cl_tipoclin") = mhh.Text
      data_reg.Recordset("cl_dircobr") = t_obs.Text
   Else
      data_reg.Recordset("cl_dircobr") = t_obs.Text
   End If
   If Check1.Value = 1 Then
      labusfin.Caption = WElusuario
   Else
      labusfin.Caption = "S/R"
   End If
   data_reg.Recordset("val1") = Check1.Value
   data_reg.Recordset("cl_descpag") = labusfin.Caption
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
If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "AGUILLEN" Or WElusuario = "DARIOH" Then
   frm_solimejoras.Show vbModal
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_mod_Click()
XAlta = 0
If WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "AGUILLEN" Or WElusuario = "DARIOH" Or XWeltipoU = "ADMINISTRADOR" Then
    Frame1.Enabled = True
    mfd.SetFocus
    txt_det.Enabled = True
    If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "AGUILLEN" Or WElusuario = "DARIOH" Then
       b_imp.Enabled = True
       mfh.Enabled = True
       mhh.Enabled = True
       t_obs.Enabled = True
    Else
       b_imp.Enabled = False
       mfh.Enabled = False
       mhh.Enabled = False
       t_obs.Enabled = False
    End If
    
    desboton
Else
   If WElusuario = "ENRIQUE" Then
      Frame1.Enabled = True
      b_can.Enabled = True
   Else
      MsgBox "No es el usuario creador de la tarea", vbCritical
   End If
End If

End Sub

Private Sub b_nue_Click()
XAlta = 1
borracua
txt_nro.Text = ""
If data_reg.Recordset.RecordCount > 0 Then
   data_reg.Recordset.MoveLast
   txt_nro.Text = data_reg.Recordset("cl_codigo") + 1
Else
   txt_nro.Text = 1
End If
data_reg.Recordset.AddNew
Frame1.Enabled = True
txt_base.Text = frm_menu.data_parse.Recordset("base")
'mfd.SetFocus
mfd.Text = Format(Date, "dd/mm/yyyy")
mhd.Text = Format(Time, "HH:mm")
txt_usua.Text = WElusuario
desboton
Combo1.SetFocus
Combo3.Enabled = False
mfh.Enabled = False
mhh.Enabled = False
t_obs.Enabled = False
Check1.Enabled = False


End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   labusfin.Caption = WElusuario
Else
   labusfin.Caption = "S/R"
End If

End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
   If WElusuario = "JFERNAN" Or WElusuario = "AGUILLEN" Or WElusuario = "SPEREZ" Or WElusuario = "DARIOH" Then
      Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " and val1 <>" & 1 & " order by cl_fnac DESC"
   Else
      Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " and val1 <>" & 1 & " and cl_nomvend ='" & WElusuario & "' order by cl_fnac DESC"
   End If
   Data1.Refresh
   borracua
Else
   If WElusuario = "JFERNAN" Or WElusuario = "AGUILLEN" Or WElusuario = "SPEREZ" Or WElusuario = "DARIOH" Then
      Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " order by cl_fnac DESC"
   Else
      Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " and cl_nomvend ='" & WElusuario & "' order by cl_fnac DESC"
   End If
   Data1.Refresh
   borracua
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_det.SetFocus
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub DBGrid1_DblClick()
data_reg.Recordset.FindFirst "cl_codigo =" & Data1.Recordset("cl_codigo")
If Not data_reg.Recordset.NoMatch Then
   borracua
   igualadat
Else
   MsgBox "No se encuentra registro, REINTENTAR!"
End If

End Sub

Private Sub Form_Load()
data_reg.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_reg.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " order by cl_codigo"
data_reg.Refresh
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'igualadat
If WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "AGUILLEN" Or XWeltipoU = "USUARIOS ADM" Or XWeltipoU = "ADMINISTRADOR" Or WElusuario = "DARIOH" Then
   Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " order by cl_codigo"
Else
   Data1.RecordSource = "Select * from solaudito where cl_tipocli =" & 1 & " and cl_nomvend ='" & WElusuario & "' order by cl_codigo"
End If

Data1.Refresh

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



Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhd.SetFocus
End If

End Sub


Private Sub mfh_GotFocus()
mfh.Text = Date

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhh.SetFocus
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

Private Sub mhh_GotFocus()
mhh.Text = Format(Time, "HH:mm")

End Sub

Private Sub mhh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_obs.SetFocus
End If

End Sub

Private Sub txt_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_usua.SetFocus
End If

End Sub

Private Sub txt_det_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_base.SetFocus
End If

End Sub

Public Function igualadat()
Dim Xnumeror As Long

On Error GoTo Quepasa

If data_reg.Recordset.RecordCount > 0 Then
   Xnumeror = 1
   txt_nro.Text = data_reg.Recordset("cl_codigo")
   Xnumeror = 2
   If IsNull(data_reg.Recordset("cl_fnac")) = False Then
      If data_reg.Recordset("cl_fnac") <> "" Then
         mfd.Text = Format(data_reg.Recordset("cl_fnac"), "dd/mm/yyyy")
      Else
         mfd.Text = "__/__/____"
      End If
   Else
      mfd.Text = "__/__/____"
   End If
   Xnumeror = 3
   If IsNull(data_reg.Recordset("cl_ruc")) = False Then
      mhd.Text = data_reg.Recordset("cl_ruc")
   Else
      mhd.Text = "__:__"
   End If
   Xnumeror = 4
   If IsNull(data_reg.Recordset("cl_etiquet")) = False Then
'      data_reg.Recordset("cl_descpag") = Combo1.Text
      Combo1.ListIndex = data_reg.Recordset("cl_etiquet")
   Else
      Combo1.ListIndex = -1
   End If
   Xnumeror = 5
   If IsNull(data_reg.Recordset("info_debit")) = False Then
      txt_det.Text = data_reg.Recordset("info_debit")
   Else
      txt_det.Text = ""
   End If
   Xnumeror = 6
   If IsNull(data_reg.Recordset("cl_cantdia")) = False Then
      txt_base.Text = data_reg.Recordset("cl_cantdia")
   Else
      txt_base.Text = 0
   End If
   Xnumeror = 7
   If IsNull(data_reg.Recordset("cl_nomvend")) = False Then
      txt_usua.Text = data_reg.Recordset("cl_nomvend")
   Else
      txt_usua.Text = ""
   End If
   Xnumeror = 8
   If IsNull(data_reg.Recordset("val1")) = False Then
      Check1.Value = data_reg.Recordset("val1")
   Else
      Check1.Value = 0
   End If
   Xnumeror = 9
   If IsNull(data_reg.Recordset("cl_fecing")) = False Then
      mfh.Text = Format(data_reg.Recordset("cl_fecing"), "dd/mm/yyyy")
   Else
      mfh.Text = "__/__/____"
   End If
   Xnumeror = 12
   If IsNull(data_reg.Recordset("cl_tipoclin")) = False Then
      mhh.Text = data_reg.Recordset("cl_tipoclin")
   Else
      mhh.Text = "__:__"
   End If
   Xnumeror = 13
   If IsNull(data_reg.Recordset("cl_dircobr")) = False Then
      t_obs.Text = data_reg.Recordset("cl_dircobr")
   Else
      t_obs.Text = ""
   End If
   Xnumeror = 14
   If IsNull(data_reg.Recordset("cl_schqmn")) = False Then
      Combo3.ListIndex = data_reg.Recordset("cl_schqmn")
   Else
      Combo3.ListIndex = -1
   End If
   If IsNull(data_reg.Recordset("cl_descpag")) = False Then
      labusfin.Caption = data_reg.Recordset("cl_descpag")
   Else
      labusfin.Caption = "S/R"
   End If
End If

Quepasa:
        If Err.Number > 0 Then
           MsgBox ("ES: " & Trim(str(Xnumeror)))
        End If
        
End Function

Public Function borracua()
txt_nro.Text = ""
mfd.Text = "__/__/____"
mhd.Text = "__:__"
Combo1.ListIndex = -1
txt_det.Text = ""
txt_base.Text = ""
txt_usua.Text = ""
Check1.Value = 0
mfh.Text = "__/__/____"
mhh.Text = "__:__"
t_obs.Text = ""
Combo3.ListIndex = -1
labusfin.Caption = ""

End Function

