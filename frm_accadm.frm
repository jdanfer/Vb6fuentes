VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_accadm 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acciones administrativas"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_accadm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9045
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Registrar Código de autorización"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Ingresar código de autorización para poder facturar en recepciones de base"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Data data_gri 
      Caption         =   "data_gri"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data data_acc 
      Caption         =   "data_acc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_accadm.frx":058A
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frm_accadm.frx":05A1
      TabIndex        =   27
      Top             =   4680
      Width           =   8775
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3000
      Picture         =   "frm_accadm.frx":1470
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Informes"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      Picture         =   "frm_accadm.frx":19FA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      Picture         =   "frm_accadm.frx":1F84
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   840
      Picture         =   "frm_accadm.frx":250E
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Editar registro"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      Picture         =   "frm_accadm.frx":2A98
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Nuevo registro"
      Top             =   4200
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos "
      Enabled         =   0   'False
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox t_contac 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1440
         MaxLength       =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   7095
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         ItemData        =   "frm_accadm.frx":3022
         Left            =   4560
         List            =   "frm_accadm.frx":3035
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2880
         Width           =   3015
      End
      Begin MSMask.MaskEdBox mfp 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_det 
         BackColor       =   &H00C0FFFF&
         Height          =   975
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1800
         Width           =   7215
      End
      Begin VB.TextBox t_desref 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   3240
         MaxLength       =   150
         TabIndex        =   13
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox t_nroref 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox t_base 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   840
         Width           =   615
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "Contacto:"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Forma de pago:"
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fecha pago:"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Detalles:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Referencia:"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label labu 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Base:"
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hora:"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fecha:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label labnom 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label labmat 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Matrícula:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   4320
      Picture         =   "frm_accadm.frx":3079
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1935
   End
End
Attribute VB_Name = "frm_accadm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
XAlta = 1
b_alta.Enabled = False
b_edita.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_imp.Enabled = False
Frame1.Enabled = True
borrar_dat
Label5.Caption = Data1.Recordset("nro_accadm") + 1
Data1.Recordset.Edit
Data1.Recordset("nro_accadm") = Data1.Recordset("nro_accadm") + 1
Data1.Recordset.Update
mf.Text = Date
mh.Text = Format(Time, "HH:mm")
t_base.Text = Data1.Recordset("base")
labu.Caption = WElusuario
If Xqueregi <> 0 Then
   t_nroref.Text = Xqueregi
   t_desref.Text = Xlehas
Else
   t_nroref.Text = ""
   t_desref.Text = ""
End If
t_det.SetFocus


End Sub

Private Sub b_cance_Click()
      XAlta = 0
      b_alta.Enabled = True
      b_edita.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_imp.Enabled = True
      borrar_dat
      Frame1.Enabled = False

End Sub

Private Sub b_edita_Click()
data_acc.RecordSource = "Select * from mant_sol where estado =" & data_gri.Recordset("estado")
data_acc.Refresh
If data_acc.Recordset.RecordCount > 0 Then
   XAlta = 0
   b_alta.Enabled = False
   b_edita.Enabled = False
   b_graba.Enabled = True
   b_cance.Enabled = True
   b_imp.Enabled = False
   Frame1.Enabled = True
   borrar_dat
   iguala_datos
   t_det.SetFocus
Else
   borrar_dat
End If

End Sub

Private Sub b_graba_Click()
If XAlta = 1 Then
   If t_det.Text <> "" And mf.Text <> "__/__/____" Then
      data_acc.Recordset.AddNew
      data_acc.Recordset("cl_fultpag") = Format(mf.Text, "dd/mm/yyyy")
      data_acc.Recordset("estado") = Label5.Caption
      data_acc.Recordset("cl_codigo") = Label5.Caption
      data_acc.Recordset("cl_ruc") = Format(mh.Text, "HH:mm")
      data_acc.Recordset("cl_atrasop") = t_base.Text
      data_acc.Recordset("cl_descpag") = labu.Caption
      data_acc.Recordset("cl_zona") = labmat.Caption
      If t_nroref.Text = "" Then
         t_nroref.Text = 0
      End If
      If IsNumeric(t_nroref.Text) = True Then
         data_acc.Recordset("cl_nro_sup") = t_nroref.Text
      Else
         data_acc.Recordset("cl_nro_sup") = 0
      End If
      If t_desref.Text <> "" Then
         data_acc.Recordset("cl_email") = t_desref.Text
      End If
      data_acc.Recordset("info_debit") = t_det.Text
      If mfp.Text <> "__/__/____" Then
         data_acc.Recordset("cl_fec1") = Format(mfp.Text, "dd/mm/yyyy")
      End If
      data_acc.Recordset("cl_val3") = Combo1.ListIndex
      If Combo1.ListIndex >= 0 Then
         data_acc.Recordset("cl_desc2") = Combo1.Text
      End If
      If t_contac.Text <> "" Then
         data_acc.Recordset("cl_desc1") = t_contac.Text
      End If
      data_acc.Recordset.Update
      data_gri.Refresh
      XAlta = 0
      b_alta.Enabled = True
      b_edita.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_imp.Enabled = True
      borrar_dat
      Frame1.Enabled = False
   End If
Else
   If t_det.Text <> "" And mf.Text <> "__/__/____" Then
      data_acc.Recordset.Edit
      If t_nroref.Text = "" Then
         t_nroref.Text = 0
      End If
      If IsNumeric(t_nroref.Text) = True Then
         data_acc.Recordset("cl_nro_sup") = t_nroref.Text
      Else
         data_acc.Recordset("cl_nro_sup") = 0
      End If
      If t_desref.Text <> "" Then
         data_acc.Recordset("cl_email") = t_desref.Text
      End If
      data_acc.Recordset("info_debit") = t_det.Text
      If mfp.Text <> "__/__/____" Then
         data_acc.Recordset("cl_fec1") = Format(mfp.Text, "dd/mm/yyyy")
      Else
         data_acc.Recordset("cl_fec1") = Null
      End If
      data_acc.Recordset("cl_val3") = Combo1.ListIndex
      If Combo1.ListIndex >= 0 Then
         data_acc.Recordset("cl_desc2") = Combo1.Text
      Else
         data_acc.Recordset("cl_desc2") = Null
      End If
      If t_contac.Text <> "" Then
         data_acc.Recordset("cl_desc1") = t_contac.Text
      Else
         data_acc.Recordset("cl_desc1") = Null
      End If
      data_acc.Recordset.Update
      data_gri.Refresh
      XAlta = 0
      b_alta.Enabled = True
      b_edita.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_imp.Enabled = True
      borrar_dat
      Frame1.Enabled = False
   End If
End If

End Sub

Private Sub b_imp_Click()
Wopsed = labmat.Caption

frm_infaccadm.Show vbModal

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_contac.SetFocus
End If

End Sub

Private Sub DBGrid1_DblClick()
data_acc.RecordSource = "Select * from mant_sol where estado =" & data_gri.Recordset("estado")
data_acc.Refresh
If data_acc.Recordset.RecordCount > 0 Then
   borrar_dat
   iguala_datos
Else
   borrar_dat
End If

End Sub

Private Sub Form_Load()
data_acc.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_acc.RecordSource = "Select * from mant_sol where cl_codigo =" & 30468 & " and estado is not null"
data_acc.Refresh

Data1.DatabaseName = App.path & "\paramb.mdb"
Data1.RecordSource = "paramb"
Data1.Refresh

data_gri.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_gri.RecordSource = "Select * from mant_sol where estado >" & 0 & " order by estado"
'data_gri.Refresh

If frm_veodeudab.Visible = True Then
   labmat.Caption = frm_veodeudab.Label2.Caption
   labnom.Caption = frm_veodeudab.Label3.Caption
   data_gri.RecordSource = "Select * from mant_sol where estado >" & 0 & " and cl_zona =" & labmat.Caption & " order by cl_fultpag"
   data_gri.Refresh
Else
   If Wveohistoadmd = 8 Then
      If frm_largador.txt_mat.Text <> "" Then
         If frm_largador.txt_mat.Text > 0 Then
            labmat.Caption = frm_largador.txt_mat.Text
            labnom.Caption = frm_largador.txt_nomb.Text
            data_gri.RecordSource = "Select * from mant_sol where estado >" & 0 & " and cl_zona =" & labmat.Caption & " order by cl_fultpag"
            data_gri.Refresh
         Else
            MsgBox "No se puede consultar historial sin matrícula"
'            data_gri.RecordSource = "Select * from mant_sol where estado >" & 0 & " and cl_zona =" & 0 & " order by estado"
'            data_gri.Refresh
         End If
      Else
         MsgBox "No se puede consultar historial sin matrícula"
'         data_gri.RecordSource = "Select * from mant_sol where estado >" & 0 & " and cl_zona =" & 0 & " order by estado"
'         data_gri.Refresh
      
      End If
   Else
      If frmabm.txt_mat.Caption = "" Then
         MsgBox "No seleccionó ningun socio"
         Unload Me
      Else
         labmat.Caption = frmabm.txt_mat.Caption
         labnom.Caption = frmabm.txt_apellid.Text
         data_gri.RecordSource = "Select * from mant_sol where estado >" & 0 & " and cl_zona =" & labmat.Caption & " order by cl_fultpag"
         data_gri.Refresh
      End If
   End If
End If
If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Then
Else
   b_alta.Enabled = False
   b_edita.Enabled = False
   Command1.Enabled = False
   b_imp.Enabled = False
End If
Wveohistoadmd = 0

End Sub

Public Sub borrar_dat()
mf.Text = "__/__/____"
mh.Text = "__:__"
t_base.Text = ""
labu.Caption = ""
t_det.Text = ""
mfp.Text = "__/__/____"
Combo1.ListIndex = -1
t_contac.Text = ""
t_nroref.Text = ""
t_desref.Text = ""


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With


End Sub

Private Sub mfp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub t_desref_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_det.SetFocus
End If

End Sub

Private Sub t_det_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfp.SetFocus
End If

End Sub

Private Sub t_nroref_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_desref.SetFocus
End If

End Sub

Public Sub iguala_datos()
If IsNull(data_acc.Recordset("estado")) = False Then
   Label5.Caption = data_acc.Recordset("estado")
Else
   Label5.Caption = ""
End If
If IsNull(data_acc.Recordset("cl_fultpag")) = False Then
   mf.Text = Format(data_acc.Recordset("cl_fultpag"), "dd/mm/yyyy")
Else
   mf.Text = "__/__/____"
End If
If IsNull(data_acc.Recordset("cl_ruc")) = False Then
   mh.Text = Format(data_acc.Recordset("cl_ruc"), "HH:mm")
Else
   mh.Text = "__:__"
End If
If IsNull(data_acc.Recordset("cl_atrasop")) = False Then
   t_base.Text = data_acc.Recordset("cl_atrasop")
Else
   t_base.Text = 0
End If
If IsNull(data_acc.Recordset("cl_descpag")) = False Then
   labu.Caption = data_acc.Recordset("cl_descpag")
Else
   labu.Caption = ""
End If
If IsNull(data_acc.Recordset("cl_nro_sup")) = False Then
   t_nroref.Text = data_acc.Recordset("cl_nro_sup")
Else
   t_nroref.Text = 0
End If
If IsNull(data_acc.Recordset("cl_email")) = False Then
   t_desref.Text = data_acc.Recordset("cl_email")
Else
   t_desref.Text = ""
End If
If IsNull(data_acc.Recordset("info_debit")) = False Then
   t_det.Text = data_acc.Recordset("info_debit")
Else
   t_det.Text = ""
End If
If IsNull(data_acc.Recordset("cl_fec1")) = False Then
   mfp.Text = Format(data_acc.Recordset("cl_fec1"), "dd/mm/yyyy")
Else
   mfp.Text = "__/__/____"
End If
If IsNull(data_acc.Recordset("cl_val3")) = False Then
   Combo1.ListIndex = data_acc.Recordset("cl_val3")
Else
   Combo1.ListIndex = -1
End If
If IsNull(data_acc.Recordset("cl_desc1")) = False Then
   t_contac.Text = data_acc.Recordset("cl_desc1")
Else
   t_contac.Text = ""
End If


End Sub
