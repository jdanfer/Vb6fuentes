VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_serv2da 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Segunda Opinión Médica"
   ClientHeight    =   8535
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
   Icon            =   "frm_serv2da.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_numera 
      Caption         =   "data_numera"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver solo registros sin ingresar conformidad "
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   6480
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
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_serv2da.frx":0442
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frm_serv2da.frx":0456
      TabIndex        =   25
      Top             =   6720
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
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      Picture         =   "frm_serv2da.frx":1331
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Informes..."
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton b_bus 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      Picture         =   "frm_serv2da.frx":18BB
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Buscar.."
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton b_eli 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      Picture         =   "frm_serv2da.frx":1E45
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Eliminar registro seleccionado"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      Picture         =   "frm_serv2da.frx":23CF
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cancelar..."
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton b_gra 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Picture         =   "frm_serv2da.frx":2959
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Grabar datos"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_serv2da.frx":2EE3
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Modificar/Editar dato"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton b_nue 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_serv2da.frx":346D
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Crear NUEVO registro"
      Top             =   5640
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Datos de solicitud"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9255
      Begin VB.Data data_cli 
         Caption         =   "data_cli"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_serv2da.frx":39F7
         Left            =   5640
         List            =   "frm_serv2da.frx":3A04
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   4560
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mhconf 
         Height          =   375
         Left            =   4800
         TabIndex        =   27
         Top             =   4560
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfconf 
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Top             =   4560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_obs 
         Height          =   720
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   3600
         Width           =   7335
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_conta 
         Height          =   375
         Left            =   6480
         MaxLength       =   25
         TabIndex        =   10
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txt_mat 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         ToolTipText     =   "Cédula del paciente"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txt_det 
         Height          =   1200
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1560
         Width           =   7095
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   6240
         TabIndex        =   5
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
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label labusufin 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   7200
         TabIndex        =   35
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label labcat 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7320
         TabIndex        =   34
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "Conformidad del cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   4560
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   7320
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contacto:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4560
         TabIndex        =   29
         Top             =   720
         Width           =   1935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   9240
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Área de CONTROLES:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   9240
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha Hora terminado:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label labnom 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   6375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Paciente:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Más detalles:"
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha/Hora:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro."
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "REGISTRO DE SOLICITUD 2da. OPINIÓN MÉDICA"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   9255
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   7560
      Picture         =   "frm_serv2da.frx":3A2B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "frm_serv2da"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_bus_Click()
'3, opt1
Dim Xingfec As String
Xingfec = InputBox("Ingrese a partir de que fecha:")
If Xingfec <> "" Then
   Data1.RecordSource = "Select * from segunda_op where fecha >=#" & Format(Xingfec, "yyyy/mm/dd") & "# order by fecha DESC"
   Data1.Refresh
Else
   Data1.RecordSource = "Select * from ssegunda_op order by fecha DESC"
   Data1.Refresh
End If
DBGrid1.SetFocus

End Sub

Private Sub b_can_Click()
XAlta = 0
Frame1.Enabled = False
habboton
borracua
'igualadat
   
End Sub

Private Sub b_eli_Click()
Dim Xqueborra As String
If WElusuario = "COMPUTOS" Then
    Xqueborra = MsgBox("Desea borrar el registro seleccionado?", vbInformation + vbYesNo, "SAPP")
    If Xqueborra = vbYes Then
       data_reg.Recordset.FindFirst "id =" & txt_nro.Text
       If Not data_reg.Recordset.NoMatch Then
          data_reg.Recordset.Delete
          data_reg.Refresh
          borracua
'          igualadat
       End If
    End If
Else
    MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_gra_Click()
On Error GoTo Algrab

If XAlta = 1 Then
   If txt_det.Text <> "" Then
      If mfd.Text <> "__/__/____" And mhd.Text <> "__:__" Then
         data_reg.Recordset.AddNew
         data_reg.Recordset("fecha") = mfd.Text
         data_reg.Recordset("hora") = mhd.Text
         data_reg.Recordset("base") = frm_menu.data_parse.Recordset("base")
         data_reg.Recordset("usuario") = Label3.Caption
         If txt_mat.Text <> "" Then
            data_reg.Recordset("matricula") = txt_mat.Text
         End If
         If txt_conta.Text <> "" Then
            data_reg.Recordset("contacto") = txt_conta.Text
         End If
         If labnom.Caption <> "" Then
            data_reg.Recordset("soc_nom") = Mid(labnom.Caption, 1, 80)
         End If
         If labcat.Caption <> "" Then
            data_reg.Recordset("soc_cat") = labcat.Caption
         End If
         data_reg.Recordset("detalle") = txt_det.Text
         If mfh.Text <> "__/__/____" Then
            data_reg.Recordset("finfecha") = mfh.Text
         End If
         If mhh.Text <> "__:__" Then
            data_reg.Recordset("finhora") = mhh.Text
         End If
         If labusufin.Caption <> "" Then
            data_reg.Recordset("finusuario") = labusufin.Caption
         End If
         If txt_obs.Text <> "" Then
            data_reg.Recordset("finobs") = txt_obs.Text
         End If
         If mfconf.Text <> "__/__/____" Then
            data_reg.Recordset("conffecha") = mfconf.Text
         End If
         If mhconf.Text <> "__:__" Then
            data_reg.Recordset("confhora") = mhconf.Text
         End If
         data_reg.Recordset("confop") = Trim(str(Combo2.ListIndex))
         data_reg.Recordset.Update
         data_reg.Refresh
         Data1.Refresh
         borracua
         habboton
'           igualadat
      Else
        MsgBox "No ingresó fecha", vbInformation
      End If
   Else
      MsgBox "No ingresó detalles", vbInformation
   End If
Else
   If txt_nro.Text <> "" Then
      data_reg.RecordSource = "Select * from segunda_op where id =" & txt_nro.Text
      data_reg.Refresh
      If data_reg.Recordset.RecordCount > 0 Then
         data_reg.Recordset.Edit
         data_reg.Recordset("fecha") = mfd.Text
         data_reg.Recordset("hora") = mhd.Text
         data_reg.Recordset("base") = frm_menu.data_parse.Recordset("base")
         data_reg.Recordset("usuario") = Label3.Caption
         If txt_mat.Text <> "" Then
            data_reg.Recordset("matricula") = txt_mat.Text
         End If
         If txt_conta.Text <> "" Then
            data_reg.Recordset("contacto") = txt_conta.Text
         End If
         If labnom.Caption <> "" Then
            data_reg.Recordset("soc_nom") = Mid(labnom.Caption, 1, 80)
         End If
         If labcat.Caption <> "" Then
            data_reg.Recordset("soc_cat") = labcat.Caption
         End If
         data_reg.Recordset("detalle") = txt_det.Text
         If mfh.Text <> "__/__/____" Then
            data_reg.Recordset("finfecha") = mfh.Text
         End If
         If mhh.Text <> "__:__" Then
            data_reg.Recordset("finhora") = mhh.Text
         End If
         If labusufin.Caption <> "" Then
            data_reg.Recordset("finusuario") = labusufin.Caption
         End If
         If txt_obs.Text <> "" Then
            data_reg.Recordset("finobs") = txt_obs.Text
         End If
         If mfconf.Text <> "__/__/____" Then
            data_reg.Recordset("conffecha") = mfconf.Text
         End If
         If mhconf.Text <> "__:__" Then
            data_reg.Recordset("confhora") = mhconf.Text
         End If
         data_reg.Recordset("confop") = Combo2.ListIndex
         data_reg.Recordset.Update
          
         Data1.Refresh
         XAlta = 0
         Frame1.Enabled = False
         habboton
         borracua
        '   igualadat
      End If
   Else
      MsgBox "No seleccionó registro"
   End If
End If

Exit Sub

Algrab:
        If Err.Number = 3155 Then
           MsgBox "Error al grabar datos, verifique información"
        Else
           MsgBox "Error al grabar"
        End If
        

End Sub

Private Sub b_imp_Click()
frm_infservap.Show vbModal

End Sub

Private Sub b_mod_Click()
XAlta = 0
Frame1.Enabled = True
mfd.SetFocus
desboton

End Sub

Private Sub b_nue_Click()
XAlta = 1
borracua
txt_nro.Text = ""
'txt_nro.Text = data_numera.Recordset("mnueva") + 1
'data_reg.Recordset.AddNew
Frame1.Enabled = True
'txt_base.Text = frm_menu.data_parse.Recordset("base")
'mfd.SetFocus
mfd.Text = Format(Date, "dd/mm/yyyy")
mhd.Text = Format(Time, "HH:mm")
Label3.Caption = WElusuario
txt_mat.SetFocus

'data_numera.Recordset.Edit
'data_numera.Recordset("mnueva") = txt_nro.Text
'data_numera.Recordset.Update
b_imp.Enabled = False
mfh.Enabled = True
mhh.Enabled = True
txt_obs.Enabled = True
desboton

End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
   Data1.RecordSource = "Select * from segunda_op where finusuario is not null order by fecha DESC"
Else
   Data1.RecordSource = "Select * from segunda_op order by fecha DESC"
End If
borracua


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_det.SetFocus
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(Data1.Recordset("id")) = False Then
   txt_nro.Text = Data1.Recordset("id")
End If
If IsNull(Data1.Recordset("fecha")) = False Then
   mfd.Text = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
End If
If IsNull(Data1.Recordset("hora")) = False Then
   mhd.Text = Format(Data1.Recordset("hora"), "HH:mm")
End If
If IsNull(Data1.Recordset("usuario")) = False Then
   Label3.Caption = Data1.Recordset("usuario")
End If
If IsNull(Data1.Recordset("matricula")) = False Then
   txt_mat.Text = Data1.Recordset("matricula")
Else
   txt_mat.Text = ""
End If
If IsNull(Data1.Recordset("contacto")) = False Then
   txt_conta.Text = Data1.Recordset("contacto")
Else
   txt_conta.Text = ""
End If
If IsNull(Data1.Recordset("soc_nom")) = False Then
   labnom.Caption = Data1.Recordset("soc_nom")
Else
   labnom.Caption = "NN"
End If
If IsNull(Data1.Recordset("soc_cat")) = False Then
   labcat.Caption = Data1.Recordset("soc_cat")
Else
   labcat.Caption = "PART"
End If
If IsNull(Data1.Recordset("detalle")) = False Then
   txt_det.Text = Data1.Recordset("detalle")
Else
   txt_det.Text = ""
End If
If IsNull(Data1.Recordset("finfecha")) = False Then
   mfh.Text = Format(Data1.Recordset("finfecha"), "dd/mm/yyyy")
Else
   mfh.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("finhora")) = False Then
   mhh.Text = Format(Data1.Recordset("finhora"), "HH:mm")
Else
   mhh.Text = "__:__"
End If
If IsNull(Data1.Recordset("finusuario")) = False Then
   labusufin.Caption = Data1.Recordset("finusuario")
Else
   labusufin.Caption = ""
End If
If IsNull(Data1.Recordset("finobs")) = False Then
   txt_obs.Text = Data1.Recordset("finobs")
Else
   txt_obs.Text = ""
End If
If IsNull(Data1.Recordset("conffecha")) = False Then
   mfconf.Text = Format(Data1.Recordset("conffecha"), "dd/mm/yyyy")
Else
   mfconf.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("confhora")) = False Then
   mhconf.Text = Format(Data1.Recordset("confhora"), "HH:mm")
Else
   mhconf.Text = "__:__"
End If
If IsNull(Data1.Recordset("confop")) = False Then
   Combo2.ListIndex = Val(Data1.Recordset("confop"))
Else
   Combo2.ListIndex = -1
End If


End Sub

Private Sub Form_Load()

'ConectarBD
'ConbdSapp.Open
'Sqlconsulta = "Select * from estudios where flia =" & 20 & " order by codest"
'With Registro1
'    .CursorLocation = adUseClient
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open Sqlconsulta, ConbdSapp, , , adCmdText
'End With
'If Registro1.RecordCount > 0 Then
'   Registro1.MoveFirst
'   Do While Not Registro1.EOF
'      Combo1.AddItem Registro1("descrip")
'      Registro1.MoveNext
'   Loop
'End If
'Registro1.Close
'ConbdSapp.Close
'ConbdSapp2.Open
'Sqlconsulta = "Select * from env_soc order by cl_fnac DESC"
'With Registro1
'    .CursorLocation = adUseClient
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open Sqlconsulta, ConbdSapp2, , , adCmdText
'End With
txt_nro.Text = ""
mfd.Text = "__/__/____"
mhd.Text = "__:__"
txt_det.Text = ""
txt_mat.Text = ""
txt_conta.Text = ""
mfh.Text = "__/__/____"
mhh.Text = "__:__"
txt_obs.Text = ""
mfconf.Text = "__/__/____"
mhconf.Text = "__:__"
Combo2.ListIndex = -1

data_reg.DatabaseName = ""
data_reg.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_reg.RecordSource = "segunda_op"
data_reg.Refresh


data_cli.Connect = "ODBC;DSN=" & Xconexrmt & ";"

Data1.Connect = "ODBC;DSN=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from segunda_op order by fecha DESC"
Data1.Refresh

data_numera.DatabaseName = App.path & "\parse.mdb"
data_numera.RecordSource = "parsec0"
data_numera.Refresh




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
    .Top = 0
    .Left = 0
    .Height = Me.Height
    .Width = Me.Width
End With

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

Private Sub mfh_GotFocus()
mfh.Text = Format(Date, "dd/mm/yyyy")
mhh.Text = Format(Time, "HH:mm")
labusufin.Caption = WElusuario

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

Private Sub txt_conta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_det.SetFocus
End If

End Sub

Private Sub txt_det_Click()
txt_det.Text = txt_det.Text

End Sub

Public Function igualadat()
        
End Function

Public Function borracua()
txt_nro.Text = ""
mfd.Text = "__/__/____"
mhd.Text = "__:__"
Label3.Caption = ""
labnom.Caption = ""
labcat.Caption = ""
labusufin.Caption = ""
txt_det.Text = ""
txt_mat.Text = ""
txt_conta.Text = ""
mfh.Text = "__/__/____"
mhh.Text = "__:__"
txt_obs.Text = ""
mfconf.Text = "__/__/____"
mhconf.Text = "__:__"
Combo2.ListIndex = -1

End Function

Private Sub txt_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_conta.SetFocus
End If

End Sub

Private Sub txt_mat_LostFocus()
If txt_mat.Text <> "" Then
   data_cli.RecordSource = "select * from clientes where cl_codigo =" & txt_mat.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      labnom.Caption = data_cli.Recordset("cl_apellid")
      labcat.Caption = data_cli.Recordset("cl_codconv")
   End If
Else
   labnom.Caption = ""
   labcat.Caption = ""
   
End If
End Sub

