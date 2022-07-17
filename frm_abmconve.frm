VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_abmconve 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Mantenimiento de convenios (FACTURACION)"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frm_abmconve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_parf 
      Caption         =   "data_parf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H008080FF&
      Height          =   615
      Left            =   6600
      Picture         =   "frm_abmconve.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H008080FF&
      Height          =   615
      Left            =   5520
      Picture         =   "frm_abmconve.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton b_elimi 
      BackColor       =   &H008080FF&
      Height          =   615
      Left            =   4440
      Picture         =   "frm_abmconve.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Eliminar registro seleccionado (BAJA)"
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      Picture         =   "frm_abmconve.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      Height          =   615
      Left            =   2280
      Picture         =   "frm_abmconve.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H008080FF&
      Height          =   615
      Left            =   1200
      Picture         =   "frm_abmconve.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Editar registro para modificación"
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H008080FF&
      Height          =   615
      Left            =   120
      Picture         =   "frm_abmconve.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Nuevo registro"
      Top             =   7320
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Datos de convenio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.Data data_pro 
         Caption         =   "data_pro"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   8400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Data data_cob 
         Caption         =   "data_cob"
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
         Top             =   4920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox t_codpro 
         Height          =   375
         Left            =   6240
         TabIndex        =   47
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox t_codcob 
         Height          =   375
         Left            =   4920
         TabIndex        =   46
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox t_codz 
         Height          =   360
         Left            =   5880
         TabIndex        =   45
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data data_zona 
         Caption         =   "data_zona"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton b_hist 
         BackColor       =   &H008080FF&
         Caption         =   "Historial"
         Height          =   495
         Left            =   8880
         MouseIcon       =   "frm_abmconve.frx":2438
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton b_fact 
         BackColor       =   &H008080FF&
         Caption         =   "Facturar"
         Height          =   495
         Left            =   7200
         MouseIcon       =   "frm_abmconve.frx":2742
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox t_obs 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   6480
         Width           =   8415
      End
      Begin VB.TextBox t_datos 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   5760
         Width           =   8415
      End
      Begin VB.ComboBox cbopro 
         Enabled         =   0   'False
         Height          =   360
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   5160
         Width           =   3135
      End
      Begin VB.ComboBox cbocob 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   5160
         Width           =   3615
      End
      Begin VB.TextBox t_der 
         Enabled         =   0   'False
         Height          =   735
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   4320
         Width           =   8415
      End
      Begin VB.TextBox t_imp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         ToolTipText     =   "SI ESTE CAMPO ESTÁ EN CERO: AL FACTURAR Extrae el importe de la factura desde la emisión SAPP"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox cbozon 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox t_correo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   6360
         MaxLength       =   30
         TabIndex        =   18
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox t_tel 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox t_dir 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   14
         Top             =   2280
         Width           =   5055
      End
      Begin MSMask.MaskEdBox mfing 
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox t_rut 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox t_razon 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox t_nom 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   6
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox t_cod 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         X1              =   6960
         X2              =   10560
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label labultano 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   9240
         TabIndex        =   35
         Top             =   840
         Width           =   855
      End
      Begin VB.Label labultmes 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   8640
         TabIndex        =   34
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ult. Pago:"
         Height          =   255
         Left            =   7320
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Observaciones:"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Datos de Facturación:"
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   10560
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Promotor:"
         Height          =   255
         Left            =   5640
         TabIndex        =   27
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cobrador:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Convenio:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   10560
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Importe Factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Zona:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         X1              =   6960
         X2              =   10560
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         X1              =   6960
         X2              =   6960
         Y1              =   120
         Y2              =   2640
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Correo:"
         Height          =   255
         Left            =   4920
         TabIndex        =   17
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ingreso:"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "RUT:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Razón Social:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NOMBRE:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ACTIVO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9000
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ESTADO:"
         Height          =   255
         Left            =   7320
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CODIGO:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2520
   End
End
Attribute VB_Name = "frm_abmconve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text12_Change()

End Sub

Public Sub Borrar()
t_cod.Text = ""
mfing.Text = "__/__/____"
t_nom.Text = ""
t_razon.Text = ""
t_rut.Text = ""
t_dir.Text = ""
t_tel.Text = ""
t_correo.Text = ""
cbozon.ListIndex = -1
t_imp.Text = ""
t_der.Text = ""
cbocob.ListIndex = -1
cbopro.ListIndex = -1
t_datos.Text = ""
t_obs.Text = ""

End Sub

Private Sub b_alta_Click()
XAlta = 1
hab_campos
Borrar
t_cod.Text = data_parf.Recordset("ultimo_soc")
t_cod.Enabled = False
t_nom.SetFocus
mfing.Text = Date
des_boton

End Sub

Private Sub b_busca_Click()
frm_buscacnvf.Show vbModal

End Sub

Private Sub b_cancela_Click()
    Borrar
    hab_boton
    des_campos
    XAlta = 0

End Sub

Private Sub b_fact_Click()
frm_quefactcnv.Show vbModal

End Sub

Private Sub b_graba_Click()
If XAlta = 1 Then
   data_cli.RecordSource = "Select * from clifact where cl_codigo =" & t_cod.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      MsgBox "Ya existe matrícula, VERIFIQUE!"
   Else
      data_cli.Recordset.AddNew
      data_cli.Recordset("cl_codigo") = t_cod.Text
      data_cli.Recordset("cl_apellid") = t_nom.Text
      data_cli.Recordset("cl_nombre") = t_razon.Text
      If mfing.Text <> "__/__/____" Then
         data_cli.Recordset("cl_fecing") = mfing.Text
      End If
      data_cli.Recordset("estado") = 1
      data_cli.Recordset("cl_nom_sup") = t_rut.Text
      data_cli.Recordset("cl_direcci") = t_dir.Text
      data_cli.Recordset("cl_telefon") = t_tel.Text
      data_cli.Recordset("cl_email") = t_correo.Text
      data_cli.Recordset("cl_zona") = cbozon.Text
      data_cli.Recordset("cl_grupo") = t_codz.Text
      data_cli.Recordset("cl_cuopaga") = t_imp.Text
      data_cli.Recordset("derechos") = t_der.Text
      data_cli.Recordset("cl_nrocobr") = t_codcob.Text
      data_cli.Recordset("cl_nomcobr") = cbocob.Text
      data_cli.Recordset("cl_nrovend") = t_codpro.Text
      data_cli.Recordset("cl_nomvend") = cbopro.Text
      data_cli.Recordset("obsfact") = t_datos.Text
      data_cli.Recordset("observa") = t_obs.Text
      data_cli.Recordset("saldo_cc") = 0
      data_cli.Recordset("cl_ultmesp") = 0
      data_cli.Recordset("cl_ultanop") = 0
      data_cli.Recordset.Update
      data_cli.Refresh
      data_parf.Recordset.Edit
      data_parf.Recordset("ultimo_soc") = data_parf.Recordset("ultimo_soc") + 1
      data_parf.Recordset.Update
      data_parf.Refresh
      Borrar
      hab_boton
      des_campos
      XAlta = 0
   End If
Else
    data_cli.Recordset.Edit
    data_cli.Recordset("cl_apellid") = t_nom.Text
    data_cli.Recordset("cl_nombre") = t_razon.Text
    If mfing.Text <> "__/__/____" Then
       data_cli.Recordset("cl_fecing") = mfing.Text
    End If
    data_cli.Recordset("cl_nom_sup") = t_rut.Text
    data_cli.Recordset("cl_direcci") = t_dir.Text
    data_cli.Recordset("cl_telefon") = t_tel.Text
    data_cli.Recordset("cl_email") = t_correo.Text
    data_cli.Recordset("cl_zona") = cbozon.Text
    data_cli.Recordset("cl_grupo") = t_codz.Text
    data_cli.Recordset("cl_cuopaga") = t_imp.Text
    data_cli.Recordset("derechos") = t_der.Text
    data_cli.Recordset("cl_nrocobr") = t_codcob.Text
    data_cli.Recordset("cl_nomcobr") = cbocob.Text
    data_cli.Recordset("cl_nrovend") = t_codpro.Text
    data_cli.Recordset("cl_nomvend") = cbopro.Text
    data_cli.Recordset("obsfact") = t_datos.Text
    data_cli.Recordset("observa") = t_obs.Text
    data_cli.Recordset.Update
    data_cli.Refresh
    Borrar
    hab_boton
    des_campos
    XAlta = 0
   
End If

End Sub

Private Sub b_hist_Click()
frm_estcnv.Show vbModal

End Sub

Private Sub b_imp_Click()
frm_inffaccnv.Show vbModal

End Sub

Private Sub b_modif_Click()
If t_cod.Text <> "" Then
   data_cli.RecordSource = "Select * from clifact where cl_codigo =" & t_cod.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      XAlta = 0
      hab_campos
      Borrar
      iguala_Datos
      des_boton
      t_nom.SetFocus
   Else
      MsgBox "Error al cargar datos, REINTENTE!!"
   End If
Else
   MsgBox "Error al buscar...SIN CODIGO!!"
End If

End Sub


Private Sub cbocob_Click()
data_cob.RecordSource = "Select * from cobrador where cb_nombre ='" & cbocob.Text & "'"
data_cob.Refresh
If data_cob.Recordset.RecordCount > 0 Then
   t_codcob.Text = data_cob.Recordset("cb_numero")
Else
   t_codcob.Text = 0
End If

End Sub

Private Sub cbocob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbopro.SetFocus
End If

End Sub

Private Sub cbopro_Click()
data_pro.RecordSource = "Select * from vendedor where vn_nombre ='" & cbopro.Text & "'"
data_pro.Refresh
If data_pro.Recordset.RecordCount > 0 Then
   t_codpro.Text = data_pro.Recordset("vn_numero")
Else
   t_codpro.Text = 0
End If

End Sub

Private Sub cbopro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_datos.SetFocus
End If

End Sub


Private Sub cbozon_Click()
data_zona.RecordSource = "Select * from zonas where zo_nombre ='" & cbozon.Text & "'"
data_zona.Refresh
If data_zona.Recordset.RecordCount > 0 Then
   t_codz.Text = data_zona.Recordset("zo_grupo")
Else
   t_codz.Text = 0
End If

End Sub

Private Sub cbozon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_imp.SetFocus
End If

End Sub

Private Sub Form_Load()
data_parf.DatabaseName = App.Path & "\parsef.mdb"
data_parf.RecordSource = "parsec0"
data_parf.Refresh
data_cli.Connect = "ODBC;DSN=facturacion;"

data_zona.DatabaseName = App.Path & "\sapp.mdb"
data_zona.RecordSource = "Select * from zonas order by zo_nombre"
data_zona.Refresh
If data_zona.Recordset.RecordCount > 0 Then
   data_zona.Recordset.MoveFirst
   Do While Not data_zona.Recordset.EOF
      cbozon.AddItem data_zona.Recordset("zo_nombre")
      data_zona.Recordset.MoveNext
   Loop
End If

data_cob.DatabaseName = App.Path & "\sapp.mdb"
data_cob.RecordSource = "Select * from cobrador order by cb_nombre"
data_cob.Refresh
If data_cob.Recordset.RecordCount > 0 Then
   data_cob.Recordset.MoveFirst
   Do While Not data_cob.Recordset.EOF
      cbocob.AddItem data_cob.Recordset("cb_nombre")
      data_cob.Recordset.MoveNext
   Loop
End If

data_pro.DatabaseName = App.Path & "\sapp.mdb"
data_pro.RecordSource = "Select * from vendedor order by vn_nombre"
data_pro.Refresh
If data_pro.Recordset.RecordCount > 0 Then
   data_pro.Recordset.MoveFirst
   Do While Not data_pro.Recordset.EOF
      cbopro.AddItem data_pro.Recordset("vn_nombre")
      data_pro.Recordset.MoveNext
   Loop
End If

End Sub

Public Sub hab_campos()
t_cod.Enabled = True
mfing.Enabled = True
t_nom.Enabled = True
t_razon.Enabled = True
t_rut.Enabled = True
t_dir.Enabled = True
t_tel.Enabled = True
t_correo.Enabled = True
cbozon.Enabled = True
t_imp.Enabled = True
t_der.Enabled = True
cbocob.Enabled = True
cbopro.Enabled = True
t_datos.Enabled = True
t_obs.Enabled = True

End Sub

Public Sub des_campos()
t_cod.Enabled = False
mfing.Enabled = False
t_nom.Enabled = False
t_razon.Enabled = False
t_rut.Enabled = False
t_dir.Enabled = False
t_tel.Enabled = False
t_correo.Enabled = False
cbozon.Enabled = False
t_imp.Enabled = False
t_der.Enabled = False
cbocob.Enabled = False
cbopro.Enabled = False
t_datos.Enabled = False
t_obs.Enabled = False

End Sub

Public Sub des_boton()
b_fact.Enabled = False
b_hist.Enabled = False
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cancela.Enabled = True
b_imp.Enabled = False
b_busca.Enabled = False
b_elimi.Enabled = False


End Sub

Public Sub hab_boton()
b_fact.Enabled = True
b_hist.Enabled = True
b_alta.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cancela.Enabled = False
b_imp.Enabled = True
b_busca.Enabled = True
b_elimi.Enabled = True

End Sub

Public Sub iguala_Datos()
If IsNull(data_cli.Recordset("cl_codigo")) = False Then
   t_cod.Text = data_cli.Recordset("cl_codigo")
Else
   t_cod.Text = 0
End If
If IsNull(data_cli.Recordset("cl_apellid")) = False Then
   t_nom.Text = data_cli.Recordset("cl_apellid")
Else
   t_nom.Text = ""
End If
If IsNull(data_cli.Recordset("cl_nombre")) = False Then
   t_razon.Text = data_cli.Recordset("cl_nombre")
Else
   t_razon.Text = ""
End If
If IsNull(data_cli.Recordset("cl_fecing")) = False Then
   mfing.Text = data_cli.Recordset("cl_fecing")
Else
   mfing.Text = "__/__/____"
End If
If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
   labultmes.Caption = data_cli.Recordset("cl_ultmesp")
Else
   labultmes.Caption = 0
End If
If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
   labultano.Caption = data_cli.Recordset("cl_ultanop")
Else
   labultano.Caption = 0
End If
If IsNull(data_cli.Recordset("cl_nom_sup")) = False Then
   t_rut.Text = data_cli.Recordset("cl_nom_sup")
Else
   t_rut.Text = 0
End If
If IsNull(data_cli.Recordset("cl_direcci")) = False Then
   t_dir.Text = data_cli.Recordset("cl_direcci")
Else
   t_dir.Text = ""
End If
If IsNull(data_cli.Recordset("cl_email")) = False Then
   t_correo.Text = data_cli.Recordset("cl_email")
Else
   t_correo.Text = ""
End If
If IsNull(data_cli.Recordset("cl_telefon")) = False Then
   t_tel.Text = data_cli.Recordset("cl_telefon")
Else
   t_tel.Text = ""
End If
If IsNull(data_cli.Recordset("cl_zona")) = False Then
   cbozon.Text = data_cli.Recordset("cl_zona")
   t_codz.Text = data_cli.Recordset("cl_grupo")
Else
   cbozon.ListIndex = -1
   t_codz.Text = 0
End If
If IsNull(data_cli.Recordset("cl_cuopaga")) = False Then
   t_imp.Text = data_cli.Recordset("cl_cuopaga")
Else
   t_imp.Text = 0
End If
If IsNull(data_cli.Recordset("derechos")) = False Then
   t_der.Text = data_cli.Recordset("derechos")
Else
   t_der.Text = ""
End If
If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
   t_codcob.Text = data_cli.Recordset("cl_nrocobr")
   cbocob.Text = data_cli.Recordset("cl_nomcobr")
Else
   t_codcob.Text = 0
   cbocob.ListIndex = -1
End If
If IsNull(data_cli.Recordset("cl_nrovend")) = False Then
   frm_abmconve.t_codpro.Text = data_cli.Recordset("cl_nrovend")
   frm_abmconve.cbopro.Text = data_cli.Recordset("cl_nomvend")
Else
   t_codpro.Text = 0
   cbopro.ListIndex = -1
End If
If IsNull(data_cli.Recordset("obsfact")) = False Then
   t_datos.Text = data_cli.Recordset("obsfact")
Else
   t_datos.Text = ""
End If
If IsNull(data_cli.Recordset("observa")) = False Then
   t_obs.Text = data_cli.Recordset("observa")
Else
   t_obs.Text = ""
End If

End Sub

Private Sub t_cli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_der.SetFocus
End If

End Sub

Private Sub Label13_Click()

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_correo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbozon.SetFocus
End If

End Sub

Private Sub t_datos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_obs.SetFocus
End If

End Sub

Private Sub t_der_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocob.SetFocus
End If

End Sub

Private Sub t_dir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_tel.SetFocus
End If

End Sub

Private Sub t_imp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_der.SetFocus
End If

End Sub

Private Sub t_imp_LostFocus()
If t_imp.Text = "" Then
   t_imp.Text = 0
End If
t_imp.Text = Format(t_imp.Text, "Standard")

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_razon.SetFocus
End If

End Sub

Private Sub t_razon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_rut.SetFocus
End If

End Sub

Private Sub t_rut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_dir.SetFocus
End If

End Sub

Private Sub t_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_correo.SetFocus
End If

End Sub
