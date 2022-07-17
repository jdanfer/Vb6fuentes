VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_atsocio 
   BackColor       =   &H00FF8080&
   Caption         =   "Sistema de administración y atención al cliente"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_atsocio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_ing2 
      Caption         =   "data_ing2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
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
      Top             =   7200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver acciones..."
      Height          =   495
      Left            =   9480
      Picture         =   "frm_atsocio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Data data_ingreso 
      Caption         =   "data_ingreso"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos del registro"
      Enabled         =   0   'False
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   11655
      Begin MSMask.MaskEdBox mfecrec 
         Height          =   375
         Left            =   10200
         TabIndex        =   52
         ToolTipText     =   "Fecha de recepción"
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_atsocio.frx":09CC
         Left            =   9240
         List            =   "frm_atsocio.frx":09E2
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Data data_numero 
         Caption         =   "data_numero"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6000
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_atsocio.frx":0A1D
         Left            =   6360
         List            =   "frm_atsocio.frx":0A3F
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox cboconf 
         Enabled         =   0   'False
         Height          =   360
         ItemData        =   "frm_atsocio.frx":0AB2
         Left            =   3240
         List            =   "frm_atsocio.frx":0ABF
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   5880
         Width           =   3135
      End
      Begin VB.TextBox txt_telef 
         Height          =   375
         Left            =   8040
         TabIndex        =   42
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txt_usufin 
         Enabled         =   0   'False
         Height          =   375
         Left            =   9120
         TabIndex        =   40
         Top             =   5400
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mhorfin 
         Height          =   375
         Left            =   6240
         TabIndex        =   38
         Top             =   5400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "HH:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfecfin 
         Height          =   375
         Left            =   3240
         TabIndex        =   36
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboest 
         Height          =   360
         ItemData        =   "frm_atsocio.frx":0AE5
         Left            =   3240
         List            =   "frm_atsocio.frx":0AF2
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox txt_det 
         Height          =   975
         Left            =   3240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   3480
         Width           =   8175
      End
      Begin VB.ComboBox cbodet 
         Height          =   360
         ItemData        =   "frm_atsocio.frx":0B18
         Left            =   1800
         List            =   "frm_atsocio.frx":0B2E
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2880
         Width           =   3615
      End
      Begin MSMask.MaskEdBox mnac 
         Height          =   375
         Left            =   5160
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ming 
         Height          =   375
         Left            =   1680
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_codced 
         Height          =   375
         Left            =   11160
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txt_ced 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9720
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_desconv 
         Height          =   375
         Left            =   3720
         MaxLength       =   60
         TabIndex        =   21
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txt_codconv 
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Buscar socio"
         Height          =   375
         Left            =   9960
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_nomb 
         Height          =   375
         Left            =   3720
         MaxLength       =   70
         TabIndex        =   17
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox txt_cliente 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mhora 
         Height          =   375
         Left            =   10440
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "HH:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfecha 
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_usua 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fec.Rec:"
         Height          =   375
         Left            =   9240
         TabIndex        =   51
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label20 
         BackColor       =   &H0080C0FF&
         Caption         =   "Recibido vía:"
         Height          =   375
         Left            =   7800
         TabIndex        =   49
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackColor       =   &H0080C0FF&
         Caption         =   "Motivo:"
         Height          =   375
         Left            =   5520
         TabIndex        =   47
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label18 
         BackColor       =   &H0080C0FF&
         Caption         =   "Datos de finalización."
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   4
         X1              =   0
         X2              =   11640
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         Caption         =   "Conformidad del cliente:"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080C0FF&
         Caption         =   "Teléfonos:"
         Height          =   375
         Left            =   6720
         TabIndex        =   41
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "Usuario:"
         Height          =   375
         Left            =   7560
         TabIndex        =   39
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hora:"
         Height          =   375
         Left            =   5160
         TabIndex        =   37
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fecha de finalizado:"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   5400
         Width           =   3135
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Estado actual:"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   4560
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "Detalle:"
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "Recepción de:"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   4
         X1              =   0
         X2              =   11640
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Nacimiento:"
         Height          =   375
         Left            =   3720
         TabIndex        =   27
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ingreso:"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cédula:"
         Height          =   375
         Left            =   8760
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Convenio:"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cliente:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hora:"
         Height          =   375
         Left            =   9240
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fecha:"
         Height          =   375
         Left            =   6360
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Usuario:"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Nro.Registro:"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar"
      Height          =   495
      Left            =   9960
      Picture         =   "frm_atsocio.frx":0B73
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informes"
      Height          =   495
      Left            =   7800
      Picture         =   "frm_atsocio.frx":10FD
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      Picture         =   "frm_atsocio.frx":1687
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      Picture         =   "frm_atsocio.frx":1C11
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modificar"
      Height          =   495
      Left            =   2040
      Picture         =   "frm_atsocio.frx":219B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   120
      Picture         =   "frm_atsocio.frx":2725
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "frm_atsocio.frx":2CAF
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1695
   End
End
Attribute VB_Name = "frm_atsocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbodet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_det.SetFocus
End If
'185
End Sub

Private Sub cboest_Click()
If cboest.ListIndex > 0 Then
   mfecfin.Enabled = True
   mhorfin.Enabled = True
   txt_usufin.Enabled = True
   cboconf.Enabled = True
Else
   mfecfin.Enabled = False
   mhorfin.Enabled = False
   txt_usufin.Enabled = False
   cboconf.Enabled = False
End If

End Sub

Private Sub Command1_Click()
Frame1.Enabled = True
borraat
txt_cliente.SetFocus
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command3.Enabled = True ' Grabar

'If data_ingreso.Recordset.RecordCount > 0 Then
'   data_ingreso.Recordset.MoveLast
txt_nro.Text = data_numero.Recordset("musada") + 1

data_numero.Recordset.Edit
data_numero.Recordset("musada") = txt_nro.Text
data_numero.Recordset.Update

mfecha.Text = Date
mhora.Text = Format(Time, "HH:mm:ss")
txt_usua.Text = WElusuario

data_ingreso.Recordset.AddNew
XAlta = 1


End Sub

Private Sub Command2_Click()
If txt_nro.Text <> "" Then
    XAlta = 0
    Frame1.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    Command3.Enabled = True ' Grabar
    
    data_ingreso.RecordSource = "Select * from ingresosat where at_nro =" & txt_nro.Text
    data_ingreso.Refresh
    If data_ingreso.Recordset.RecordCount > 0 Then
       borraat
       igualaat
    Else
       MsgBox "ATENCION!! No se encontró el registro, verifique!!", vbCritical
       Command4_Click
    End If
End If

End Sub

Private Sub Command3_Click()
Dim XFaltadato As Integer
XFaltadato = 0

On Error GoTo Quepasa

If txt_cliente.Text = "" Then
   XFaltadato = 1
End If
If Trim(txt_nomb.Text) = "" Then
   XFaltadato = 1
End If
If Trim(txt_codconv.Text) = "" Then
   XFaltadato = 1
End If
If Trim(txt_ced.Text) = "" Then
   XFaltadato = 1
End If
If Trim(txt_telef.Text) = "" Then
   XFaltadato = 1
Else
   If Len(txt_telef.Text) < 4 Then
      XFaltadato = 1
   End If
End If
If cbodet.ListIndex < 0 Then
   XFaltadato = 1
End If
If Combo1.ListIndex < 0 Then
   XFaltadato = 1
End If
If mfecrec.Text = "__/__/____" Then
   XFaltadato = 1
End If
If Trim(txt_det.Text) = "" Then
   XFaltadato = 1
End If
If cboest.ListIndex < 0 Then
   XFaltadato = 1
End If
If Combo2.ListIndex < 0 Then
   XFaltadato = 1
End If

If XFaltadato = 0 Then

    If XAlta = 1 Then
       If txt_cliente.Text = "" Then
          txt_cliente.Text = 0
       End If
       data_ingreso.Recordset("at_cliente") = txt_cliente.Text
       data_ingreso.Recordset("at_nomb") = txt_nomb.Text
       If txt_codconv.Text = "" Then
          txt_codconv.Text = "AABBCC"
       End If
       data_ingreso.Recordset("at_codconv") = txt_codconv.Text
       If txt_desconv.Text = "" Then
          txt_desconv.Text = "AABBCC"
       End If
       data_ingreso.Recordset("at_nomconv") = txt_desconv.Text
       If txt_ced.Text = "" Then
          txt_ced.Text = 0
       End If
       data_ingreso.Recordset("at_ced") = txt_ced.Text
       If txt_codced.Text = "" Then
          txt_codced.Text = 0
       End If
       data_ingreso.Recordset("at_via") = Combo2.ListIndex
       data_ingreso.Recordset("at_codced") = txt_codced.Text
       If ming.Text <> "__/__/____" Then
          data_ingreso.Recordset("at_ing") = ming.Text
       End If
       If mfecrec.Text <> "__/__/____" Then
          data_ingreso.Recordset("at_fecrec") = mfecrec.Text
       End If
       If mnac.Text <> "__/__/____" Then
          data_ingreso.Recordset("at_nac") = mnac.Text
       End If
       data_ingreso.Recordset("at_telef") = txt_telef.Text
       data_ingreso.Recordset("at_nro") = txt_nro.Text
       data_ingreso.Recordset("at_usuario") = txt_usua.Text
       data_ingreso.Recordset("at_fecha") = mfecha.Text
       data_ingreso.Recordset("at_hora") = mhora.Text
       data_ingreso.Recordset("at_categ") = cbodet.ListIndex
       If cbodet.ListIndex = 0 Then
          data_ingreso.Recordset("at_descat") = "QUEJA"
       Else
          If cbodet.ListIndex = 1 Then
             data_ingreso.Recordset("at_descat") = "RECLAMO"
          Else
             If cbodet.ListIndex = 2 Then
                data_ingreso.Recordset("at_descat") = "SUGERENCIA"
             Else
                If cbodet.ListIndex = 3 Then
                   data_ingreso.Recordset("at_descat") = "CONSULTA"
                Else
                   If cbodet.ListIndex = 4 Then
                      data_ingreso.Recordset("at_descat") = "AGRADECIMIENTO"
                   Else
                      If cbodet.ListIndex = 5 Then
                         data_ingreso.Recordset("at_descat") = "SOLICITUD"
                      End If
                   End If
                End If
             End If
          End If
       End If
       data_ingreso.Recordset("at_detal") = txt_det.Text
       If cboest.ListIndex >= 0 Then
          data_ingreso.Recordset("at_estado") = cboest.ListIndex
       Else
          data_ingreso.Recordset("at_estado") = 0
       End If
       If mfecfin.Text = "__/__/____" Then
       Else
          data_ingreso.Recordset("at_fecfin") = mfecfin.Text
       End If
       If mhorfin.Text = "__:__:__" Then
       Else
          data_ingreso.Recordset("at_horfin") = mhorfin.Text
       End If
       data_ingreso.Recordset("at_usufin") = txt_usufin.Text
       data_ingreso.Recordset("at_confor") = cboconf.ListIndex
       data_ingreso.Recordset("at_motiind") = Combo1.ListIndex
       If Combo1.ListIndex >= 0 Then
          data_ingreso.Recordset("at_moti") = Combo1.Text
       End If
       data_ingreso.Recordset.Update
       data_ingreso.RecordSource = "Select * from ingresosat where at_nro =" & txt_nro.Text
       data_ingreso.Refresh
       XAlta = 0
    Else
       data_ingreso.Recordset.Edit
       If txt_cliente.Text = "" Then
          txt_cliente.Text = 0
       End If
       data_ingreso.Recordset("at_cliente") = txt_cliente.Text
       data_ingreso.Recordset("at_nomb") = txt_nomb.Text
       If txt_codconv.Text = "" Then
          txt_codconv.Text = "AABBCC"
       End If
       data_ingreso.Recordset("at_codconv") = txt_codconv.Text
       If txt_desconv.Text = "" Then
          txt_desconv.Text = "AABBCC"
       End If
       data_ingreso.Recordset("at_nomconv") = txt_desconv.Text
       If txt_ced.Text = "" Then
          txt_ced.Text = 0
       End If
       data_ingreso.Recordset("at_ced") = txt_ced.Text
       If txt_codced.Text = "" Then
          txt_codced.Text = 0
       End If
       data_ingreso.Recordset("at_codced") = txt_codced.Text
       If mfecrec.Text <> "__/__/____" Then
          data_ingreso.Recordset("at_fecrec") = mfecrec.Text
       End If
       If ming.Text <> "__/__/____" Then
          data_ingreso.Recordset("at_ing") = ming.Text
       End If
       If mnac.Text <> "__/__/____" Then
          data_ingreso.Recordset("at_nac") = mnac.Text
       Else
          data_ingreso.Recordset("at_nac") = Null
       End If
       data_ingreso.Recordset("at_via") = Combo2.ListIndex
       data_ingreso.Recordset("at_telef") = txt_telef.Text
    '   data_ingreso.Recordset("at_nro") = txt_nro.Text
       data_ingreso.Recordset("at_usuario") = txt_usua.Text
       data_ingreso.Recordset("at_fecha") = mfecha.Text
       data_ingreso.Recordset("at_hora") = mhora.Text
       data_ingreso.Recordset("at_categ") = cbodet.ListIndex
       If cbodet.ListIndex = 0 Then
          data_ingreso.Recordset("at_descat") = "QUEJA"
       Else
          If cbodet.ListIndex = 1 Then
             data_ingreso.Recordset("at_descat") = "RECLAMO"
          Else
             If cbodet.ListIndex = 2 Then
                data_ingreso.Recordset("at_descat") = "SUGERENCIA"
             Else
                If cbodet.ListIndex = 3 Then
                   data_ingreso.Recordset("at_descat") = "CONSULTA"
                Else
                   If cbodet.ListIndex = 4 Then
                      data_ingreso.Recordset("at_descat") = "AGRADECIMIENTO"
                   Else
                      If cbodet.ListIndex = 5 Then
                         data_ingreso.Recordset("at_descat") = "SOLICITUD"
                      End If
                   End If
                End If
             End If
          End If
       End If
       data_ingreso.Recordset("at_detal") = txt_det.Text
       If cboest.ListIndex >= 0 Then
          data_ingreso.Recordset("at_estado") = cboest.ListIndex
       Else
          data_ingreso.Recordset("at_estado") = 0
       End If
       If mfecfin.Text = "__/__/____" Then
       Else
          data_ingreso.Recordset("at_fecfin") = mfecfin.Text
       End If
       If mhorfin.Text = "__:__:__" Then
       Else
          data_ingreso.Recordset("at_horfin") = mhorfin.Text
       End If
       data_ingreso.Recordset("at_usufin") = txt_usufin.Text
       data_ingreso.Recordset("at_confor") = cboconf.ListIndex
       If Combo1.ListIndex >= 0 Then
          data_ingreso.Recordset("at_motiind") = Combo1.ListIndex
       End If
       If Combo1.ListIndex >= 0 Then
          data_ingreso.Recordset("at_moti") = Combo1.Text
       End If
       data_ingreso.Recordset.Update
       data_ingreso.RecordSource = "Select * from ingresosat where at_nro =" & txt_nro.Text
       data_ingreso.Refresh
    
    End If
Else
    MsgBox "Faltan datos para poder grabar el registro, VERIFIQUE!", vbCritical
    
End If
Exit Sub

XAlta = 0
Command1.Enabled = True 'Nuevo
Command2.Enabled = True ' Modif
Command4.Enabled = False ' Cancelar
Command5.Enabled = True ' Informes
Command6.Enabled = True ' Buscar
Command8.Enabled = True ' Acciones
Command3.Enabled = False ' Grabar
borraat
If data_ingreso.Recordset.RecordCount > 0 Then
   igualaat
End If
Frame1.Enabled = False

Quepasa:
        If Err.Number = 3197 Then
           MsgBox "No hay datos para modificar, verifique o presione cancelar.", vbInformation
        Else
           MsgBox "Error al grabar, verifique datos", vbCritical
        End If

End Sub

Private Sub Command4_Click()
If XAlta = 1 Then
   data_ingreso.Recordset.CancelUpdate
   XAlta = 0
   Command1.Enabled = True 'Nuevo
   Command2.Enabled = True ' Modif
   Command4.Enabled = False ' Cancelar
   Command5.Enabled = True ' Informes
   Command6.Enabled = True ' Buscar
   Command8.Enabled = True ' Acciones
   Command3.Enabled = False ' Grabar
   borraat
   Frame1.Enabled = False
Else
   XAlta = 0
   Command1.Enabled = True 'Nuevo
   Command2.Enabled = True ' Modif
   Command4.Enabled = False ' Cancelar
   Command5.Enabled = True ' Informes
   Command6.Enabled = True ' Buscar
   Command8.Enabled = True ' Acciones
   Command3.Enabled = False ' Grabar
   borraat
   Frame1.Enabled = False

End If


End Sub

Private Sub Command5_Click()
frm_infatsoc.Show vbModal

End Sub

Private Sub Command6_Click()
frm_buscaatsoc.Show vbModal

End Sub

Private Sub Command7_Click()
frm_buscasocat.Show vbModal

End Sub

Private Sub Command8_Click()
frm_accat.Show vbModal

End Sub

Private Sub Form_Load()
data_ingreso.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ingreso.RecordSource = "Select * from ingresosat where at_cliente =" & 20809
data_ingreso.Refresh

data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_ing2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ing2.RecordSource = "Select * from cobrador where cb_recatra =" & 2
data_ing2.Refresh

data_numero.DatabaseName = App.path & "\parse.mdb"
data_numero.RecordSource = "parsec0"
data_numero.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub mfecfin_GotFocus()
mfecfin.Text = Date

End Sub

Private Sub mhorfin_GotFocus()
mhorfin.Text = Format(Time, "HH:mm:ss")

End Sub

Private Sub ming_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mnac.SetFocus
End If

End Sub

Private Sub mnac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_telef.SetFocus
End If

End Sub

Private Sub txt_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codced.SetFocus
End If

End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomb.SetFocus
End If

End Sub

Private Sub txt_cliente_LostFocus()
If txt_cliente.Text <> "" Then
   If txt_cliente.Text > 0 Then
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & txt_cliente.Text
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         txt_nomb.Text = data_cli.Recordset("cl_apellid")
         txt_codconv.Text = data_cli.Recordset("cl_codconv")
         txt_desconv.Text = data_cli.Recordset("cl_nomconv")
         If IsNull(data_cli.Recordset("cl_cedula")) = False Then
            txt_ced.Text = data_cli.Recordset("cl_cedula")
         Else
            txt_ced.Text = 0
         End If
         If IsNull(data_cli.Recordset("cl_codced")) = False Then
            txt_codced.Text = data_cli.Recordset("cl_codced")
         Else
            txt_codced.Text = 0
         End If
         If IsNull(data_cli.Recordset("cl_fecing")) = False Then
            ming.Text = Format(data_cli.Recordset("cl_fecing"), "dd/mm/yyyy")
         Else
            ming.Text = "__/__/____"
         End If
         If IsNull(data_cli.Recordset("cl_fnac")) = False Then
            mnac.Text = Format(data_cli.Recordset("cl_fnac"), "dd/mm/yyyy")
         Else
            mnac.Text = "__/__/____"
         End If
         If IsNull(data_cli.Recordset("cl_telefon")) = False Then
            txt_telef.Text = data_cli.Recordset("cl_telefon")
         Else
            txt_telef.Text = ""
         End If
      End If
   End If
Else
   txt_cliente.Text = 0
End If

End Sub

Public Function borraat()
txt_cliente.Text = ""
txt_nomb.Text = ""
txt_codconv.Text = ""
txt_desconv.Text = ""
txt_ced.Text = ""
txt_codced.Text = ""
ming.Text = "__/__/____"
mnac.Text = "__/__/____"
txt_telef.Text = ""
txt_nro.Text = ""
txt_usua.Text = ""
mfecha.Text = "__/__/____"
mhora.Text = "__:__:__"
cbodet.ListIndex = -1
txt_det.Text = ""
cboest.ListIndex = -1
mfecfin.Text = "__/__/____"
mfecrec.Text = "__/__/____"
mhorfin.Text = "__:__:__"
txt_usufin.Text = ""
cboconf.ListIndex = -1
Combo1.ListIndex = -1
Combo2.ListIndex = -1

End Function

Public Function igualaat()
If IsNull(data_ingreso.Recordset("at_cliente")) = False Then
   txt_cliente.Text = data_ingreso.Recordset("at_cliente")
Else
   txt_cliente.Text = 0
End If
If IsNull(data_ingreso.Recordset("at_nomb")) = False Then
   txt_nomb.Text = data_ingreso.Recordset("at_nomb")
Else
   txt_nomb.Text = ""
End If
If IsNull(data_ingreso.Recordset("at_codconv")) = False Then
   txt_codconv.Text = data_ingreso.Recordset("at_codconv")
Else
   txt_codconv.Text = ""
End If
If IsNull(data_ingreso.Recordset("at_nomconv")) = False Then
   txt_desconv.Text = data_ingreso.Recordset("at_nomconv")
Else
   txt_desconv.Text = ""
End If
If IsNull(data_ingreso.Recordset("at_ced")) = False Then
   txt_ced.Text = data_ingreso.Recordset("at_ced")
Else
   txt_ced.Text = 0
End If
If IsNull(data_ingreso.Recordset("at_codced")) = False Then
   txt_codced.Text = data_ingreso.Recordset("at_codced")
Else
   txt_codced.Text = 0
End If
If IsNull(data_ingreso.Recordset("at_ing")) = False Then
   ming.Text = Format(data_ingreso.Recordset("at_ing"), "dd/mm/yyyy")
Else
   ming.Text = "__/__/____"
End If
If IsNull(data_ingreso.Recordset("at_fecrec")) = False Then
   mfecrec.Text = Format(data_ingreso.Recordset("at_fecrec"), "dd/mm/yyyy")
Else
   mfecrec.Text = "__/__/____"
End If
If IsNull(data_ingreso.Recordset("at_nac")) = False Then
   mnac.Text = Format(data_ingreso.Recordset("at_nac"), "dd/mm/yyyy")
Else
   mnac.Text = "__/__/____"
End If
If IsNull(data_ingreso.Recordset("at_via")) = False Then
   Combo2.ListIndex = data_ingreso.Recordset("at_via")
Else
   Combo2.ListIndex = -1
End If

If IsNull(data_ingreso.Recordset("at_telef")) = False Then
   txt_telef.Text = data_ingreso.Recordset("at_telef")
Else
   txt_telef.Text = ""
End If
If IsNull(data_ingreso.Recordset("at_nro")) = False Then
   txt_nro.Text = data_ingreso.Recordset("at_nro")
Else
   txt_nro.Text = 9999999
End If
If IsNull(data_ingreso.Recordset("at_fecha")) = False Then
   mfecha.Text = data_ingreso.Recordset("at_fecha")
Else
   mfecha.Text = Date
End If
If IsNull(data_ingreso.Recordset("at_hora")) = False Then
   mhora.Text = data_ingreso.Recordset("at_hora")
Else
   mhora.Text = Format(Time, "HH:mm:ss")
End If
If IsNull(data_ingreso.Recordset("at_usuario")) = False Then
   txt_usua.Text = data_ingreso.Recordset("at_usuario")
Else
   txt_usua.Text = "YO"
End If
If IsNull(data_ingreso.Recordset("at_categ")) = False Then
   cbodet.ListIndex = data_ingreso.Recordset("at_categ")
Else
   cbodet.ListIndex = 0
End If
If IsNull(data_ingreso.Recordset("at_detal")) = False Then
   txt_det.Text = data_ingreso.Recordset("at_detal")
Else
   txt_det.Text = ""
End If
If IsNull(data_ingreso.Recordset("at_estado")) = False Then
   cboest.ListIndex = data_ingreso.Recordset("at_estado")
Else
   cboest.ListIndex = -1
End If
If IsNull(data_ingreso.Recordset("at_fecfin")) = False Then
   mfecfin.Text = data_ingreso.Recordset("at_fecfin")
Else
   mfecfin.Text = "__/__/____"
End If
If IsNull(data_ingreso.Recordset("at_horfin")) = False Then
   mhorfin.Text = data_ingreso.Recordset("at_horfin")
Else
   mhorfin.Text = "__:__:__"
End If
If IsNull(data_ingreso.Recordset("at_usufin")) = False Then
   txt_usufin.Text = data_ingreso.Recordset("at_usufin")
Else
   txt_usufin.Text = ""
End If
If IsNull(data_ingreso.Recordset("at_confor")) = False Then
   cboconf.ListIndex = data_ingreso.Recordset("at_confor")
Else
   cboconf.ListIndex = -1
End If
If IsNull(data_ingreso.Recordset("at_motiind")) = False Then
   Combo1.ListIndex = data_ingreso.Recordset("at_motiind")
Else
   Combo1.ListIndex = -1
End If


End Function

Private Sub txt_codced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ming.SetFocus
End If

End Sub

Private Sub txt_codconv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_desconv.SetFocus
End If

End Sub

Private Sub txt_desconv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ced.SetFocus
End If

End Sub

Private Sub txt_det_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboest.SetFocus
End If

End Sub

Private Sub txt_nomb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codconv.SetFocus
End If

End Sub

Private Sub txt_telef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbodet.SetFocus
End If

End Sub

Private Sub txt_usufin_GotFocus()
txt_usufin.Text = WElusuario

End Sub
