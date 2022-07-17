VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_servap 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Servicios A.Protegidas"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_servap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo3 
      Height          =   360
      ItemData        =   "frm_servap.frx":0442
      Left            =   5880
      List            =   "frm_servap.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sin acciones de:"
      Height          =   255
      Left            =   3840
      TabIndex        =   43
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Data data_usuar 
      Caption         =   "data_usuar"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acciones"
      Height          =   495
      Left            =   8280
      Picture         =   "frm_servap.frx":0486
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Registrar o ver acciones para el registro seleccionado"
      Top             =   7320
      Width           =   2055
   End
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
      Top             =   7560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sin conformidad "
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   7920
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   9120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_servap.frx":0A10
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frm_servap.frx":0A24
      TabIndex        =   30
      Top             =   8280
      Width           =   10215
   End
   Begin VB.Data data_reg 
      Caption         =   "data_reg"
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      Picture         =   "frm_servap.frx":18F3
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Informes..."
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton b_bus 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      Picture         =   "frm_servap.frx":1E7D
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Buscar.."
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton b_eli 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      Picture         =   "frm_servap.frx":2407
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Eliminar registro seleccionado"
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      Picture         =   "frm_servap.frx":2991
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cancelar..."
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton b_gra 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Picture         =   "frm_servap.frx":2F1B
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Grabar datos"
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_servap.frx":34A5
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Modificar/Editar dato"
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton b_nue 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_servap.frx":3A2F
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Crear NUEVO registro"
      Top             =   7320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de solicitud"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Presupuesto Aceptado"
         Height          =   495
         Left            =   7800
         TabIndex        =   45
         ToolTipText     =   "Si no está marcado se toma como NO aceptado."
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Data data_users 
         Caption         =   "data_users"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Height          =   1020
         Left            =   5640
         TabIndex        =   40
         ToolTipText     =   "Doble click para borrar"
         Top             =   3360
         Width           =   4455
      End
      Begin VB.ComboBox cbousuarios 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3360
         Width           =   3975
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_servap.frx":3FB9
         Left            =   4200
         List            =   "frm_servap.frx":3FC6
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   6240
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mhconf 
         Height          =   375
         Left            =   3240
         TabIndex        =   33
         Top             =   6240
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
         Left            =   1800
         TabIndex        =   32
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Conformidad"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   6240
         Width           =   1695
      End
      Begin VB.TextBox txt_obs 
         Enabled         =   0   'False
         Height          =   720
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   5280
         Width           =   8415
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   4800
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   5520
         TabIndex        =   18
         Top             =   4800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Prog. SAPP"
         Height          =   255
         Left            =   6360
         TabIndex        =   16
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Correo Elect"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_usua 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         MaxLength       =   25
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_base 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txt_det 
         Height          =   1560
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1320
         Width           =   8295
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_servap.frx":3FED
         Left            =   1800
         List            =   "frm_servap.frx":3FEF
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   8295
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label labsi 
         Height          =   255
         Left            =   1920
         TabIndex        =   42
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seleccione usuarios a compartir:"
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   120
         TabIndex        =   38
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label labutermina 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   8040
         TabIndex        =   37
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   10200
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Área de CONTROLES:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   10200
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha y Hora de Terminado:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Solicitado vía:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Más detalles:"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Servicio:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hora:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   735
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
   Begin VB.Label estagrabado 
      Height          =   375
      Left            =   2880
      TabIndex        =   46
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REGISTRO DE SERVICIOS A.PROTEGIDAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2160
      TabIndex        =   35
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   9000
      Picture         =   "frm_servap.frx":3FF1
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1815
   End
End
Attribute VB_Name = "frm_servap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_bus_Click()
'3, opt1
Dim Xingfec As String
Dim Xbuscanro As String

Xbuscanro = MsgBox("Desea buscar por número de registro?", vbInformation + vbYesNo, "AP")
If Xbuscanro = vbYes Then
   Xingfec = InputBox("Ingrese número de registro:")
   If Xingfec <> "" Then
      Data1.RecordSource = "Select * from env_soc where cl_codigo >=" & Val(Xingfec) & " order by cl_codigo"
      Data1.Refresh
   Else
      Data1.RecordSource = "Select * from env_soc order by cl_fnac DESC"
      Data1.Refresh
   End If
Else
    Xingfec = InputBox("Ingrese a partir de que fecha:")
    If Xingfec <> "" Then
       Data1.RecordSource = "Select * from env_soc where cl_fnac >=#" & Format(Xingfec, "yyyy/mm/dd") & "# order by cl_fnac"
       Data1.Refresh
    Else
       Data1.RecordSource = "Select * from env_soc order by cl_fnac DESC"
       Data1.Refresh
    End If
End If
DBGrid1.SetFocus

End Sub

Private Sub b_can_Click()
If XAlta = 1 Then
   data_reg.Recordset.CancelUpdate
   List1.Clear
   XAlta = 0
   Frame1.Enabled = False
   habboton
Else
   XAlta = 0
   Frame1.Enabled = False
   List1.Clear
   habboton
End If
borracua
igualadat
   
End Sub

Private Sub b_eli_Click()
Dim Xqueborra As String
If WElusuario = "COMPUTOS" Then
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
    MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_gra_Click()
Dim X As Integer
Dim CorreoAP As String
Dim EnviarCorreo As String
Dim textocorreo, Cerrado As String
textocorreo = ""
EnviarCorreo = ""
CorreoAP = ""
Cerrado = "N"
X = 0
On Error GoTo Vererror
estagrabado.Caption = "0"

If List1.ListCount > 0 Then
    If XAlta = 1 Then
       If Combo1.ListIndex >= 0 Then
          If txt_det.Text <> "" Then
            If mfd.Text <> "__/__/____" And mhd.Text <> "__:__" Then
               data_reg.Recordset("cl_codigo") = txt_nro.Text
               data_reg.Recordset("cl_fnac") = mfd.Text
               data_reg.Recordset("cl_ruc") = mhd.Text
               data_reg.Recordset("cl_ter_vto") = 7
               data_reg.Recordset("cl_descpag") = Mid(Combo1.Text, 1, 100)
               data_reg.Recordset("estado") = Combo1.ListIndex
               data_reg.Recordset("info_debit") = txt_det.Text
               data_reg.Recordset("cl_nrovend") = txt_base.Text
               data_reg.Recordset("cl_nom_sup") = WElusuario
               data_reg.Recordset("cl_nro_sup") = Check1.Value
               data_reg.Recordset("cl_atrasoa") = Check2.Value
               data_reg.Recordset("cl_atrasop") = Check3.Value
               data_reg.Recordset("cl_cuopaga") = Check7.Value
               data_reg.Recordset("ci_tarj") = 0
               data_reg.Recordset("ultmespmut") = 0
               data_reg.Recordset("ultanopmut") = 0
               data_reg.Recordset("cl_mes_ant") = 0
               If mfh.Text <> "__/__/____" Then
                  data_reg.Recordset("cl_fultmov") = mfh.Text
                  data_reg.Recordset("cl_fax") = mhh.Text
                  data_reg.Recordset("cl_email") = txt_obs.Text
               End If
               data_reg.Recordset("cl_etiquet") = 0
               If Check4.Value = 1 Then
                  MsgBox "Se guardará la conformidad con USUARIO: " & WElusuario
                  data_reg.Recordset("cl_grupo") = Check4.Value
                  data_reg.Recordset("cl_fultpag") = mfconf.Text
                  data_reg.Recordset("cl_codconv") = mhconf.Text
                  data_reg.Recordset("cl_numero") = Combo2.ListIndex
                  data_reg.Recordset("cl_zona") = Combo2.Text
                  data_reg.Recordset("cl_nomcobr") = WElusuario
               Else
                  data_reg.Recordset("cl_grupo") = 0
               End If
               data_reg.Recordset.Update
               If mfh.Text = "__/__/____" Then
                  textocorreo = "Nro. de registro: " & txt_nro & vbCrLf & "Título: " & Combo1.Text & vbCrLf & "Rte: Atención al socio"
               Else
                  textocorreo = "Nro. de registro: " & txt_nro & vbCrLf & "Título: " & Combo1.Text & "--->Ha sido cerrado." & vbCrLf & "Rte: Atención al socio"
                  Cerrado = "S"
               End If
               List1.ListIndex = 0
               For X = 1 To List1.ListCount
                   data_usuar.RecordSource = "Select * from usuarios where nombre ='" & List1.List(List1.ListIndex) & "' and serv_ap in ('S')"
                   data_usuar.Refresh
                   If data_usuar.Recordset.RecordCount > 0 Then
                      data_users.RecordSource = "select * from serap_users"
                      data_users.Refresh
                      data_users.Recordset.AddNew
                      data_users.Recordset("nro_acc") = txt_nro.Text
                      data_users.Recordset("usuario") = data_usuar.Recordset("usuario")
                      data_users.Recordset.Update
                   End If
                   If List1.ListCount - 1 = List1.ListIndex Then
                   Else
                      List1.ListIndex = List1.ListIndex + 1
                   End If
               Next
               data_reg.Refresh
               Data1.Refresh
               XAlta = 0
               Frame1.Enabled = False
               borracua
               habboton
            Else
               MsgBox "El registro no se grabó porque falta fecha y hora"
            End If
          Else
            MsgBox "No ingresó detalle de la solicitud", vbInformation
          End If
       Else
          MsgBox "No seleccionó grupo de solicitud", vbInformation
       End If
    Else
       data_reg.Recordset.Edit
       data_reg.Recordset("cl_descpag") = Mid(Combo1.Text, 1, 100)
       data_reg.Recordset("estado") = Combo1.ListIndex
       data_reg.Recordset("info_debit") = txt_det.Text
       data_reg.Recordset("cl_ter_vto") = 7
       data_reg.Recordset("cl_nrovend") = txt_base.Text
       data_reg.Recordset("cl_nro_sup") = Check1.Value
       data_reg.Recordset("cl_atrasoa") = Check2.Value
       data_reg.Recordset("cl_atrasop") = Check3.Value
       data_reg.Recordset("cl_cuopaga") = Check7.Value
       If mfh.Text <> "__/__/____" Then
          data_reg.Recordset("cl_fultmov") = mfh.Text
          data_reg.Recordset("cl_fax") = mhh.Text
          data_reg.Recordset("cl_email") = txt_obs.Text
       End If
       If Check4.Value = 1 Then
          MsgBox "Se guardará la conformidad con USUARIO: " & WElusuario
          data_reg.Recordset("cl_grupo") = Check4.Value
          data_reg.Recordset("cl_fultpag") = mfconf.Text
          data_reg.Recordset("cl_codconv") = mhconf.Text
          data_reg.Recordset("cl_numero") = Combo2.ListIndex
          data_reg.Recordset("cl_zona") = Combo2.Text
          data_reg.Recordset("cl_nomcobr") = WElusuario
       Else
          data_reg.Recordset("cl_grupo") = 0
          data_reg.Recordset("cl_fultpag") = Null
          data_reg.Recordset("cl_codconv") = Null
          data_reg.Recordset("cl_numero") = Null
          data_reg.Recordset("cl_zona") = Null
          data_reg.Recordset("cl_nomcobr") = Null
       End If
       data_reg.Recordset.Update
       If mfh.Text = "__/__/____" Then
          textocorreo = "Nro. de registro: " & txt_nro & vbCrLf & "Título: " & Combo1.Text & vbCrLf & "Rte: Atención al socio"
       Else
          textocorreo = "Nro. de registro: " & txt_nro & vbCrLf & "Título: " & Combo1.Text & "--->Ha sido cerrado." & vbCrLf & "Rte: Atención al socio"
          Cerrado = "S"
       End If
       
       data_users.RecordSource = "select * from serap_users where nro_acc =" & txt_nro.Text
       data_users.Refresh
       If data_users.Recordset.RecordCount > 0 Then
          data_users.Recordset.MoveFirst
          Do While Not data_users.Recordset.EOF
             data_users.Recordset.Delete
             data_users.Recordset.MoveNext
          Loop
       End If
       List1.ListIndex = 0
       For X = 1 To List1.ListCount
           data_usuar.RecordSource = "Select * from usuarios where nombre ='" & List1.List(List1.ListIndex) & "' and serv_ap in ('S')"
           data_usuar.Refresh
           If data_usuar.Recordset.RecordCount > 0 Then
              data_users.RecordSource = "select * from serap_users"
              data_users.Refresh
              data_users.Recordset.AddNew
              data_users.Recordset("nro_acc") = txt_nro.Text
              data_users.Recordset("usuario") = data_usuar.Recordset("usuario")
              data_users.Recordset.Update
           End If
           If List1.ListCount - 1 = List1.ListIndex Then
           Else
              List1.ListIndex = List1.ListIndex + 1
           End If
       Next
       
       Data1.Refresh
       XAlta = 0
       Frame1.Enabled = False
       habboton
       borracua
    
    End If
'-----
'     .servidor = "smtp.office365.com"
'     .puerto = 25
'     .UseAuntentificacion = True
'     .ssl = True
'     .Usuario = "jefedepartamentoti@sapp.com.uy"
'     .PassWord = "DptotiJunio2021"
'     .de = "jefedepartamentoti@sapp.com.uy"
    
    
'------
    yaesta_Grabado
    If estagrabado.Caption = "0" Then
        EnviarCorreo = MsgBox("Desea enviar correo a los destinatarios?", vbInformation + vbYesNo)
        If EnviarCorreo = vbYes Then
           If Trim(textocorreo) <> "" Then
              frm_servap.MousePointer = 11
              Dim MenCorreo As String
              Dim oMail As Class1
              List1.ListIndex = 0
              For X = 1 To List1.ListCount
                  data_usuar.RecordSource = "Select * from usuarios where nombre ='" & List1.List(List1.ListIndex) & "' and serv_ap in ('S')"
                  data_usuar.Refresh
                  If data_usuar.Recordset.RecordCount > 0 Then
                     If IsNull(data_usuar.Recordset("correo_ap")) = False Then
                        CorreoAP = data_usuar.Recordset("correo_ap")
                     Else
                        CorreoAP = "jdanfer@gmail.com"
                     End If
                  Else
                     CorreoAP = "jdanfer@gmail.com"
                  End If
                  Set oMail = New Class1
                      With oMail
                        .servidor = "smtp.office365.com"
                        .puerto = 25
                        .UseAuntentificacion = True
                        .ssl = True
                        .Usuario = "jefedepartamentoti@sapp.com.uy"
                        .PassWord = "DptotiJunio2021"
                        If Cerrado = "S" Then
                           .Asunto = List1.List(List1.ListIndex) & " Se ha completado un registro de servicios."
                        Else
                           .Asunto = List1.List(List1.ListIndex) & " Se asignó nuevo Servicio de AREA PROTEGIDA"
                        End If
                        .de = "jefedepartamentoti@sapp.com.uy"
                        .para = CorreoAP
            '             .Adjunto = Xarchtex
                        .Mensaje = textocorreo
                        .Enviar_Backup ' manda el mail
                      End With
                     Set oMail = Nothing
                  If List1.ListCount - 1 = List1.ListIndex Then
                  Else
                     List1.ListIndex = List1.ListIndex + 1
                  End If
              Next
              frm_servap.MousePointer = 0
              MsgBox "Correos enviados!", vbInformation
           Else
              MsgBox "No hay texto para enviar correo, verifique si hay datos ingresados.", vbCritical
           End If
        End If
    End If
    
    frm_servap.MousePointer = 0
        
Else
    MsgBox "Debe seleccionar al menos un destinatario para poder GRABAR.", vbCritical
End If

Exit Sub

Vererror:
        If Err.Number = 3150 Then
           MsgBox "No hay datos para grabar, verifique!" & Err.Number, vbCritical
        Else
           MsgBox "No hay datos para grabar, verifique!" & Err.Number, vbCritical
        End If
End Sub

Private Sub b_imp_Click()
frm_infservap.Show vbModal

End Sub

Private Sub b_mod_Click()
'If ControlUsuario("edit_servap") = 1 Then
'   XAlta = 0
'   Frame1.Enabled = True
'   Combo1.SetFocus
'   desboton
'   mfh.Enabled = True
'   mhh.Enabled = True
'   txt_obs.Enabled = True
'Else
'   If ControlUsuario("ver_servap") = 1 Then
'      XAlta = 0
'      Frame1.Enabled = True
'      txt_det.SetFocus
'      desboton
'      b_gra.Enabled = False
'   Else
If txt_usua.Text = WElusuario Then
      XAlta = 0
      Frame1.Enabled = True
      Combo1.SetFocus
      desboton
      mfh.Enabled = True
      mhh.Enabled = True
      txt_obs.Enabled = True
Else
    If ControlUsuario("ver_servap") = 1 Then
       XAlta = 0
       Frame1.Enabled = True
       txt_det.SetFocus
       desboton
       b_gra.Enabled = False
    Else
        MsgBox "No es el usuario creador del registro.", vbCritical
    
    End If
End If
'   End If
'End If

End Sub

Private Sub b_nue_Click()
If ControlUsuario("alta_servap") = 1 Then
    XAlta = 1
    borracua
    txt_nro.Text = ""
    txt_nro.Text = data_numera.Recordset("p_servap") + 1
    data_reg.Recordset.AddNew
    Frame1.Enabled = True
    txt_base.Text = frm_menu.data_parse.Recordset("base")
    'mfd.SetFocus
    mfd.Text = Format(Date, "dd/mm/yyyy")
    mhd.Text = Format(Time, "HH:mm")
    txt_usua.Text = WElusuario
    Check3.Value = 1
    List1.Clear
    Combo1.SetFocus
    data_numera.Recordset.Edit
    data_numera.Recordset("p_servap") = txt_nro.Text
    data_numera.Recordset.Update
    b_imp.Enabled = False
    mfh.Enabled = True
    mhh.Enabled = True
    txt_obs.Enabled = True
    desboton
End If

End Sub

Private Sub cbousuarios_Click()
Dim XX, Xban As Integer
Xban = 0
XX = 0
If List1.ListCount >= 1 Then
   For XX = 1 To List1.ListCount
       List1.ListIndex = XX - 1
       If List1.List(List1.ListIndex) = cbousuarios.Text Then
          Xban = 1
       End If
   Next
Else
   Xban = 0
End If
If Xban = 0 Then
   List1.AddItem cbousuarios.Text
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
   If mfh.Text = "__/__/____" Then
      MsgBox "No está cerrado."
      Check4.Value = 0
   End If
End If

End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
   Data1.RecordSource = "Select * from env_soc where cl_grupo <>" & 1 & " order by cl_fnac DESC"
Else
   Data1.RecordSource = "Select * from env_soc order by cl_fnac DESC"
End If
borracua


End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
   Combo3.SetFocus
Else
   Combo3.ListIndex = -1
   Data1.RecordSource = "Select * from env_soc order by cl_fnac DESC"
   Data1.Refresh
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_det.SetFocus
End If

End Sub

Private Sub Combo3_Click()
If Combo3.Text = "TESORERIA" Then
   Data1.RecordSource = "Select * from env_soc where ci_tarj in (0) and cl_fnac >=#" & Format("01/02/2022", "yyyy/mm/dd") & "# order by cl_fnac DESC"
   Data1.Refresh
Else
   If Combo3.Text = "HABILITACION DT" Then
      Data1.RecordSource = "Select * from env_soc where ultmespmut in (0) and cl_fnac >=#" & Format("01/02/2022", "yyyy/mm/dd") & "# order by cl_fnac DESC"
      Data1.Refresh
   Else
      If Combo3.Text = "PRESUPUESTO" Then
         Data1.RecordSource = "Select * from env_soc where ultanopmut in (0) and cl_fnac >=#" & Format("01/02/2022", "yyyy/mm/dd") & "# order by cl_fnac DESC"
         Data1.Refresh
      Else
         If Combo3.Text = "OTROS" Then
            Data1.RecordSource = "Select * from env_soc where cl_mes_ant in (0) and cl_fnac >=#" & Format("01/02/2022", "yyyy/mm/dd") & "# order by cl_fnac DESC"
            Data1.Refresh
         Else
            Data1.RecordSource = "Select * from env_soc order by cl_fnac DESC"
            Data1.Refresh
         End If
      End If
   End If
End If

End Sub

Private Sub Command1_Click()
Dim Xsi As String
Dim X As Integer
Dim UsuarioAcc As String
UsuarioAcc = ""
Xsi = ""
If txt_nro.Text <> "" Then
    If List1.ListCount > 0 Then
       List1.ListIndex = 0
       For X = 1 To List1.ListCount
           data_usuar.RecordSource = "Select * from usuarios where nombre ='" & List1.List(List1.ListIndex) & "' and serv_ap in ('S')"
           data_usuar.Refresh
           If data_usuar.Recordset.RecordCount > 0 Then
              UsuarioAcc = data_usuar.Recordset("usuario")
              If Trim(WElusuario) = Trim(UsuarioAcc) Then
                 Xsi = "S"
              Else
                 If Trim(Xsi) <> "S" Then
                    Xsi = "N"
                 End If
              End If
           Else
              If Trim(Xsi) <> "S" Then
                 Xsi = "N"
              End If
           End If
           If List1.ListCount - 1 = List1.ListIndex Then
           Else
              List1.ListIndex = List1.ListIndex + 1
           End If
       
       Next
       If Trim(Xsi) = "S" Then
          labsi.Caption = "S"
          frm_servapacc.Show vbModal
       Else
          If ControlUsuario("alta_servap") = 1 Then
             labsi.Caption = "N"
             frm_servapacc.Show vbModal
          Else
             labsi.Caption = "N"
             MsgBox "El registro seleccionado no está compartido con su usuario.", vbCritical
          End If
       End If
    Else
       labsi.Caption = "N"
       MsgBox "No hay destinatarios seleccionados.", vbCritical
    End If
Else
    labsi.Caption = "N"
    MsgBox "Debe seleccionar registro.", vbCritical
End If
End Sub

Private Sub DBGrid1_DblClick()
data_reg.Recordset.FindFirst "cl_codigo =" & Data1.Recordset("cl_codigo")
If Not data_reg.Recordset.NoMatch Then
   borracua
   txt_nro.Text = data_reg.Recordset("cl_codigo")
   igualadat
End If

End Sub

Private Sub Form_Load()

ConectarBD
ConbdSapp.Open
Sqlconsulta = "Select * from estudios where flia =" & 20 & " order by codest"
With Registro1
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Sqlconsulta, ConbdSapp, , , adCmdText
End With
If Registro1.RecordCount > 0 Then
   Registro1.MoveFirst
   Do While Not Registro1.EOF
      Combo1.AddItem Registro1("descrip")
      Registro1.MoveNext
   Loop
End If
Registro1.Close

Sqlconsulta = "Select * from usuarios where serv_ap in ('S') order by nombre"
With Registro1
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Sqlconsulta, ConbdSapp, , , adCmdText
End With
If Registro1.RecordCount > 0 Then
   Registro1.MoveFirst
   Do While Not Registro1.EOF
      cbousuarios.AddItem Registro1("nombre")
      Registro1.MoveNext
   Loop
End If
Registro1.Close

ConbdSapp.Close
'ConbdSapp2.Open
'Sqlconsulta = "Select * from env_soc order by cl_fnac DESC"
'With Registro1
'    .CursorLocation = adUseClient
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open Sqlconsulta, ConbdSapp2, , , adCmdText
'End With

data_reg.DatabaseName = ""
data_reg.Connect = "ODBC;DSN=sappespecial;"
data_reg.RecordSource = "Select * from env_soc order by cl_codigo DESC"
data_reg.Refresh

data_usuar.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_users.Connect = "odbc;dsn=sappespecial;"

Data1.Connect = "ODBC;DSN=sappespecial;"
Data1.RecordSource = "Select * from env_soc order by cl_fnac DESC"
Data1.Refresh
'data_numera.DatabaseName = App.path & "\parse.mdb"
data_numera.Connect = "odbc;dsn=sappnew;"
data_numera.RecordSource = "param_gral"
data_numera.Refresh

igualadat



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

Private Sub List1_DblClick()
List1.RemoveItem List1.ListIndex

End Sub

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

Private Sub mfh_GotFocus()
mfh.Text = Format(Date, "dd/mm/yyyy")

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

Private Sub mhh_GotFocus()
mhh.Text = Format(Time, "HH:mm")

End Sub

Private Sub txt_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_usua.SetFocus
End If

End Sub

Private Sub txt_det_Click()
txt_det.Text = txt_det.Text

End Sub

Private Sub txt_det_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbousuarios.SetFocus
End If

End Sub

Public Function igualadat()

If Data1.Recordset.RecordCount > 0 Then
   If txt_nro.Text <> "" Then
      Data1.Recordset.FindFirst "cl_codigo =" & txt_nro.Text
      If Not Data1.Recordset.NoMatch Then
         List1.Clear
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
         If IsNull(data_reg.Recordset("cl_descpag")) = False Then
        '      data_reg.Recordset("cl_descpag") = Combo1.Text
            Combo1.ListIndex = data_reg.Recordset("estado")
         Else
            Combo1.ListIndex = -1
         End If
         If IsNull(data_reg.Recordset("info_debit")) = False Then
            txt_det.Text = data_reg.Recordset("info_debit")
         Else
            txt_det.Text = ""
         End If
         If IsNull(data_reg.Recordset("cl_cuopaga")) = False Then
            Check7.Value = data_reg.Recordset("cl_cuopaga")
         Else
            Check7.Value = 0
         End If
         If IsNull(data_reg.Recordset("cl_nrovend")) = False Then
            txt_base.Text = data_reg.Recordset("cl_nrovend")
         Else
            txt_base.Text = 0
         End If
         If IsNull(data_reg.Recordset("cl_nom_sup")) = False Then
            txt_usua.Text = data_reg.Recordset("cl_nom_sup")
         Else
            txt_usua.Text = ""
         End If
         If IsNull(data_reg.Recordset("cl_nro_sup")) = False Then
            Check1.Value = data_reg.Recordset("cl_nro_sup")
         Else
            Check1.Value = 0
         End If
         If IsNull(data_reg.Recordset("cl_atrasoa")) = False Then
            Check2.Value = data_reg.Recordset("cl_atrasoa")
         Else
            Check2.Value = 0
         End If
         If IsNull(data_reg.Recordset("cl_atrasop")) = False Then
            Check3.Value = data_reg.Recordset("cl_atrasop")
         Else
            Check3.Value = 0
         End If
         If IsNull(data_reg.Recordset("cl_fultmov")) = False Then
            mfh.Text = Format(data_reg.Recordset("cl_fultmov"), "dd/mm/yyyy")
         Else
            mfh.Text = "__/__/____"
         End If
         If IsNull(data_reg.Recordset("cl_fax")) = False Then
            mhh.Text = data_reg.Recordset("cl_fax")
         Else
            mhh.Text = "__:__"
         End If
         If IsNull(data_reg.Recordset("cl_email")) = False Then
            txt_obs.Text = data_reg.Recordset("cl_email")
         Else
            txt_obs.Text = ""
         End If
         data_users.RecordSource = "select * from serap_users where nro_acc =" & txt_nro.Text
         data_users.Refresh
         If data_users.Recordset.RecordCount > 0 Then
            data_users.Recordset.MoveFirst
            Do While Not data_users.Recordset.EOF
               data_usuar.RecordSource = "select * from usuarios where usuario ='" & data_users.Recordset("usuario") & "'"
               data_usuar.Refresh
               If data_usuar.Recordset.RecordCount > 0 Then
                  List1.AddItem data_usuar.Recordset("nombre")
               End If
               data_users.Recordset.MoveNext
            Loop
         End If
         If IsNull(data_reg.Recordset("cl_grupo")) = False Then
            If data_reg.Recordset("cl_grupo") = 1 Then
               Check4.Value = data_reg.Recordset("cl_grupo")
               If IsNull(data_reg.Recordset("cl_fultpag")) = False Then
                  mfconf.Text = Format(data_reg.Recordset("cl_fultpag"), "dd/mm/yyyy")
               Else
                  mfconf.Text = "__/__/____"
               End If
               If IsNull(data_reg.Recordset("cl_codconv")) = False Then
                  mhconf.Text = data_reg.Recordset("cl_codconv")
               Else
                  mhconf.Text = "__:__"
               End If
               If IsNull(data_reg.Recordset("cl_numero")) = False Then
                  Combo2.ListIndex = data_reg.Recordset("cl_numero")
               Else
                  Combo2.ListIndex = -1
               End If
            Else
               Check4.Value = 0
               mfconf.Text = "__/__/____"
               mhconf.Text = "__:__"
               Combo2.ListIndex = -1
            End If
         Else
            Check4.Value = 0
            mfconf.Text = "__/__/____"
            mhconf.Text = "__:__"
            Combo2.ListIndex = -1
         End If
      Else
         borracua
      End If
   Else
      borracua
   End If
Else
   borracua
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
Check2.Value = 0
Check3.Value = 0
Check7.Value = 0
mfh.Text = "__/__/____"
mhh.Text = "__:__"
txt_obs.Text = ""
Check4.Value = 0
mfconf.Text = "__/__/____"
mhconf.Text = "__:__"
Combo2.ListIndex = -1
cbousuarios.ListIndex = -1

End Function

Public Sub yaesta_Grabado()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD

ConbdSapp.Open

If Trim(txt_nro.Text) <> "" Then
   Xsqlpromo = "Select * from env_soc where cl_codigo =" & txt_nro.Text
   With Xrecclii
      .CursorLocation = adUseClient
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   If Xrecclii.RecordCount > 0 Then
      estagrabado.Caption = "0"
   Else
      estagrabado.Caption = "1"
   End If
   Xrecclii.Close
Else
   estagrabado.Caption = "1"
End If

ConbdSapp.Close

End Sub
