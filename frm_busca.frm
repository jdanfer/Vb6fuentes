VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_busca 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Búsqueda de Datos"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sin datos de reservas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   8520
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid dbgrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8705
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   7680
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin VB.Data data_cnvmut 
      Caption         =   "data_cnvmut"
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
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
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
      Left            =   120
      Picture         =   "frm_busca.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   6120
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
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
      Height          =   405
      ItemData        =   "frm_busca.frx":058A
      Left            =   1800
      List            =   "frm_busca.frx":05A0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txt_buscacli 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5055
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "BUSCAR...."
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   8280
      Picture         =   "frm_busca.frx":05F2
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1935
   End
End
Attribute VB_Name = "frm_busca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Funcion Api que obtiene información sobre el estado de Red
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

'Constantes para obtener la información
Private Const INTERNET_CONNECTION_MODEM As Long = &H1
Private Const INTERNET_CONNECTION_LAN As Long = &H2
Private Const INTERNET_CONNECTION_PROXY As Long = &H4
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Private dwflags As Long

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_buscacli.SetFocus
End If

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
Dim Xgpoconv As String
Dim Xfecvolver As Date
Dim RsP As New ADODB.Recordset
Dim SqlP As String

On Error GoTo Quepasaal


'   frmabm.data_clientes.Recordset.FindFirst "cl_codigo =" & Data1.Recordset("cl_codigo")
   frmabm.data_clientes.RecordSource = "Select * from clientes where cl_codigo =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0)
   frmabm.data_clientes.Refresh
'   If Not frmabm.data_clientes.Recordset.NoMatch Then
   If frmabm.data_clientes.Recordset.RecordCount > 0 Then
      Consulta_Notas (Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 0)))
    frm_busca.Hide
    If frmabm.data_clientes.Recordset("estado") = 2 Or frmabm.data_clientes.Recordset("estado") = 3 Then
       frmabm.labestado.Caption = "BAJA"
    Else
       If frmabm.data_clientes.Recordset("fecha_baja") <> "" Then
          frmabm.labestado.Caption = "BAJA"
       Else
          frmabm.labestado.Caption = "ACTIVO"
       End If
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_fultvta")) = False Then
       If IsNull(frmabm.data_clientes.Recordset("cl_tipocli")) = False Then
          If frmabm.data_clientes.Recordset("cl_tipocli") = 1 Or frmabm.data_clientes.Recordset("cl_tipocli") = 2 Then
             frmabm.Image1.Visible = True
          Else
             frmabm.Image1.Visible = False
          End If
       Else
          frmabm.Image1.Visible = False
       End If
    Else
       frmabm.Image1.Visible = False
    End If
    If frmabm.Image1.Visible = False Then
       If IsNull(frmabm.data_clientes.Recordset("cl_fultpag")) = False Then
          frmabm.Image1.Visible = True
       End If
    End If
    frmabm.txt_mat.Caption = frmabm.data_clientes.Recordset("cl_codigo")
    If IsNull(frmabm.data_clientes.Recordset("cl_codconv")) = True Then
       MsgBox "Verifique el convenio", vbCritical, "Mensaje"
       frmabm.txt_codcnv.Text = ""
       Xgpoconv = ""
    Else
       frmabm.txt_codcnv.Text = frmabm.data_clientes.Recordset("cl_codconv")
    End If
    data_cnvmut.RecordSource = "Select * from convenio where cnv_codigo ='" & frmabm.txt_codcnv.Text & "'"
    data_cnvmut.Refresh
    If data_cnvmut.Recordset.RecordCount > 0 Then
       If IsNull(data_cnvmut.Recordset("cnv_entre")) = False Then
          If Trim(data_cnvmut.Recordset("cnv_entre")) <> "" Then
             If Val(data_cnvmut.Recordset("cnv_cuenta")) = Val(frmabm.txt_mat.Caption) Then
                frmabm.t_rs.Text = data_cnvmut.Recordset("cnv_entre")
             Else
                frmabm.t_rs.Text = ""
             End If
          Else
             frmabm.t_rs.Text = ""
          End If
       Else
          frmabm.t_rs.Text = ""
       End If
       If IsNull(data_cnvmut.Recordset("cnv_grupo")) = False Then
          If Trim(data_cnvmut.Recordset("cnv_grupo")) <> "" Then
             Xgpoconv = data_cnvmut.Recordset("cnv_grupo")
          Else
             Xgpoconv = ""
          End If
       Else
          Xgpoconv = ""
       End If
    Else
       Xgpoconv = ""
    End If
    data_cnvmut.Recordset.Close
    frmabm.txt_nomcnv.Enabled = True
    If IsNull(frmabm.data_clientes.Recordset("cl_nomconv")) = True Then
       frmabm.txt_nomcnv.Text = ""
    Else
       frmabm.txt_nomcnv.Text = frmabm.data_clientes.Recordset("cl_nomconv")
    End If
    frmabm.txt_nomcnv.Enabled = False
    If IsNull(frmabm.data_clientes.Recordset("cl_apellid")) = False Then
       frmabm.txt_apellid.Text = frmabm.data_clientes.Recordset("cl_apellid")
    Else
       frmabm.txt_apellid.Text = "NN"
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_ruc")) = False Then
       frmabm.t_otrocnv.Text = frmabm.data_clientes.Recordset("cl_ruc")
    Else
       frmabm.t_otrocnv.Text = ""
    End If
     
    If IsNull(frmabm.data_clientes.Recordset("cl_tipoced")) = False Then
       frmabm.cbotipoced.ListIndex = frmabm.data_clientes.Recordset("cl_tipoced")
       If frmabm.data_clientes.Recordset("cl_tipoced") = 0 Then
          frmabm.txt_ced2.Visible = True
          If frmabm.data_clientes.Recordset("cl_cedula") <> "" Then
             frmabm.txt_ced.Text = frmabm.data_clientes.Recordset("cl_cedula")
          Else
             frmabm.txt_ced.Text = ""
          End If
          If frmabm.data_clientes.Recordset("cl_codced") <> "" Then
             frmabm.txt_ced2.Text = frmabm.data_clientes.Recordset("cl_codced")
          Else
             frmabm.txt_ced2.Text = 0
          End If
       Else
          frmabm.txt_ced2.Visible = False
          If frmabm.data_clientes.Recordset("cl_cedula") <> "" Then
             frmabm.txt_ced.Text = frmabm.data_clientes.Recordset("cl_cedula")
          Else
             frmabm.txt_ced.Text = ""
          End If
          frmabm.txt_ced2.Text = 0
       End If
    Else
       frmabm.cbotipoced.ListIndex = 0
       frmabm.txt_ced2.Visible = True
       If frmabm.data_clientes.Recordset("cl_cedula") <> "" Then
          frmabm.txt_ced.Text = frmabm.data_clientes.Recordset("cl_cedula")
       Else
          frmabm.txt_ced.Text = ""
       End If
       If frmabm.data_clientes.Recordset("cl_codced") <> "" Then
          frmabm.txt_ced2.Text = frmabm.data_clientes.Recordset("cl_codced")
       Else
          frmabm.txt_ced2.Text = 0
       End If
    End If
     If IsNull(frmabm.data_clientes.Recordset("cl_fnac")) = False Then
        frmabm.txt_nac.Text = Format(frmabm.data_clientes.Recordset("cl_fnac"), "dd/mm/yyyy")
     Else
        frmabm.txt_nac.Text = "__/__/____"
        frmabm.labedad.Caption = 0
        frmabm.labunie.Caption = 0
        frmabm.labdias.Caption = 0
     End If
     If Not IsDate(frmabm.txt_nac.Text) Then
        '   MsgBox "Digite una fecha válida"
        
     Else
        CalculaEdad (frmabm.txt_nac.Text)
     End If
     If IsNull(frmabm.data_clientes.Recordset("cl_codruta")) = True Then
        frmabm.t_ruta.Text = ""
     Else
        frmabm.t_ruta.Text = frmabm.data_clientes.Recordset("cl_codruta")
     End If
     
     If IsNull(frmabm.data_clientes.Recordset("cl_decuota")) = False Then
        If frmabm.data_clientes.Recordset("cl_decuota") = 1 Then
           frmabm.Option1.Value = True
        Else
           If frmabm.data_clientes.Recordset("cl_decuota") = 2 Then
              frmabm.Option2.Value = True
           Else
              If frmabm.data_clientes.Recordset("cl_decuota") = 3 Then
                 frmabm.Option3.Value = True
              Else
                 If frmabm.data_clientes.Recordset("cl_decuota") = 4 Then
                    frmabm.Option4.Value = True
                 Else
                    If frmabm.data_clientes.Recordset("cl_decuota") = 5 Then
                       frmabm.Option5.Value = True
                    Else
                       frmabm.Option1.Value = False
                       frmabm.Option2.Value = False
                       frmabm.Option3.Value = False
                       frmabm.Option4.Value = False
                       frmabm.Option5.Value = False
                    End If
                 End If
              End If
           End If
        End If
     Else
        frmabm.Option1.Value = False
        frmabm.Option2.Value = False
        frmabm.Option3.Value = False
        frmabm.Option4.Value = False
        frmabm.Option5.Value = False
     
     End If
     If IsNull(frmabm.data_clientes.Recordset("fecha_reac")) = False Then
        frmabm.mfcarta.Text = Format(frmabm.data_clientes.Recordset("fecha_reac"), "dd/mm/yyyy")
     Else
        frmabm.mfcarta.Text = "__/__/____"
     End If
     If IsNull(frmabm.data_clientes.Recordset("saldo_chc2")) = False Then
        frmabm.cbosrv.ListIndex = frmabm.data_clientes.Recordset("saldo_chc2")
     Else
        frmabm.cbosrv.ListIndex = -1
     End If
     If frmabm.data_clientes.Recordset("cl_ultmesp") <> "" Then
        frmabm.labump.Caption = frmabm.data_clientes.Recordset("cl_ultmesp")
     Else
        frmabm.labump.Caption = ""
     End If
     If frmabm.data_clientes.Recordset("cl_ultanop") <> "" Then
        If frmabm.data_clientes.Recordset("cl_ultanop") = 0 Then
           frmabm.labuap.Caption = frmabm.data_clientes.Recordset("cl_ultanop")
        Else
           frmabm.labuap.Caption = "/" + str(frmabm.data_clientes.Recordset("cl_ultanop"))
        End If
     Else
        frmabm.labuap.Caption = ""
     End If
     If frmabm.data_clientes.Recordset("cl_atrasoa") <> "" Then
        frmabm.labatra.Caption = frmabm.data_clientes.Recordset("cl_atrasoa")
     Else
        frmabm.labatra.Caption = ""
     End If
     If frmabm.data_clientes.Recordset("saldo_cc") <> "" Then
        frmabm.labdeudap.Caption = frmabm.data_clientes.Recordset("saldo_cc")
     Else
        frmabm.labdeudap.Caption = ""
     End If
     If frmabm.data_clientes.Recordset("cl_direcci") <> "" Then
        frmabm.txt_direcc1.Text = frmabm.data_clientes.Recordset("cl_direcci")
     Else
        frmabm.txt_direcc1.Text = ""
     End If
     If IsNull(frmabm.data_clientes.Recordset("cl_dpto")) = False Then
        frmabm.t_cel.Text = frmabm.data_clientes.Recordset("cl_dpto")
     Else
        frmabm.t_cel.Text = ""
     End If
     If IsNull(frmabm.data_clientes.Recordset("cl_referen")) = False Then
        frmabm.t_correo.Text = frmabm.data_clientes.Recordset("cl_referen")
     Else
        frmabm.t_correo.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_entre") <> "" Then
        frmabm.txt_direcc2.Text = frmabm.data_clientes.Recordset("cl_entre")
     Else
        frmabm.txt_direcc2.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_grupo") <> "" Then
        frmabm.txt_codzon.Text = frmabm.data_clientes.Recordset("cl_grupo")
     Else
        frmabm.txt_codzon.Text = 0
     End If
     If frmabm.data_clientes.Recordset("cl_zona") <> "" Then
        frmabm.cbolocalid.Text = frmabm.data_clientes.Recordset("cl_zona")
     Else
        frmabm.cbolocalid.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_sexo") = 2 Then
        frmabm.cbosexo.Text = "FEMENINO"
     Else
        frmabm.cbosexo.Text = "MASCULINO"
     End If
    If frmabm.data_clientes.Recordset("cl_telefon") <> "" Then
       frmabm.txt_telef.Text = frmabm.data_clientes.Recordset("cl_telefon")
    Else
       frmabm.txt_telef.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_dircobr") <> "" Then
       frmabm.txt_dircob.Text = frmabm.data_clientes.Recordset("cl_dircobr")
    Else
       frmabm.txt_dircob.Text = ""
    End If
    If IsNull(frmabm.data_clientes.Recordset("cl_nombre")) = False Then
       frmabm.txt_conmut.Text = frmabm.data_clientes.Recordset("cl_nombre")
    Else
       frmabm.txt_conmut.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_socmnom") <> "" Then
       frmabm.cbomutual.Text = frmabm.data_clientes.Recordset("cl_socmnom")
    Else
       frmabm.cbomutual.Text = ""
    End If
    If frmabm.data_clientes.Recordset("cl_nrosocm") <> "" Then
       frmabm.txt_matmut.Text = frmabm.data_clientes.Recordset("cl_nrosocm")
    Else
       frmabm.txt_matmut.Text = ""
    End If
     If frmabm.data_clientes.Recordset("cl_fecing") <> "" Then
        frmabm.txt_fecing.Text = Format(frmabm.data_clientes.Recordset("cl_fecing"), "dd/mm/yyyy")
     Else
        frmabm.txt_fecing.Text = "__/__/____"
     End If
     If frmabm.data_clientes.Recordset("fecha_baja") <> "" Then
        frmabm.txt_fecbaj.Text = Format(frmabm.data_clientes.Recordset("fecha_baja"), "dd/mm/yyyy")
     Else
        frmabm.txt_fecbaj.Text = "__/__/____"
     End If
     If IsNull(frmabm.data_clientes.Recordset("idpromos")) = False Then
        frmabm.labidpromo.Caption = frmabm.data_clientes.Recordset("idpromos")
        If Val(frmabm.labidpromo.Caption) > 0 Then
           BuscaPromosId
        Else
           frmabm.cbopromos.Text = ""
        End If
     Else
        frmabm.labidpromo.Caption = 0
        frmabm.cbopromos.Text = ""
     End If
     If IsNull(frmabm.data_clientes.Recordset("mesproxemi")) = False Then
        frmabm.t_pmemi.Text = frmabm.data_clientes.Recordset("mesproxemi")
        frmabm.t_paemi.Text = frmabm.data_clientes.Recordset("anoproxemi")
     Else
        frmabm.t_pmemi.Text = 0
        frmabm.t_paemi.Text = 0
     End If
     
     If frmabm.data_clientes.Recordset("cl_nrovend") <> "" Then
        frmabm.txt_codpro.Text = frmabm.data_clientes.Recordset("cl_nrovend")
     Else
        frmabm.txt_codpro.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_nomvend") <> "" Then
        frmabm.cbonompro.Text = frmabm.data_clientes.Recordset("cl_nomvend")
     Else
        frmabm.cbonompro.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_nrocobr") <> "" Then
        frmabm.txt_codcob.Text = frmabm.data_clientes.Recordset("cl_nrocobr")
     Else
        frmabm.txt_codcob.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_nomcobr") <> "" Then
        frmabm.cbonomcob.Text = frmabm.data_clientes.Recordset("cl_nomcobr")
     Else
        frmabm.cbonomcob.Text = ""
     End If
     frmabm.Veoladeuda (frmabm.data_clientes.Recordset("cl_codigo"))
          
     If IsNull(frmabm.data_clientes.Recordset("cl_descpag")) = True Then
        frmabm.cbopago.Text = "Abono Mensual"
     Else
        If UCase(frmabm.data_clientes.Recordset("cl_descpag")) = "DEBITO AUTOMATICO" Then
           frmabm.cbopago.Text = "Debito Automatico"
        Else
           frmabm.cbopago.Text = "Abono Mensual"
        End If
     End If
     If frmabm.data_clientes.Recordset("cl_diacobr") <> "" Then
        frmabm.txt_diacob.Text = frmabm.data_clientes.Recordset("cl_diacobr")
     Else
        frmabm.txt_diacob.Text = ""
     End If
     If frmabm.data_clientes.Recordset("tit_tarj") <> "" Then
        frmabm.txt_nomtarj.Text = frmabm.data_clientes.Recordset("tit_tarj")
     Else
        frmabm.txt_nomtarj.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_nrotarj") <> "" Then
        frmabm.txt_nrotarj.Text = frmabm.data_clientes.Recordset("cl_nrotarj")
     Else
        frmabm.txt_nrotarj.Text = ""
     End If
     If frmabm.data_clientes.Recordset("ci_tarj") <> "" Then
        frmabm.txt_cedtarj.Text = frmabm.data_clientes.Recordset("ci_tarj")
     Else
        frmabm.txt_cedtarj.Text = ""
     End If
     If frmabm.data_clientes.Recordset("codcitarj") <> "" Then
        frmabm.txt_codtarj.Text = frmabm.data_clientes.Recordset("codcitarj")
     Else
        frmabm.txt_codtarj.Text = ""
     End If
     
     If frmabm.data_clientes.Recordset("cl_tjemi_c") <> "" Then
        frmabm.txt_codemisor.Text = frmabm.data_clientes.Recordset("cl_tjemi_c")
     Else
        frmabm.txt_codemisor.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_tjemi_n") <> "" Then
        frmabm.cbotarj.Text = frmabm.data_clientes.Recordset("cl_tjemi_n")
     Else
        frmabm.cbotarj.Text = ""
     End If
     If frmabm.data_clientes.Recordset("cl_tj_venc") <> "" Then
        frmabm.txt_vence.Text = Format(frmabm.data_clientes.Recordset("cl_tj_venc"), "dd/mm/yyyy")
     Else
        frmabm.txt_vence.Text = "__/__/____"
     End If
     frmabm.labmr.Caption = ""
     Veoladeuda (DBGrid1.TextMatrix(DBGrid1.RowSel, 0))
     If frmabm.cbopromos.Text = "Grupo de 3 o más" Then
        VerPromoCliNew
     Else
        VerPromocion (DBGrid1.TextMatrix(DBGrid1.RowSel, 0))
     End If
     
     Dim Xtex As String
     Dim Xqfees As Date
     Xqfees = Date + 1
     Xtex = ""
     If Check1.Value = 1 Then
        frmabm.labavis.Caption = ""
     Else
        If frmabm.txt_mat.Caption <> "" Then
           If CBool(Online()) = True Then
              frmabm.data_fecped.Connect = "ODBC;DSN=" & Xconexrmt & ";"
              frmabm.data_fecped.RecordSource = "Select * from t_fechas where mat_pac =" & frmabm.txt_mat.Caption & " and cdate(fecha) >=#" & Format(Xqfees, "yyyy/mm/dd") & "#"
              frmabm.data_fecped.Refresh
              If frmabm.data_fecped.Recordset.RecordCount > 0 Then
                 Do While Not frmabm.data_fecped.Recordset.EOF
                    If Xtex = "" Then
                       Xtex = "Anotado para: " & frmabm.data_fecped.Recordset("especial") & " DIA:" & frmabm.data_fecped.Recordset("fecha") & " H." & frmabm.data_fecped.Recordset("hora") & " BASE:" & frmabm.data_fecped.Recordset("base")
                    Else
                       Xtex = Xtex & chr(13) & "Anotado para: " & frmabm.data_fecped.Recordset("especial") & " DIA:" & frmabm.data_fecped.Recordset("fecha") & " H." & frmabm.data_fecped.Recordset("hora") & " BASE:" & frmabm.data_fecped.Recordset("base")
                    End If
                    frmabm.data_fecped.Recordset.MoveNext
                 Loop
              End If
              If Xtex <> "" Then
                 frmabm.labavis.Caption = Xtex
              Else
                 frmabm.labavis.Caption = ""
              End If
           Else
              frmabm.labavis.Caption = ""
           End If
        End If
        Dim Xelaviso As String
        Xgpoconv = ""
        If Xgpoconv <> "" Then
           Xelaviso = frmabm.labavis.Caption
           frmabm.labavis.Caption = ""
           If Xgpoconv = "CCOU" Or Xgpoconv = "SMI" Or Xgpoconv = "UNIVERSAL" Or _
              Xgpoconv = "H.EVANGELICO" Or Xgpoconv = "CASA DE GALICIA" Then
              ConectarBD
              ConbdSapp.Open
              SqlP = "Select * from prestamo where nom1 ='" & Trim(str(frmabm.txt_mat.Caption)) & "' and nomc ='" & "MEDICO DE REFERENCIA" & "' order by fecing DESC"
              With RsP
                 .CursorLocation = adUseClient
                 .CursorType = adOpenKeyset
                 .LockType = adLockOptimistic
                 .Open SqlP, ConbdSapp, , , adCmdText
              End With
              If RsP.RecordCount > 0 Then
                 RsP.MoveFirst
                 frmabm.labmr.Caption = "Med.Ref:" & RsP("desccar") & " FECHA:" & RsP("fecing")
              End If
              ConbdSapp.Close
              If Val(frmabm.labedad.Caption) = 0 Then
                 If Val(frmabm.labunie.Caption) = 0 Then
                    If Val(frmabm.labdias.Caption) > 0 And Val(frmabm.labdias.Caption) <= 10 Then
                       data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190001
                       data_lin.Refresh
                       If data_lin.Recordset.RecordCount > 0 Then
                          frmabm.labavis.Caption = "METAS: " & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " " & data_lin.Recordset("base")
                       Else
                          frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CAPTACIÓN RECIÉN NACIDO-"
                       End If
                       data_lin.Recordset.Close
                    End If
                 Else
                    If Val(frmabm.labunie.Caption) <= 11 Then
                       data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190003 & " order by fecha"
                       data_lin.Refresh
                       If data_lin.Recordset.RecordCount > 0 Then
                          data_lin.Recordset.MoveFirst
                          frmabm.labavis.Caption = "METAS: "
                          Do While Not data_lin.Recordset.EOF
                             frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " " & data_lin.Recordset("base")
                             data_lin.Recordset.MoveNext
                          Loop
                          data_lin.Recordset.MovePrevious
                          Xfecvolver = data_lin.Recordset("fecha") + 36
                          frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                       Else
                          frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.1ER.AÑO DE VIDA-"
                       End If
                       data_lin.Recordset.Close
                    End If
                 End If
              Else
                 If Val(frmabm.labedad.Caption) = 1 And Val(frmabm.labunie.Caption) >= 0 Then
                    data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190003 & " order by fecha"
                    data_lin.Refresh
                    If data_lin.Recordset.RecordCount > 0 Then
                       data_lin.Recordset.MoveFirst
                       frmabm.labavis.Caption = "METAS: "
                       Do While Not data_lin.Recordset.EOF
                          frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                          data_lin.Recordset.MoveNext
                       Loop
                       data_lin.Recordset.MovePrevious
                       Xfecvolver = data_lin.Recordset("fecha") + 36
                       frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                    Else
                       frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.1er.AÑO DE VIDA-"
                    End If
                 Else
                    If Val(frmabm.labedad.Caption) = 2 And Val(frmabm.labunie.Caption) >= 0 Then
                       data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190004 & " order by fecha"
                       data_lin.Refresh
                       If data_lin.Recordset.RecordCount > 0 Then
                          data_lin.Recordset.MoveFirst
                          frmabm.labavis.Caption = "METAS: "
                          Do While Not data_lin.Recordset.EOF
                             frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                             data_lin.Recordset.MoveNext
                          Loop
                          data_lin.Recordset.MovePrevious
                          Xfecvolver = data_lin.Recordset("fecha") + 91
                          frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                       Else
                          frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.2do.AÑO DE VIDA-"
                       End If
                       data_lin.Recordset.Close
                    Else
                       'cambiar el codigo de facturación
                       If Val(frmabm.labedad.Caption) = 3 And Val(frmabm.labunie.Caption) >= 0 Then
                          data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190004 & " order by fecha"
                          data_lin.Refresh
                          If data_lin.Recordset.RecordCount > 0 Then
                             data_lin.Recordset.MoveFirst
                             frmabm.labavis.Caption = "METAS: "
                             Do While Not data_lin.Recordset.EOF
                                frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                data_lin.Recordset.MoveNext
                             Loop
                             data_lin.Recordset.MovePrevious
                             Xfecvolver = data_lin.Recordset("fecha") + 122
                             frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                          Else
                             frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.3er.AÑO DE VIDA-"
                          End If
                          data_lin.Recordset.Close
                       Else
                          'cambiar el codigo de facturación
                          If Val(frmabm.labedad.Caption) = 4 And Val(frmabm.labunie.Caption) >= 0 Then
                             data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190030 & " order by fecha"
                             data_lin.Refresh
                             If data_lin.Recordset.RecordCount > 0 Then
                                data_lin.Recordset.MoveFirst
                                frmabm.labavis.Caption = "METAS: "
                                Do While Not data_lin.Recordset.EOF
                                   frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                   data_lin.Recordset.MoveNext
                                Loop
                                data_lin.Recordset.MovePrevious
                                Xfecvolver = data_lin.Recordset("fecha") + 183
                                frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                             Else
                                frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.4to.AÑO DE VIDA-"
                             End If
                             data_lin.Recordset.Close
                          Else
                             'cambiar el codigo de facturación
                             If Val(frmabm.labedad.Caption) = 5 And Val(frmabm.labunie.Caption) >= 0 Then
                                data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190031 & " order by fecha"
                                data_lin.Refresh
                                If data_lin.Recordset.RecordCount > 0 Then
                                   data_lin.Recordset.MoveFirst
                                   frmabm.labavis.Caption = "METAS: "
                                   Do While Not data_lin.Recordset.EOF
                                      frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                      data_lin.Recordset.MoveNext
                                   Loop
                                   data_lin.Recordset.MovePrevious
                                   Xfecvolver = data_lin.Recordset("fecha") + 183
                                   frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                                Else
                                   frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -CTRL.5to.AÑO DE VIDA-"
                                End If
                                data_lin.Recordset.Close
                             Else
                                If Val(frmabm.labedad.Caption) >= 15 And Val(frmabm.labedad.Caption) <= 100 Then
                                   If frmabm.cbosexo.Text = "FEMENINO" Then
                                      data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod =" & 190010 & " order by fecha"
                                      data_lin.Refresh
                                      If data_lin.Recordset.RecordCount > 0 Then
                                         data_lin.Recordset.MoveFirst
                                         frmabm.labavis.Caption = "METAS: "
                                         Do While Not data_lin.Recordset.EOF
                                            frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                            data_lin.Recordset.MoveNext
                                         Loop
                                         data_lin.Recordset.MovePrevious
                                         Xfecvolver = data_lin.Recordset("fecha") + 365
                                         frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                                      Else
                                         frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 1 -PESQUISA V.DOMESTICA-"
                                      End If
                                      data_lin.Recordset.Close
                                   End If
                                Else
                                   If Val(frmabm.labedad.Caption) >= 12 And Val(frmabm.labedad.Caption) <= 19 Then
                                      data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod in (190011,190012)"
                                      data_lin.Refresh
                                      If data_lin.Recordset.RecordCount > 0 Then
                                         data_lin.Recordset.MoveFirst
                                         frmabm.labavis.Caption = "METAS: "
                                         Do While Not data_lin.Recordset.EOF
                                            frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                            data_lin.Recordset.MoveNext
                                         Loop
                                      Else
                                         frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 2 -MÉDICO DE REF. 12 A 19AÑOS"
                                      End If
                                      data_lin.Recordset.Close
                                   Else
                                      If Val(frmabm.labedad.Caption) >= 45 And Val(frmabm.labedad.Caption) <= 64 Then
                                         data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod in (190013,190014)"
                                         data_lin.Refresh
                                         If data_lin.Recordset.RecordCount > 0 Then
                                            data_lin.Recordset.MoveFirst
                                            frmabm.labavis.Caption = "METAS: "
                                            Do While Not data_lin.Recordset.EOF
                                               frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                               data_lin.Recordset.MoveNext
                                            Loop
                                         Else
                                            frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 2 -MÉDICO DE REF. 45 A 64AÑOS"
                                         End If
                                         data_lin.Recordset.Close
                                         If Val(frmabm.labedad.Caption) > 50 Then
                                            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod in (30063,30067)"
                                            data_lin.Refresh
                                            If data_lin.Recordset.RecordCount > 0 Then
                                               data_lin.Recordset.MoveFirst
                                               Do While Not data_lin.Recordset.EOF
                                                  frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                                  data_lin.Recordset.MoveNext
                                               Loop
                                            Else
                                               frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "FALTA FECATEST"
                                            End If
                                            data_lin.Recordset.Close
                                         End If
                                      Else
                                         If Val(frmabm.labedad.Caption) >= 65 And Val(frmabm.labedad.Caption) <= 74 Then
                                            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod in (190018,190019,190023,190027) order by fecha"
                                            data_lin.Refresh
                                            If data_lin.Recordset.RecordCount > 0 Then
                                               data_lin.Recordset.MoveFirst
                                               frmabm.labavis.Caption = "METAS: "
                                               Do While Not data_lin.Recordset.EOF
                                                  frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                                  data_lin.Recordset.MoveNext
                                               Loop
                                               data_lin.Recordset.MovePrevious
                                               Xfecvolver = data_lin.Recordset("fecha") + 275
                                               frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                                            Else
                                               frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 3 -MÉDICO DE REF. 65 A 74AÑOS"
                                            End If
                                            data_lin.Recordset.Close
                                         Else
                                            If Val(frmabm.labedad.Caption) >= 75 And Val(frmabm.labedad.Caption) <= 115 Then
                                               data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.txt_mat.Caption & " and cod_prod in (190020,190021) order by fecha"
                                               data_lin.Refresh
                                               If data_lin.Recordset.RecordCount > 0 Then
                                                  data_lin.Recordset.MoveFirst
                                                  frmabm.labavis.Caption = "METAS: "
                                                  Do While Not data_lin.Recordset.EOF
                                                     frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & data_lin.Recordset("fecha") & " " & data_lin.Recordset("nom_prod") & " Base:" & data_lin.Recordset("base")
                                                     data_lin.Recordset.MoveNext
                                                  Loop
                                                  data_lin.Recordset.MovePrevious
                                                  Xfecvolver = data_lin.Recordset("fecha") + 91
                                                  frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & "PROX.CONSULTA:" & Xfecvolver
                                               Else
                                                  frmabm.labavis.Caption = "METAS: " & "DEBE REALIZAR META 3 -MÉDICO DE REF. >75 AÑOS"
                                               End If
                                               data_lin.Recordset.Close
                                            Else
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                 End If
              End If
           Else
            
           End If
           If Trim(frmabm.labavis.Caption) = "" Then
              frmabm.labavis.Caption = Xelaviso
           Else
              frmabm.labavis.Caption = frmabm.labavis.Caption & vbCrLf & Xelaviso
           End If
        End If
    End If
    
   Else
     MsgBox "Atención!!! ERROR en la búsqueda", vbCritical, "Búsqueda..."
     txt_buscacli.SetFocus
   End If
Exit Sub

Quepasaal:
          If Err.Number = 3157 Then
             frmabm.MousePointer = 0
             MsgBox "Error al buscar"
          Else
             frmabm.MousePointer = 0
             MsgBox "Error al conectar"
          End If
          
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Deactivate()
'frm_busca.Hide

End Sub

Private Sub Form_Initialize()
'Data1.Recordset.MoveLast

End Sub

Private Sub Form_Load()

'Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.RecordSource = "select top 500, * from clientes"
'Data1.Refresh
Data1.ConnectionString = "dsn=" & Xconexrmt
Combo1.ListIndex = 0
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cnvmut.Connect = "odbc;dsn=" & Xconexrmt & ";"
DBGrid1.rows = 2
DBGrid1.Cols = 7
DBGrid1.TextMatrix(0, 0) = "MATRICULA"
DBGrid1.ColWidth(0) = 1500
DBGrid1.TextMatrix(0, 1) = "NOMBRE"
DBGrid1.ColWidth(1) = 4500
DBGrid1.TextMatrix(0, 2) = "DIRECCION"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "CEDULA"
DBGrid1.ColWidth(3) = 1500
DBGrid1.TextMatrix(0, 4) = "CONVENIO"
DBGrid1.ColWidth(4) = 1200
DBGrid1.TextMatrix(0, 5) = "TELEFONO"
DBGrid1.ColWidth(5) = 1500
DBGrid1.TextMatrix(0, 6) = "CELULAR"
DBGrid1.ColWidth(6) = 1500

End Sub

Private Sub Form_Resize()
With Image1
     .Top = 0
     .Left = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub txt_buscacli_Change()
''''     Data1.RecordSource = "select top 500, * from clientes where cl_apellid like '" & data & "*' order by cl_apellid "

End Sub

Private Sub txt_buscacli_KeyPress(KeyAscii As Integer)
Dim Xcann As Integer
Xcann = 1
KeyAscii = Asc(UCase(chr(KeyAscii)))

If KeyAscii = 13 Then
   DBGrid1.Clear
   DBGrid1.rows = 2
    DBGrid1.TextMatrix(0, 0) = "MATRICULA"
    DBGrid1.ColWidth(0) = 1500
    DBGrid1.TextMatrix(0, 1) = "NOMBRE"
    DBGrid1.ColWidth(1) = 4500
    DBGrid1.TextMatrix(0, 2) = "DIRECCION"
    DBGrid1.ColWidth(2) = 2900
    DBGrid1.TextMatrix(0, 3) = "CEDULA"
    DBGrid1.ColWidth(3) = 1500
    DBGrid1.TextMatrix(0, 4) = "CONVENIO"
    DBGrid1.ColWidth(4) = 1200
    DBGrid1.TextMatrix(0, 5) = "TELEFONO"
    DBGrid1.ColWidth(5) = 1500
    DBGrid1.TextMatrix(0, 6) = "CELULAR"
    DBGrid1.ColWidth(6) = 1500
    
    If Combo1.ListIndex = 0 Then
       Data1.RecordSource = "select cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv,cl_telefon,cl_dpto from clientes where cl_apellid >='" & txt_buscacli.Text & "' limit 3000"
'       Data1.RecordSource = "select top 50, cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv from clientes where cl_apellid like  '%" & txt_buscacli.Text & "%' order by cl_apellid"
       Data1.Refresh
    Else
       If Combo1.ListIndex = 1 Then
          Data1.RecordSource = "select cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv,cl_telefon,cl_dpto from clientes where cl_cedula =" & Val(txt_buscacli.Text)
          Data1.Refresh
          If Data1.Recordset.RecordCount > 0 Then
          Else
             MsgBox "No se encontró este número de cédula, se mostrará el más cercano", vbExclamation, "SAPP"
             Data1.RecordSource = "select cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv,cl_telefon,cl_dpto from clientes where cl_cedula >=" & Val(txt_buscacli.Text) & " order by cl_cedula limit 20"
             Data1.Refresh
          End If
       Else
          If Combo1.ListIndex = 2 Then
             Data1.RecordSource = "select cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv,cl_telefon,cl_dpto from clientes where cl_codigo =" & Val(txt_buscacli.Text)
             Data1.Refresh
          Else
             If Combo1.ListIndex = 3 Then
                Data1.RecordSource = "select cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv,cl_telefon,cl_dpto from clientes where cl_codruta >=" & Val(txt_buscacli.Text)
                Data1.Refresh
             Else
                If Combo1.ListIndex = 4 Then
                   Data1.RecordSource = "select cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv,cl_telefon,cl_dpto from clientes where cl_telefon ='" & txt_buscacli.Text & "'"
                   Data1.Refresh
                Else
                   Data1.RecordSource = "select cl_codigo,cl_apellid,cl_direcci,cl_cedula,cl_codconv,cl_telefon,cl_dpto from clientes where cl_dpto ='" & txt_buscacli.Text & "'"
                   Data1.Refresh
                End If
             End If
          End If
      End If
    End If
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("cl_codigo")
         If IsNull(Data1.Recordset("cl_apellid")) = False Then
            DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("cl_apellid")
         End If
         If IsNull(Data1.Recordset("cl_direcci")) = False Then
            DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("cl_direcci")
         End If
         If IsNull(Data1.Recordset("cl_cedula")) = False Then
            DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("cl_cedula")
         End If
         If IsNull(Data1.Recordset("cl_codconv")) = False Then
            DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("cl_codconv")
         End If
         If IsNull(Data1.Recordset("cl_telefon")) = False Then
            DBGrid1.TextMatrix(Xcann, 5) = Data1.Recordset("cl_telefon")
         End If
         If IsNull(Data1.Recordset("cl_dpto")) = False Then
            DBGrid1.TextMatrix(Xcann, 6) = Data1.Recordset("cl_dpto")
         End If
         
         DBGrid1.rows = DBGrid1.rows + 1
         Data1.Recordset.MoveNext
         Xcann = Xcann + 1
      Loop
   End If
   DBGrid1.SetFocus
End If


End Sub

Private Sub CalculaEdad(ByVal FNaci As Date)
Dim FAct As String
Dim Anios As String
Dim Meses As String
Dim Dias As String
Dim newday As String
Dim newmonth As String
Dim newyear As String

FAct = Format(Now, "dd/MM/yyyy")
FNaci = Format(FNaci, "dd/MM/yyyy")

'Calcula los años
Anios = DateDiff("yyyy", CDate(Format(FNaci, "dd/MM/yyyy")), CDate(FAct))
'Si el mes actual es menor que el mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) < Month(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 newmonth = Month(CDate(FAct)) + 12
 Else
 'Deje el mes actual tal y como estan
 newmonth = Month(CDate(FAct))
 End If

 'Si el mes actual es igual al mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) = Month(CDate(FNaci)) Then
 'Si el día de la fecha actual es menor al día de la fecha de nacimiento
 If Day(CDate(FAct)) < Day(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 End If
End If

If Day(CDate(FAct)) < Day(CDate(FNaci)) Then

   If Month(FNaci) = 1 Or Month(FNaci) = 3 Or Month(FNaci) = 5 Or _
      Month(FNaci) = 7 Or Month(FNaci) = 8 Or Month(FNaci) = 10 Or _
      Month(FNaci) = 12 Then
      newday = Day(CDate(FAct)) + 31
   Else
      If Month(FNaci) = 2 Then
         newday = Day(CDate(FAct)) + 28
      Else
         newday = Day(CDate(FAct)) + 30
      End If
   End If
   newmonth = newmonth - 1
Else
   newday = Day(CDate(FAct))
End If

If Month(CDate(FNaci)) = Month(Date) Then
   
   Meses = 0
Else
   Meses = newmonth - Month(CDate(FNaci))
End If

If Meses < 0 And Anios = 0 Then
   Meses = Meses + 12
End If

Dias = newday - Day(CDate(FNaci))

If FNaci <= FAct Then

'Me.TextBox3.Text = Anios & " Años, " & Meses & " Meses, " & Dias & " Dias."
   frmabm.labedad.Caption = Anios
   If Month(Date) = Month(FNaci) Then
      If Day(Date) > Day(FNaci) Then
         Meses = Meses
      Else
         If Day(Date) = Day(FNaci) Then
            Meses = 0
         Else
            Meses = 11
         End If
      End If
   End If
   frmabm.labunie.Caption = Meses
   frmabm.labdias.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   frmabm.labedad.Caption = 0
   frmabm.labunie.Caption = 0
   frmabm.labdias.Caption = 0
End If

End Sub
Private Function Lan() As Boolean
   
   Call InternetGetConnectedState(dwflags, 0&)
   Lan = dwflags And INTERNET_CONNECTION_LAN
End Function
Private Function Online() As Boolean
   Online = InternetGetConnectedState(0&, 0&)
End Function

Public Sub Veoladeuda(ByVal Xmatricula As Long)

Dim Xsubt As Double
Dim Xcant As Long
Dim Xmes, Xano As Integer
Dim Xsqldeuda As String
Dim Xrecdeuda As New ADODB.Recordset
Xcant = 0
Xsubt = 0
Xmes = 0
Xano = 0
ConectarBDDeuda
ConbdSappDeu.Open
             
Xsqldeuda = "Select * from deudas where cliente =" & Xmatricula & " and fecha_pago is null order by ano,mes"
With Xrecdeuda
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqldeuda, ConbdSappDeu, , , adCmdText
End With

'Set Basesapp = Sesionsapp.OpenDatabase(App.Path & "\sapp.mdb")

'Set Recdeudas = Basesapp.OpenRecordset("Select * from deudas where cliente =" & Xmatricula & " and fecha_pago is null order by ano,mes")

If Xrecdeuda.RecordCount > 0 Then
   Xrecdeuda.MoveFirst
   Do While Not Xrecdeuda.EOF
      If Xrecdeuda("mes") = 0 Then
         Xsubt = Xsubt + Xrecdeuda("total")
      Else
         Xsubt = Xsubt + Xrecdeuda("total")
         If Xmes = 0 Then
            Xmes = Xrecdeuda("mes")
            Xano = Xrecdeuda("ano")
         End If
         Xcant = Xcant + 1
      End If
      Xrecdeuda.MoveNext
   Loop
   frmabm.labump.Caption = Xmes
   frmabm.labuap.Caption = Xano
   frmabm.labatra.Caption = Xcant
   frmabm.labdeudap.Caption = Format(Xsubt, "0.00")
   Xrecdeuda.Close
   
   Xsqldeuda = "Select * from deudas where cliente =" & Xmatricula & " and mes not in (0) and fecha_pago is not null order by fecha DESC limit 1"
   With Xrecdeuda
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqldeuda, ConbdSappDeu, , , adCmdText
   End With
   If Xrecdeuda.RecordCount > 0 Then
      Xmes = Xrecdeuda("mes")
      Xano = Xrecdeuda("ano")
      frmabm.labump.Caption = Xmes
      frmabm.labuap.Caption = Xano
   Else
      Xmes = 0
      Xano = 0
      frmabm.labump.Caption = Xmes
      frmabm.labuap.Caption = Xano
   End If
   Xrecdeuda.Close
   ConbdSappDeu.Close
Else
   frmabm.labatra.Caption = 0
   frmabm.labdeudap.Caption = 0
   Xrecdeuda.Close
   Xsqldeuda = "Select * from deudas where cliente =" & Xmatricula & " and mes not in (0) and fecha_pago is not null order by fecha DESC limit 1"
   With Xrecdeuda
       .CursorLocation = adUseClient
       .CursorType = adOpenKeyset
       .LockType = adLockOptimistic
       .Open Xsqldeuda, ConbdSappDeu, , , adCmdText
   End With
   If Xrecdeuda.RecordCount > 0 Then
      Xmes = Xrecdeuda("mes")
      Xano = Xrecdeuda("ano")
      frmabm.labump.Caption = Xmes
      frmabm.labuap.Caption = Xano
   Else
      Xmes = 0
      Xano = 0
      frmabm.labump.Caption = Xmes
      frmabm.labuap.Caption = Xano
   End If
   Xrecdeuda.Close
   ConbdSappDeu.Close
End If


End Sub

Public Sub VerPromocion(ByVal Xmatricula As Long)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & Xmatricula
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   frmabm.Label42.Caption = "Tiene promo X" & Xrecclii.RecordCount
   frmabm.t_ruta.Enabled = False
Else
   frmabm.Label42.Caption = ""
   frmabm.t_ruta.Enabled = True
End If
Xrecclii.Close
ConbdSapp.Close


End Sub
Public Sub BuscaPromosId()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from promocion_gpo where id =" & Val(frmabm.labidpromo.Caption)
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   frmabm.cbopromos.Text = Xrecclii("descrip")
Else
   MsgBox "No se encuentra promoción. Verifique!", vbCritical
   frmabm.cbopromos.Text = ""
   frmabm.labidpromo.Caption = 0
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub VerPromoCliNew()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim cedruta As String
If frmabm.txt_ced.Text <> "" Then
   cedruta = Trim(frmabm.txt_ced.Text) & Trim(frmabm.txt_ced2.Text)
Else
   cedruta = "0"
End If
ConectarBD
ConbdSapp.Open
If frmabm.t_ruta.Text = "" Then
   Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & Val(cedruta)
Else
   Xsqlpromo = "Select cl_codruta,cl_codigo from clientes where cl_codruta =" & frmabm.t_ruta.Text
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   frmabm.Label42.Caption = "Grupo: " & Xrecclii.RecordCount + 1
   frmabm.t_ruta.Enabled = False
Else
   frmabm.Label42.Caption = ""
   frmabm.t_ruta.Enabled = True
End If
Xrecclii.Close
ConbdSapp.Close

End Sub


Public Function ConectarBDDeuda()
ConbdSappDeu.ConnectionString = "dsn=" & Xconexrmt

End Function

Public Sub Consulta_Notas(ByVal XmatNotas As Long)
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
Xsqlpromo = "Select * from notas_med where matricula =" & XmatNotas
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   frmabm.Image4.Visible = True
   frmabm.Image3.Visible = False
Else
   frmabm.Image3.Visible = True
   frmabm.Image4.Visible = False
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

