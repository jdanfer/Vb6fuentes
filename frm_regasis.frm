VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_regasis 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de solicitud asistencia técnica informática"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_regasis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   9525
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   39
      Top             =   6480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3625
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   5880
      Top             =   5640
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver solo registros sin ingresar conformidad"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   6240
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pendiente"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_regasis.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Informes..."
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton b_bus 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_regasis.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Buscar..."
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton b_eli 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_regasis.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Eliminar registro seleccionado"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_regasis.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cancelar..."
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton b_gra 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_regasis.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Grabar datos"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      Picture         =   "frm_regasis.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Modificar datos"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton b_nue 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_regasis.frx":257E
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Nuevo registro"
      Top             =   5520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de solicitud"
      Enabled         =   0   'False
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.ComboBox Combo3 
         Height          =   360
         ItemData        =   "frm_regasis.frx":2B08
         Left            =   120
         List            =   "frm_regasis.frx":2B12
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Data data_numera 
         Caption         =   "data_numera"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSAdodcLib.Adodc data_reg 
         Height          =   375
         Left            =   360
         Top             =   2280
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "data_reg"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_regasis.frx":2B2B
         Left            =   5760
         List            =   "frm_regasis.frx":2B38
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   4920
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mhconf 
         Height          =   375
         Left            =   4800
         TabIndex        =   34
         Top             =   4920
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
         TabIndex        =   33
         Top             =   4920
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
         Caption         =   "Conformidad de usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   5040
         Width           =   3015
      End
      Begin VB.TextBox txt_obs 
         Enabled         =   0   'False
         Height          =   360
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   23
         Top             =   4320
         Width           =   4935
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   8040
         TabIndex        =   21
         Top             =   3840
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
         TabIndex        =   20
         Top             =   3840
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
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C00000&
         Caption         =   "Prog. SAPP"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   18
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Correo Elect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Teléfono"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txt_usua 
         Height          =   375
         Left            =   6960
         MaxLength       =   25
         TabIndex        =   14
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txt_base 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txt_det 
         BackColor       =   &H00C0FFFF&
         Height          =   1200
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Width           =   7095
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_regasis.frx":2B5F
         Left            =   1800
         List            =   "frm_regasis.frx":2B78
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   3975
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   360
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
         Top             =   360
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
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C00000&
         Caption         =   "Clasificación:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C00000&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   36
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   9240
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "CONTROLES(Área Dpto. de informática):"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3600
         Width           =   8895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   9240
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "Solucionado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Solicitado vía:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Base:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Solicita:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Más detalles:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Asistencia en:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Hora:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Nro."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   6120
      Picture         =   "frm_regasis.frx":2C07
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2175
   End
End
Attribute VB_Name = "frm_regasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_bus_Click()
Dim Xingfec As String
Xingfec = InputBox("Ingrese a partir de que fecha:")
If Xingfec <> "" Then
   If ControlUsuario("Informática") = 1 Then
      Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " And cl_fnac >='" & Format(Xingfec, "yyyy-mm-dd") & "' order by cl_fnac"
      Data1.Refresh
   Else
      Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " And cl_fnac >='" & Format(Xingfec, "yyyy-mm-dd") & "' and cl_nom_sup ='" & WElusuario & "' order by cl_fnac"
      Data1.Refresh
   End If
   DBGrid1.SetFocus
Else
   Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_nom_sup ='" & WElusuario & "' order by cl_fnac"
   Data1.Refresh
   DBGrid1.SetFocus
End If
DBGrid1.Clear
DBGrid1.rows = 2
DBGrid1.Cols = 5
DBGrid1.TextMatrix(0, 0) = "NUMERO"
DBGrid1.ColWidth(0) = 1500
DBGrid1.TextMatrix(0, 1) = "FECHA"
DBGrid1.ColWidth(1) = 1300
DBGrid1.TextMatrix(0, 2) = "MOTIVO"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "USUARIO"
DBGrid1.ColWidth(3) = 1500
DBGrid1.TextMatrix(0, 4) = "OBSERVACION"
DBGrid1.ColWidth(4) = 2500
Dim Xcann As Integer
Xcann = 1
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("cl_codigo")
      If IsNull(Data1.Recordset("cl_fnac")) = False Then
         DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("cl_fnac")
      End If
      If IsNull(Data1.Recordset("cl_descpag")) = False Then
         DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("cl_descpag")
      End If
      If IsNull(Data1.Recordset("cl_nom_sup")) = False Then
         DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("cl_nom_sup")
      End If
      If IsNull(Data1.Recordset("cl_email")) = False Then
         DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("cl_email")
      End If
      DBGrid1.rows = DBGrid1.rows + 1
      Data1.Recordset.MoveNext
      Xcann = Xcann + 1
   Loop
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
If ControlUsuario("Informática") = 1 Then
   Xqueborra = MsgBox("Desea borrar el registro seleccionado?", vbInformation + vbYesNo, "SAPP")
    If Xqueborra = vbYes Then
       data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_codigo =" & txt_nro.Text
       data_reg.Refresh
'       data_reg.RecordSource = "Select * from env_soc"
       If data_reg.Recordset.RecordCount > 0 Then
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
Dim Xcann As Integer
Xcann = 1
If XAlta = 1 Then
   If Combo1.ListIndex >= 0 Then
      If txt_det.Text <> "" Then
        If mfd.Text <> "__/__/____" And mhd.Text <> "__:__" Then
           data_reg.Recordset("cl_codigo") = txt_nro.Text
           data_reg.Recordset("cl_fnac") = mfd.Text
           data_reg.Recordset("cl_ruc") = mhd.Text
           data_reg.Recordset("cl_ter_vto") = 1
           data_reg.Recordset("cl_descpag") = Combo1.Text
           data_reg.Recordset("estado") = Combo1.ListIndex
           data_reg.Recordset("info_debit") = txt_det.Text
           data_reg.Recordset("cl_nrovend") = txt_base.Text
           data_reg.Recordset("cl_nom_sup") = WElusuario
           data_reg.Recordset("cl_nro_sup") = Check1.Value
           data_reg.Recordset("cl_atrasoa") = Check2.Value
           data_reg.Recordset("cl_atrasop") = Check3.Value
           data_reg.Recordset("saldo_cc") = Combo3.ListIndex
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
           data_reg.Refresh
           Data1.Refresh
           XAlta = 0
           Frame1.Enabled = False
           borracua
           habboton
           igualadat
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
   data_reg.Recordset("cl_descpag") = Combo1.Text
   data_reg.Recordset("estado") = Combo1.ListIndex
   data_reg.Recordset("info_debit") = txt_det.Text
   data_reg.Recordset("cl_ter_vto") = 1
   data_reg.Recordset("cl_nrovend") = txt_base.Text
   data_reg.Recordset("cl_nro_sup") = Check1.Value
   data_reg.Recordset("cl_atrasoa") = Check2.Value
   data_reg.Recordset("cl_atrasop") = Check3.Value
   data_reg.Recordset("saldo_cc") = Combo3.ListIndex
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
   XAlta = 0
   Frame1.Enabled = False
   habboton
   borracua
   igualadat
End If

End Sub

Private Sub b_imp_Click()
'If XWeltipoU = "ADMINISTRADOR" Then
If ControlUsuario("Informática") = 1 Then
   frm_infreg.Show vbModal
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_mod_Click()
XAlta = 0
If ControlUsuario("Informática") = 1 Then
    Frame1.Enabled = True
    mfd.SetFocus
    If ControlUsuario("Informática") = 1 Then
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
    If WElusuario = txt_usua.Text Then
       Frame1.Enabled = True
       
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
   
End If
    

End Sub

Private Sub b_nue_Click()
XAlta = 1
borracua
txt_nro.Text = ""
txt_nro.Text = data_numera.Recordset("musada") + 1
data_reg.Recordset.AddNew
Frame1.Enabled = True
txt_base.Text = frm_menu.data_parse.Recordset("base")
'mfd.SetFocus
mfd.Text = Format(Date, "dd/mm/yyyy")
mhd.Text = Format(Time, "HH:mm")
txt_usua.Text = WElusuario
Check3.Value = 1
Combo1.SetFocus
data_numera.Recordset.Edit
data_numera.Recordset("musada") = txt_nro.Text
data_numera.Recordset.Update
If ControlUsuario("Informática") = 1 Then
   b_imp.Enabled = True
   mfh.Enabled = True
   mhh.Enabled = True
   txt_obs.Enabled = True
   Combo3.Enabled = True
Else
   b_imp.Enabled = False
   mfh.Enabled = False
   mhh.Enabled = False
   txt_obs.Enabled = False
   Combo3.Enabled = False
End If
desboton

End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
   If ControlUsuario("Informática") = 1 Then
      Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_grupo <>" & 1 & " order by cl_fnac DESC"
   Else
      Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_grupo <>" & 1 & " and cl_nom_sup ='" & WElusuario & "' order by cl_fnac DESC"
   End If
   Data1.Refresh
   borracua
Else
   If ControlUsuario("Informática") = 1 Then
      Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " order by cl_fnac DESC"
   Else
      Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_nom_sup ='" & WElusuario & "' order by cl_fnac DESC"
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



Private Sub DBGrid1_DblClick()
data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_codigo =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0) & " and cl_ruc ='" & DBGrid1.TextMatrix(DBGrid1.RowSel, 5) & "'"
'data_reg.Recordset.FindFirst "cl_codigo =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0)
data_reg.Refresh
If data_reg.Recordset.RecordCount > 0 Then
   borracua
   igualadat
End If

End Sub

Private Sub Form_Load()
Dim Xdesdef As Date
Xdesdeff = Date - 400
'data_reg.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_reg.ConnectionString = "dsn=" & Xconexrmt
If ControlUsuario("Informática") = 1 Then
   data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_fnac >='" & Format(Xdesdef, "yyyy-mm-dd") & "' order by cl_codigo DESC"
Else
   data_reg.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_fnac >='" & Format(Xdesdef, "yyyy-mm-dd") & "' order by cl_codigo DESC"
End If
data_reg.Refresh
igualadat
Data1.ConnectionString = "dsn=" & Xconexrmt
If ControlUsuario("Informática") = 1 Then
   Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_fnac >='" & Format(Xdesdef, "yyyy-mm-dd") & "' order by cl_fnac DESC"
Else
   Data1.RecordSource = "Select * from env_soc where cl_ter_vto =" & 1 & " and cl_nom_sup ='" & WElusuario & "' and cl_fnac >='" & Format(Xdesdef, "yyyy-mm-dd") & "' order by cl_fnac DESC"
End If
Data1.Refresh
'data_numera.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\parse.mdb"
data_numera.DatabaseName = App.path & "\parse.mdb"
data_numera.RecordSource = "parsec0"
data_numera.Refresh

DBGrid1.rows = 2
DBGrid1.Cols = 6
DBGrid1.TextMatrix(0, 0) = "NUMERO"
DBGrid1.ColWidth(0) = 1500
DBGrid1.TextMatrix(0, 1) = "FECHA"
DBGrid1.ColWidth(1) = 1300
DBGrid1.TextMatrix(0, 2) = "MOTIVO"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "USUARIO"
DBGrid1.ColWidth(3) = 1500
DBGrid1.TextMatrix(0, 4) = "OBSERVACION"
DBGrid1.ColWidth(4) = 2500
DBGrid1.TextMatrix(0, 5) = "HORA"
DBGrid1.ColWidth(5) = 1000

Dim Xcann As Integer
Xcann = 1
If Data1.Recordset.RecordCount > 0 Then
    Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
       DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("cl_codigo")
       If IsNull(Data1.Recordset("cl_fnac")) = False Then
          DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("cl_fnac")
       End If
       If IsNull(Data1.Recordset("cl_descpag")) = False Then
          DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("cl_descpag")
       End If
       If IsNull(Data1.Recordset("cl_nom_sup")) = False Then
          DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("cl_nom_sup")
       End If
       If IsNull(Data1.Recordset("cl_email")) = False Then
          DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("cl_email")
       End If
       If IsNull(Data1.Recordset("cl_ruc")) = False Then
          DBGrid1.TextMatrix(Xcann, 5) = Data1.Recordset("cl_ruc")
       End If
       DBGrid1.rows = DBGrid1.rows + 1
       Data1.Recordset.MoveNext
       Xcann = Xcann + 1
    Loop
End If


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
   If IsNull(data_reg.Recordset("cl_descpag")) = False Then
'      data_reg.Recordset("cl_descpag") = Combo1.Text
      Combo1.ListIndex = data_reg.Recordset("estado")
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
   If IsNull(data_reg.Recordset("cl_nrovend")) = False Then
      txt_base.Text = data_reg.Recordset("cl_nrovend")
   Else
      txt_base.Text = 0
   End If
   Xnumeror = 7
   If IsNull(data_reg.Recordset("cl_nom_sup")) = False Then
      txt_usua.Text = data_reg.Recordset("cl_nom_sup")
   Else
      txt_usua.Text = ""
   End If
   Xnumeror = 8
   If IsNull(data_reg.Recordset("cl_nro_sup")) = False Then
      Check1.Value = data_reg.Recordset("cl_nro_sup")
   Else
      Check1.Value = 0
   End If
   Xnumeror = 9
   If IsNull(data_reg.Recordset("cl_atrasoa")) = False Then
      Check2.Value = data_reg.Recordset("cl_atrasoa")
   Else
      Check2.Value = 0
   End If
   Xnumeror = 10
   If IsNull(data_reg.Recordset("cl_atrasop")) = False Then
      Check3.Value = data_reg.Recordset("cl_atrasop")
   Else
      Check3.Value = 0
   End If
   Xnumeror = 11
   If IsNull(data_reg.Recordset("cl_fultmov")) = False Then
      mfh.Text = Format(data_reg.Recordset("cl_fultmov"), "dd/mm/yyyy")
   Else
      mfh.Text = "__/__/____"
   End If
   Xnumeror = 12
   If IsNull(data_reg.Recordset("cl_fax")) = False Then
      mhh.Text = data_reg.Recordset("cl_fax")
   Else
      mhh.Text = "__:__"
   End If
   Xnumeror = 13
   If IsNull(data_reg.Recordset("saldo_cc")) = False Then
      Combo3.ListIndex = data_reg.Recordset("saldo_cc")
   Else
      Combo3.ListIndex = -1
   End If
   If IsNull(data_reg.Recordset("cl_email")) = False Then
      txt_obs.Text = data_reg.Recordset("cl_email")
   Else
      txt_obs.Text = ""
   End If
   Xnumeror = 14
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
Check2.Value = 0
Check3.Value = 0
mfh.Text = "__/__/____"
mhh.Text = "__:__"
txt_obs.Text = ""
Check4.Value = 0
mfconf.Text = "__/__/____"
mhconf.Text = "__:__"
Combo2.ListIndex = -1
Combo3.ListIndex = -1


End Function

