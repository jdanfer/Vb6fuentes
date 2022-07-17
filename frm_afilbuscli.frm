VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_afilbuscli 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar socios"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9420
   Icon            =   "frm_afilbuscli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9420
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_deuda 
      Caption         =   "data_deuda"
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8640
      Picture         =   "frm_afilbuscli.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cerrar sin seleccionar"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_afilbuscli.frx":0B14
      Height          =   3135
      Left            =   240
      OleObjectBlob   =   "frm_afilbuscli.frx":0B28
      TabIndex        =   4
      Top             =   600
      Width           =   9015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Picture         =   "frm_afilbuscli.frx":2407
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox t_busca 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_afilbuscli.frx":2991
      Left            =   2280
      List            =   "frm_afilbuscli.frx":299E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Doble click para seleccionar el registro"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frm_afilbuscli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If t_busca.Text <> "" Then
   If Combo1.ListIndex = 0 Then
      Data1.RecordSource = "select * from clientes where cl_apellid >='" & t_busca.Text & "' order by cl_apellid"
      Data1.Refresh
   Else
      If Combo1.ListIndex = 1 Then
         Data1.RecordSource = "select * from clientes where cl_telefon ='" & t_busca.Text & "'"
         Data1.Refresh
      Else
         If Combo1.ListIndex = 2 Then
            Data1.RecordSource = "select * from clientes where cl_dpto ='" & t_busca.Text & "'"
            Data1.Refresh
         End If
      End If
   End If
End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
Dim Xbajasi As Integer
Dim Xfecha_baja As Date
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo, Xantnro As Long
Dim Xcedok, Xcodcedok As Long
Dim XserviciosRest As Integer
XserviciosRest = 0
Xcedok = 0
Xcodcedok = 0

Xfecha_baja = Date - 180
Xbajasi = 0

frm_afilia.labcl_codigo.Caption = Data1.Recordset("cl_codigo")
If IsNull(Data1.Recordset("fecha_baja")) = False Then
   If Format(Data1.Recordset("fecha_baja"), "yyyy/mm/dd") > Format(Xfecha_baja, "yyyy/mm/dd") Then
      Xbajasi = 1
      MsgBox "Socio de baja con fecha menor a 6 meses, NO INCLUYE PAGO DE PROMOCION AL VENDEDOR.", vbCritical
      frm_afilia.labpromocion.Caption = "NO"
   Else
      Xbajasi = 0
   End If
Else
   Xbajasi = 0
End If
data_deuda.RecordSource = "Select * from deudas where cliente =" & Data1.Recordset("cl_codigo") & " and fecha_pago is null and total >" & 0
data_deuda.Refresh
If data_deuda.Recordset.RecordCount > 0 Then
   data_deuda.Recordset.MoveLast
   If data_deuda.Recordset.RecordCount >= 1 Then
      If Xbajasi = 1 Then
         MsgBox "Socio con deuda pendiente de pago y fecha de baja menor a seis meses.", vbCritical
         MsgBox "La afiliación no podrá ser ingresada al padrón hasta que sea autorizada por administración.", vbCritical
         frm_afilia.labauto.Caption = "NO"
      Else
         If data_deuda.Recordset.RecordCount >= 2 Then
            MsgBox "Socio con deuda pendiente de pago.", vbCritical
            MsgBox "La afiliación no podrá ser ingresada al padrón hasta que sea autorizada por administración.", vbCritical
            frm_afilia.labauto.Caption = "NO"
         End If
      End If
   Else
      frm_afilia.labauto.Caption = "SI"
   End If
End If
   
If frm_afilia.labauto.Caption = "SI" Then
   If Data1.Recordset("cl_codconv") = "CCNOS" Or Data1.Recordset("cl_codconv") = "SMIN" Or _
      Data1.Recordset("cl_codconv") = "UNIVS" Or _
      Data1.Recordset("cl_codconv") = "HEVANO" Or _
      Data1.Recordset("cl_codconv") = "GANOS" Or _
      Data1.Recordset("cl_codconv") = "CASANO" Then
      data_deuda.RecordSource = "Select * from linmmdd where cod_cli =" & Data1.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806) and fecha >='" & Format("31/12/2019", "yyyy/mm/dd") & "'"
      data_deuda.Refresh
      If data_deuda.Recordset.RecordCount > 0 Then
      Else
         MsgBox "No tiene realizada la carta mutual, la afiliación deberá ser autorizada por administración.", vbCritical
         frm_afilia.labauto.Caption = "NOC"
      End If
   End If
End If

If IsNull(Data1.Recordset("estado")) = False Then
   If Data1.Recordset("estado") = 1 Then
      data_deuda.RecordSource = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cl_codconv") & "' and cnv_grupo in ('CCOU','UNIVERSAL','SMI','H.EVANGELICO','CASA DE GALICIA')"
      data_deuda.Refresh
      If data_deuda.Recordset.RecordCount > 0 Then
         MsgBox "Categoría sugerida: COMPLEMENTO", vbInformation
      End If
   End If
   If frm_afilia.labauto.Caption <> "NO" Then
      If Data1.Recordset("cl_codconv") = "SEMM1" Then
         frm_afilia.labauto.Caption = "NO"
      End If
   End If
End If
If IsNull(Data1.Recordset("cl_cedula")) = False Then
   If Data1.Recordset("cl_cedula") > 0 Then
      frm_afilia.t_ced.Text = Data1.Recordset("cl_cedula")
      Xcedok = Data1.Recordset("cl_cedula")
   Else
      frm_afilia.t_ced.Text = ""
      Xcedok = 0
   End If
   If IsNull(Data1.Recordset("cl_codced")) = False Then
      frm_afilia.t_codced.Text = Data1.Recordset("cl_codced")
      Xcodcedok = Data1.Recordset("cl_codced")
   Else
      Xcodcedok = 0
   End If

   Xn1 = 2
   Xn2 = 9
   Xn3 = 8
   Xn4 = 7
   Xn5 = 6
   Xn6 = 3
   Xn7 = 4
   Xpond = 10
   If Xcedok > 0 Then
      Xcedtex = Trim(str(Xcedok))
      Xlargo = Len(Xcedtex)
      If Xlargo = 6 Then
         Xcedtex = "0" & Trim(Xcedtex)
      End If
      Xced1 = Val(Mid(Trim(Xcedtex), 1, 1))
      Xced2 = Val(Mid(Xcedtex, 2, 1))
      Xced3 = Val(Mid(Xcedtex, 3, 1))
      Xced4 = Val(Mid(Xcedtex, 4, 1))
      Xced5 = Val(Mid(Xcedtex, 5, 1))
      Xced6 = Val(Mid(Xcedtex, 6, 1))
      Xced7 = Val(Mid(Xcedtex, 7, 1))
      Xced1 = Xced1 * Xn1
      Xced2 = Xced2 * Xn2
      Xced3 = Xced3 * Xn3
      Xced4 = Xced4 * Xn4
      Xced5 = Xced5 * Xn5
      Xced6 = Xced6 * Xn6
      Xced7 = Xced7 * Xn7
      Xtot = Xced1 + Xced2 + Xced3 + Xced4 + Xced5 + Xced6 + Xced7
      If Len(Trim(str(Xtot))) = 1 Then
         Xtottex = "0000" & Trim(str(Xtot))
      End If
      If Len(Trim(str(Xtot))) = 2 Then
         Xtottex = "000" & Trim(str(Xtot))
      End If
      If Len(Trim(str(Xtot))) = 3 Then
         Xtottex = "00" & Trim(str(Xtot))
      End If
      If Len(Trim(str(Xtot))) = 4 Then
         Xtottex = "0" & Trim(str(Xtot))
      End If
      Xtot = Val(Mid(Xtottex, 5, 1))
      If Xtot <> 0 Then
         Xtot = Xpond - Xtot
      Else
         Xtot = 0
      End If
      If Xtot <> Xcodcedok Then
'             MsgBox "Hay un error en la cédula, verifique!", vbCritical
         frm_afilia.t_codced.Text = Xtot
      End If
   End If
Else
   frm_afilia.t_ced.Text = ""
   frm_afilia.t_codced.Text = ""
End If

If IsNull(Data1.Recordset("saldo_chc2")) = False Then
   If Data1.Recordset("saldo_chc2") = 1 Then
      frm_afilia.labauto.Caption = "NO"
      MsgBox "ATENCION!! Socio con servicios restringidos, no se puede afiliar. Consulte con Administración.", vbCritical
      XserviciosRest = 1
   End If
End If

If XserviciosRest <> 1 Then
    If IsNull(Data1.Recordset("cl_codconv")) = False Then
       frm_afilia.labcodconv.Caption = Data1.Recordset("cl_codconv")
       frm_afilia.labnomconv.Caption = Data1.Recordset("cl_nomconv")
    Else
       frm_afilia.labcodconv.Caption = ""
       frm_afilia.labnomconv.Caption = ""
    End If
    If IsNull(Data1.Recordset("cl_apellid")) = False Then
       frm_afilia.t_nom1.Text = Data1.Recordset("cl_apellid")
    Else
       frm_afilia.t_nom1.Text = ""
    End If
    If IsNull(Data1.Recordset("cl_apellid")) = False Then
       frm_afilia.t_ape1.Text = Data1.Recordset("cl_apellid")
    Else
       frm_afilia.t_ape1.Text = ""
    End If
    If IsNull(Data1.Recordset("cl_fnac")) = False Then
       frm_afilia.mfnac.Text = Data1.Recordset("cl_fnac")
    Else
       frm_afilia.mfnac.Text = "__/__/____"
    End If
    If IsNull(Data1.Recordset("cl_sexo")) = False Then
       If Data1.Recordset("cl_sexo") = 1 Then
          frm_afilia.cbosexo.ListIndex = 1
       Else
          If Data1.Recordset("cl_sexo") = 2 Then
             frm_afilia.cbosexo.ListIndex = 0
          Else
             frm_afilia.cbosexo.ListIndex = -1
          End If
       End If
    End If
    If IsNull(Data1.Recordset("cl_telefon")) = False Then
       If Data1.Recordset("cl_telefon") = "NO APLICA" Then
          If Trim(frm_afilia.t_telef.Text) <> "" Then
          Else
             frm_afilia.t_telef.Text = ""
          End If
       Else
          If Trim(frm_afilia.t_telef.Text) <> "" Then
          Else
             frm_afilia.t_telef.Text = Data1.Recordset("cl_telefon")
          End If
       End If
    Else
       If Trim(frm_afilia.t_telef.Text) <> "" Then
       Else
          frm_afilia.t_telef.Text = ""
       End If
    End If
    If IsNull(Data1.Recordset("cl_dpto")) = False Then
       If Data1.Recordset("cl_dpto") = "NO APLICA" Then
          frm_afilia.t_celu.Text = ""
       Else
          frm_afilia.t_celu.Text = Data1.Recordset("cl_dpto")
       End If
    Else
       frm_afilia.t_celu.Text = ""
    End If
    If IsNull(Data1.Recordset("cl_referen")) = False Then
       If Data1.Recordset("cl_referen") = "NO APLICA" Then
          frm_afilia.t_correo.Text = ""
       Else
          frm_afilia.t_correo.Text = Data1.Recordset("cl_referen")
       End If
    Else
       frm_afilia.t_correo.Text = ""
    End If
    If IsNull(Data1.Recordset("cl_direcci")) = False Then
       If Trim(frm_afilia.t_calle.Text) <> "" Then
       Else
          frm_afilia.t_calle.Text = Data1.Recordset("cl_direcci")
       End If
    End If
    frm_afilia.Frame1.Enabled = True
    
    Unload Me
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deuda.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub t_busca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   Command1.SetFocus
   
End If
End Sub
