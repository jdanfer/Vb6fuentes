VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_labo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratorios y comercios"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11730
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_labou.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Importar datos"
      Height          =   375
      Left            =   3360
      Picture         =   "frm_labou.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   "C:\mutuales\cliente.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "'clientes y proveedores tabla Me'$"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Corresponde a ECONOMATO"
      Height          =   255
      Left            =   7920
      TabIndex        =   34
      ToolTipText     =   "Al seleccionar esta opción, el registro podrá ser visto por economato"
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox t_cp 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      TabIndex        =   33
      Top             =   1560
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7680
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      DataSourceName  =   "sapp"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.Data data_otra 
      Caption         =   "data_otra"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_labou.frx":09CC
      Left            =   2160
      List            =   "frm_labou.frx":0A0C
      TabIndex        =   31
      Text            =   "Combo1"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox t_local 
      Height          =   375
      Left            =   8760
      MaxLength       =   100
      TabIndex        =   29
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton b_selec 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Seleccionar"
      Height          =   375
      Left            =   120
      Picture         =   "frm_labou.frx":0ADD
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox t_rut 
      Height          =   360
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   26
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox t_rsoc 
      Height          =   360
      Left            =   2160
      MaxLength       =   25
      TabIndex        =   25
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_labou.frx":1067
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Borrar registro seleccionado"
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton b_grab 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_labou.frx":15F1
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   960
      Picture         =   "frm_labou.frx":1B7B
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton b_nuev 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   120
      Picture         =   "frm_labou.frx":2105
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox t_bus 
      Height          =   360
      Left            =   2520
      TabIndex        =   18
      Top             =   3240
      Width           =   4695
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   6360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Historial de evaluaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Picture         =   "frm_labou.frx":268F
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_labou.frx":2C19
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprimir laboratorios y comercios ingresados"
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      Picture         =   "frm_labou.frx":31A3
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cancelar"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox t_correo 
      Enabled         =   0   'False
      Height          =   360
      Left            =   2160
      MaxLength       =   300
      TabIndex        =   13
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox t_conta 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   11
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox t_direc 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      MaxLength       =   60
      TabIndex        =   9
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox T_TELEF 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      MaxLength       =   250
      TabIndex        =   7
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox T_NOM 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      MaxLength       =   200
      TabIndex        =   5
      Top             =   120
      Width           =   5415
   End
   Begin VB.TextBox T_COD 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   10680
      Picture         =   "frm_labou.frx":372D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_labou.frx":3CB7
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frm_labou.frx":3CCB
      TabIndex        =   0
      Top             =   3600
      Width           =   11415
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FF80&
      Caption         =   "CP:"
      Height          =   375
      Left            =   8160
      TabIndex        =   32
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      X1              =   0
      X2              =   11760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "Departamento:"
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Caption         =   "Localidad:"
      Height          =   375
      Left            =   6840
      TabIndex        =   28
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FF80&
      Caption         =   "RUT:"
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "RAZON SOCIAL:"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Buscar por nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Caption         =   "CORRE@ E."
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "CONTACTO:"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "DIRECCION:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "TELEFONO/S:"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "NOMBRE:"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "CODIGO:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   9000
      Picture         =   "frm_labou.frx":49EE
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2415
   End
End
Attribute VB_Name = "frm_labo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_canc_Click()
t_cod.Text = ""
t_nom.Text = ""
t_direc.Text = ""
t_telef.Text = ""
t_conta.Text = ""
t_correo.Text = ""
t_rut.Text = ""
t_rsoc.Text = ""
t_local.Text = ""
t_cp.Text = ""
Combo1.ListIndex = -1

XAlta = 0
t_cod.Enabled = False
t_nom.Enabled = False
t_direc.Enabled = False
t_telef.Enabled = False
t_conta.Enabled = False
t_correo.Enabled = False
t_rut.Enabled = False
t_rsoc.Enabled = False
t_local.Enabled = False
Combo1.Enabled = False
t_cp.Enabled = False

b_grab.Enabled = False
b_canc.Enabled = False
b_nuev.Enabled = True
b_mod.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command1.Enabled = True


End Sub

Private Sub b_grab_Click()
If XAlta = 1 Then
   Data1.Recordset.AddNew
   If t_local.Text = "" Then
      t_local.Text = "Sin datos"
   End If
   If t_direc.Text = "" Then
      t_direc.Text = "Sin Datos"
   End If
   
   Data1.Recordset("nro") = t_cod.Text
   Data1.Recordset("obsmot") = t_rsoc.Text
   Data1.Recordset("usuario") = t_local.Text
   Data1.Recordset("base") = 0
   Data1.Recordset("obs") = t_direc.Text
   If t_cp.Text = "" Then
      t_cp.Text = "0"
   End If
   Data1.Recordset("accion") = t_cp.Text
'   Data1.Recordset("accion") = t_conta.Text
   If t_correo.Text = "" Then
      t_correo.Text = "S/D"
   End If
'   Data1.Recordset("referen") = t_correo.Text
   Data1.Recordset("referen") = t_telef.Text
   Data1.Recordset("mat") = Combo1.ListIndex
   Data1.Recordset.Update
   Data1.Refresh
'    adous.Recordset.Close
'    adous.Recordset.Open "insert into us (nro,obsmot,usuario,horarec,hora,base,obs,accion,referen,nom_us,cargo,nombre,apellidos,documento,cp,pass_us,pin,correo,acceso,caduca,fec_crea)" & _
'     " values ('" & t_id.Text & "','" & t_nomu.Text & "', '" & Combo1.ListIndex & "','" & T_NOM.Text & "','" & t_apel.Text & "'," & _
'     " '" & t_ced.Text & T_COD.Text & "','" & t_cp.Text & "', AES_ENCRYPT('" & t_pass.Text & "','historiascl'), AES_ENCRYPT('" & t_pin.Text & "','historiascl')," & _
'     " '" & t_correo.Text & "','" & Check1.Value & "','" & Format(Xladatecad, "yyyy-mm-dd") & "','" & Format(Date, "yyyy-mm-dd") & "')"
   
   Data1.Recordset.AddNew
   Data1.Recordset("nro") = t_cod.Text + 1
   Data1.Recordset("base") = 97
   Data1.Recordset("mat") = t_cod.Text
   Data1.Recordset("usuario") = Combo1.Text
   Data1.Recordset("obsmot") = t_rut.Text
   Data1.Recordset("referen") = t_correo.Text
   Data1.Recordset("accion") = t_conta.Text
   Data1.Recordset("obs") = t_nom.Text
   Data1.Recordset.Update
   Data1.Refresh
Else
   Data1.Recordset.Edit
   Data1.Recordset("obsmot") = t_rsoc.Text
   Data1.Recordset("usuario") = t_local.Text
   Data1.Recordset("base") = 0
   Data1.Recordset("obs") = t_direc.Text
   If t_cp.Text = "" Then
      t_cp.Text = "0"
   End If
   Data1.Recordset("accion") = t_cp.Text
'   Data1.Recordset("accion") = t_conta.Text
   If t_correo.Text = "" Then
      t_correo.Text = "S/D"
   End If
'   Data1.Recordset("referen") = t_correo.Text
   Data1.Recordset("referen") = t_telef.Text
   Data1.Recordset("mat") = Combo1.ListIndex
   Data1.Recordset.Update
   Data1.Refresh
   data_otra.RecordSource = "Select * from abmdesp where base =" & 97 & " and mat =" & t_cod.Text
   data_otra.Refresh
   If data_otra.Recordset.RecordCount > 0 Then
        data_otra.Recordset.Edit
        Data1.Recordset("base") = 97
        Data1.Recordset("mat") = t_cod.Text
        Data1.Recordset("usuario") = Combo1.Text
        Data1.Recordset("obsmot") = t_rut.Text
        Data1.Recordset("referen") = t_correo.Text
        Data1.Recordset("accion") = t_conta.Text
        Data1.Recordset("obs") = t_nom.Text
        data_otra.Recordset.Update
        data_otra.Refresh
        Data1.Refresh
   Else
        Data1.RecordSource = "Select * from abmdesp order by nro"
        Data1.Refresh
        Data1.Recordset.MoveLast
        data_otra.Recordset.AddNew
        data_otra.Recordset("nro") = Data1.Recordset("nro") + 1
        Data1.Recordset("base") = 97
        Data1.Recordset("mat") = t_cod.Text
        Data1.Recordset("usuario") = Combo1.Text
        Data1.Recordset("obsmot") = t_rut.Text
        Data1.Recordset("referen") = t_correo.Text
        Data1.Recordset("accion") = t_conta.Text
        Data1.Recordset("obs") = t_nom.Text
        data_otra.Recordset.Update
        data_otra.Refresh
        Data1.RecordSource = "Select * from abmdesp where base <>" & 99 & " and base <>" & 97 & " order by nro"
        Data1.Refresh
   End If
End If

t_cod.Text = ""
t_nom.Text = ""
t_direc.Text = ""
t_telef.Text = ""
t_conta.Text = ""
t_correo.Text = ""
t_rut.Text = ""
t_rsoc.Text = ""
t_local.Text = ""
Combo1.ListIndex = -1
t_cp.Text = ""

XAlta = 0
t_cod.Enabled = False
t_nom.Enabled = False
t_direc.Enabled = False
t_telef.Enabled = False
t_conta.Enabled = False
t_correo.Enabled = False
t_rut.Enabled = False
t_rsoc.Enabled = False
t_local.Enabled = False
Combo1.Enabled = False
t_cp.Enabled = False

b_grab.Enabled = False
b_canc.Enabled = False
b_nuev.Enabled = True
b_mod.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command1.Enabled = True


End Sub

Private Sub b_mod_Click()

t_cod.Enabled = True
t_nom.Enabled = True
t_direc.Enabled = True
t_telef.Enabled = True
t_conta.Enabled = True
t_correo.Enabled = True
t_rut.Enabled = True
t_rsoc.Enabled = True
t_local.Enabled = True
t_cp.Enabled = True
Combo1.Enabled = True

b_grab.Enabled = True
b_canc.Enabled = True
b_nuev.Enabled = False
b_mod.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command1.Enabled = False

End Sub

Private Sub b_nuev_Click()

Data1.RecordSource = "Select * from abmdesp order by nro"
Data1.Refresh
t_cod.Enabled = True
t_nom.Enabled = True
t_direc.Enabled = True
t_telef.Enabled = True
t_conta.Enabled = True
t_correo.Enabled = True
t_rut.Enabled = True
t_rsoc.Enabled = True
t_local.Enabled = True
Combo1.Enabled = True
t_cp.Enabled = True

t_cod.Text = ""
t_nom.Text = ""
t_direc.Text = ""
t_telef.Text = ""
t_conta.Text = ""
t_correo.Text = ""
t_rut.Text = ""
t_rsoc.Text = ""
t_local.Text = ""
Combo1.ListIndex = -1
t_cp.Text = ""

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
   t_cod.Text = Data1.Recordset("nro") + 1
Else
   t_cod.Text = 1
End If
Data1.RecordSource = "Select * from abmdesp where base <>" & 99 & " and base <>" & 97 & " order by nro"
Data1.Refresh

XAlta = 1
t_nom.SetFocus
b_grab.Enabled = True
b_canc.Enabled = True
b_nuev.Enabled = False
b_mod.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command1.Enabled = False

End Sub

Private Sub b_selec_Click()
If t_cod.Text <> "" Then
   If Xestaok = 7 Then ' Tesorería
      frm_teso.t_cli.Text = t_cod.Text
      frm_teso.labnomcl.Caption = t_rsoc.Text
      Unload Me
      Xestaok = 0
   Else
      MsgBox "Opción no permitida"
   End If
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cp.SetFocus
End If

End Sub

Private Sub Command1_Click()
Unload Me

End Sub



Private Sub Command3_Click()
'T_COD.Text = ""
'T_NOM.Text = ""
't_direc.Text = ""
'T_TELEF.Text = ""
't_conta.Text = ""
't_correo.Text = ""
't_rut.Text = ""
't_rsoc.Text = ""
't_local.Text = ""
'Combo1.ListIndex = -1
't_cp.Text = ""
'DBGrid1.SetFocus
'Data2.DatabaseName = "c:\mutuales"
'Data2.Refresh
If Data2.Recordset.RecordCount > 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      Data2.Recordset.MoveNext
   Loop
   MsgBox "Datos importados correctamente"
End If
End Sub

Private Sub Command2_Click()
'Dim elmensa As String
'elmensa = MsgBox("Desea borrar el registro seleccionado?", vbInformation + vbYesNo, "BORRAR")
'If elmensa = vbYes Then
'   Data1.Recordset.Delete
'   Data1.Refresh
'End If
MsgBox "Usuario no habilitado"

End Sub

Private Sub Command4_Click()
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "inflla"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
   data_inf.Refresh
End If
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      data_inf.Recordset.AddNew
      data_inf.Recordset("nombre") = Mid(Data1.Recordset("obsmot"), 1, 70)
      data_inf.Recordset("nomcat") = Mid(Data1.Recordset("obsmot"), 1, 50)
      data_inf.Recordset("direcc") = Data1.Recordset("obs")
      data_inf.Recordset("fecha") = Data1.Recordset("fecha")
      data_inf.Recordset.Update
      Data1.Recordset.MoveNext
   Loop
   data_inf.Refresh
   cr1.ReportFileName = App.path & "\infcomer.rpt"
   cr1.Action = 1
End If

End Sub

Private Sub Command5_Click()
frm_evaluacom.Show vbModal

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(Data1.Recordset("nro")) = False Then
   t_cod.Text = Data1.Recordset("nro")
   If IsNull(Data1.Recordset("obsmot")) = False Then
      t_rsoc.Text = Data1.Recordset("obsmot")
   Else
      t_rsoc.Text = ""
   End If
   data_otra.RecordSource = "Select * from abmdesp where base =" & 97 & " and mat =" & Data1.Recordset("nro")
   data_otra.Refresh
   If data_otra.Recordset.RecordCount > 0 Then
      If IsNull(data_otra.Recordset("obs")) = False Then
         t_nom.Text = data_otra.Recordset("obs")
      Else
         t_nom.Text = Data1.Recordset("obsmot")
      End If
      If IsNull(data_otra.Recordset("usuario")) = False Then
         Combo1.Text = data_otra.Recordset("usuario")
      Else
         Combo1.Text = "S/REGISTRAR"
      End If
'   Data1.Recordset("usuario") = Combo1.Text
      If IsNull(data_otra.Recordset("referen")) = False Then
         t_correo.Text = data_otra.Recordset("referen")
      Else
         t_correo.Text = "S/D"
      End If
      If IsNull(data_otra.Recordset("accion")) = False Then
         t_conta.Text = data_otra.Recordset("accion")
      Else
         t_conta.Text = "S/D"
      End If
      If IsNull(data_otra.Recordset("obsmot")) = False Then
         t_rut.Text = data_otra.Recordset("obsmot")
      Else
         t_rut.Text = "0"
      End If
   Else
      If IsNull(Data1.Recordset("obsmot")) = False Then
         t_nom.Text = Data1.Recordset("obsmot")
      Else
         t_nom.Text = "S/D"
      End If
'   Data1.Recordset("usuario") = Combo1.Text
      t_correo.Text = "S/D"
      t_conta.Text = "S/D"
      Combo1.Text = "S/REGISTRAR"
      t_rut.Text = 0
   End If
   If IsNull(Data1.Recordset("usuario")) = False Then
      t_local.Text = Data1.Recordset("usuario")
   Else
      t_local.Text = "S/D"
   End If
'   Data1.Recordset("base") = 0
   If IsNull(Data1.Recordset("obs")) = False Then
      t_direc.Text = Data1.Recordset("obs")
   Else
      t_direc.Text = "S/D"
   End If
   If IsNull(Data1.Recordset("accion")) = False Then
      t_cp.Text = Data1.Recordset("accion")
   Else
      t_cp.Text = "0"
   End If
'   Data1.Recordset("accion") = t_conta.Text
'   Data1.Recordset("referen") = t_correo.Text
   If IsNull(Data1.Recordset("referen")) = False Then
      t_telef.Text = Data1.Recordset("referen")
   Else
      t_telef.Text = "0"
   End If
   If IsNull(Data1.Recordset("mat")) = False Then
      If Data1.Recordset("mat") > 19 Then
         Combo1.ListIndex = 0
      Else
         Combo1.ListIndex = Data1.Recordset("mat")
      End If
   Else
      Combo1.ListIndex = -1
   End If
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.DatabaseName = App.Path & "\sapp.mdb"
'data_otra.DatabaseName = App.Path & "\sapp.mdb"
data_otra.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_otra.RecordSource = "Select * from abmdesp where base =" & 97
data_otra.Refresh
adodc1.ConnectionString = "dsn=" & Xconexrmt
'Adodc1.ConnectionString = "dsn=sapp"
'Adodc1.RecordSource = "Select * from abmdesp"
'Adodc1.Refresh

Data1.RecordSource = "Select * from abmdesp where base <>" & 99 & " and base <>" & 97 & " order by nro"
Data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_bus_Change()
Data1.RecordSource = "Select * from abmdesp where base <>" & 99 & " and base <>" & 97 & " and obsmot like '*" & t_bus.Text & "*' order by nro"
Data1.Refresh

End Sub

Private Sub t_bus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_conta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_correo.SetFocus
End If

End Sub

Private Sub t_correo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command2.SetFocus
End If

End Sub

Private Sub t_cp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_telef.SetFocus
End If

End Sub

Private Sub t_direc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_local.SetFocus
End If

End Sub

Private Sub t_local_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_rsoc.SetFocus
End If

End Sub

Private Sub t_rsoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_rut.SetFocus
End If

End Sub

Private Sub t_rut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_direc.SetFocus
End If

End Sub

Private Sub t_telef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_conta.SetFocus
End If

End Sub
