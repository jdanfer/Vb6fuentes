VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infcaja 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprime Caja"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frm_infcaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr3 
      Left            =   2160
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
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1560
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc data_caja 
      Height          =   375
      Left            =   3120
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
      Caption         =   "data_caja"
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
   Begin Crystal.CrystalReport cr2 
      Left            =   4560
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PARSEC0"
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5400
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox hashora 
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "HH:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox deshora 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "HH:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton btn_cerrar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Picture         =   "frm_infcaja.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton btn_acep 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "frm_infcaja.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txt_base 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin MSMask.MaskEdBox hasta 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
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
   Begin MSMask.MaskEdBox desde 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
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
   Begin VB.Label labnombre 
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label labvalor 
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label labgrupo 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label labcodigo 
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label labcedula 
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label labcodconv 
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808000&
      BorderWidth     =   3
      X1              =   0
      X2              =   5880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   3
      X1              =   0
      X2              =   5880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "BASE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
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
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Rango de hora:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Rango de fechas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2760
      Picture         =   "frm_infcaja.frx":0F56
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frm_infcaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_acep_Click()
Dim Xsaldoca As Long
Dim Xacre As String
Dim BaseaListar As String

Xacre = "CREDITO"
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcaja"
data_inf.RecordSource = "infcaja"
data_inf.Refresh
BaseaListar = "SIN BASE"

If Trim(txt_base.Text) = "" Then
   txt_base.Text = 0
End If

If Val(txt_base.Text) = 91 Then
   BaseaListar = "CENTRAL TOLEDO"
End If
If Val(txt_base.Text) = 92 Then
   BaseaListar = "SALINAS"
End If
If Val(txt_base.Text) = 93 Then
   BaseaListar = "BARROS BLANCOS"
End If
If Val(txt_base.Text) = 1 Then
   BaseaListar = "BASE 1"
End If
If Val(txt_base.Text) = 2 Then
   BaseaListar = "BASE 2"
End If
If Val(txt_base.Text) = 4 Then
   BaseaListar = "BASE 4"
End If
If Val(txt_base.Text) = 6 Then
   BaseaListar = "BASE 6"
End If
If Val(txt_base.Text) = 8 Then
   BaseaListar = "BASE 8"
End If
If Val(txt_base.Text) = 12 Then
   BaseaListar = "BASE 12"
End If
If Val(txt_base.Text) = 13 Then
   BaseaListar = "BASE 13"
End If

data_caja.RecordSource = "Select * from caja where fecha >= '" & Format(desde.Text, "yyyy-mm-dd") & "' And fecha <= '" & Format(hasta.Text, "yyyy-mm-dd") & "' and hora >= '" & deshora.Text & "' and hora <= '" & hashora.Text & "' and usuario ='" & WElusuario & "' And base =" & txt_base.Text & " order by numero,fecha,hora"
data_caja.Refresh
If data_caja.Recordset.RecordCount > 0 Then
   data_caja.Recordset.MoveFirst
   Do While Not data_caja.Recordset.EOF
      If data_caja.Recordset("movimiento") = "EGRESO" Then
         Xsaldoca = Xsaldoca - data_caja.Recordset("imp_fact")
      Else
         If data_caja.Recordset("movimiento") = "INGRESO" Then
            Xsaldoca = Xsaldoca + data_caja.Recordset("imp_fact")
         End If
      End If
      data_inf.Recordset.AddNew
      data_inf.Recordset("fecha") = data_caja.Recordset("fecha")
      data_inf.Recordset("numero") = data_caja.Recordset("numero")
      data_inf.Recordset("nombre") = data_caja.Recordset("nombre")
      data_inf.Recordset("movimiento") = data_caja.Recordset("movimiento")
      data_inf.Recordset("imp_fact") = data_caja.Recordset("imp_fact")
      data_inf.Recordset("documento") = data_caja.Recordset("documento")
      data_inf.Recordset("usuario") = data_caja.Recordset("usuario")
      data_inf.Recordset("hora") = data_caja.Recordset("hora")
      data_inf.Recordset("saldo_user") = Xsaldoca
      data_inf.Recordset("base") = data_caja.Recordset("base")
      data_inf.Recordset("cod_serv") = data_caja.Recordset("cod_serv")
      data_inf.Recordset("nom_serv") = data_caja.Recordset("nom_serv")
      data_inf.Recordset("cod_socio") = data_caja.Recordset("cod_socio")
      data_inf.Recordset("nom_socio") = data_caja.Recordset("nom_socio")
      data_inf.Recordset("caja_mesp") = data_caja.Recordset("caja_mesp")
      data_inf.Recordset("caja_anop") = data_caja.Recordset("caja_anop")
      data_inf.Recordset("saldo") = data_caja.Recordset("saldo")
      data_inf.Recordset("observ") = data_caja.Recordset("observ")
      data_inf.Recordset.Update
      data_caja.Recordset.MoveNext
   Loop
Else
   MsgBox "No existen registros", vbInformation, "Caja"
End If

If Val(txt_base.Text) = 78 Then
   BaseaListar = BaseaListar & " TEST"
End If

data_caja.RecordSource = "select * from pedidos_facturar where fecha_fact is null and fecha >='" & Format(desde.Text, "yyyy-mm-dd") & "' and lugar ='" & Trim(BaseaListar) & "'"
data_caja.Refresh
If data_caja.Recordset.RecordCount > 0 Then
   data_caja.Recordset.MoveFirst
   Do While Not data_caja.Recordset.EOF
      labcedula.Caption = data_caja.Recordset("matricula")
      Retorna_cliente
      Codigo_facturacion
      Devuelve_valor
      Xsaldoca = Xsaldoca + Val(labvalor.Caption)
      data_inf.Recordset.AddNew
      data_inf.Recordset("fecha") = data_caja.Recordset("fecha")
      data_inf.Recordset("numero") = 99999888
      data_inf.Recordset("nombre") = "QUEBRANTO MEDICACION"
      data_inf.Recordset("imp_fact") = Val(labvalor.Caption)
      data_inf.Recordset("documento") = Val(data_caja.Recordset("matricula"))
      data_inf.Recordset("usuario") = WElusuario
      data_inf.Recordset("hora") = Mid(data_caja.Recordset("updated_at"), 12, 5)
      data_inf.Recordset("saldo_user") = Xsaldoca
      data_inf.Recordset("base") = Val(data_caja.Recordset("base"))
      data_inf.Recordset("cod_serv") = Val(labcodigo.Caption)
      data_inf.Recordset("nom_serv") = "MEDICACION NOFACT"
      data_inf.Recordset("cod_socio") = Val(data_caja.Recordset("matricula"))
      data_inf.Recordset("nom_socio") = labnombre.Caption
      data_inf.Recordset("saldo") = 0
      data_inf.Recordset("observ") = Mid(data_caja.Recordset("nom_medicacion"), 1, 120)
      data_inf.Recordset.Update
      data_caja.Recordset.MoveNext
   Loop
End If

data_caja.Recordset.Close

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt
Data1.RecordSource = "Select * from linmmdd where tipo ='" & "CREDITO" & "' and fecha >= '" & Format(desde.Text, "yyyy-mm-dd") & "' And fecha <= '" & Format(hasta.Text, "yyyy-mm-dd") & "' and operador ='" & WElusuario & "' And base =" & txt_base.Text & " order by cod_prod"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      data_inf.Recordset.AddNew
      data_inf.Recordset("fecha") = Data1.Recordset("fecha")
      data_inf.Recordset("factura") = Data1.Recordset("factura")
      data_inf.Recordset("tipo") = Data1.Recordset("tipo")
      data_inf.Recordset("cod_cli") = Data1.Recordset("cod_cli")
      data_inf.Recordset("nom_cli") = Data1.Recordset("nom_cli")
      data_inf.Recordset("nom_prod") = Data1.Recordset("nom_prod")
      data_inf.Recordset("cod_prod") = Data1.Recordset("cod_prod")
      data_inf.Recordset("operador") = Data1.Recordset("operador")
      data_inf.Recordset("convenio") = Data1.Recordset("convenio")
      data_inf.Recordset("tot_lin") = Data1.Recordset("tot_lin")
      data_inf.Recordset("base") = Data1.Recordset("base")
      If IsNull(Data1.Recordset("pendiente")) = False Then
         If Data1.Recordset("pendiente") = "C" Then
            data_inf.Recordset("nom_superv") = "Nota Crédito e-Tck"
         Else
            If Data1.Recordset("pendiente") = "T" Then
               data_inf.Recordset("nom_superv") = "e-Ticket"
            Else
               data_inf.Recordset("nom_superv") = "Documento"
            End If
         End If
      End If
      data_inf.Recordset.Update
      Data1.Recordset.MoveNext
   Loop
End If

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt
Data1.RecordSource = "Select * from linmmdd where cod_prod in (802,803,804,805,806) and fecha >= '" & Format(desde.Text, "yyyy-mm-dd") & "' And fecha <= '" & Format(hasta.Text, "yyyy-mm-dd") & "' and operador ='" & WElusuario & "' And base =" & txt_base.Text & " order by cod_prod"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      data_inf.Recordset.AddNew
      data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
      data_inf.Recordset("cl_codigo") = Data1.Recordset("factura")
      data_inf.Recordset("cl_dpto") = Data1.Recordset("tipo")
      data_inf.Recordset("cl_cedula") = Data1.Recordset("cod_cli")
      data_inf.Recordset("cl_apellid") = Data1.Recordset("nom_cli")
      data_inf.Recordset("cl_direcci") = Data1.Recordset("nom_prod")
      data_inf.Recordset("cl_codced") = Data1.Recordset("cod_prod")
      data_inf.Recordset("cl_nombre") = Data1.Recordset("operador")
      data_inf.Recordset("cl_codconv") = Data1.Recordset("convenio")
      data_inf.Recordset("cl_nrovend") = Data1.Recordset("base")
      data_inf.Recordset.Update
      Data1.Recordset.MoveNext
   Loop
End If

Data1.Recordset.Close


data_inf.RecordSource = "Select * from infcaja order by fecha"
data_inf.Refresh
CrystalReport1.ReportTitle = "Fecha: " & desde.Text & " Hasta: " & hasta.Text & "  BASE: " & txt_base.Text & "  USUARIO : " & WElusuario
'CrystalReport1.ParameterFields(xdesde) = Date
CrystalReport1.Action = 1

data_inf.RecordSource = "Select * from infvtas order by cod_prod"
data_inf.Refresh
cr2.ReportFileName = App.path & "\infcredcaj.rpt"
cr2.ReportTitle = "INFORME DE VENTAS CREDITO DESDE: " & desde.Text & " HASTA: " & hasta.Text & "  BASE: " & txt_base.Text & "  USUARIO : " & WElusuario
cr2.Action = 1

data_inf.RecordSource = "Select * from infcli order by cl_codced"
data_inf.Refresh
cr3.ReportFileName = App.path & "\infcartacaj.rpt"
cr3.ReportTitle = "INFORME DE CARTAS FACTURADAS DE:" & desde.Text & " AL:" & hasta.Text & " BASE:" & txt_base.Text & " USUARIO:" & WElusuario
cr3.Action = 1


End Sub

Private Sub btn_cerrar_Click()
frm_infcaja.Hide

End Sub



Private Sub desde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   hasta.SetFocus
End If

End Sub

Private Sub deshora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   hashora.SetFocus
End If

End Sub

Private Sub Form_Activate()
txt_base.Text = data_parsec.Recordset("base")
desde.SetFocus
desde.Text = Format(Date, "dd/mm/yyyy")
hasta.Text = Format(Date, "dd/mm/yyyy")
deshora.Text = Format("00:00", "HH:mm")
hashora.Text = Format(Time, "HH:mm")

End Sub

Private Sub Form_Load()
Dim Xladesde As Date
Xladesde = Date - 10

'data_caja.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_caja.ConnectionString = "dsn=" & Xconexrmt
data_caja.RecordSource = "Select * from caja where fecha >='" & Format(Xladesde, "yyyy-mm-dd") & "'"
data_caja.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_parsec.DatabaseName = App.path & "\parse.mdb"

CrystalReport1.ReportFileName = App.path & "\infcajabien.rpt"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub hashora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_base.SetFocus
End If

End Sub

Private Sub hasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   deshora.SetFocus
End If

End Sub

Private Sub txt_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btn_acep.SetFocus
End If

End Sub

Public Sub Retorna_cliente()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from clientes where cl_cedula =" & Val(labcedula.Caption)

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labcodconv.Caption = Xrecclii("cl_codconv")
   labnombre.Caption = Mid(Xrecclii("cl_apellid"), 1, 25)
Else
   labcodconv.Caption = "CCSD"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Codigo_facturacion()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from convenio where cnv_codigo ='" & labcodconv.Caption & "'"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cnv_aran")) = False Then
      labgrupo.Caption = Xrecclii("cnv_aran")
   Else
      labgrupo.Caption = 0
   End If
   If IsNull(Xrecclii("cnv_grupo")) = False Then
      If Xrecclii("cnv_grupo") = "CCOU" Then
         labcodigo.Caption = 60103
      Else
         If Xrecclii("cnv_grupo") = "H.EVANGELICO" Then
            labcodigo.Caption = 60107
         Else
            If Xrecclii("cnv_grupo") = "SMI" Then
               labcodigo.Caption = 60106
            Else
               If Xrecclii("cnv_grupo") = "UNIVERSAL" Then
                  labcodigo.Caption = 60108
               Else
                  labcodigo.Caption = 60103
               End If
            End If
         End If
      End If
   Else
      labcodigo.Caption = 60103
   End If
Else
   labcodigo.Caption = 60103
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Devuelve_valor()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xdescuento As Double
Dim XPorcen As Integer
Dim XPrec As Integer
Xdescuento = 0
XPorcen = 0
XPrec = 0
If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from estudios where codest =" & Val(labcodigo.Caption)

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labvalor.Caption = Xrecclii("cons")
Else
   labvalor.Caption = 0
End If

Xrecclii.Close

ConbdSapp.Close

End Sub
