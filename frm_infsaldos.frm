VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infsaldos 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Saldos de clientes"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infsaldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
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
      Top             =   2160
      Visible         =   0   'False
      Width           =   2940
   End
   Begin Crystal.CrystalReport cr22 
      Left            =   5400
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      Picture         =   "frm_infsaldos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_infsaldos.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Procesar"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Seleccione los datos a informar..."
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin MSAdodcLib.Adodc data1 
         Height          =   375
         Left            =   3600
         Top             =   2640
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
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data_deudas 
         Height          =   375
         Left            =   3240
         Top             =   2280
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "data_deudas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C00000&
         Caption         =   "Por edades"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C00000&
         Caption         =   "Solo deudas de emisión"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Socios ACTIVOS y BAJAS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Informe sin detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_infsaldos.frx":0F56
         Left            =   2160
         List            =   "frm_infsaldos.frx":0F66
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infsaldos.frx":0F8F
         Left            =   2160
         List            =   "frm_infsaldos.frx":0F9F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Cobrador:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Deuda >= a:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1800
      Picture         =   "frm_infsaldos.frx":0FC4
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "frm_infsaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xquequiere As String
Dim Xquema, Xcuantosm As Long
Dim Xdias As Double
'data_deudas.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deudas.ConnectionString = "dsn=" & Xconexrmt

If Combo2.ListIndex = 3 Then
   Xquequiere = InputBox("Ingrese número de cobrador", "Cobrador a listar")
   data_deudas.RecordSource = "Select deudas.cliente,deudas.nombre,deudas.tiquet,deudas.servi,deudas.deudas,deudas.cod_cnv,deudas.importe,deudas.iva,deudas.total,deudas.mes,deudas.ano,deudas.estado_cta," & _
   "deudas.nom_cobr,deudas.nro_cobr,deudas.fecha_pago,clientes.estado,clientes.cl_codigo,clientes.cl_telefon,clientes.cl_dpto,clientes.cl_cedula,clientes.cl_codced,clientes.fecha_baja " & _
   "from deudas inner join clientes on deudas.cliente=clientes.cl_codigo where clientes.fecha_baja is null and deudas.estado_cta =" & 1 & " and deudas.nro_cobr =" & Val(Xquequiere) & " and deudas.mes <>" & 0 & " and deudas.fecha_pago is null order by cliente,ano,mes"
   
'   data_deudas.RecordSource = "Select * from deudas where nro_cobr =" & Val(Xquequiere) & " order by cliente,ano,mes"
'   data_deudas.Refresh
Else
   Xquequiere = ""
End If
frm_infsaldos.MousePointer = 11
Command1.Enabled = False
Command2.Enabled = False

If Xquequiere = "" Then
   If Combo2.ListIndex = 0 Then
      data_deudas.RecordSource = "Select deudas.cliente,deudas.nombre,deudas.tiquet,deudas.servi,deudas.deudas,deudas.cod_cnv,deudas.importe,deudas.iva,deudas.total,deudas.mes,deudas.ano,deudas.estado_cta," & _
      "deudas.nom_cobr,deudas.nro_cobr,deudas.fecha_pago,clientes.estado,clientes.cl_codigo,clientes.cl_telefon,clientes.cl_dpto,clientes.cl_cedula,clientes.cl_codced,clientes.fecha_baja " & _
      "from deudas inner join clientes on deudas.cliente=clientes.cl_codigo where clientes.fecha_baja is null and deudas.estado_cta =" & 1 & " and deudas.nro_cobr not in (0) and deudas.mes <>" & 0 & " and deudas.fecha_pago is null order by cliente,ano,mes"
      data_deudas.Refresh
      
'      data_deudas.RecordSource = "Select * from deudas where nro_cobr >" & 0 & " and fecha_pago is null order by cliente,ano,mes"
'      data_deudas.Refresh
   End If
   If Combo2.ListIndex = 1 Then 'bases
      data_deudas.RecordSource = "Select deudas.cliente,deudas.nombre,deudas.tiquet,deudas.servi,deudas.deudas,deudas.cod_cnv,deudas.importe,deudas.iva,deudas.total,deudas.mes,deudas.ano,deudas.estado_cta," & _
      "deudas.nom_cobr,deudas.nro_cobr,deudas.fecha_pago,clientes.estado,clientes.cl_codigo,clientes.cl_telefon,clientes.cl_dpto,clientes.cl_cedula,clientes.cl_codced,clientes.fecha_baja " & _
      "from deudas inner join clientes on deudas.cliente=clientes.cl_codigo where clientes.fecha_baja is null and deudas.estado_cta =" & 1 & " and deudas.nro_cobr in (208,209,10,615,616,602,1,113,201,635,636,653,685) and deudas.mes <>" & 0 & " and deudas.fecha_pago is null order by cliente,ano,mes"
      data_deudas.Refresh
   End If
   If Combo2.ListIndex = 2 Then
      data_deudas.RecordSource = "Select deudas.cliente,deudas.nombre,deudas.tiquet,deudas.servi,deudas.deudas,deudas.cod_cnv,deudas.importe,deudas.iva,deudas.total,deudas.mes,deudas.ano,deudas.estado_cta," & _
      "deudas.nom_cobr,deudas.nro_cobr,deudas.fecha_pago,clientes.estado,clientes.cl_codigo,clientes.cl_telefon,clientes.cl_dpto,clientes.cl_cedula,clientes.cl_codced,clientes.fecha_baja " & _
      "from deudas inner join clientes on deudas.cliente=clientes.cl_codigo where clientes.fecha_baja is null and deudas.estado_cta =" & 1 & " and deudas.nro_cobr not in (208,209,10,615,616,602,1,113,201,635,636,653,685) and deudas.mes <>" & 0 & " and deudas.fecha_pago is null order by cliente,ano,mes"
      
'      data_deudas.RecordSource = "Select * from deudas where nro_cobr =" & 601 & _
'      " or nro_cobr =" & 605 & " or nro_cobr =" & 699 & " and fecha_pago is null order by cliente,ano,mes"
      data_deudas.Refresh
   End If

Else
   If Combo2.ListIndex = 3 Then 'seleccion
      data_deudas.RecordSource = "Select deudas.cliente,deudas.nombre,deudas.tiquet,deudas.servi,deudas.deudas,deudas.cod_cnv,deudas.importe,deudas.iva,deudas.total,deudas.mes,deudas.ano,deudas.estado_cta," & _
      "deudas.nom_cobr,deudas.nro_cobr,deudas.fecha_pago,clientes.estado,clientes.cl_codigo,clientes.cl_telefon,clientes.cl_dpto,clientes.cl_cedula,clientes.cl_codced,clientes.fecha_baja " & _
      "from deudas inner join clientes on deudas.cliente=clientes.cl_codigo where clientes.fecha_baja is null and deudas.estado_cta =" & 1 & " and deudas.nro_cobr =" & Xquequiere & " and deudas.mes <>" & 0 & " and deudas.fecha_pago is null order by cliente,ano,mes"
      data_deudas.Refresh
   End If

End If

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infemis"
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infemis"
data_inf.Refresh

If data_deudas.Recordset.RecordCount > 0 Then
   data_deudas.Recordset.MoveFirst
   Xquema = data_deudas.Recordset("cliente")
   Xcuantosm = 0
   Do While Not data_deudas.Recordset.EOF
      If Xquema = data_deudas.Recordset("cliente") Then
         Xcuantosm = Xcuantosm + 1
         data_inf.Recordset.AddNew
         data_inf.Recordset("cliente") = data_deudas.Recordset("cliente")
         data_inf.Recordset("apellidos") = data_deudas.Recordset("nombre")
         data_inf.Recordset("cod_cnv") = data_deudas.Recordset("cod_cnv")
         data_inf.Recordset("importe") = data_deudas.Recordset("importe")
         data_inf.Recordset("tiquet") = data_deudas.Recordset("tiquet")
         data_inf.Recordset("servi") = data_deudas.Recordset("servi")
         data_inf.Recordset("deudas") = data_deudas.Recordset("deudas")
         data_inf.Recordset("iva") = data_deudas.Recordset("iva")
         data_inf.Recordset("total") = data_deudas.Recordset("total")
         data_inf.Recordset("mes") = data_deudas.Recordset("mes")
         data_inf.Recordset("ano") = data_deudas.Recordset("ano")
         data_inf.Recordset("nro_cobr") = data_deudas.Recordset("nro_cobr")
         data_inf.Recordset("nom_cobr") = data_deudas.Recordset("nom_cobr")
'         Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_deudas.Recordset("cliente")
'         Data1.Refresh
'         If Data1.Recordset.RecordCount > 0 Then
         If IsNull(data_deudas.Recordset("cl_dpto")) = False Then
            data_inf.Recordset("origen") = data_deudas.Recordset("cl_dpto")
         Else
            If IsNull(data_deudas.Recordset("cl_telefon")) = False Then
               data_inf.Recordset("origen") = data_deudas.Recordset("cl_telefon")
            Else
               data_inf.Recordset("origen") = "Sin Datos"
            End If
         End If
'            If IsNull(Data1.Recordset("cl_cedula")) = False Then
'               data_inf.Recordset("ruc") = Trim(Str(Data1.Recordset("cl_cedula"))) & Trim(Str(Data1.Recordset("cl_codced")))
'            Else
'               data_inf.Recordset("ruc") = "Sin CI"
'            End If
'            If IsNull(Data1.Recordset("cl_fnac")) = False Then
'               Xdias = Date - CDate(Data1.Recordset("cl_fnac"))
'               data_inf.Recordset("numero") = Xdias / 365
'               data_inf.Recordset("fecha_nac") = Data1.Recordset("cl_fnac")
'            Else
'               data_inf.Recordset("numero") = 0
'            End If
            
'         End If
         data_inf.Recordset.Update
         data_inf.Refresh
         Xquema = data_deudas.Recordset("cliente")
         data_deudas.Recordset.MoveNext
      Else
         data_deudas.Recordset.MovePrevious
         If Combo1.ListIndex = 1 Then
            If Xcuantosm >= 2 Then
            Else
               data_inf.RecordSource = "Select * from infemis where cliente =" & Xquema
               data_inf.Refresh
               If data_inf.Recordset.RecordCount > 0 Then
                  data_inf.Recordset.MoveFirst
                  Do While Not data_inf.Recordset.EOF
                     data_inf.Recordset.Delete
                     data_inf.Recordset.MoveNext
                  Loop
                  data_inf.RecordSource = "infemis"
                  data_inf.Refresh
               End If
            End If
         End If
         If Combo1.ListIndex = 2 Then
'            Data1.RecordSource = "Select * from clientes where cl_codigo =" & Xquema
'            Data1.Refresh
'            If Data1.Recordset.RecordCount > 0 Then
               If data_deudas.Recordset("estado") = 1 Then
                  If Xcuantosm >= 3 Then
'                     Xcuantosm = 0
                  Else
                      MiBaseact.Execute "Delete * from infemis where cliente =" & Xquema
                      data_inf.RecordSource = "infemis"
                      data_inf.Refresh
                      data_inf.RecordSource = "infemis"
                      data_inf.Refresh
        '                  Xcuantosm = 0
                  End If
               Else
                  MiBaseact.Execute "Delete * from infemis where cliente =" & Xquema
                  data_inf.RecordSource = "infemis"
                  data_inf.Refresh
                  data_inf.RecordSource = "infemis"
                  data_inf.Refresh
               End If
'            Else
'               MiBaseact.Execute "Delete * from infemis where cliente =" & Xquema
'               data_inf.RecordSource = "infemis"
'               data_inf.Refresh
'            End If
         End If
         If Combo1.ListIndex = 3 Then
            If Xcuantosm >= 4 Then
            Else
               MiBaseact.Execute "Delete * from infemis where cliente =" & Xquema
               data_inf.RecordSource = "infemis"
               data_inf.Refresh
            End If
         End If
         Xcuantosm = 0
         data_deudas.Recordset.MoveNext
         Xquema = data_deudas.Recordset("cliente")
      End If
   Loop
   data_deudas.Recordset.MovePrevious
     If Combo1.ListIndex = 1 Then
        If Xcuantosm >= 2 Then
        Else
           data_inf.RecordSource = "Select * from infemis where cliente =" & Xquema
           data_inf.Refresh
           If data_inf.Recordset.RecordCount > 0 Then
              data_inf.Recordset.MoveFirst
              Do While Not data_inf.Recordset.EOF
                 data_inf.Recordset.Delete
                 data_inf.Recordset.MoveNext
              Loop
              data_inf.RecordSource = "infemis"
              data_inf.Refresh
           End If
        End If
     End If
     If Combo1.ListIndex = 2 Then
'        Data1.RecordSource = "Select * from clientes where cl_codigo =" & Xquema
'        Data1.Refresh
'        If Data1.Recordset.RecordCount > 0 Then
           If data_deudas.Recordset("estado") = 1 Then
              If Xcuantosm >= 3 Then
    '                     Xcuantosm = 0
              Else
                  data_inf.RecordSource = "Select * from infemis where cliente =" & Xquema
                  data_inf.Refresh
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveFirst
                     Do While Not data_inf.Recordset.EOF
                        data_inf.Recordset.Delete
                        data_inf.Recordset.MoveNext
                     Loop
                     data_inf.RecordSource = "infemis"
                     data_inf.Refresh
    '                  Xcuantosm = 0
                  End If
              End If
           Else
              data_inf.RecordSource = "Select * from infemis where cliente =" & Xquema
              data_inf.Refresh
              If data_inf.Recordset.RecordCount > 0 Then
                 data_inf.Recordset.MoveFirst
                 Do While Not data_inf.Recordset.EOF
                    data_inf.Recordset.Delete
                    data_inf.Recordset.MoveNext
                 Loop
                 data_inf.RecordSource = "infemis"
                 data_inf.Refresh
              End If
           End If
''        End If
     End If
     If Combo1.ListIndex = 3 Then
        If Xcuantosm >= 4 Then
        Else
           data_inf.RecordSource = "Select * from infemis where cliente =" & Xquema
           data_inf.Refresh
           If data_inf.Recordset.RecordCount > 0 Then
              data_inf.Recordset.MoveFirst
              Do While Not data_inf.Recordset.EOF
                 data_inf.Recordset.Delete
                 data_inf.Recordset.MoveNext
              Loop
              data_inf.RecordSource = "infemis"
              data_inf.Refresh
           End If
        End If
     End If
     Xcuantosm = 0
End If

''''If data_inf.Recordset.RecordCount > 0 Then
''''   data_inf.Recordset.MoveFirst
''''   Do While Not data_inf.Recordset.EOF
''''      Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_inf.Recordset("cliente")
''''      Data1.Refresh
''''      If Data1.Recordset.RecordCount > 0 Then
''''         If IsNull(Data1.Recordset("fecha_baja")) = True Then
''''         Else
 ''''           data_inf.Recordset.Delete
''''         End If
''''      End If
''''      data_inf.Recordset.MoveNext
''''   Loop
''''End If
If Check4.value = 1 Then
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         If data_inf.Recordset("numero") >= 18 And data_inf.Recordset("numero") <= 65 Then
         Else
            data_inf.Recordset.Delete
         End If
         data_inf.Recordset.MoveNext
      Loop
   End If
End If

frm_infsaldos.MousePointer = 0
MsgBox "Proceso terminado..."
data_inf.RecordSource = "Select * from infemis"
data_inf.Refresh
If Check1.value = 1 Then
   cr22.ReportFileName = App.Path & "\infdeuda2.rpt"
Else
   If Check4.value = 1 Then
      cr22.ReportFileName = App.Path & "\infdeuda22.rpt"
   Else
      cr22.ReportFileName = App.Path & "\infdeuda2.rpt"
   End If
End If
cr22.Action = 1

frm_infsaldos.MousePointer = 0

Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
