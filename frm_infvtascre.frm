VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infvtascre 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes ventas crédito"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infvtascre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inflla 
      Caption         =   "data_inflla"
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
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_inf2 
      Caption         =   "data_inf2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   2040
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      Picture         =   "frm_infvtascre.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infvtascre.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Procesar"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Opciones de informe"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   330
         Left            =   1800
         Top             =   3360
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "data_cli"
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
      Begin MSAdodcLib.Adodc data_lla 
         Height          =   375
         Left            =   2760
         Top             =   2280
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "data_lla"
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
      Begin MSAdodcLib.Adodc data_lin2 
         Height          =   330
         Left            =   2760
         Top             =   1920
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "data_lin2"
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
      Begin MSAdodcLib.Adodc data_cab 
         Height          =   495
         Left            =   120
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
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
         Caption         =   "data_cab"
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
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   330
         Left            =   2520
         Top             =   3000
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Caption         =   "data_lin"
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
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Controles"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Ordenar por cliente > 2"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Ordenar por cliente $ > a..."
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Resumen"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   2160
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infvtascre.frx":109E
         Left            =   1920
         List            =   "frm_infvtascre.frx":10B4
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Si no ingresa rango de fecha se emite TODO."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Selección"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "FECHA:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1680
      Picture         =   "frm_infvtascre.frx":1146
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2295
   End
End
Attribute VB_Name = "frm_infvtascre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
   If Check2.Value = 1 Then
      Check2.Value = 0
   End If
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   If Check1.Value = 1 Then
      Check1.Value = 0
   End If
End If

End Sub

Private Sub Command1_Click()
Command1.Enabled = False

frm_infvtascre.MousePointer = 11
Xaltaaa = 0
Dim Xbandeu As Integer
Dim Xdiasinf, Xcantdeu As Integer
Dim Xmatdeu As Double

Xdiasinf = 0
Xbandeu = 0
Xcantdeu = 0

If mfd.Text = "__/__/____" And mfh.Text = "__/__/____" Then
   mfd.Text = Date - 1600
   mfh.Text = Date
End If
pb.Visible = False
If mfd.Text = "__/__/____" And mfh.Text = "__/__/____" Then
   data_lin.RecordSource = "Select * from deudas"
   data_lin.Refresh
Else
   If Combo1.ListIndex = 0 Then
      data_lin.RecordSource = "select deudas.fecha,deudas.cliente,deudas.fecha_pago,deudas.nro_superv,deudas.nro_vende,deudas.nom_cnv," & _
      "deudas.tipodoc,deudas.nombre,deudas.documento,deudas.cod_cnv,deudas.importe,clientes.cl_codigo,clientes.cl_direcci," & _
      "clientes.cl_dpto,clientes.cl_telefon,clientes.cl_cedula,clientes.cl_codced,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_prod,linmmdd.tipo from deudas " & _
      "inner join clientes on deudas.cliente=clientes.cl_codigo inner join linmmdd on deudas.documento=linmmdd.factura and deudas.cliente=linmmdd.cod_cli where deudas.tipodoc ='" & "CRE" & "'" & _
      " and deudas.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and deudas.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and linmmdd.tipo='" & "CREDITO" & "' order by deudas.fecha"
      
'      data_lin.RecordSource = "Select * from deudas where tipodoc ='" & "CRE" & "' and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
      data_lin.Refresh
   Else
      If Combo1.ListIndex = 1 Then
         data_lin.RecordSource = "select deudas.fecha,deudas.cliente,deudas.fecha_pago,deudas.nro_superv,deudas.nro_vende,deudas.nom_cnv," & _
         "deudas.tipodoc,deudas.nombre,deudas.documento,deudas.cod_cnv,deudas.importe,clientes.cl_codigo,clientes.cl_direcci," & _
         "clientes.cl_dpto,clientes.cl_telefon,clientes.cl_cedula,clientes.cl_codced,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_prod,linmmdd.cod_prod,linmmdd.tipo from deudas " & _
         "inner join clientes on deudas.cliente=clientes.cl_codigo inner join linmmdd on deudas.documento=linmmdd.factura and deudas.cliente=linmmdd.cod_cli where deudas.tipodoc ='" & "CRE" & "'" & _
         " and deudas.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and deudas.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and deudas.fecha_pago is null and linmmdd.tipo='" & "CREDITO" & "' and linmmdd.cod_prod not in (10009,10010,10011,10012,10013,10014,10015,10016,997,999) order by deudas.fecha"
         
'         data_lin.RecordSource = "Select * from deudas where tipodoc ='" & "CRE" & "' and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and fecha_pago is null order by fecha"
         data_lin.Refresh
      Else
         If Combo1.ListIndex = 2 Then
            data_lin.RecordSource = "select deudas.fecha,deudas.cliente,deudas.fecha_pago,deudas.nro_superv,deudas.nro_vende,deudas.nom_cnv," & _
            "deudas.tipodoc,deudas.nombre,deudas.documento,deudas.cod_cnv,deudas.importe,clientes.cl_codigo,clientes.cl_direcci," & _
            "clientes.cl_dpto,clientes.cl_telefon,clientes.cl_cedula,clientes.cl_codced,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_prod,linmmdd.tipo from deudas " & _
            "inner join clientes on deudas.cliente=clientes.cl_codigo inner join linmmdd on deudas.documento=linmmdd.factura and deudas.cliente=linmmdd.cod_cli where deudas.tipodoc ='" & "CRE" & "'" & _
            " and deudas.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and deudas.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and deudas.fecha_pago is not null and linmmdd.tipo='" & "CREDITO" & "' order by deudas.fecha"
'            data_lin.RecordSource = "Select * from deudas where tipodoc ='" & "CRE" & "' and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and fecha_pago >='" & CDate("01/01/2011") & "# order by fecha"
            data_lin.Refresh
         Else
            If Combo1.ListIndex = 3 Then
               data_lin.RecordSource = "select deudas.fecha,deudas.cliente,deudas.fecha_pago,deudas.nro_superv,deudas.nro_vende,deudas.nom_cnv," & _
               "deudas.tipodoc,deudas.nombre,deudas.documento,deudas.cod_cnv,deudas.importe,clientes.cl_codigo,clientes.cl_direcci," & _
               "clientes.cl_dpto,clientes.cl_telefon,clientes.cl_cedula,clientes.cl_codced,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_prod,linmmdd.tipo from deudas " & _
               "inner join clientes on deudas.cliente=clientes.cl_codigo inner join linmmdd on deudas.documento=linmmdd.factura and deudas.cliente=linmmdd.cod_cli where deudas.tipodoc ='" & "CRE" & "'" & _
               " and deudas.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and deudas.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and deudas.fecha_pago is null and linmmdd.tipo='" & "CREDITO" & "' order by deudas.fecha"
               data_lin.Refresh
            Else
               If Combo1.ListIndex = 4 Then
                  data_lin.RecordSource = "select deudas.fecha,deudas.cliente,deudas.fecha_pago,deudas.nro_superv,deudas.nro_vende,deudas.nom_cnv," & _
                  "deudas.tipodoc,deudas.nombre,deudas.documento,deudas.cod_cnv,deudas.importe,clientes.cl_codigo,clientes.cl_direcci," & _
                  "clientes.cl_dpto,clientes.cl_telefon,clientes.cl_cedula,clientes.cl_codced,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_prod,linmmdd.tipo from deudas " & _
                  "inner join clientes on deudas.cliente=clientes.cl_codigo inner join linmmdd on deudas.documento=linmmdd.factura and deudas.cliente=linmmdd.cod_cli where deudas.tipodoc ='" & "CRE" & "'" & _
                  " and deudas.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and deudas.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and deudas.fecha_pago is null and linmmdd.tipo='" & "CREDITO" & "' order by deudas.fecha"
                  data_lin.Refresh
               Else
                  If Combo1.ListIndex = 5 Then
                     data_lin.RecordSource = "Select * from deudas where tipodoc ='" & "CRE" & "' limit 1"
                     data_lin.Refresh
                     data_lla.RecordSource = "Select distinct fecha,nro_flia,cod_cli,nom_cli,convenio,nom_prod from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia not in (8,19,18,2) and cod_prod not in (10008) and convenio not in " & _
                     "('911','911B','EVENTO','CAAMEP','APNORE','MSP','CASH','UCM','TING','SEMM','SEMM1','PLAYA','PART','CAUTE','CCASMU','CASMU','UDEMM','CERSEM','711','55','TATA','MUCAMT','DISCO','PREFE2','PREFE1','PEAJE1','SUAT')"
                     data_lla.Refresh
'                  Else
'                     If Combo1.ListIndex = 6 Then
'                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod <>" & 999 & " and cod_prod <>" & 997 & " and cod_prod <>" & 994 & " and cod_prod <>" & 993 & " order by fecha"
'                        data_lin.Refresh
'                     Else
'                        Xaltaaa = 8
'                        frm_vtallama.Show vbModal
'                        Command1.Enabled = True
'                        frm_infvtascre.MousePointer = 0
'                        Exit Sub
                     'End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
pb.Value = 0

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

MiBaseact.Execute "Delete * from inflla"

data_inflla.RecordSource = "inflla"
data_inflla.Refresh

data_inf2.DatabaseName = App.path & "\informes.mdb"
data_inf2.RecordSource = "infcli"
data_inf2.Refresh

If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
        If Combo1.ListIndex = 5 Then
        Else
            Do While Not data_lin.Recordset.EOF
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_lin.Recordset("cliente")
               data_inf.Recordset("cl_apellid") = Mid(data_lin.Recordset("nombre"), 1, 60)
               data_inf.Recordset("cl_fnac") = data_lin.Recordset("fecha")
               data_inf.Recordset("cl_direcci") = data_lin.Recordset("cl_direcci")
               data_inf.Recordset("cl_localid") = data_lin.Recordset("cl_dpto") & " " & data_lin.Recordset("cl_telefon") 'tel
               data_inf.Recordset("cl_nombre") = Mid(data_lin.Recordset("nombre"), 1, 30)
               data_inf.Recordset("cl_cantdia") = data_lin.Recordset("cl_cedula") 'ced
               data_inf.Recordset("cl_forpago") = data_lin.Recordset("cl_codced") 'codced
               data_inf.Recordset("cl_nomconv") = Mid(data_lin.Recordset("nom_cnv"), 1, 30)
               data_inf.Recordset("cl_codconv") = data_lin.Recordset("cod_cnv")
               data_inf.Recordset("saldo_cc") = data_lin.Recordset("importe")
               data_inf.Recordset("cl_atrasoa") = data_lin.Recordset("documento") 'factura
               data_inf.Recordset("cl_fecing") = data_lin.Recordset("fecha_pago")
               data_inf.Recordset("cl_nrovend") = data_lin.Recordset("nro_superv")
               If data_lin.Recordset("nro_superv") > 0 Then
                  data_inf.Recordset("cl_fultmov") = data_lin.Recordset("fecha") + data_lin.Recordset("nro_superv")
               Else
                  data_inf.Recordset("cl_fultmov") = data_lin.Recordset("fecha") + 30
               End If
               data_inf.Recordset("cl_nro_sup") = 1
               data_inf.Recordset("cl_nom_sup") = Mid(data_lin.Recordset("nom_prod"), 1, 25)
               data_inf.Recordset("cl_nrocobr") = 0
               data_inf.Recordset.Update
               data_lin.Recordset.MoveNext
            Loop
        End If
   'End If
   data_inf.Refresh
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Database.Execute "Delete from infcli where saldo_cc =" & 0
      data_inf.RecordSource = "Select * from infcli order by cl_atrasoa"
      data_inf.Refresh
      data_inf.Recordset.MoveFirst
      Xcantdeu = 0
      Xmatdeu = 0
      Do While Not data_inf.Recordset.EOF
         If data_inf.Recordset("cl_atrasoa") = Xmatdeu Then
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_codced") = 8
            data_inf.Recordset.Update
         Else
            Xcantdeu = 0
         End If
         Xmatdeu = data_inf.Recordset("cl_atrasoa")
         data_inf.Recordset.MoveNext
      Loop
      data_inf.Refresh
      data_inf.Database.Execute "Delete from infcli where cl_codced =" & 8
      data_inf.RecordSource = "Select * from infcli"
      data_inf.Refresh
      
   End If
   If Combo1.ListIndex = 5 Then
      pb.Visible = True
      If data_lla.Recordset.RecordCount > 0 Then
         data_lla.Recordset.MoveLast
         pb.Max = data_lla.Recordset.RecordCount + 1
         data_lla.Recordset.MoveFirst
         Xdiasinf = DateDiff("d", Format(mfd.Text, "dd/mm/yyyy"), Format(mfh.Text, "dd/mm/yyyy"))
         If Xdiasinf > 5 Then
            Dim Silocambia As String
            frm_infvtascre.MousePointer = 0
            Silocambia = MsgBox("El rango de fecha ingresado puede demorar más de 5 minutos el proceso, desea continuar?", vbInformation + vbYesNo)
            If Silocambia = vbYes Then
                frm_infvtascre.MousePointer = 11
                Do While Not data_lla.Recordset.EOF
                   Xbandeu = 0
                   If Xbandeu <> 8 Then
                      data_cab.RecordSource = "select count(*) Total_creds from deudas where cliente =" & data_lla.Recordset("cod_cli") & " and fecha_pago is null and mes =" & 0 & " and origen <'" & "Refinanc" & "' and fecha <'" & Format(data_lla.Recordset("fecha"), "yyyy-mm-dd") & "' and importe >" & 0
                      data_cab.Refresh
                      If data_cab.Recordset("Total_creds") > 0 Then
                         data_inflla.Recordset.AddNew
                         data_inflla.Recordset("matric") = data_lla.Recordset("cod_cli")
                         data_inflla.Recordset("fecha") = data_lla.Recordset("fecha")
'''                         data_inflla.Recordset("hora") = data_lla.Recordset("hora")
                         data_inflla.Recordset("nombre") = data_lla.Recordset("nom_cli")
                         data_inflla.Recordset("categ") = data_lla.Recordset("convenio")
'''                         data_inflla.Recordset("codmed") = data_lla.Recordset("nro_med_a")
'''                         data_inflla.Recordset("nommed") = data_lla.Recordset("nom_med_a")
                         data_inflla.Recordset("direcc") = data_lla.Recordset("nom_prod")
'''                         data_inflla.Recordset("movilpas") = data_lla.Recordset("base")
                         data_inflla.Recordset.Update
                      End If
                   End If
                   data_lla.Recordset.MoveNext
                   pb.Value = pb.Value + 1
                Loop
            End If
         Else
             Do While Not data_lla.Recordset.EOF
                Xbandeu = 0
                
                If Xbandeu <> 8 Then
                   data_cab.RecordSource = "select count(*) Total_creds from deudas where cliente =" & data_lla.Recordset("cod_cli") & " and fecha_pago is null and mes =" & 0 & " and origen <'" & "Refinanc" & "' and fecha <'" & Format(data_lla.Recordset("fecha"), "yyyy-mm-dd") & "' and importe >" & 0
                   data_cab.Refresh
                   If data_cab.Recordset("Total_creds") > 0 Then
                      data_inflla.Recordset.AddNew
                      data_inflla.Recordset("matric") = data_lla.Recordset("cod_cli")
                      data_inflla.Recordset("fecha") = data_lla.Recordset("fecha")
'                      data_inflla.Recordset("hora") = data_lla.Recordset("hora")
                      data_inflla.Recordset("nombre") = data_lla.Recordset("nom_cli")
                      data_inflla.Recordset("categ") = data_lla.Recordset("convenio")
'                      data_inflla.Recordset("codmed") = data_lla.Recordset("nro_med_a")
'                      data_inflla.Recordset("nommed") = data_lla.Recordset("nom_med_a")
                      data_inflla.Recordset("direcc") = data_lla.Recordset("nom_prod")
'                      data_inflla.Recordset("movilpas") = data_lla.Recordset("base")
                      data_inflla.Recordset.Update
                   End If
                End If
                data_lla.Recordset.MoveNext
                pb.Value = pb.Value + 1
             Loop
         End If
      End If
   End If
   If Combo1.ListIndex = 3 Then
   
   End If
   If Combo1.ListIndex = 4 Then
   
   End If
   If Check1.Value = 1 And Xvalornum > 0 Then
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.RecordSource = "Select * from infcli order by cl_codigo"
         data_inf.Refresh
         data_inf.Recordset.MoveFirst
         Dim Xmat, Xtot As Double
         Xmat = data_inf.Recordset("cl_codigo")
         Xtot = 0
         Do While Not data_inf.Recordset.EOF
            If Xmat = data_inf.Recordset("cl_codigo") Then
               Xtot = Xtot + data_inf.Recordset("saldo_cc")
            Else
               If Xtot >= Xvalornum Then
               Else
                  data_inf2.Database.Execute "Delete from infcli where cl_codigo =" & Xmat
               End If
               Xtot = 0
               Xtot = data_inf.Recordset("saldo_cc")
            End If
            Xmat = data_inf.Recordset("cl_codigo")
            data_inf.Recordset.MoveNext
         Loop
         If Xtot >= Xvalornum Then
         Else
            data_inf2.Database.Execute "Delete from infcli where cl_codigo =" & Xmat
         End If
         data_inf.Refresh
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
         End If
      End If
   End If
   
   If Check2.Value = 1 Then
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.RecordSource = "Select * from infcli order by cl_codigo"
         data_inf.Refresh
         data_inf.Recordset.MoveFirst
         Dim Xmat2, Xtot2 As Double
         Xmat2 = data_inf.Recordset("cl_codigo")
         Xtot2 = 0
         Do While Not data_inf.Recordset.EOF
            If Xmat2 = data_inf.Recordset("cl_codigo") Then
               If data_inf.Recordset("saldo_cc") > 0 Then
                  Xtot2 = Xtot2 + 1
               End If
            Else
               If Xtot2 > 2 Then
               Else
                  data_inf2.Database.Execute "Delete from infcli where cl_codigo =" & Xmat2
               End If
               Xtot2 = 0
               If data_inf.Recordset("saldo_cc") > 0 Then
                  Xtot2 = 1
               End If
            End If
            Xmat2 = data_inf.Recordset("cl_codigo")
            data_inf.Recordset.MoveNext
         Loop
         If Xtot2 > 2 Then
         Else
            data_inf2.Database.Execute "Delete from infcli where cl_codigo =" & Xmat2
         End If
         data_inf.Refresh
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
         End If
      End If
   End If
   If Check3.Value = 1 Then
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            If Combo1.ListIndex = 7 Then
               If data_inf.Recordset("cl_cedula") = 10002 Or data_inf.Recordset("cl_cedula") = 10004 Or _
                  data_inf.Recordset("cl_cedula") = 10006 Or data_inf.Recordset("cl_cedula") = 10008 Or _
                  data_inf.Recordset("cl_cedula") = 10009 Or data_inf.Recordset("cl_cedula") = 10010 Then
                  data_inf.Recordset.Delete
               End If
            End If
            data_inf.Recordset.MoveNext
         Loop
      End If
   End If
   
   frm_infvtascre.MousePointer = 0
   
   MsgBox "Proceso terminado"
   If Combo1.ListIndex <= 4 Then
      If Option1.Value = True Then
         If Check1.Value = 1 Or Check2.Value = 1 Then
            cr1.ReportFileName = App.path & "\infvtascre1.rpt"
         Else
            cr1.ReportFileName = App.path & "\infvtascre2.rpt"
         End If
      Else
         If Option2.Value = True Then
            If Check1.Value = 1 Or Check2.Value = 1 Then
               cr1.ReportFileName = App.path & "\infvtascre1n.rpt"
            Else
                cr1.ReportFileName = App.path & "\infvtascre2n.rpt"
            End If
         End If
      End If
   Else
      If Combo1.ListIndex = 5 Then
         cr1.ReportFileName = App.path & "\infvtascre33.rpt"
      End If
   End If
   If Format(mfd.Text, "yyyy/mm/dd") = Format("01/01/1900", "yyyy/mm/dd") Then
      mfd.Text = "__/__/____"
      mfh.Text = "__/__/____"
   End If
   If Combo1.ListIndex = 0 Then
      cr1.ReportTitle = "INFORME VENTAS CREDITO FECHA: " & mfd.Text & " HASTA: " & mfh.Text
      cr1.Action = 1
   Else
      If Combo1.ListIndex = 1 Then
         cr1.ReportTitle = "INFORME VENTAS CREDITO PENDIENTES DE PAGO FECHA: " & mfd.Text & " HASTA: " & mfh.Text
         cr1.Action = 1
      Else
         If Combo1.ListIndex = 2 Then
            cr1.ReportTitle = "INFORME VENTAS CREDITO SALDADAS FECHA: " & mfd.Text & " HASTA: " & mfh.Text
            cr1.Action = 1
         Else
            If Combo1.ListIndex = 3 Then
               cr1.ReportTitle = "INFORME VENTAS CREDITO VENCIDAS FECHA: " & mfd.Text & " HASTA: " & mfh.Text
               cr1.Action = 1
            Else
               If Combo1.ListIndex = 4 Then
                  cr1.ReportTitle = "INFORME VENTAS CREDITO A VENCER FECHA: " & mfd.Text & " HASTA: " & mfh.Text
                  cr1.Action = 1
               Else
                  If Combo1.ListIndex = 5 Then
                     cr1.ReportTitle = "INFORME DE LLAMADOS CON CREDITOS PENDIENTES FECHA: " & mfd.Text & " HASTA: " & mfh.Text
                     cr1.Action = 1
                  End If
               End If
            End If
         End If
      End If
   End If
End If

Command1.Enabled = True
frm_infvtascre.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_lin2.ConnectionString = "dsn=" & Xconexrmt
data_cab.ConnectionString = "dsn=" & Xconexrmt
data_lla.ConnectionString = "dsn=" & Xconexrmt
data_inflla.DatabaseName = App.path & "\informes.mdb"
data_cli.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

