VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtasxfaccp 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de ventas por tipo de factura"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frm_vtasxfaccp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   1080
      Top             =   2760
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
   Begin VB.Data data_notas 
      Caption         =   "data_notas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox t_tipo 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5640
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      Picture         =   "frm_vtasxfaccp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_vtasxfaccp.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos para el informe"
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
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc data_cab 
         Height          =   375
         Left            =   600
         Top             =   600
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
         Caption         =   "data_cab"
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Informe desde historial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox t_b 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "99"
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_vtasxfaccp.frx":0F56
         Left            =   1920
         List            =   "frm_vtasxfaccp.frx":0F72
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   360
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
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo Factura:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   2880
      Picture         =   "frm_vtasxfaccp.frx":0FD7
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2655
   End
End
Attribute VB_Name = "frm_vtasxfaccp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
frm_vtasxfac.MousePointer = 11
If Combo1.ListIndex = 0 Then
   t_tipo.Text = "T"
Else
   If Combo1.ListIndex = 1 Then
      t_tipo.Text = "F"
   Else
      If Combo1.ListIndex = 2 Then
         t_tipo.Text = "C"
      Else
         If Combo1.ListIndex = 3 Then
            t_tipo.Text = "N"
         Else
            If Combo1.ListIndex = 4 Then
               t_tipo.Text = "B"
            Else
               If Combo1.ListIndex = 5 Then
                  t_tipo.Text = "A"
               Else
                  If Combo1.ListIndex = 6 Then
                     t_tipo.Text = "Z"
                  Else
                     If Combo1.ListIndex = 7 Then
                        t_tipo.Text = "R"
                     Else
                        t_tipo.Text = "X"
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If
      
If md.Text <> "__/__/____" Then
   If t_b.Text = 99 Then
      If Check1.Value = 1 Then
         data1.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And tipo in ('CONTADO','CREDITO') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by cod_prod"
         data1.Refresh
      Else
         data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And pendiente ='" & t_tipo.Text & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by cod_prod"
         data1.Refresh
      End If
   Else
      If Check1.Value = 1 Then
         data1.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And tipo in ('CONTADO','CREDITO') and base =" & t_b.Text & " and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by cod_prod"
         data1.Refresh
      Else
         data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And pendiente ='" & t_tipo.Text & "' and base =" & t_b.Text & " and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by cod_prod"
         data1.Refresh
      End If
   End If
   
    Dim MiBaseact As Database
    Dim Unasesact As Workspace
    Set Unasesact = Workspaces(0)
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
    
    MiBaseact.Execute "Delete * from infvtas"
   
   data_inf.RecordSource = "infvtas"
   data_inf.Refresh
   If data1.Recordset.RecordCount > 0 Then
      data1.Recordset.MoveFirst
      Do While Not data1.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("fecha") = data1.Recordset("fecha")
         data_inf.Recordset("factura") = data1.Recordset("factura")
         data_inf.Recordset("cod_cli") = data1.Recordset("cod_cli")
         data_inf.Recordset("nom_cli") = data1.Recordset("nom_cli")
         data_inf.Recordset("cod_prod") = data1.Recordset("cod_prod")
         data_inf.Recordset("nom_prod") = data1.Recordset("nom_prod")
         data_inf.Recordset("nro_flia") = data1.Recordset("nro_flia")
         data_inf.Recordset("nom_flia") = data1.Recordset("nom_flia")
         data_inf.Recordset("convenio") = data1.Recordset("convenio")
'         If IsNull(Data1.Recordset("valor_iva")) = False Then
         data_inf.Recordset("tot_lin") = data1.Recordset("tot_lin")
         data_inf.Recordset("costo_prod") = data1.Recordset("imp_iva")
         data_inf.Recordset("costo") = data1.Recordset("tot_lin") - data1.Recordset("imp_iva")
'         Else
'            data_inf.Recordset("tot_lin") = Data1.Recordset("tot_lin")
'            data_inf.Recordset("costo_prod") = Data1.Recordset("tot_lin") / 1.1 * 0.1
''            data_inf.Recordset("costo") = Data1.Recordset("tot_lin") - data_inf.Recordset("costo_prod")
'         End If
         data_inf.Recordset("nro_med_a") = data1.Recordset("nro_med_a")
         data_inf.Recordset("nom_med_a") = data1.Recordset("nom_med_a")
         data_inf.Recordset("mes_paga") = data1.Recordset("mes_paga")
         data_inf.Recordset("ano_paga") = data1.Recordset("ano_paga")
         data_inf.Recordset("base") = data1.Recordset("base")
         data_inf.Recordset("pendiente") = data1.Recordset("pendiente")
         data_inf.Recordset("ced_socio") = data1.Recordset("ced_socio")
         data_inf.Recordset("libro_rub") = data1.Recordset("unidad")
         data_inf.Recordset("nom_superv") = data1.Recordset("tipo")
         If IsNull(data1.Recordset("pendiente")) = False Then
            If data1.Recordset("pendiente") = "T" Then
               data_inf.Recordset("tipo") = "e-Ticket"
            Else
               If data1.Recordset("pendiente") = "F" Then
                  data_inf.Recordset("tipo") = "e-Factura"
               Else
                  If data1.Recordset("pendiente") = "C" Then
                     data_inf.Recordset("tipo") = "NC e-Tck"
                     data_inf.Recordset("tot_lin") = data1.Recordset("tot_lin") * -1
                     data_inf.Recordset("costo_prod") = data_inf.Recordset("costo_prod") * -1
                     data_inf.Recordset("costo") = data_inf.Recordset("costo") * -1
                  Else
                     If data1.Recordset("pendiente") = "N" Then
                        data_inf.Recordset("tipo") = "NC e-Fct"
                        data_inf.Recordset("tot_lin") = data1.Recordset("tot_lin") * -1
                        data_inf.Recordset("costo_prod") = data_inf.Recordset("costo_prod") * -1
                        data_inf.Recordset("costo") = data_inf.Recordset("costo") * -1
                     Else
                        If data1.Recordset("pendiente") = "A" Then
                           data_inf.Recordset("tipo") = "ND e-Fct"
                        Else
                           If data1.Recordset("pendiente") = "B" Then
                              data_inf.Recordset("tipo") = "ND e-Tck"
                           Else
                              If data1.Recordset("pendiente") = "R" Then
                                 data_inf.Recordset("tipo") = "Dev.Recibo"
                                 data_inf.Recordset("tot_lin") = data1.Recordset("tot_lin") * -1
                              Else
                                 If data1.Recordset("pendiente") = "Z" Then
                                    data_inf.Recordset("tipo") = "Recibo"
                                 Else
                                    data_inf.Recordset("tipo") = "Registro"
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         Else
            data_inf.Recordset("tipo") = data1.Recordset("tipo")
         End If
'         If Check1.Value = 1 Then
'            data_inf.Recordset("reg_cab") = Data1.Recordset("reg_cab")
'         End If
         If IsNull(data1.Recordset("imp_iva")) = False Then
            data_inf.Recordset("imp_iva") = Format(data1.Recordset("imp_iva"), "Standard")
         Else
            data_inf.Recordset("imp_iva") = 0
         End If
         data_inf.Recordset("rub_cont") = data1.Recordset("rub_cont")
         data_inf.Recordset("rub_nomb") = data1.Recordset("rub_nomb")
         data_inf.Recordset("nro_superv") = data1.Recordset("grupo")
         data_inf.Recordset("operador") = data1.Recordset("operador")
         
         Data2.RecordSource = "Select * from convenio where cnv_codigo ='" & data1.Recordset("convenio") & "'"
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
            If IsNull(Data2.Recordset("cnv_grupo")) = False Then
               If Data2.Recordset("cnv_grupo") = "UNIVERSAL" Or _
                  Data2.Recordset("cnv_grupo") = "CCOU" Or _
                  Data2.Recordset("cnv_grupo") = "SMI" Or _
                  Data2.Recordset("cnv_grupo") = "H.EVANGELICO" Or _
                  Data2.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                  Data2.Recordset("cnv_grupo") = "IMPASA" Then
                  data_inf.Recordset("usa_timbre") = "S"
               Else
                  data_inf.Recordset("usa_timbre") = "N"
               End If
            Else
               data_inf.Recordset("usa_timbre") = "N"
            End If
         Else
            data_inf.Recordset("usa_timbre") = "N"
         End If
         
         data_inf.Recordset.Update
         data1.Recordset.MoveNext
      Loop
      MiBaseact.Execute "Delete * from infvtas where base in (101,102)"
      MiBaseact.Execute "Delete * from infvtas where usa_timbre ='" & "S" & "'"
      data_inf.RecordSource = "Select * from infvtas order by fecha"
      data_inf.Refresh
      cr1.ReportFileName = App.path & "\infvtasxtipocp.rpt"
      cr1.ReportTitle = "INFORME DE VENTAS TIPO DE FACT. " & Combo1.Text & " FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
      cr1.Action = 1
   Else
      MsgBox "No existen registros"
   End If
End If
Command1.Enabled = True
Command2.Enabled = True
frm_vtasxfac.MousePointer = 0


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data1.ConnectionString = "dsn=" & Xconexrmt
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_inf.DatabaseName = App.path & "\informes.mdb"
data_cab.ConnectionString = "dsn=" & Xconexrmt
data_notas.DatabaseName = App.path & "\notascr.mdb"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
