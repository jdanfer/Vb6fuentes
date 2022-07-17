VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_veodeudab 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Estado deuda del cliente"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   330
      Left            =   7680
      Top             =   2160
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
      Caption         =   "data_cli"
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
   Begin MSAdodcLib.Adodc data_deudas 
      Height          =   375
      Left            =   720
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Acciones administración"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_deubus 
      Caption         =   "data_deubus"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_u 
      Caption         =   "data_u"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Refinanciar"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   2535
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3480
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton btn_imphis 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Imprimir Deuda"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton btn_cerrar 
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
      Left            =   9960
      Picture         =   "frm_veodeudab.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Text            =   "Mes"
         Object.Width           =   883
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Text            =   "Año"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "d"
         Text            =   "Descripción"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "o"
         Text            =   "Vencimiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "e"
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "f"
         Text            =   "Fecha PAGO"
         Object.Width           =   2539
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "g"
         Text            =   "Saldos"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Linea"
         Object.Width           =   776
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Estado deuda del socio:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   6480
      Picture         =   "frm_veodeudab.frx":058A
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2055
   End
End
Attribute VB_Name = "frm_veodeudab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cerrar_Click()
'frm_veodeudab.Hide
Unload Me

End Sub

Private Sub btn_imphis_Click()
Dim Xsaldoo As Double
Dim Xvto As Date

'Me.PrintForm

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"

Data1.DatabaseName = App.path & "\informes.mdb"
Data1.RecordSource = "infvtas"
Data1.Refresh

If data_deudas.Recordset.RecordCount > 0 Then
   data_deudas.Recordset.MoveFirst
    Do While Not data_deudas.Recordset.EOF
       If IsNull(data_deudas.Recordset("fecha_pago")) = False Then
       Else
            Data1.Recordset.AddNew
            Data1.Recordset("cod_cli") = frmabm.data_clientes.Recordset("cl_codigo")
            Data1.Recordset("nom_cli") = Mid(frmabm.data_clientes.Recordset("cl_apellid"), 1, 30)
            Data1.Recordset("fecha") = data_deudas.Recordset("fecha")
            Data1.Recordset("mes_paga") = data_deudas.Recordset("mes")
            Data1.Recordset("ano_paga") = data_deudas.Recordset("ano")
            Data1.Recordset("factura") = data_deudas.Recordset("documento")
            Data1.Recordset("nom_prod") = data_deudas.Recordset("origen")
            Data1.Recordset("imp_timbre") = data_deudas.Recordset("total")
            If data_deudas.Recordset("tipodoc") = "CRE" Then
               If IsNull(data_deudas.Recordset("nro_superv")) = False Then
                  Xvto = data_deudas.Recordset("fecha") + data_deudas.Recordset("nro_superv")
               Else
                  Xvto = data_deudas.Recordset("fecha") + 15
               End If
               Data1.Recordset("vto") = Xvto
            End If
'            Xsaldoo = Xsaldoo + data_deudas.Recordset("total")
            If IsNull(data_deudas.Recordset("fecha_pago")) = True Then
               Xsaldoo = Xsaldoo + data_deudas.Recordset("total")
            End If
            Data1.Recordset("realizada") = data_deudas.Recordset("fecha_pago")
'            Xsaldoo = Xsaldoo + data_deudas.Recordset("total")
            Data1.Recordset("tot_lin") = Xsaldoo
            Data1.Recordset.Update
       End If
       data_deudas.Recordset.MoveNext
    Loop
    Data1.RecordSource = "Select * from infvtas"
    Data1.Refresh
    cr1.ReportFileName = App.path & "\infestacta.rpt"
    cr1.Action = 1
End If

End Sub

Private Sub Command1_Click()
If IsNull(frmabm.data_clientes.Recordset("cl_nrocobr")) = False Then
   If frmabm.data_clientes.Recordset("cl_nrocobr") = 615 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 616 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 635 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 602 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 113 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 653 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 672 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 1 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 10 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 201 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 512 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 636 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 685 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 208 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 209 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 8 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 0 Or _
       frmabm.data_clientes.Recordset("cl_nrocobr") = 999 Then
       MsgBox "Debe figurar con cobrador a domicilio", vbExclamation
   Else
       If WElusuario = "MCOSTA" Or WElusuario = "JFERNAN" Or XWeltipoU = "USUARIOS ADM" Or WElusuario = "NELIDA" Or WElusuario = "PAOLA" Or XWeltipoU = "ADMINISTRADOR" Or WElusuario = "MSANCHEZ" Then
           Dim Xcountt, Xind, Xnrodoc, Xlinn As Long
           Dim Xmensaj, Xmsgcantcuo, Xcodau, Xvencee As String
            Dim XImp, Ximpcuo As Double
            XImp = 0
            Xcountt = 1
            Xmensaj = MsgBox("Desea realizar refinanciación de los registros seleccionados?", vbInformation + vbYesNo, "Deudas")
            Xind = 0
            If Xmensaj = vbYes Then
               
               Xmsgcantcuo = InputBox("INGRESE CANTIDAD DE CUOTAS")
               
            '   Xvencee = InputBox("INGRESE VENCIMIENTO DE PAGO(EN DIAS)")
               Xvencee = "30"
               If Xmsgcantcuo <> "" And Xvencee <> "" Then
                  
                  For Xind = 1 To ListView1.ListItems.count
                      ListView1.ListItems(Xind).Selected = True
                      If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
               '       MsgBox "Chequeado"
                         XImp = XImp + ListView1.SelectedItem.ListSubItems(5).Text
                      End If
                  Next Xind
                  Xind = 0
                  MsgBox "TOTAL A REFINANCIAR $..:" & str(XImp) & " EN..:" & Xmsgcantcuo & " CUOTAS" & " VENCIMIENTO CADA: " & Xvencee & " DIAS.", vbInformation
                  
                  data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Label2.Caption
                  data_cli.Refresh
                  Ximpcuo = XImp / Val(Xmsgcantcuo)
                  Dim Xeliv, Xvenn As Double
                  Xeliv = 0
                  Xvenn = Val(Xvencee)
                  For Xind = 1 To Val(Xmsgcantcuo)
        '''              ListView1.ListItems(Xind).Selected = True
        '''              If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
                      If XImp > 0 Then
            '     MsgBox "Chequeado"
                         Xnrodoc = ListView1.SelectedItem.ListSubItems(8).Text
                         Xlinn = ListView1.SelectedItem.ListSubItems(9).Text
        '                 data_deubus.Database.Execute "Delete from deudas where cliente =" & Label2.Caption & " and documento =" & Xnrodoc & " and nro_vende =" & Xlinn
                         data_par.Recordset.Edit
                         data_par.Recordset("contado") = data_par.Recordset("contado") + 1
                         data_par.Recordset.Update
                         data_par.Refresh
                         data_deudas.Recordset.AddNew
                         data_deudas.Recordset("cod_cnv") = data_cli.Recordset("cl_codconv")
                         data_deudas.Recordset("nom_cnv") = Mid(data_cli.Recordset("cl_nomconv"), 1, 20)
                         data_deudas.Recordset("cliente") = Label2.Caption
                         data_deudas.Recordset("nombre") = data_cli.Recordset("cl_apellid")
                         data_deudas.Recordset("fecha") = Date
                         data_deudas.Recordset("tipodoc") = "CRE"
                         data_deudas.Recordset("nro_superv") = Xvenn
                         data_deudas.Recordset("documento") = data_par.Recordset("contado")
                         data_deudas.Recordset("importe") = Ximpcuo
                         data_deudas.Recordset("moneda") = 1
                         data_deudas.Recordset("origen") = "Refinanciación CUOTA " & Trim(str(Xind)) & " De:" & Trim(Xmsgcantcuo)
                         data_deudas.Recordset("saldo_cc") = 0
                         data_deudas.Recordset("mes") = 0
                         data_deudas.Recordset("ano") = 0
                         data_deudas.Recordset("estado_cta") = 1
                         data_deudas.Recordset("tiquet") = 0
                         data_deudas.Recordset("deudas") = 0
                         data_deudas.Recordset("total") = Ximpcuo
                         Xeliv = Ximpcuo * 0.1
                         Xeliv = Xeliv / 1.1
                         data_deudas.Recordset("iva") = Xeliv
                         data_deudas.Recordset("servi") = 0
                         data_deudas.Recordset("nro_vende") = Xind
                         data_deudas.Recordset.Update
                         Xvenn = Xvenn + Val(Xvencee)
                      Else
                         MsgBox "Importe en CERO"
                      End If
                  Next Xind
                  Dim Xdocp As Long
                  Xind = 0
                  For Xind = 1 To ListView1.ListItems.count
                      ListView1.ListItems(Xind).Selected = True
                      If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                '     MsgBox "Chequeado"
                         Xdocp = Val(ListView1.SelectedItem.ListSubItems(8).Text)
                         data_deudas.RecordSource = "Select * from deudas where documento =" & Xdocp & " and cliente =" & Label2.Caption
                         data_deudas.Refresh
                         If data_deudas.Recordset.RecordCount > 0 Then
                            If IsNull(data_deudas.Recordset("fecha_pago")) = True Then
'                               data_deudas.Recordset.Edit
                               data_deudas.Recordset("fecha_pago") = Date
                               data_deudas.Recordset.Update
                            End If
                         End If
                      End If
                  Next Xind
                  data_deudas.RecordSource = "Select * from deudas where cliente =" & frmabm.data_clientes.Recordset("cl_codigo") & " order by fecha"
                  data_deudas.Refresh
                  MsgBox "Proceso de REFINANCIACION terminado."
                  Unload Me
               End If
            End If
       Else
           MsgBox "Usuario no autorizado", vbExclamation
       End If
   End If
Else
   MsgBox "No figura cobrador.", vbExclamation
End If


End Sub

Private Sub Command2_Click()
If XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS ADM" Then
   Dim Xcountt, Xind, Xnrodoc, Xlinn As Long
   Xcountt = 1
   Xqueregi = 999
   Wopsed = 0
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       '       MsgBox "Chequeado"
          Xqueregi = ListView1.SelectedItem.ListSubItems(8).Text ' documento
          Xlehas = ListView1.SelectedItem.ListSubItems(3).Text ' descripcion

       End If
   Next Xind
   If Xqueregi = 999 Then
      Xqueregi = 0
   End If
   Wopsed = Label2.Caption
   frm_accadm.Show vbModal
Else
   MsgBox "Usuario no autorizado"
   
End If

End Sub

Private Sub Form_Activate()
Dim Xcount, Xsaldo As Long
Dim a, b, c, d, e, f, g, h As String
Dim Xven As Date
If Trim(frmabm.txt_mat.Caption) <> "" Then
    Label2.Caption = frmabm.data_clientes.Recordset("cl_codigo")
    Label3.Caption = frmabm.data_clientes.Recordset("cl_apellid")
    data_deudas.ConnectionString = "dsn=" & Xconexrmt
    a = "a"
    b = "b"
    c = "c"
    d = "d"
    e = "e"
    f = "f"
    g = "g"
    h = "h"
    Xcount = 1
    ListView1.ListItems.Clear
    data_deudas.RecordSource = "Select * from deudas where cliente =" & frmabm.data_clientes.Recordset("cl_codigo") & " order by ano DESC,mes DESC"
    data_deudas.Refresh
    If data_deudas.Recordset.RecordCount <> 0 Then
       data_deudas.Recordset.MoveFirst
        Do While Not data_deudas.Recordset.EOF
           ListView1.ListItems.Add Xcount, , Format(data_deudas.Recordset("fecha"), "dd/mm/yyyy")
           If IsNull(data_deudas.Recordset("mes")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("mes")
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           If IsNull(data_deudas.Recordset("ano")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("ano")
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           If IsNull(data_deudas.Recordset("mes")) = False Then
              If IsNull(data_deudas.Recordset("ano")) = False Then
                 ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("origen")
              Else
                 ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("origen")
              End If
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("origen")
           End If
           If data_deudas.Recordset("tipodoc") = "CRE" Then
                If IsNull(data_deudas.Recordset("nro_superv")) = False Then
                   Xven = data_deudas.Recordset("fecha") + data_deudas.Recordset("nro_superv")
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xven
                Else
                   Xven = data_deudas.Recordset("fecha") + 15
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xven
                End If
           Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("total")
           Xsaldo = Xsaldo + data_deudas.Recordset("total")
           If IsNull(data_deudas.Recordset("fecha_pago")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("fecha_pago")
              Xsaldo = Xsaldo - data_deudas.Recordset("total")
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xsaldo
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("documento")
           If IsNull(data_deudas.Recordset("nro_vende")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("nro_vende")
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
           End If
           data_deudas.Recordset.MoveNext
           Xcount = Xcount + 1
        Loop
    Else
        MsgBox "No existe deuda", vbInformation, "Ver Deudas"
    End If
    btn_cerrar.SetFocus
Else
    MsgBox "No seleccionó cliente."
End If
End Sub

Private Sub Form_Load()
data_deudas.ConnectionString = "dsn=" & Xconexrmt
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_par.DatabaseName = App.path & "\parse.mdb"
data_par.RecordSource = "parsec0"
data_par.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

