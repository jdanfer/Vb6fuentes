VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_envfaccnv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envío de correos facturas convenios"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9315
   Icon            =   "frm_envfaccnv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc data_lin3 
      Height          =   375
      Left            =   480
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "data_lin3"
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
   Begin VB.Data data_imagen 
      Caption         =   "data_imagen"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc data_lincab 
      Height          =   330
      Left            =   360
      Top             =   1920
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
      Caption         =   "data_lincab"
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
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   480
      Top             =   1440
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
      Caption         =   "data_lin"
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
   Begin VB.Data data_faccab 
      Caption         =   "data_faccab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_fac 
      Caption         =   "data_fac"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   1920
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2880
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
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
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton b_envios 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enviar facturas"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6000
      Picture         =   "frm_envfaccnv.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro.Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre/Razón Social"
         Object.Width           =   6421
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Matrícula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Importe"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Serie"
         Object.Width           =   776
      EndProperty
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox md 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Fecha de realizado el documento"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RECORDAR establecer como impresora predeterminada PDFCreator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   6000
      TabIndex        =   8
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPORTANTE! El envío de las facturas se debe realizar desde el equipo dónde se realizaron las mismas."
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Rango de Fecha:"
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
Attribute VB_Name = "frm_envfaccnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub b_busca_Click()
Dim Xcountt, Xcantfact As Integer
Dim Xnrofac As Double

Xcountt = 1

b_busca.Enabled = False
b_imp.Enabled = False
b_envios.Enabled = False
frm_envfaccnv.MousePointer = 11

Data1.RecordSource = "Select * from linmmdd where base =" & 101 & " and fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by factura"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   ListView1.ListItems.Clear
   Data1.Recordset.MoveFirst
   Xnrofac = Data1.Recordset("factura")
   Xcantfact = 0
   Do While Not Data1.Recordset.EOF
      If Xnrofac = Data1.Recordset("factura") Then
         If Xcantfact = 0 Then
            data_lincab.RecordSource = "Select * from clirespl where cl_numero =" & Data1.Recordset("factura") & " and cl_codigo =" & Data1.Recordset("cod_cli")
            data_lincab.Refresh
            If data_lincab.Recordset.RecordCount > 0 Then
               ListView1.ListItems.Add Xcountt, , data_lincab.Recordset("cl_numero")
               If IsNull(data_lincab.Recordset("info_debit")) = False Then
                  ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lincab.Recordset("info_debit")
               Else
                  ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
               End If
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lincab.Recordset("cl_codigo")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lincab.Recordset("cl_fnac")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lincab.Recordset("saldo_doc")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lincab.Recordset("cl_socmnro")
            '       ListView1.ListItems.Item(Xcountt).Checked = True
               Xcountt = Xcountt + 1
               Xcantfact = Xcantfact + 1
            Else
               ListView1.ListItems.Add Xcountt, , Data1.Recordset("factura")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Data1.Recordset("nom_cli")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Data1.Recordset("cod_cli")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Data1.Recordset("fecha")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , Data1.Recordset("tot_lin")
               ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "A"
            '       ListView1.ListItems.Item(Xcountt).Checked = True
               Xcountt = Xcountt + 1
               Xcantfact = Xcantfact + 1
            End If
         Else
            Xcantfact = Xcantfact + 1
         End If
         Xnrofac = Data1.Recordset("factura")
         Data1.Recordset.MoveNext
      Else
         Xcantfact = 0
         Xnrofac = Data1.Recordset("factura")
      End If
   Loop
   frm_envfaccnv.MousePointer = 0
   
End If

b_imp.Enabled = True
b_busca.Enabled = True
b_envios.Enabled = True
frm_envfaccnv.MousePointer = 0

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub b_envios_Click()
Dim MenCorreo As String
Dim oMail As Class1
Dim XX, Xind, Xcant As Integer
Dim Factura As String
Dim matricula, serie As String
Dim lafecha As Date

frm_envfaccnv.MousePointer = 11
b_envios.Enabled = False

For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       Xcant = Xcant + 1
       Factura = ListView1.ListItems.Item(ListView1.SelectedItem.index).Text
       matricula = ListView1.ListItems(Xind).SubItems(2)
       serie = ListView1.ListItems(Xind).SubItems(5)
       data_faccab.RecordSource = "cabezados"
       data_faccab.Refresh
       If data_faccab.Recordset.RecordCount > 0 Then
          data_faccab.Recordset.MoveFirst
          Do While Not data_faccab.Recordset.EOF
             data_faccab.Recordset.Delete
             data_faccab.Recordset.MoveNext
          Loop
       End If

       data_fac.RecordSource = "lineas2"
       data_fac.Refresh
       If data_fac.Recordset.RecordCount > 0 Then
          data_fac.Recordset.MoveFirst
          Do While Not data_fac.Recordset.EOF
             data_fac.Recordset.Delete
             data_fac.Recordset.MoveNext
          Loop
       End If

       If matricula = "" Then
          MsgBox "No ingresó número de matrícula"
          data_lincab.RecordSource = "Select * from clirespl where cl_numero =" & Val(Factura) & " and cl_socmnro ='" & serie & "'"
          data_lincab.Refresh
       Else
          data_lincab.RecordSource = "Select * from clirespl where cl_numero =" & Val(Factura) & " and cl_codigo =" & matricula
          data_lincab.Refresh
       End If
       If data_lincab.Recordset.RecordCount > 0 Then
          data_faccab.Recordset.AddNew
          data_faccab.Recordset("cl_socmnro") = data_lincab.Recordset("cl_socmnro")
          data_faccab.Recordset("cl_numero") = data_lincab.Recordset("cl_numero")
          data_imagen.RecordSource = "Select * from qr where nrofact =" & Val(Factura) & " and serie ='" & serie & "'"
          data_imagen.Refresh
          If data_imagen.Recordset.RecordCount > 0 Then
             data_faccab.Recordset("qr") = data_imagen.Recordset("qr")
          End If
          data_faccab.Recordset("cl_tipocli") = data_lincab.Recordset("cl_tipocli")
          data_faccab.Recordset("cl_socmnro") = data_lincab.Recordset("cl_socmnro")
          data_faccab.Recordset("cl_numero") = data_lincab.Recordset("cl_numero")
          data_faccab.Recordset("cl_fnac") = data_lincab.Recordset("cl_fnac")
          data_faccab.Recordset("fecha_reac") = data_lincab.Recordset("fecha_reac")
          data_faccab.Recordset("cl_tj_venc") = data_lincab.Recordset("cl_tj_venc")
          data_faccab.Recordset("cl_nrovend") = data_lincab.Recordset("cl_nrovend")
          data_faccab.Recordset("cl_forpago") = data_lincab.Recordset("cl_forpago")
          data_faccab.Recordset("cl_celular") = data_lincab.Recordset("cl_celular")
          data_faccab.Recordset("fecha_modi") = data_lincab.Recordset("fecha_modi")
          data_faccab.Recordset("cl_diacobr") = data_lincab.Recordset("cl_diacobr")
          data_faccab.Recordset("cl_nrotarj") = data_lincab.Recordset("cl_nrotarj")
          data_faccab.Recordset("cl_tjemi_n") = data_lincab.Recordset("cl_tjemi_n")
          data_faccab.Recordset("cl_tjemi_c") = data_lincab.Recordset("cl_tjemi_c")
          data_faccab.Recordset("cl_referen") = data_lincab.Recordset("cl_referen")
          data_faccab.Recordset("tit_tarj") = data_lincab.Recordset("tit_tarj")
          data_faccab.Recordset("cl_nomconv") = data_lincab.Recordset("cl_nomconv")
          data_faccab.Recordset("cl_nro_sup") = data_lincab.Recordset("cl_nro_sup")
          data_faccab.Recordset("hora_baja") = data_lincab.Recordset("hora_baja")
          data_faccab.Recordset("cl_nom_sup") = data_lincab.Recordset("cl_nom_sup")
          data_faccab.Recordset("info_debit") = data_lincab.Recordset("info_debit")
          data_faccab.Recordset("cl_direcci") = data_lincab.Recordset("cl_direcci")
          data_faccab.Recordset("cl_zona") = data_lincab.Recordset("cl_zona")
          data_faccab.Recordset("cl_localid") = data_lincab.Recordset("cl_localid")
          data_faccab.Recordset("cl_codigo") = data_lincab.Recordset("cl_codigo")
          data_faccab.Recordset("usu_baja") = data_lincab.Recordset("usu_baja")
          data_faccab.Recordset("saldo_chc2") = data_lincab.Recordset("saldo_chc2")
          data_faccab.Recordset("saldo_cc") = data_lincab.Recordset("saldo_cc")
          data_faccab.Recordset("saldo_cc2") = data_lincab.Recordset("saldo_cc2")
          data_faccab.Recordset("cl_atrasoa") = data_lincab.Recordset("cl_atrasoa")
          data_faccab.Recordset("cl_cedula") = data_lincab.Recordset("cl_cedula")
          data_faccab.Recordset("saldo_doc2") = data_lincab.Recordset("saldo_doc2")
          data_faccab.Recordset("cl_atrasop") = data_lincab.Recordset("cl_atrasop")
          data_faccab.Recordset("cl_decuota") = data_lincab.Recordset("cl_decuota")
          data_faccab.Recordset("saldo_doc") = data_lincab.Recordset("saldo_doc")
          data_faccab.Recordset("cl_grupo") = data_lincab.Recordset("cl_grupo")
          data_faccab.Recordset("saldo_chc") = data_lincab.Recordset("saldo_chc")
          data_faccab.Recordset("cl_telefon") = data_lincab.Recordset("cl_telefon")
          data_faccab.Recordset("cl_fultpag") = data_lincab.Recordset("cl_fultpag")
          data_faccab.Recordset("cl_ultmesp") = data_lincab.Recordset("cl_ultmesp")
          data_faccab.Recordset("cl_nomvend") = data_lincab.Recordset("cl_nomvend")
          data_faccab.Recordset("cl_fax") = data_lincab.Recordset("cl_fax")
          data_faccab.Recordset("cl_nombre") = data_lincab.Recordset("cl_nombre")
          data_lin3.RecordSource = "Select * from indica_enfc where idhc =" & Val(Factura) & " and in_dosis =" & 1
          data_lin3.Refresh
          If data_lin3.Recordset.RecordCount > 0 Then
             If IsNull(data_lin3.Recordset("in_obs")) = False Then
                data_faccab.Recordset("obsp") = data_lin3.Recordset("in_obs")
             End If
          End If
          data_faccab.Recordset.Update
        'fin de cabezal
       End If

       If matricula = "" Then
          data_lin.RecordSource = "Select * from linmmdd where factura =" & Val(Factura)
       Else
          data_lin.RecordSource = "Select * from linmmdd where factura =" & Val(Factura) & " and cod_cli =" & Val(matricula)
       End If
       data_lin.Refresh
       If data_lin.Recordset.RecordCount > 0 Then
          data_lin.Recordset.MoveFirst
          Do While Not data_lin.Recordset.EOF
             data_fac.Recordset.AddNew
             data_fac.Recordset("fecha") = data_lin.Recordset("fecha")
             data_fac.Recordset("reg_cab") = data_lin.Recordset("reg_cab")
             data_fac.Recordset("factura") = data_lin.Recordset("factura")
             data_fac.Recordset("moneda") = data_lin.Recordset("moneda")
             data_fac.Recordset("servicio") = data_lin.Recordset("servicio")
             data_fac.Recordset("tipo") = data_lin.Recordset("tipo")
             data_fac.Recordset("realizada") = data_lin.Recordset("realizada")
             data_fac.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
             data_fac.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
             data_fac.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
             data_fac.Recordset("tcambio") = data_lin.Recordset("tcambio")
             data_fac.Recordset("fact") = data_lin.Recordset("fact")
             data_fac.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
             data_fac.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
             data_fac.Recordset("cantidad") = data_lin.Recordset("cantidad")
             data_fac.Recordset("operador") = data_lin.Recordset("operador")
             data_fac.Recordset("hora") = data_lin.Recordset("hora")
             data_fac.Recordset("ruc") = data_lin.Recordset("ruc")
             data_fac.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
             data_fac.Recordset("nom_flia") = data_lin.Recordset("nom_flia")
             data_fac.Recordset("nro_superv") = data_lin.Recordset("nro_superv")
             data_fac.Recordset("nom_superv") = data_lin.Recordset("nom_superv")
             data_fac.Recordset("convenio") = data_lin.Recordset("convenio")
             data_fac.Recordset("unidad") = data_lin.Recordset("unidad")
             data_fac.Recordset("grupo") = data_lin.Recordset("grupo")
             data_fac.Recordset("rub_cont") = data_lin.Recordset("rub_cont")
             data_fac.Recordset("arancel") = data_lin.Recordset("arancel")
             data_fac.Recordset("usa_timbre") = data_lin.Recordset("usa_timbre")
             data_fac.Recordset("imp_timbre") = data_lin.Recordset("imp_timbre")
             data_fac.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
             data_fac.Recordset("rub_nomb") = data_lin.Recordset("rub_nomb")
             data_fac.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
             data_fac.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
             data_fac.Recordset("nom_med_s") = data_lin.Recordset("nom_med_s")
             data_fac.Recordset("precio_est") = data_lin.Recordset("precio_est")
             data_fac.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
             data_fac.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
             data_fac.Recordset("base") = data_lin.Recordset("base")
             data_fac.Recordset("imp_iva") = data_lin.Recordset("imp_iva")
             data_fac.Recordset("linea") = data_lin.Recordset("linea")
             data_fac.Recordset("dias") = data_lin.Recordset("dias")
             data_fac.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
             data_fac.Recordset("pre_civa") = data_lin.Recordset("pre_civa")
             data_fac.Recordset("porce_est") = data_lin.Recordset("porce_est")
             data_fac.Recordset("rub_nomb") = data_lin.Recordset("rub_nomb")
             data_fac.Recordset("solicitant") = data_lin.Recordset("solicitant")
             data_lin3.RecordSource = "Select * from indica_enfc where idhc =" & Val(Factura) & " and in_hora ='" & serie & "' and in_dosis =" & 3 & " and in_uni =" & data_lin.Recordset("linea")
             data_lin3.Refresh
             If data_lin3.Recordset.RecordCount > 0 Then
                If IsNull(data_lin3.Recordset("in_obs")) = False Then
                   data_fac.Recordset("obsp") = data_lin3.Recordset("in_obs")
                End If
             End If
             data_fac.Recordset.Update
             data_lin.Recordset.MoveNext
          Loop
       End If
       data_faccab.RecordSource = "Select * from cabezados"
       data_faccab.Refresh
       lafecha = data_faccab.Recordset("cl_fnac")
       data_fac.RecordSource = "Select * from lineas2"
       data_fac.Refresh
           
       cr1.ReportFileName = App.path & "\facemail.rpt"
       cr1.Action = 1
       
'       Timer1.Enabled = True
       
       If Dir("c:\planillas\facturas" & "\" & Trim(matricula) & ".pdf") <> "" Then
          Kill ("c:\planillas\facturas" & "\" & Trim(matricula) & ".pdf")
       End If
       Sleep 4000
       
       Name "c:\planillas\facturas\inftr.pdf" As "c:\planillas\facturas\" & Trim(matricula) & ".pdf"
'       Shell ("ren " & "c:\planillas\inftr.pdf " & Trim(Matricula) & ".pdf"), vbMaximizedFocus
       Data2.Recordset.AddNew
       Data2.Recordset("fecha") = lafecha
       Data2.Recordset("cuenta") = Val(matricula)
       data_conv.RecordSource = "Select * from convenio where cnv_cuenta =" & Val(matricula)
       data_conv.Refresh
       If data_conv.Recordset.RecordCount > 0 Then
          If IsNull(data_conv.Recordset("cnv_correoe")) = False Then
             Data2.Recordset("correo") = data_conv.Recordset("cnv_correoe")
          End If
       End If
       Data2.Recordset("envio") = "NO"
       Data2.Recordset.Update
       
    End If
 Next Xind
 Xind = 0
 If Xcant >= 1 Then

    Data2.RecordSource = "select * from envios where envio ='" & "NO" & "'"
    Data2.Refresh

    If Data2.Recordset.RecordCount > 0 Then
       Data2.Recordset.MoveFirst
       Set oMail = New Class1
       Do While Not Data2.Recordset.EOF
          If IsNull(Data2.Recordset("correo")) = False Then
             With oMail
                  .servidor = "smtp.gmail.com"
                  .puerto = 465
                  .UseAuntentificacion = True
                  .ssl = True
                  .Usuario = "sappfacturacion@gmail.com"
                  .PassWord = "sapp1987"
                  .Asunto = "ENVIO Factura SAPP S.A. " & Format(Data2.Recordset("fecha"), "dd/mm/yyyy")
                  .de = "sappfacturacion@gmail.com"
                  .para = Data2.Recordset("correo")
                  .Adjunto = "C:\planillas\facturas" & "\" & Trim(str(Data2.Recordset("cuenta"))) & ".pdf"
                  .Mensaje = "Adjuntamos factura SAPP S.A."
                  .Enviar_Backup ' manda el mail
              End With
              Data2.Recordset.Edit
              Data2.Recordset("env_fecha") = Date
              Data2.Recordset("env_hora") = Format(Time, "HH:mm")
              Data2.Recordset("env_usu") = WElusuario
              Data2.Recordset("envio") = "SI"
              Data2.Recordset.Update
          End If
          Data2.Recordset.MoveNext
       Loop
       Set oMail = Nothing
       frm_envfaccnv.MousePointer = 0
       MsgBox "Proceso de envío de facturas terminado!", vbInformation
    Else
       MsgBox "No hay archivos para enviar"
    End If
Else
    MsgBox "No hay datos seleccionados para enviar"
End If
frm_envfaccnv.MousePointer = 0
b_envios.Enabled = True



End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_fac.DatabaseName = App.path & "\factura.mdb"
data_faccab.DatabaseName = App.path & "\factura.mdb"
data_lin.ConnectionString = "DSN=" & Xconexrmt
data_lincab.ConnectionString = "DSN=" & Xconexrmt
data_imagen.DatabaseName = App.path & "\imagen.mdb"
data_lin3.ConnectionString = "DSN=" & Xconexrmt
Data2.DatabaseName = App.path & "\env_fact.mdb"
Data2.RecordSource = "envios"
Data2.Refresh

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_busca.SetFocus
End If

End Sub

