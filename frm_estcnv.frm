VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_estcnv 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Estadísticas"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data data_linanula1 
      Caption         =   "data_linanula1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver Facturas anuladas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anular Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Envío de facturas"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Envío de facturas por correo electrónico"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc data_lineas 
      Height          =   330
      Left            =   840
      Top             =   2640
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
      Caption         =   "data_lineas"
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
   Begin MSAdodcLib.Adodc data_cab 
      Height          =   330
      Left            =   1320
      Top             =   4200
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Re-imprimir Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   2175
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   7680
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton btn_imphis 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir Historial"
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
      Left            =   240
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   2055
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
      Height          =   375
      Left            =   9000
      Picture         =   "frm_estcnv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Text            =   "Hora"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Text            =   "Servicio"
         Object.Width           =   8185
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "d"
         Text            =   "Pago Cuota"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "e"
         Text            =   "RUT"
         Object.Width           =   3177
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "f"
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "g"
         Text            =   "Base"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "h"
         Text            =   "Usuario"
         Object.Width           =   2188
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "i"
         Text            =   "Nro.Fact."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "j"
         Text            =   "F.PAGO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Tipo FACT."
         Object.Width           =   2540
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
      Width           =   4815
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
      Caption         =   "Estadísticas del socio:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   6480
      Picture         =   "frm_estcnv.frx":058A
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "frm_estcnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cerrar_Click()
frm_estcnv.Hide

End Sub

Private Sub btn_imphis_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"

Data1.RecordSource = "infvtas"
Data1.Refresh

If data_lineas.Recordset.RecordCount > 0 Then
   data_lineas.Recordset.MoveFirst
   Do While Not data_lineas.Recordset.EOF
      Data1.Recordset.AddNew
      Data1.Recordset("cod_cli") = data_lineas.Recordset("cod_cli")
      Data1.Recordset("nom_cli") = data_lineas.Recordset("nom_cli")
      Data1.Recordset("fecha") = data_lineas.Recordset("fecha")
      Data1.Recordset("nom_prod") = data_lineas.Recordset("nom_prod")
      Data1.Recordset("operador") = data_lineas.Recordset("operador")
      Data1.Recordset("tot_lin") = data_lineas.Recordset("tot_lin")
      Data1.Recordset("base") = data_lineas.Recordset("base")
      Data1.Recordset("factura") = data_lineas.Recordset("factura")
      Data1.Recordset("tipo") = data_lineas.Recordset("tipo")
      Data1.Recordset.Update
      data_lineas.Recordset.MoveNext
   Loop
   Data1.RecordSource = "Select * from infvtas order by fecha"
   Data1.Refresh
   cr1.ReportFileName = App.path & "\infestad.rpt"
   cr1.Action = 1
   
End If
End Sub


Private Sub Command1_Click()
frm_impfaccnv.Show vbModal

End Sub

Private Sub Command2_Click()
frm_envfaccnv.Show vbModal

End Sub

Private Sub Command3_Click()
Dim Xlafactura As String
Xlafactura = InputBox("Ingrese número de factura a anular para: " & Label3.Caption, "Anulación de Factura")
If Xlafactura <> "" Then
   frm_estcnv.MousePointer = 11
   data_lineas.RecordSource = "Select * from clirespl where cl_codigo =" & frm_estcnv.Label2.Caption & " and cl_numero =" & Xlafactura
   data_lineas.Refresh
   If data_lineas.Recordset.RecordCount > 0 Then
      data_lineas.Recordset.MoveFirst
      Do While Not data_lineas.Recordset.EOF
         data_lineas.Recordset.Delete
         data_lineas.Recordset.MoveNext
      Loop
   End If
   data_lineas.RecordSource = "Select * from linmmdd where cod_cli =" & frm_estcnv.Label2.Caption & "  and factura =" & Xlafactura
   data_lineas.Refresh
   If data_lineas.Recordset.RecordCount > 0 Then
      data_lineas.Recordset.MoveFirst
      Do While Not data_lineas.Recordset.EOF
         data_linanula1.Recordset.AddNew
         data_linanula1.Recordset("factura") = data_lineas.Recordset("factura")
         data_linanula1.Recordset("tipo") = data_lineas.Recordset("tipo")
         data_linanula1.Recordset("realizada") = data_lineas.Recordset("realizada")
         data_linanula1.Recordset("fecha") = data_lineas.Recordset("fecha")
         data_linanula1.Recordset("cod_cli") = data_lineas.Recordset("cod_cli")
         data_linanula1.Recordset("nom_cli") = data_lineas.Recordset("nom_cli")
         data_linanula1.Recordset("nom_cli") = data_lineas.Recordset("nom_cli")
         data_linanula1.Recordset("cod_prod") = data_lineas.Recordset("cod_prod")
         data_linanula1.Recordset("nom_prod") = data_lineas.Recordset("nom_prod")
         data_linanula1.Recordset("valor_iva") = data_lineas.Recordset("valor_iva")
         data_linanula1.Recordset("operador") = data_lineas.Recordset("operador")
         data_linanula1.Recordset("hora") = data_lineas.Recordset("hora")
         data_linanula1.Recordset("convenio") = data_lineas.Recordset("convenio")
         data_linanula1.Recordset("linea") = data_lineas.Recordset("linea")
         data_linanula1.Recordset("pendiente") = data_lineas.Recordset("pendiente")
         data_linanula1.Recordset("tot_lin") = data_lineas.Recordset("tot_lin")
         data_linanula1.Recordset("mes_paga") = data_lineas.Recordset("mes_paga")
         data_linanula1.Recordset("ano_paga") = data_lineas.Recordset("ano_paga")
         data_linanula1.Recordset("base") = data_lineas.Recordset("base")
         data_linanula1.Recordset("ruc") = data_lineas.Recordset("ruc")
         data_linanula1.Recordset("fec_a") = Date
         data_linanula1.Recordset("hora_a") = Format(Time, "HH:mm")
         data_linanula1.Recordset("usua_a") = WElusuario
         data_linanula1.Recordset.Update
         data_lineas.Recordset.MoveNext
      Loop
      data_lineas.Recordset.MoveFirst
      Do While Not data_lineas.Recordset.EOF
         data_lineas.Recordset.Delete
         data_lineas.Recordset.MoveNext
      Loop
      frm_estcnv.MousePointer = 0
      MsgBox "Proceso terminado"
      Unload Me
   Else
      frm_estcnv.MousePointer = 0
      MsgBox "No se encuentran facturas"
   End If
Else
   MsgBox "No ingresó factura"
End If
End Sub

Private Sub Command4_Click()
frm_factanula.Show vbModal

End Sub

Private Sub Form_Activate()
Dim Xcount As Long
Dim a, b, c, d, e, f, g, h, i, j As String
Dim Xlafdesde As Date
Xlafdesde = Date - 400
Label2.Caption = frm_convenios.data_conv.Recordset("cnv_cuenta")
Label3.Caption = frm_convenios.data_conv.Recordset("cnv_desc")
data_linanula1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_linanula1.RecordSource = "select * from lin_anula where cod_cli =" & Label2.Caption
data_linanula1.Refresh

data_cab.ConnectionString = "dsn=" & Xconexrmt

a = "a"
b = "b"
c = "c"
d = "d"
e = "e"
f = "f"
g = "g"
h = "h"
i = "i"
j = "j"
Xcount = 1
ListView1.ListItems.Clear
frm_estcnv.MousePointer = 11
btn_imphis.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
btn_cerrar.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

data_lineas.RecordSource = "Select * from linmmdd where cod_cli =" & frm_convenios.data_conv.Recordset("cnv_cuenta") & "  and fecha >='" & Format(Xlafdesde, "yyyy-mm-dd") & "' order by fecha DESC"
data_lineas.Refresh
If data_lineas.Recordset.RecordCount <> 0 Then
   data_lineas.Recordset.MoveFirst
    Do While Not data_lineas.Recordset.EOF
       If IsNull(data_lineas.Recordset("fecha")) = False Then
          ListView1.ListItems.Add Xcount, , Format(data_lineas.Recordset("fecha"), "dd/mm/yyyy")
       Else
          ListView1.ListItems.Add Xcount, , " "
       End If
       If IsNull(data_lineas.Recordset("hora")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("hora")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , " "
       End If
       If IsNull(data_lineas.Recordset("nom_prod")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN DATOS"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_prod")
       End If
       If IsNull(data_lineas.Recordset("mes_paga")) = False Then
          If data_lineas.Recordset("mes_paga") <> 0 Then
             If IsNull(data_lineas.Recordset("ano_paga")) = False Then
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Trim(str(data_lineas.Recordset("mes_paga"))) + "/" + Trim(str(data_lineas.Recordset("ano_paga")))
             Else
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Trim(str(data_lineas.Recordset("mes_paga"))) + "/00"
             End If
          Else
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
          End If
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data_lineas.Recordset("ruc")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN RUT"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("ruc")
       End If
       If IsNull(data_lineas.Recordset("tot_lin")) = False Then
          If IsNull(data_lineas.Recordset("valor_iva")) = False Then
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Val(data_lineas.Recordset("tot_lin")) + Val(data_lineas.Recordset("valor_iva"))
          Else
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
          End If
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data_lineas.Recordset("base")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("base")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data_lineas.Recordset("operador")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("operador")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data_lineas.Recordset("factura")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("factura")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data_lineas.Recordset("tipo")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tipo")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
       End If
       If IsNull(data_lineas.Recordset("pendiente")) = False Then
          If data_lineas.Recordset("pendiente") = "F" Then
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Factura"
          Else
             If data_lineas.Recordset("pendiente") = "T" Then
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Ticket"
             Else
                If data_lineas.Recordset("pendiente") = "N" Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NC e-Factura"
                Else
                   If data_lineas.Recordset("pendiente") = "C" Then
                      ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NC e-Ticket"
                   Else
                      data_cab.RecordSource = "Select * from clirespl where cl_numero =" & data_lineas.Recordset("factura") & " and cl_codigo =" & data_lineas.Recordset("cod_cli")
                      data_cab.Refresh
                      If data_cab.Recordset.RecordCount > 0 Then
                         If IsNull(data_cab.Recordset("cl_telefon")) = False Then
                            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_cab.Recordset("cl_telefon")
                         Else
                            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
                         End If
                      Else
                         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                      End If
                   End If
                End If
             End If
          End If
       Else
          data_cab.RecordSource = "Select * from clirespl where cl_numero =" & data_lineas.Recordset("factura") & " and cl_codigo =" & data_lineas.Recordset("cod_cli")
          data_cab.Refresh
          If data_cab.Recordset.RecordCount > 0 Then
             If IsNull(data_cab.Recordset("cl_telefon")) = False Then
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_cab.Recordset("cl_telefon")
             Else
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
             End If
          Else
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
          End If
       End If
       data_lineas.Recordset.MoveNext
       Xcount = Xcount + 1
    Loop
    frm_estcnv.MousePointer = 0
    btn_imphis.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    btn_cerrar.Enabled = True

Else
    MsgBox "No existe historial", vbInformation, "Ver historial"
End If
frm_estcnv.MousePointer = 0
btn_imphis.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
btn_cerrar.Enabled = True
Command3.Enabled = True
Command4.Enabled = True

data_lineas.Recordset.Close

btn_cerrar.SetFocus


End Sub

Private Sub Form_Load()
'data_lineas.DatabaseName = App.Path & "\sapp.mdb"
'data_lineas.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lineas.ConnectionString = "dsn=" & Xconexrmt
'data_lineas.RecordSource = "linmmdd"
'data_lineas.Refresh
'data1.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
Data1.DatabaseName = App.path & "\informes.mdb"
'data1.RecordSource = "infvtas"
'data1.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

