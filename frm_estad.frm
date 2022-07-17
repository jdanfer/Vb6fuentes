VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_estad 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Estadísticas"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data data_lin_borra 
      Caption         =   "data_lin_borra"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_lin_anula 
      Caption         =   "data_lin_anula"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton b_anular_reg 
      BackColor       =   &H0000FF00&
      Caption         =   "Anular Registro sin valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Picture         =   "frm_estad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver registros anulados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc data_lineas 
      Height          =   330
      Left            =   3360
      Top             =   3000
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
   Begin MSAdodcLib.Adodc data_linlab 
      Height          =   375
      Left            =   3240
      Top             =   3120
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
      Caption         =   "data_linlab"
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
      Caption         =   "Ver registros desde respaldos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   3855
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
      BackColor       =   &H00FFFF80&
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
      Height          =   495
      Left            =   240
      MaskColor       =   &H008080FF&
      Picture         =   "frm_estad.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1935
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
      Left            =   9120
      Picture         =   "frm_estad.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Fecha"
         Object.Width           =   2293
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
         Text            =   "Médico"
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
         Text            =   "F.Pago"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "MEDICACION"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Tipo Fact."
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
      Width           =   5055
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
      Height          =   735
      Left            =   6120
      Picture         =   "frm_estad.frx":109E
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1935
   End
End
Attribute VB_Name = "frm_estad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_anular_reg_Click()
Dim Xcountt, Xind, Xnrodoc, Xlinn, Xnromat As Long
Dim Xmensaj, Xmsgcantcuo As String
Dim Xtotales As Double
Xtotales = 0

Xcountt = 1

Xmensaj = MsgBox("Desea anular el documento número " & ListView1.SelectedItem.ListSubItems(8).Text & " ?", vbInformation + vbYesNo, "Anulación")
Xind = 0
Xnromat = Val(frmabm.txt_mat.Caption)

If Xmensaj = vbYes Then
   Xmsgcantcuo = InputBox("INGRESE MOTIVO DE ANULACIÓN")
   Xind = 0
   If Trim(Xmsgcantcuo) <> "" Then
      For Xind = 1 To ListView1.ListItems.count
          ListView1.ListItems(Xind).Selected = True
          If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
             Xnrodoc = Val(ListView1.SelectedItem.ListSubItems(8).Text)
             frm_estad.MousePointer = 11
             data_lin_anula.RecordSource = "Select * from linmmdd where factura =" & Xnrodoc & " and cod_cli =" & Xnromat
             data_lin_anula.Refresh
             If data_lin_anula.Recordset.RecordCount > 0 Then
                data_lin_anula.Recordset.MoveFirst
                Do While Not data_lin_anula.Recordset.EOF
                   Xtotales = Xtotales + data_lin_anula.Recordset("tot_lin")
                   data_lin_anula.Recordset.MoveNext
                Loop
                If Xtotales = 0 Then
                   data_lin_anula.Recordset.MoveFirst
                   Do While Not data_lin_anula.Recordset.EOF
                      data_lin_borra.Recordset.AddNew
                      data_lin_borra.Recordset("factura") = data_lin_anula.Recordset("factura")
                      data_lin_borra.Recordset("fecha") = data_lin_anula.Recordset("fecha")
                      data_lin_borra.Recordset("cod_cli") = data_lin_anula.Recordset("cod_cli")
                      data_lin_borra.Recordset("nom_cli") = data_lin_anula.Recordset("nom_cli")
                      data_lin_borra.Recordset("cod_prod") = data_lin_anula.Recordset("cod_prod")
                      data_lin_borra.Recordset("nom_prod") = data_lin_anula.Recordset("nom_prod")
                      data_lin_borra.Recordset("operador") = data_lin_anula.Recordset("operador")
                      data_lin_borra.Recordset("nro_flia") = data_lin_anula.Recordset("nro_flia")
                      data_lin_borra.Recordset("linea") = data_lin_anula.Recordset("linea")
                      data_lin_borra.Recordset("convenio") = data_lin_anula.Recordset("convenio")
                      data_lin_borra.Recordset("pendiente") = data_lin_anula.Recordset("pendiente")
                      data_lin_borra.Recordset("ced_socio") = data_lin_anula.Recordset("ced_socio")
                      data_lin_borra.Recordset("tot_lin") = data_lin_anula.Recordset("tot_lin")
                      data_lin_borra.Recordset("fact") = data_lin_anula.Recordset("fact")
                      data_lin_borra.Recordset("nro_med_a") = data_lin_anula.Recordset("nro_med_a")
                      data_lin_borra.Recordset("nom_med_a") = data_lin_anula.Recordset("nom_med_a")
                      data_lin_borra.Recordset("base") = data_lin_anula.Recordset("base")
                      data_lin_borra.Recordset("contact_tel") = data_lin_anula.Recordset("contact_tel")
                      data_lin_borra.Recordset("nom_med_s") = Mid(Xmsgcantcuo, 1, 40)
                      data_lin_borra.Recordset.Update
                      data_lin_anula.Recordset.MoveNext
                   Loop
                   data_lin_anula.Recordset.MoveFirst
                   Do While Not data_lin_anula.Recordset.EOF
                      data_lin_anula.Recordset.Delete
                      data_lin_anula.Recordset.MoveNext
                   Loop
                   frm_estad.MousePointer = 0
                   MsgBox "Proceso terminado.", vbInformation
                Else
                   frm_estad.MousePointer = 0
                   MsgBox "El documento no puede ser anulado porque no es un registro.", vbCritical
                End If
             Else
                frm_estad.MousePointer = 0
                MsgBox "No se encuentra el documento.", vbCritical
                
             End If
          End If
      Next Xind
      ListView1.ListItems.Clear
      Carga_grid
   
   Else
     frm_estad.MousePointer = 0
      MsgBox "Debe ingresar motivo de anulación.", vbCritical
   End If
End If


End Sub

Private Sub btn_cerrar_Click()
Unload Me


End Sub

Private Sub btn_imphis_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
Data1.DatabaseName = App.path & "\informes.mdb"
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
      Data1.Recordset("nom_med_a") = data_lineas.Recordset("nom_med_a")
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

Private Sub Check1_Click()
frm_estad.MousePointer = 11
If Check1.Value = 1 Then
    Dim Xcount As Long
    Dim a, b, c, d, e, f, g, h, i, j As String
    Label2.Caption = frmabm.data_clientes.Recordset("cl_codigo")
    Label3.Caption = frmabm.data_clientes.Recordset("cl_apellid")
    
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
    data_lineas.RecordSource = "Select * from resplin where cod_cli =" & frmabm.data_clientes.Recordset("cl_codigo") & " order by fecha DESC"
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
           If IsNull(data_lineas.Recordset("nom_med_a")) = True Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN MEDICO"
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_med_a")
           End If
           If IsNull(data_lineas.Recordset("tot_lin")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
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
           If data_lineas.Recordset("nro_flia") = 6 Then
              If IsNull(data_lineas.Recordset("nom_medic")) = False Then
                 ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_medic")
              Else
                 ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
              End If
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           data_lineas.Recordset.MoveNext
           Xcount = Xcount + 1
        Loop
    
    Else
        MsgBox "No existe historial", vbInformation, "Ver historial"
    End If
    btn_cerrar.SetFocus
Else
    Label2.Caption = frmabm.data_clientes.Recordset("cl_codigo")
    Label3.Caption = frmabm.data_clientes.Recordset("cl_apellid")
    
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
    data_lineas.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.data_clientes.Recordset("cl_codigo") & " order by fecha DESC"
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
           If IsNull(data_lineas.Recordset("nom_med_a")) = True Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN MEDICO"
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_med_a")
           End If
           If IsNull(data_lineas.Recordset("tot_lin")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
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
           If data_lineas.Recordset("nro_flia") = 6 Then
              If IsNull(data_lineas.Recordset("nom_medic")) = False Then
                 ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_medic")
              Else
                 ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
              End If
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           
           data_lineas.Recordset.MoveNext
           Xcount = Xcount + 1
        Loop
    
    Else
        MsgBox "No existe historial", vbInformation, "Ver historial"
    End If
    btn_cerrar.SetFocus

End If
frm_estad.MousePointer = 0

End Sub



Private Sub Check2_Click()
If Check2.Value = 1 Then
   Carga_grid_cancelado
Else
   Carga_grid
End If

End Sub

Private Sub Form_Load()
data_lineas.ConnectionString = "DSN=" & Xconexrmt
data_linlab.ConnectionString = "DSN=" & Xconexrmt
Label2.Caption = frmabm.data_clientes.Recordset("cl_codigo")
If IsNull(frmabm.data_clientes.Recordset("cl_apellid")) = False Then
   Label3.Caption = frmabm.data_clientes.Recordset("cl_apellid")
Else
   Label3.Caption = "NN"
End If

data_linlab.ConnectionString = "dsn=" & Xconexrmt
data_lin_anula.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin_borra.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin_borra.RecordSource = "select * from lin_reg_anula"
data_lin_borra.Refresh

Carga_grid

'btn_cerrar.SetFocus
If ControlUsuario(b_anular_reg.Name) = 1 Then
   b_anular_reg.Enabled = True
   Check2.Enabled = True
Else
   b_anular_reg.Enabled = False
   Check2.Enabled = False
End If
'data_lineas.RecordSource = "linmmdd"
'data_lineas.Refresh
'data1.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
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

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'Dim Xfact As Long
'Dim Xserv As String


'Xfact = 0
'Xserv = ""
'If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
'   Xfact = ListView1.SelectedItem.ListSubItems(8).Text
'   Xserv = ListView1.SelectedItem.ListSubItems(2).Text
'   frm_estad.MousePointer = 11
'   data_linlab.RecordSource = "Select * from linmmdd where nom_prod ='" & Xserv & "' and factura =" & Xfact & " and nro_flia =" & 3
'   data_linlab.Refresh
'   If data_linlab.Recordset.RecordCount > 0 Then
'      data_linlab.Recordset.Edit
''      data_linlab.Recordset("tcambio") = 8
'      data_linlab.Recordset("vto") = Format(Date, "dd/mm/yyyy")
'      data_linlab.Recordset.Update
'   Else
'      ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = False
'   End If
'   frm_estad.MousePointer = 0
'Else
'   frm_estad.MousePointer = 11
'   Xfact = ListView1.SelectedItem.ListSubItems(8).Text
'   Xserv = ListView1.SelectedItem.ListSubItems(2).Text
'   data_linlab.RecordSource = "Select * from linmmdd where nom_prod ='" & Xserv & "' and factura =" & Xfact
'   data_linlab.RecordSource = "Select * from linmmdd where nom_prod ='" & Xserv & "' and factura =" & Xfact & " and nro_flia =" & 3
'   data_linlab.Refresh
'   If data_linlab.Recordset.RecordCount > 0 Then
'      If IsNull(data_linlab.Recordset("tcambio")) = False Then
'         data_linlab.Recordset.Edit
'         data_linlab.Recordset("tcambio") = Null
'         data_linlab.Recordset.Update
'      End If
'      If IsNull(data_linlab.Recordset("vto")) = False Then
'         data_linlab.Recordset.Edit
'         data_linlab.Recordset("vto") = Null
'         data_linlab.Recordset.Update
'      End If
'   End If
'   frm_estad.MousePointer = 0
'End If


End Sub



Public Sub Carga_grid()
Dim Xcount As Long
Dim a, b, c, d, e, f, g, h, i, j, k As String
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
k = "k"
Xcount = 1
ListView1.ListItems.Clear
data_lineas.RecordSource = "Select * from linmmdd where cod_cli =" & frmabm.data_clientes.Recordset("cl_codigo") & " order by fecha DESC, hora DESC"
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
       If IsNull(data_lineas.Recordset("nom_med_a")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN MEDICO"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_med_a")
       End If
       If IsNull(data_lineas.Recordset("tot_lin")) = False Then
          If IsNull(data_lineas.Recordset("pendiente")) = False Then
             If data_lineas.Recordset("pendiente") = "F" Or data_lineas.Recordset("pendiente") = "N" Then
                If IsNull(data_lineas.Recordset("valor_iva")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Val(data_lineas.Recordset("tot_lin")) + Val(data_lineas.Recordset("valor_iva"))
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
                End If
             Else
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
             End If
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
       If data_lineas.Recordset("nro_flia") = 6 Then
          If IsNull(data_lineas.Recordset("nom_medic")) = False Then
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_medic")
          Else
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
          End If
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
                      If data_lineas.Recordset("pendiente") = "A" Then
                         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "ND e-Factura"
                      Else
                         If data_lineas.Recordset("pendiente") = "B" Then
                            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "ND e-Ticket"
                         Else
                            If data_lineas.Recordset("pendiente") = "X" Then
                               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "REG."
                            Else
                               If data_lineas.Recordset("pendiente") = "R" Then
                                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "DEV.Recibo"
                               Else
                                  If data_lineas.Recordset("pendiente") = "Z" Then
                                     ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Recibo"
                                  Else
                                     ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Anterior"
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
             End If
          End If
       Else
          If Format(data_lineas.Recordset("fecha"), "yyyy/mm/dd") >= Format("01/07/2016", "yyyy/mm/dd") Then
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Ticket"
          Else
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Anterior"
          End If
       End If
          
       data_lineas.Recordset.MoveNext
       Xcount = Xcount + 1
   Loop
Else
    MsgBox "No existe historial", vbInformation, "Ver historial"
End If
data_lineas.Recordset.Close


End Sub
Public Sub Carga_grid_cancelado()

Dim Xcount As Long
Dim a, b, c, d, e, f, g, h, i, j, k As String
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
k = "k"
Xcount = 1
ListView1.ListItems.Clear
data_lineas.RecordSource = "Select * from lin_reg_anula where cod_cli =" & frmabm.data_clientes.Recordset("cl_codigo") & " order by fecha DESC, hora DESC"
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
       If IsNull(data_lineas.Recordset("nom_med_a")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN MEDICO"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_med_a")
       End If
       If IsNull(data_lineas.Recordset("tot_lin")) = False Then
          If IsNull(data_lineas.Recordset("pendiente")) = False Then
             If data_lineas.Recordset("pendiente") = "F" Or data_lineas.Recordset("pendiente") = "N" Then
                If IsNull(data_lineas.Recordset("valor_iva")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Val(data_lineas.Recordset("tot_lin")) + Val(data_lineas.Recordset("valor_iva"))
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
                End If
             Else
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
             End If
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
       If data_lineas.Recordset("nro_flia") = 6 Then
          If IsNull(data_lineas.Recordset("nom_medic")) = False Then
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nom_medic")
          Else
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
          End If
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
                      If data_lineas.Recordset("pendiente") = "A" Then
                         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "ND e-Factura"
                      Else
                         If data_lineas.Recordset("pendiente") = "B" Then
                            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "ND e-Ticket"
                         Else
                            If data_lineas.Recordset("pendiente") = "X" Then
                               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "REG."
                            Else
                               If data_lineas.Recordset("pendiente") = "R" Then
                                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "DEV.Recibo"
                               Else
                                  If data_lineas.Recordset("pendiente") = "Z" Then
                                     ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Recibo"
                                  Else
                                     ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Anterior"
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
             End If
          End If
       Else
          If Format(data_lineas.Recordset("fecha"), "yyyy/mm/dd") >= Format("01/07/2016", "yyyy/mm/dd") Then
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "e-Ticket"
          Else
             ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Anterior"
          End If
       End If
          
       data_lineas.Recordset.MoveNext
       Xcount = Xcount + 1
   Loop
Else
    MsgBox "No existe historial", vbInformation, "Ver historial"
End If


End Sub

