VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_veodeudall 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Datos del socio"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data_deudas 
      Height          =   375
      Left            =   1200
      Top             =   3360
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
   Begin MSAdodcLib.Adodc data_lineas 
      Height          =   495
      Left            =   6360
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   495
      Left            =   3600
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   2055
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
         Text            =   "FECHA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "HORA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SERVICIO"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PAGO"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "MEDICO"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "TOTAL $"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "BASE"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "USUARIO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "NRO.FACTURA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "TIPO FACT."
         Object.Width           =   1764
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6600
      Top             =   6840
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
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   9000
      Picture         =   "frm_veodeudall.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   6720
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "d"
         Text            =   "Descripción"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vencimiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "e"
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "f"
         Text            =   "Fecha PAGO"
         Object.Width           =   2539
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "g"
         Text            =   "Saldos"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label12 
      BackColor       =   &H000080FF&
      Caption         =   "Historial de consultas/servicios:"
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
      TabIndex        =   13
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackColor       =   &H000080FF&
      Caption         =   "Estado actual de la deuda"
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
      TabIndex        =   12
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label labcnvn 
      BackColor       =   &H0080FFFF&
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
      Left            =   3840
      TabIndex        =   11
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label labcnv 
      BackColor       =   &H0080FFFF&
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
      Left            =   2280
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Caption         =   "Convenio:"
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
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label labced 
      BackColor       =   &H0080FFFF&
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
      Left            =   7680
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
      Caption         =   "Cédula:"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label labing 
      BackColor       =   &H0080FFFF&
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
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "Fecha ingreso:"
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
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
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
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
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
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos del socio:"
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
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   4680
      Picture         =   "frm_veodeudall.frx":058A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frm_veodeudall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cerrar_Click()
'frm_veodeudab.Hide
Unload Me

End Sub


Private Sub Form_Activate()
Dim Xcount, Xsaldo As Long
Dim a, b, c, d, e, f, g, h, i, j As String
Dim Xven As Date
If frm_largador.txt_mat.Text <> "" Then
   Label2.Caption = frm_largador.txt_mat.Text
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Label2.Caption
   data_cli.Refresh
   data_lineas.RecordSource = "Select * from linmmdd where cod_cli =" & Label2.Caption
   data_lineas.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      Label3.Caption = data_cli.Recordset("cl_apellid")
      If IsNull(data_cli.Recordset("cl_fecing")) = False Then
         labing.Caption = data_cli.Recordset("cl_fecing")
      Else
         labing.Caption = "__/__/____"
      End If
      If IsNull(data_cli.Recordset("cl_cedula")) = False Then
         If IsNull(data_cli.Recordset("cl_codced")) = False Then
            labced.Caption = data_cli.Recordset("cl_cedula") & "-" & data_cli.Recordset("cl_codced")
         Else
            labced.Caption = data_cli.Recordset("cl_cedula") & "-" & "0"
         End If
      Else
         labced.Caption = "0"
      End If
      If IsNull(data_cli.Recordset("cl_codconv")) = False Then
         labcnv.Caption = data_cli.Recordset("cl_codconv")
      Else
         labcnv.Caption = ""
      End If
      If IsNull(data_cli.Recordset("cl_nomconv")) = False Then
         labcnvn.Caption = data_cli.Recordset("cl_nomconv")
      Else
         labcnvn.Caption = ""
      End If
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
      data_deudas.RecordSource = "Select * from deudas where cliente =" & Label2.Caption & " order by fecha"
      data_deudas.Refresh
      If data_deudas.Recordset.RecordCount <> 0 Then
         data_deudas.Recordset.MoveFirst
         Do While Not data_deudas.Recordset.EOF
            ListView1.ListItems.Add Xcount, , Format(data_deudas.Recordset("fecha"), "dd/mm/yyyy")
            If IsNull(data_deudas.Recordset("mes")) = False Then
               If IsNull(data_deudas.Recordset("ano")) = False Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("origen")
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("origen")
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("origen")
            End If
            If IsNull(data_deudas.Recordset("nro_superv")) = False Then
               Xven = data_deudas.Recordset("fecha") + data_deudas.Recordset("nro_superv")
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xven
            Else
               Xven = data_deudas.Recordset("fecha") + 15
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xven
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
            data_deudas.Recordset.MoveNext
            Xcount = Xcount + 1
         Loop
      Else
         MsgBox "No existe deuda", vbInformation, "Ver Deudas"
      End If
   
      Dim Xcounttt As Long
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
      Xcounttt = 1
      ListView2.ListItems.Clear
      data_lineas.RecordSource = "Select * from linmmdd where cod_cli =" & data_cli.Recordset("cl_codigo") & " order by fecha DESC"
      data_lineas.Refresh
      If data_lineas.Recordset.RecordCount <> 0 Then
         data_lineas.Recordset.MoveFirst
         Do While Not data_lineas.Recordset.EOF
            If IsNull(data_lineas.Recordset("fecha")) = False Then
               ListView2.ListItems.Add Xcounttt, , Format(data_lineas.Recordset("fecha"), "dd/mm/yyyy")
            Else
               ListView2.ListItems.Add Xcounttt, , " "
            End If
            If IsNull(data_lineas.Recordset("hora")) = False Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("hora")
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , " "
            End If
            If IsNull(data_lineas.Recordset("nom_prod")) = True Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , "SIN DATOS"
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("nom_prod")
            End If
            If IsNull(data_lineas.Recordset("mes_paga")) = False Then
               If data_lineas.Recordset("mes_paga") <> 0 Then
                  If IsNull(data_lineas.Recordset("ano_paga")) = False Then
                     ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , Trim(Str(data_lineas.Recordset("mes_paga"))) + "/" + Trim(Str(data_lineas.Recordset("ano_paga")))
                  Else
                     ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , Trim(Str(data_lineas.Recordset("mes_paga"))) + "/00"
                  End If
               Else
                  ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , ""
               End If
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , ""
            End If
            If IsNull(data_lineas.Recordset("nom_med_a")) = True Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , "SIN MEDICO"
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("nom_med_a")
            End If
            If IsNull(data_lineas.Recordset("tot_lin")) = False Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("tot_lin")
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , ""
            End If
            If IsNull(data_lineas.Recordset("base")) = False Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("base")
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , ""
            End If
            If IsNull(data_lineas.Recordset("operador")) = False Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("operador")
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , ""
            End If
            If IsNull(data_lineas.Recordset("factura")) = False Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("factura")
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , ""
            End If
            If IsNull(data_lineas.Recordset("tipo")) = False Then
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , data_lineas.Recordset("tipo")
            Else
               ListView2.ListItems.Item(Xcounttt).ListSubItems.Add , , ""
            End If
            data_lineas.Recordset.MoveNext
            Xcounttt = Xcounttt + 1
         Loop
      Else
         MsgBox "No existe historial", vbInformation, "Ver historial"
      End If
   Else
      MsgBox "No se encuentra deuda"
   End If
Else

End If

btn_cerrar.SetFocus

End Sub

Private Sub Form_Load()
data_deudas.ConnectionString = "dsn=" & Xconexrmt
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_lineas.ConnectionString = "dsn=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
