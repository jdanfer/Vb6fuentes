VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_veodeucnv 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Estado deuda del cliente"
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton btn_imphis 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Imprimir Deuda"
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
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.Data data_deudas 
      Caption         =   "data_deudas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton btn_cerrar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Cerrar"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
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
         Key             =   "b"
         Text            =   "Mes"
         Object.Width           =   707
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
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "e"
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "f"
         Text            =   "Saldo"
         Object.Width           =   2187
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
      Caption         =   "Estado deuda del cliente:"
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
End
Attribute VB_Name = "frm_veodeucnv"
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
'Me.PrintForm
Data1.DatabaseName = App.Path & "\informes.mdb"
Data1.RecordSource = "infvtas"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If

If data_deudas.Recordset.RecordCount > 0 Then
   data_deudas.Recordset.MoveFirst
    Do While Not data_deudas.Recordset.EOF
       Data1.Recordset.AddNew
       Data1.Recordset("cod_cli") = frmabm.data_clientes.Recordset("cl_codigo")
       Data1.Recordset("nom_cli") = frmabm.data_clientes.Recordset("cl_apellid")
       Data1.Recordset("fecha") = data_deudas.Recordset("fecha")
       Data1.Recordset("mes_paga") = data_deudas.Recordset("mes")
       Data1.Recordset("ano_paga") = data_deudas.Recordset("ano")
       Data1.Recordset("nom_prod") = data_deudas.Recordset("origen")
       Data1.Recordset("imp_timbre") = data_deudas.Recordset("total")
       Xsaldoo = Xsaldoo + data_deudas.Recordset("total")
       Data1.Recordset("tot_lin") = Xsaldoo
       Data1.Recordset.Update
       data_deudas.Recordset.MoveNext
    Loop
    Data1.RecordSource = "Select * from infvtas"
    Data1.Refresh
    cr1.ReportFileName = App.Path & "\infestacta.rpt"
    cr1.Action = 1
End If

End Sub

Private Sub Form_Activate()
Dim Xcount, Xsaldo As Long
Dim a, b, c, d, e, f, g As String
Label2.Caption = frm_abmconve.data_cli.Recordset("cl_codigo")
Label3.Caption = frm_abmconve.data_cli.Recordset("cl_apellid")

a = "a"
b = "b"
c = "c"
d = "d"
e = "e"
f = "f"
g = "g"
Xcount = 1
ListView1.ListItems.Clear

data_deudas.RecordSource = "Select * from deudas where cliente =" & frm_abmconve.data_cli.Recordset("cl_codigo") & " order by ano, mes"
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
       ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_deudas.Recordset("total")
       Xsaldo = Xsaldo + data_deudas.Recordset("total")
       ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xsaldo
       data_deudas.Recordset.MoveNext
       Xcount = Xcount + 1
    Loop
Else
    MsgBox "No existe deuda", vbInformation, "Ver Deudas"
End If
btn_cerrar.SetFocus

End Sub

Private Sub Form_Load()
data_deudas.DatabaseName = App.Path & "\sapp.mdb"

End Sub

