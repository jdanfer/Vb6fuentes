VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_estaudito 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Estadísticas"
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
      Left            =   3600
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_lineas 
      Caption         =   "data_lineas"
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
      BackColor       =   &H00FFFFC0&
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
      NumItems        =   10
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
         Text            =   "TIPO FACT"
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
End
Attribute VB_Name = "frm_estaudito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cerrar_Click()
Unload Me


End Sub



Private Sub Form_Activate()
Dim Xcount As Long
Dim a, b, c, d, e, f, g, h, i, j As String
Label2.Caption = frm_solaudito.txt_mat.Text
Label3.Caption = frm_solaudito.txt_nom.Text

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
If frm_solaudito.Combo1.ListIndex = 1 Then
   data_lineas.RecordSource = "Select * from linmmdd where cod_cli =" & frm_solaudito.txt_mat.Text & " order by fecha DESC"
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
                    ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Trim(Str(data_lineas.Recordset("mes_paga"))) + "/" + Trim(Str(data_lineas.Recordset("ano_paga")))
                 Else
                    ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Trim(Str(data_lineas.Recordset("mes_paga"))) + "/00"
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
           
           data_lineas.Recordset.MoveNext
           Xcount = Xcount + 1
        Loop
    
    Else
        MsgBox "No existe historial", vbInformation, "Ver historial"
    End If
Else
   data_lineas.RecordSource = "Select * from llamado where matric =" & frm_solaudito.txt_mat.Text & " order by fecha DESC"
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
           If IsNull(data_lineas.Recordset("motcon")) = True Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN DATOS"
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("motcon")
           End If
           If IsNull(data_lineas.Recordset("categ")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("categ")
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           If IsNull(data_lineas.Recordset("nommed")) = True Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN MEDICO"
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("nommed")
           End If
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           If IsNull(data_lineas.Recordset("movilpas")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("movilpas")
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           If IsNull(data_lineas.Recordset("usuario")) = False Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("usuario")
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           End If
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
           
           data_lineas.Recordset.MoveNext
           Xcount = Xcount + 1
        Loop
    
    Else
        MsgBox "No existe historial", vbInformation, "Ver historial"
    End If


End If

btn_cerrar.SetFocus

End Sub

Private Sub Form_Load()
data_lineas.DatabaseName = App.Path & "\sapp.mdb"
'data_lineas.RecordSource = "linmmdd"
'data_lineas.Refresh

End Sub
